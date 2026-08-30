/**
 * OOXML Reconciliation Pipeline - Diff Engine
 * 
 * Word-level diffing with offset tracking for precise run splitting.
 */

import { diff_match_patch } from 'diff-match-patch';
import { DiffOp } from '../core/types.js';

export const TOKEN_CODE_POINT_BASE = 0x10000;
export const MAX_DIFF_TOKENS = 0x40000;

// diff-match-patch's JavaScript port operates on UTF-16 code units, not
// Unicode code points. Keep its private encoding in the BMP while skipping
// surrogates; larger token sets take the deterministic token-array fallback.
const BMP_FIRST_CODE = 1;
const BMP_BEFORE_SURROGATES = 0xD800 - BMP_FIRST_CODE;
const BMP_AFTER_SURROGATES = 0x10000 - 0xE000;
const DMP_SAFE_TOKEN_LIMIT = BMP_BEFORE_SURROGATES + BMP_AFTER_SURROGATES;

export class DiffTokenLimitError extends Error {
    constructor(limit = MAX_DIFF_TOKENS) {
        super(`Word diff exceeds the safe limit of ${limit} unique tokens.`);
        this.name = 'DiffTokenLimitError';
        this.code = 'DIFF_TOKEN_LIMIT';
        this.limit = limit;
    }
}

export function isDiffTokenLimitError(error) {
    return error?.code === 'DIFF_TOKEN_LIMIT';
}

export function createDiffEngine(options = {}) {
    const timeout = options.diffTimeoutSeconds ?? 0;
    if (!Number.isFinite(timeout) || timeout < 0) {
        throw new TypeError('diffTimeoutSeconds must be a finite non-negative number.');
    }
    const engine = new diff_match_patch();
    engine.Diff_Timeout = timeout;
    return engine;
}

function tokenize(text) {
    const tokens = [];
    const leading = text.match(/^\s+/);
    if (leading) tokens.push(leading[0]);

    const regex = /(\S+)(\s*)/g;
    regex.lastIndex = leading?.[0].length || 0;
    let match;
    while ((match = regex.exec(text)) !== null) {
        if (match[1]) tokens.push(match[1]);
        if (match[2]) tokens.push(match[2]);
    }
    return tokens;
}

/**
 * Converts text into word tokens represented as unique characters.
 * This allows DMP to diff at word-level instead of character-level.
 * 
 * @param {string} text1 - First text to tokenize
 * @param {string} text2 - Second text to tokenize
 * @param {{ maxTokens?: number }} [options={}] - Internal/test capacity override
 * @returns {{ chars1: string, chars2: string, wordArray: string[], tokenIds1: number[], tokenIds2: number[] }}
 */
export function wordsToChars(text1, text2, options = {}) {
    const wordArray = [];
    const wordHash = new Map();
    const maxTokens = options.maxTokens ?? MAX_DIFF_TOKENS;
    if (!Number.isInteger(maxTokens) || maxTokens < 1 || maxTokens > MAX_DIFF_TOKENS) {
        throw new RangeError(`maxTokens must be an integer from 1 to ${MAX_DIFF_TOKENS}.`);
    }

    function mapTokensToChars(tokens) {
        let chars = '';
        const tokenIds = [];
        for (const token of tokens) {
            let tokenId = wordHash.get(token);
            if (tokenId === undefined) {
                if (wordArray.length >= maxTokens) throw new DiffTokenLimitError(maxTokens);
                tokenId = wordArray.length;
                wordArray.push(token);
                wordHash.set(token, tokenId);
            }
            tokenIds.push(tokenId);
            chars += String.fromCodePoint(TOKEN_CODE_POINT_BASE + tokenId);
        }
        return { chars, tokenIds };
    }

    const tokens1 = tokenize(text1);
    const tokens2 = tokenize(text2);
    const encoded1 = mapTokensToChars(tokens1);
    const encoded2 = mapTokensToChars(tokens2);

    return {
        chars1: encoded1.chars,
        chars2: encoded2.chars,
        wordArray,
        tokenIds1: encoded1.tokenIds,
        tokenIds2: encoded2.tokenIds
    };
}

/**
 * Converts character-encoded diffs back to actual word diffs.
 * 
 * @param {Array} diffs - DMP diff array with character codes
 * @param {string[]} wordArray - Array mapping char codes to words
 * @returns {Array} DMP-style diff array with actual words
 */
export function charsToWords(diffs, wordArray) {
    const originalChars = diffs.filter(([op]) => op !== 1).map(([, chars]) => chars).join('');
    const modifiedChars = diffs.filter(([op]) => op !== -1).map(([, chars]) => chars).join('');

    const decode = chars => {
        const tokenIds = [];
        for (let i = 0; i < chars.length;) {
            const codePoint = chars.codePointAt(i);
            const tokenId = codePoint - TOKEN_CODE_POINT_BASE;
            if (!Number.isInteger(tokenId) || tokenId < 0 || tokenId >= wordArray.length) {
                throw new RangeError(`Diff token code point U+${codePoint.toString(16).toUpperCase()} has no mapping.`);
            }
            tokenIds.push(tokenId);
            i += codePoint > 0xFFFF ? 2 : 1;
        }
        return tokenIds;
    };

    // Reconstruct both encoded sides before decoding. A UTF-16-based diff
    // engine may split a surrogate pair across adjacent diff tuples even
    // though the complete original/modified streams remain valid.
    return deterministicLargeTokenDiff(decode(originalChars), decode(modifiedChars), wordArray);
}

function tokenIdToBmpChar(tokenId) {
    const code = tokenId < BMP_BEFORE_SURROGATES
        ? BMP_FIRST_CODE + tokenId
        : 0xE000 + (tokenId - BMP_BEFORE_SURROGATES);
    return String.fromCharCode(code);
}

function bmpCharToTokenId(charCode) {
    if (charCode >= BMP_FIRST_CODE && charCode < 0xD800) return charCode - BMP_FIRST_CODE;
    if (charCode >= 0xE000 && charCode <= 0xFFFF) return BMP_BEFORE_SURROGATES + charCode - 0xE000;
    throw new RangeError(`BMP diff token U+${charCode.toString(16).toUpperCase()} has no mapping.`);
}

function encodeTokenIdsForDmp(tokenIds) {
    let chars = '';
    for (const tokenId of tokenIds) chars += tokenIdToBmpChar(tokenId);
    return chars;
}

function decodeBmpDiffs(diffs, wordArray) {
    return diffs.map(([op, chars]) => {
        const parts = [];
        for (let index = 0; index < chars.length; index++) {
            const tokenId = bmpCharToTokenId(chars.charCodeAt(index));
            if (tokenId >= wordArray.length) {
                throw new RangeError(`BMP diff token ${tokenId} has no mapping.`);
            }
            parts.push(wordArray[tokenId]);
        }
        return [op, parts.join('')];
    });
}

function deterministicLargeTokenDiff(tokenIds1, tokenIds2, wordArray) {
    let prefixLength = 0;
    const sharedLength = Math.min(tokenIds1.length, tokenIds2.length);
    while (prefixLength < sharedLength && tokenIds1[prefixLength] === tokenIds2[prefixLength]) {
        prefixLength++;
    }

    let suffixLength = 0;
    while (
        suffixLength < sharedLength - prefixLength
        && tokenIds1[tokenIds1.length - 1 - suffixLength] === tokenIds2[tokenIds2.length - 1 - suffixLength]
    ) {
        suffixLength++;
    }

    const joinTokens = ids => ids.map(id => wordArray[id]).join('');
    const diffs = [];
    if (prefixLength) diffs.push([0, joinTokens(tokenIds1.slice(0, prefixLength))]);

    const deleted = tokenIds1.slice(prefixLength, tokenIds1.length - suffixLength);
    const inserted = tokenIds2.slice(prefixLength, tokenIds2.length - suffixLength);
    if (deleted.length) diffs.push([-1, joinTokens(deleted)]);
    if (inserted.length) diffs.push([1, joinTokens(inserted)]);
    if (suffixLength) diffs.push([0, joinTokens(tokenIds1.slice(tokenIds1.length - suffixLength))]);
    return diffs;
}

/**
 * Computes word-level diff tuples using a shared diff engine instance.
 *
 * @param {string} originalText - Original text
 * @param {string} newText - New text
 * @param {{ cleanupSemantic?: boolean, diffTimeoutSeconds?: number, maxTokens?: number }} [options={}] - Diff options
 * @returns {Array<[number, string]>}
 */
export function computeWordDiffs(originalText, newText, options = {}) {
    if (originalText === newText) {
        return [[0, originalText]];
    }

    if (!originalText) {
        return [[1, newText]];
    }

    if (!newText) {
        return [[-1, originalText]];
    }

    const { cleanupSemantic = true } = options;

    const { wordArray, tokenIds1, tokenIds2 } = wordsToChars(originalText, newText, options);
    if (wordArray.length > DMP_SAFE_TOKEN_LIMIT) {
        return deterministicLargeTokenDiff(tokenIds1, tokenIds2, wordArray);
    }

    const dmp = createDiffEngine(options);
    const charDiffs = dmp.diff_main(encodeTokenIdsForDmp(tokenIds1), encodeTokenIdsForDmp(tokenIds2));
    if (cleanupSemantic) {
        dmp.diff_cleanupSemantic(charDiffs);
    }

    return decodeBmpDiffs(charDiffs, wordArray);
}

/**
 * Computes word-level diff operations with offset tracking.
 * 
 * @param {string} originalText - Original text
 * @param {string} newText - New text
 * @param {{ cleanupSemantic?: boolean }} [options={}] - Diff options
 * @returns {import('../core/types.js').DiffOperation[]}
 */
export function computeWordLevelDiffOps(originalText, newText, options = {}) {
    // Handle edge cases
    if (originalText === newText) {
        return [{
            type: DiffOp.EQUAL,
            startOffset: 0,
            endOffset: originalText.length,
            text: originalText
        }];
    }

    if (!originalText) {
        return [{
            type: DiffOp.INSERT,
            startOffset: 0,
            endOffset: 0,
            text: newText
        }];
    }

    if (!newText) {
        return [{
            type: DiffOp.DELETE,
            startOffset: 0,
            endOffset: originalText.length,
            text: originalText
        }];
    }

    const wordDiffs = computeWordDiffs(originalText, newText, options);

    // Convert to operations with offsets
    const operations = [];
    let originalOffset = 0;

    for (const [op, text] of wordDiffs) {
        if (op === 0) { // EQUAL
            operations.push({
                type: DiffOp.EQUAL,
                startOffset: originalOffset,
                endOffset: originalOffset + text.length,
                text
            });
            originalOffset += text.length;
        } else if (op === -1) { // DELETE
            operations.push({
                type: DiffOp.DELETE,
                startOffset: originalOffset,
                endOffset: originalOffset + text.length,
                text
            });
            originalOffset += text.length;
        } else if (op === 1) { // INSERT
            operations.push({
                type: DiffOp.INSERT,
                startOffset: originalOffset,
                endOffset: originalOffset, // Insertions don't span original text
                text
            });
            // Don't advance originalOffset for insertions
        }
    }

    return operations;
}

/**
 * Collects all unique boundary offsets from diff operations.
 * 
 * @param {import('../core/types.js').DiffOperation[]} diffOps - Diff operations
 * @returns {Set<number>}
 */
export function collectDiffBoundaries(diffOps) {
    const boundaries = new Set();
    for (const op of diffOps) {
        boundaries.add(op.startOffset);
        boundaries.add(op.endOffset);
    }
    return boundaries;
}

