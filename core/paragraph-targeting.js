/**
 * Shared paragraph-targeting helpers for standalone/add-in consumers.
 *
 * This module centralizes target parsing and matching used by callers that
 * apply per-paragraph operations (for example chat redlines/comments/highlights).
 */

function toArray(nodeList) {
    return Array.from(nodeList || []);
}

function createTargetNotFoundError(message) {
    const error = new Error(message);
    error.code = 'TARGET_NOT_FOUND';
    return error;
}

function createTargetError(code, message, candidates = null) {
    const error = new Error(message);
    error.code = code;
    if (Array.isArray(candidates)) error.candidates = candidates;
    return error;
}

export const WORD_MAIN_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

function getElementsByLocalName(node, localName) {
    if (!node) return [];

    if (typeof node.getElementsByTagNameNS === 'function') {
        const namespaced = toArray(node.getElementsByTagNameNS('*', localName));
        if (namespaced.length > 0) return namespaced;
    }

    if (typeof node.getElementsByTagName !== 'function') return [];

    const prefixed = toArray(node.getElementsByTagName(`w:${localName}`));
    if (prefixed.length > 0) return prefixed;

    return toArray(node.getElementsByTagName(localName));
}

import { extractCanonicalParagraphText } from './paragraph-text.js';

/**
 * Reads visible text from a paragraph by concatenating `w:t` nodes and mapping
 * structural `w:tab` elements to a literal tab character.
 *
 * @param {Element|null|undefined} paragraph - OOXML paragraph node
 * @returns {string}
 */
export function getParagraphText(paragraph) {
    if (!paragraph) return '';
    return extractCanonicalParagraphText(paragraph);
}

/**
 * Returns body paragraphs for a document, or all paragraphs as fallback.
 *
 * @param {Document|Element|null|undefined} xmlDoc - OOXML document root
 * @returns {Element[]}
 */
export function getDocumentParagraphNodes(xmlDoc) {
    if (!xmlDoc) return [];
    const bodies = getElementsByLocalName(xmlDoc, 'body');
    const searchRoot = bodies.length > 0 ? bodies[0] : xmlDoc;
    return getElementsByLocalName(searchRoot, 'p');
}

export function getParagraphId(paragraph) {
    if (!paragraph) return null;
    const attribute = toArray(paragraph.attributes).find(candidate =>
        String(candidate?.localName || candidate?.name || '').replace(/^.*:/, '') === 'paraId'
    );
    return attribute?.value || null;
}

export function createParagraphFingerprint(paragraph) {
    if (!paragraph) return null;
    const text = getParagraphText(paragraph);
    const documentRoot = paragraph.ownerDocument || paragraph;
    const documentIndex = getDocumentParagraphNodes(documentRoot).indexOf(paragraph) + 1;
    const identity = `${getParagraphId(paragraph) || ''}\u001f${documentIndex}\u001f${isParagraphInTable(paragraph) ? 'table' : 'body'}\u001f${text}`;
    let hash = 0x811c9dc5;
    for (let i = 0; i < identity.length; i++) {
        hash ^= identity.charCodeAt(i);
        hash = Math.imul(hash, 0x01000193) >>> 0;
    }
    return `fnv1a32:${hash.toString(16).padStart(8, '0')}`;
}

/**
 * Normalizes whitespace for paragraph-comparison matching.
 *
 * @param {string} text - Input text
 * @returns {string}
 */
export function normalizeWhitespaceForTargeting(text) {
    return String(text || '').replace(/\s+/g, ' ').trim();
}

/**
 * Detects markdown table syntax used for table reconciliation.
 *
 * @param {string} text - Candidate markdown text
 * @returns {boolean}
 */
export function isMarkdownTableText(text) {
    const trimmed = String(text || '').trim();
    return /^\|.+\|/.test(trimmed) && trimmed.includes('\n');
}

/**
 * Parses paragraph references such as `P12`, `[P12]`, `12`, or `P12.3`.
 *
 * @param {string|number|null|undefined} rawValue - Reference input
 * @returns {number|null}
 */
export function parseParagraphReference(rawValue) {
    if (rawValue == null) return null;
    if (typeof rawValue === 'number' && Number.isInteger(rawValue) && rawValue > 0) return rawValue;

    const text = String(rawValue).trim();
    if (!text) return null;

    const prefixed = text.match(/^\[?P(\d+)(?:\.\d+)?\]?$/i);
    if (prefixed) return Number.parseInt(prefixed[1], 10);

    const numeric = text.match(/^(\d+)$/);
    if (numeric) return Number.parseInt(numeric[1], 10);

    return null;
}

/**
 * Removes leading paragraph labels (for example `[P12]`) from text fields.
 *
 * @param {string|null|undefined} text - Input text
 * @returns {string}
 */
export function stripLeadingParagraphMarker(text) {
    if (text == null) return '';
    return String(text).replace(/^\s*\[P\d+(?:\.\d+)?\]\s*/i, '').trim();
}

/**
 * Splits a leading paragraph label from text.
 *
 * @param {string|null|undefined} text - Input text
 * @returns {{ text: string, targetRef: number|null }}
 */
export function splitLeadingParagraphMarker(text) {
    const raw = String(text || '');
    const marker = raw.match(/^\s*\[P(\d+)(?:\.\d+)?\]\s*/i);
    if (!marker) return { text: raw.trim(), targetRef: null };

    return {
        text: raw.replace(/^\s*\[P\d+(?:\.\d+)?\]\s*/i, '').trim(),
        targetRef: Number.parseInt(marker[1], 10)
    };
}

/**
 * Resolves a paragraph by 1-based paragraph index.
 *
 * @param {Document|Element|null|undefined} xmlDoc - OOXML document
 * @param {number|null|undefined} targetRef - 1-based paragraph number
 * @returns {Element|null}
 */
export function findParagraphByReference(xmlDoc, targetRef) {
    if (!Number.isInteger(targetRef) || targetRef < 1) return null;
    const paragraphs = getDocumentParagraphNodes(xmlDoc);
    return paragraphs[targetRef - 1] || null;
}

/**
 * Finds the closest ancestor matching Word namespace + localName.
 *
 * @param {Node|null|undefined} node - Start node
 * @param {string} localName - WordprocessingML local element name (for example `tbl`, `tc`)
 * @param {string} [namespaceUri] - Namespace URI to match
 * @returns {Element|null}
 */
export function findContainingWordElement(node, localName, namespaceUri = WORD_MAIN_NS) {
    let current = node;
    while (current) {
        if (
            current.nodeType === 1 &&
            current.namespaceURI === namespaceUri &&
            current.localName === localName
        ) {
            return current;
        }
        current = current.parentNode;
    }
    return null;
}

/**
 * Finds paragraph by exact/normalized text equality.
 *
 * @param {Document|Element|null|undefined} xmlDoc - OOXML document
 * @param {string} targetText - Target paragraph text
 * @returns {Element|null}
 */
export function findParagraphByStrictText(xmlDoc, targetText) {
    const paragraphs = getDocumentParagraphNodes(xmlDoc);
    const normalizedTarget = String(targetText || '').trim();
    if (!normalizedTarget) return null;

    const exact = paragraphs.find(p => getParagraphText(p).trim() === normalizedTarget);
    if (exact) return exact;

    const normTarget = normalizeWhitespaceForTargeting(normalizedTarget);
    return paragraphs.find(p => normalizeWhitespaceForTargeting(getParagraphText(p)) === normTarget) || null;
}

/**
 * Finds paragraph by strict match, then fuzzy fallback heuristics.
 *
 * @param {Document|Element|null|undefined} xmlDoc - OOXML document
 * @param {string} targetText - Target paragraph text
 * @param {{ onInfo?: (msg:string)=>void }} [options] - Optional logger callbacks
 * @returns {Element|null}
 */
export function findParagraphByBestTextMatch(xmlDoc, targetText, options = {}) {
    const onInfo = typeof options.onInfo === 'function' ? options.onInfo : () => {};
    const paragraphs = getDocumentParagraphNodes(xmlDoc);
    const normalizedTarget = String(targetText || '').trim();
    if (!normalizedTarget) return null;

    const strictMatch = findParagraphByStrictText(xmlDoc, normalizedTarget);
    if (strictMatch) return strictMatch;

    const normTarget = normalizeWhitespaceForTargeting(normalizedTarget);

    const startsWithMatch = paragraphs.find(p => {
        const paragraphText = normalizeWhitespaceForTargeting(getParagraphText(p));
        return paragraphText.length > 10 && normTarget.startsWith(paragraphText);
    });
    if (startsWithMatch) {
        onInfo(`[Fuzzy] Prefix match (target starts with paragraph): "${getParagraphText(startsWithMatch).trim().slice(0, 60)}..."`);
        return startsWithMatch;
    }

    const containsMatch = paragraphs.find(p => {
        const paragraphText = normalizeWhitespaceForTargeting(getParagraphText(p));
        return paragraphText.length > 15 && normTarget.includes(paragraphText);
    });
    if (containsMatch) {
        onInfo(`[Fuzzy] Contains match: "${getParagraphText(containsMatch).trim().slice(0, 60)}..."`);
        return containsMatch;
    }

    let bestScore = 0;
    let bestParagraph = null;
    const targetWords = new Set(normTarget.toLowerCase().split(/\s+/).filter(word => word.length > 2));
    for (const paragraph of paragraphs) {
        const paragraphText = getParagraphText(paragraph).trim();
        if (!paragraphText) continue;

        const paragraphWords = normalizeWhitespaceForTargeting(paragraphText)
            .toLowerCase()
            .split(/\s+/)
            .filter(word => word.length > 2);
        const overlap = paragraphWords.filter(word => targetWords.has(word)).length;
        const score = overlap / Math.max(targetWords.size, 1);
        if (score > bestScore && score > 0.5) {
            bestScore = score;
            bestParagraph = paragraph;
        }
    }

    if (bestParagraph) {
        onInfo(`[Fuzzy] Best word-overlap match (${(bestScore * 100).toFixed(0)}%): "${getParagraphText(bestParagraph).trim().slice(0, 60)}..."`);
        return bestParagraph;
    }

    return null;
}

/**
 * Resolves a target paragraph from `targetRef` + `targetText`.
 *
 * Resolution order:
 * 1) `targetRef` when provided and valid
 * 2) strict text match
 * 3) fuzzy text match
 *
 * @param {Document|Element|null|undefined} xmlDoc - OOXML document
 * @param {{
 *   targetText?: string,
 *   targetRef?: string|number|null,
 *   opType?: string,
 *   onInfo?: (msg:string)=>void,
 *   onWarn?: (msg:string)=>void
 * }} options - Resolution options
 * @returns {{ paragraph: Element, resolvedBy: 'ref'|'strict_text'|'fuzzy_text'|'strict_text_after_ref_drift'|'fuzzy_text_after_ref_drift' }}
 */
export function resolveTargetParagraph(xmlDoc, options = {}) {
    const onInfo = typeof options.onInfo === 'function' ? options.onInfo : () => {};
    const onWarn = typeof options.onWarn === 'function' ? options.onWarn : () => {};
    const opType = options.opType || 'operation';
    const descriptor = options.targetDescriptor && typeof options.targetDescriptor === 'object'
        ? options.targetDescriptor
        : null;
    const cleanTargetText = String(descriptor?.text ?? options.targetText ?? '').trim();
    const parsedRef = parseParagraphReference(descriptor?.index ?? options.targetRef);
    const strictAmbiguity = options.strictAmbiguity === true;

    if (descriptor?.paragraphId) {
        const byId = findParagraphById(xmlDoc, descriptor.paragraphId);
        if (!byId) {
            throw createTargetError(
                'TARGET_NOT_FOUND',
                `Target paragraphId not found: "${descriptor.paragraphId}".`
            );
        }
        const actualFingerprint = createParagraphFingerprint(byId);
        if (descriptor.fingerprint && descriptor.fingerprint !== actualFingerprint) {
            throw createTargetError(
                'TARGET_FINGERPRINT_MISMATCH',
                `Target paragraphId "${descriptor.paragraphId}" no longer matches its source fingerprint.`
            );
        }
        if (typeof descriptor.inTable === 'boolean' && descriptor.inTable !== isParagraphInTable(byId)) {
            throw createTargetError(
                'TARGET_CONTEXT_MISMATCH',
                `Target paragraphId "${descriptor.paragraphId}" does not match the requested table context.`
            );
        }
        if (cleanTargetText) {
            const actualText = normalizeWhitespaceForTargeting(getParagraphText(byId));
            if (actualText !== normalizeWhitespaceForTargeting(cleanTargetText)) {
                throw createTargetError(
                    'TARGET_TEXT_MISMATCH',
                    `Target paragraphId "${descriptor.paragraphId}" no longer matches the supplied text.`
                );
            }
        }
        return { paragraph: byId, resolvedBy: 'paragraph_id' };
    }

    if (cleanTargetText) {
        const unfilteredCandidates = findStrictTargetCandidates(xmlDoc, cleanTargetText);
        const candidates = filterTargetCandidates(unfilteredCandidates, descriptor);

        if (descriptor?.fingerprint && unfilteredCandidates.length > 0 && candidates.length === 0) {
            throw createTargetError(
                'TARGET_FINGERPRINT_MISMATCH',
                'Target text matched, but no paragraph matched the supplied source fingerprint.',
                unfilteredCandidates.map(serializeTargetCandidate)
            );
        }

        if (
            typeof descriptor?.inTable === 'boolean'
            && unfilteredCandidates.length > 0
            && candidates.length === 0
        ) {
            throw createTargetError(
                'TARGET_CONTEXT_MISMATCH',
                'Target text matched, but no paragraph matched the requested table context.',
                unfilteredCandidates.map(serializeTargetCandidate)
            );
        }

        if (descriptor?.occurrence) {
            const occurrenceMatch = candidates[descriptor.occurrence - 1] || null;
            if (!occurrenceMatch) {
                throw createTargetError(
                    'TARGET_NOT_FOUND',
                    `Target occurrence ${descriptor.occurrence} was not found.`,
                    candidates.map(serializeTargetCandidate)
                );
            }
            return { paragraph: occurrenceMatch.paragraph, resolvedBy: 'occurrence' };
        }

        if (strictAmbiguity) {
            if (parsedRef) {
                const byReference = candidates.find(candidate => candidate.index === parsedRef) || null;
                if (byReference) return { paragraph: byReference.paragraph, resolvedBy: 'ref' };
                if (descriptor?.fingerprint && candidates.length > 0) {
                    throw createTargetError(
                        'TARGET_FINGERPRINT_MISMATCH',
                        `Target fingerprint does not match paragraph reference [P${parsedRef}].`,
                        candidates.map(serializeTargetCandidate)
                    );
                }
                if (candidates.length === 1) {
                    return { paragraph: candidates[0].paragraph, resolvedBy: 'strict_text_after_ref_drift' };
                }
            }
            if (candidates.length > 1) {
                throw createTargetError(
                    'AMBIGUOUS_TARGET',
                    `Target text matched ${candidates.length} paragraphs; provide paragraphId, index, occurrence, or fingerprint.`,
                    candidates.map(serializeTargetCandidate)
                );
            }
            if (candidates.length === 0) {
                throw createTargetNotFoundError(`Target paragraph not found: "${cleanTargetText}"`);
            }
        }

        if (!parsedRef && candidates.length === 1) {
            const candidate = candidates[0];
            return {
                paragraph: candidate.paragraph,
                resolvedBy: descriptor?.fingerprint ? 'fingerprint' : 'strict_text'
            };
        }
    }

    if (parsedRef) {
        const byRef = findParagraphByReference(xmlDoc, parsedRef);
        if (byRef) {
            if (cleanTargetText) {
                const strictMatch = findParagraphByStrictText(xmlDoc, cleanTargetText);
                const byRefText = getParagraphText(byRef).trim();
                const byRefNorm = normalizeWhitespaceForTargeting(byRefText);
                const targetNorm = normalizeWhitespaceForTargeting(cleanTargetText);
                const hasDrift = byRefNorm !== targetNorm;

                if (hasDrift && strictMatch && strictMatch !== byRef) {
                    onInfo(`[Target] [P${parsedRef}] drifted for ${opType}; using strict text rematch.`);
                    return { paragraph: strictMatch, resolvedBy: 'strict_text_after_ref_drift' };
                }

                if (hasDrift) {
                    const fuzzyMatch = findParagraphByBestTextMatch(xmlDoc, cleanTargetText, { onInfo });
                    if (fuzzyMatch && fuzzyMatch !== byRef) {
                        onInfo(`[Target] [P${parsedRef}] drifted for ${opType}; using fuzzy text rematch.`);
                        return { paragraph: fuzzyMatch, resolvedBy: 'fuzzy_text_after_ref_drift' };
                    }
                    onInfo(`[Target] Using [P${parsedRef}] fallback for ${opType}; target text drifted.`);
                } else if (strictMatch && strictMatch !== byRef) {
                    onInfo(`[Target] [P${parsedRef}] disambiguated duplicate target text for ${opType}.`);
                }
            } else {
                onInfo(`[Target] Using [P${parsedRef}] fallback for ${opType}.`);
            }
            return { paragraph: byRef, resolvedBy: 'ref' };
        }

        onWarn(`[WARN] Target reference [P${parsedRef}] not found; falling back to text matching for ${opType}.`);
    }

    if (cleanTargetText && !strictAmbiguity) {
        const strictMatch = findParagraphByStrictText(xmlDoc, cleanTargetText);
        if (strictMatch) return { paragraph: strictMatch, resolvedBy: 'strict_text' };

        const fuzzyMatch = findParagraphByBestTextMatch(xmlDoc, cleanTargetText, { onInfo });
        if (fuzzyMatch) return { paragraph: fuzzyMatch, resolvedBy: 'fuzzy_text' };
    }

    if (cleanTargetText) throw createTargetNotFoundError(`Target paragraph not found: "${cleanTargetText}"`);
    if (parsedRef) throw createTargetNotFoundError(`Target paragraph reference not found: [P${parsedRef}]`);
    throw createTargetNotFoundError('Operation target missing: provide "target" text or "targetRef" ([P#]).');
}

function isParagraphInTable(paragraph) {
    return !!findContainingWordElement(paragraph, 'tbl');
}

export function findStrictTargetCandidates(xmlDoc, targetText) {
    const normalizedTarget = normalizeWhitespaceForTargeting(targetText);
    if (!normalizedTarget) return [];

    const paragraphs = getDocumentParagraphNodes(xmlDoc);
    const candidates = [];
    for (let i = 0; i < paragraphs.length; i++) {
        const paragraph = paragraphs[i];
        const paragraphText = getParagraphText(paragraph).trim();
        if (!paragraphText) continue;
        if (normalizeWhitespaceForTargeting(paragraphText) !== normalizedTarget) continue;
        candidates.push({
            paragraph,
            index: i + 1,
            inTable: isParagraphInTable(paragraph),
            paragraphId: getParagraphId(paragraph),
            fingerprint: createParagraphFingerprint(paragraph),
            text: getParagraphText(paragraph)
        });
    }
    return candidates;
}

function serializeTargetCandidate(candidate) {
    return {
        index: candidate.index,
        paragraphId: candidate.paragraphId || null,
        text: candidate.text,
        inTable: candidate.inTable,
        fingerprint: candidate.fingerprint
    };
}

function filterTargetCandidates(candidates, descriptor) {
    let scoped = candidates.slice();
    if (typeof descriptor?.inTable === 'boolean') {
        scoped = scoped.filter(candidate => candidate.inTable === descriptor.inTable);
    }
    if (descriptor?.fingerprint) {
        scoped = scoped.filter(candidate => candidate.fingerprint === descriptor.fingerprint);
    }
    return scoped;
}

function findParagraphById(xmlDoc, paragraphId) {
    if (!paragraphId) return null;
    return getDocumentParagraphNodes(xmlDoc).find(paragraph => getParagraphId(paragraph) === paragraphId) || null;
}

function selectBestTargetCandidate(candidates, parsedRef, expectedInTable = null) {
    if (!Array.isArray(candidates) || candidates.length === 0) return null;

    let scoped = candidates.slice();
    if (typeof expectedInTable === 'boolean') {
        const sameContext = scoped.filter(candidate => candidate.inTable === expectedInTable);
        if (sameContext.length > 0) scoped = sameContext;
    }

    if (Number.isInteger(parsedRef) && parsedRef > 0) {
        scoped.sort((a, b) => Math.abs(a.index - parsedRef) - Math.abs(b.index - parsedRef));
    }

    return scoped[0] || null;
}

/**
 * Builds a turn-start paragraph snapshot keyed by 1-based paragraph index.
 *
 * Intended for callers that apply multiple operations sequentially and need to
 * detect `targetRef` drift after earlier structural edits.
 *
 * @param {Document|Element|null|undefined} xmlDoc - OOXML document root
 * @returns {Map<number, { text: string, normalizedText: string, inTable: boolean }>}
 */
export function buildTargetReferenceSnapshot(xmlDoc) {
    const paragraphs = getDocumentParagraphNodes(xmlDoc);
    const snapshot = new Map();
    for (let i = 0; i < paragraphs.length; i++) {
        const paragraph = paragraphs[i];
        const text = getParagraphText(paragraph).trim();
        snapshot.set(i + 1, {
            text,
            normalizedText: normalizeWhitespaceForTargeting(text),
            inTable: isParagraphInTable(paragraph)
        });
    }
    return snapshot;
}

/**
 * Resolves a paragraph using the standard resolver, then corrects stale
 * `targetRef` mappings via strict rematch when a turn-start snapshot is provided.
 *
 * @param {Document|Element|null|undefined} xmlDoc - OOXML document
 * @param {{
 *   targetText?: string,
 *   targetRef?: string|number|null,
 *   opType?: string,
 *   targetRefSnapshot?: Map<number, { text?: string, inTable?: boolean }>|null,
 *   onInfo?: (msg:string)=>void,
 *   onWarn?: (msg:string)=>void
 * }} options - Resolution options
 * @returns {{ paragraph: Element, resolvedBy: 'ref'|'strict_text'|'fuzzy_text'|'strict_text_after_ref_drift' }}
 */
export function resolveTargetParagraphWithSnapshot(xmlDoc, options = {}) {
    const onInfo = typeof options.onInfo === 'function' ? options.onInfo : () => {};
    const resolved = resolveTargetParagraph(xmlDoc, options);

    const parsedRef = parseParagraphReference(options.targetRef);
    if (!parsedRef || resolved?.resolvedBy !== 'ref') return resolved;

    const snapshotEntry = options.targetRefSnapshot instanceof Map
        ? (options.targetRefSnapshot.get(parsedRef) || null)
        : null;
    if (!snapshotEntry) return resolved;

    const cleanTargetText = String(options.targetText || '').trim();
    const expectedText = cleanTargetText || snapshotEntry.text || '';
    const expectedNorm = normalizeWhitespaceForTargeting(expectedText);
    if (!expectedNorm) return resolved;

    const resolvedNorm = normalizeWhitespaceForTargeting(getParagraphText(resolved.paragraph));
    if (resolvedNorm === expectedNorm) return resolved;

    const candidateTexts = [];
    if (cleanTargetText) candidateTexts.push(cleanTargetText);
    if (snapshotEntry.text) {
        const snapshotNorm = normalizeWhitespaceForTargeting(snapshotEntry.text);
        if (snapshotNorm && !candidateTexts.some(text => normalizeWhitespaceForTargeting(text) === snapshotNorm)) {
            candidateTexts.push(snapshotEntry.text);
        }
    }

    let bestCandidate = null;
    for (const candidateText of candidateTexts) {
        const candidates = findStrictTargetCandidates(xmlDoc, candidateText);
        const selected = selectBestTargetCandidate(candidates, parsedRef, snapshotEntry.inTable);
        if (!selected) continue;
        if (!bestCandidate) bestCandidate = selected;
        if (selected.paragraph !== resolved.paragraph) {
            bestCandidate = selected;
            break;
        }
    }

    if (bestCandidate && bestCandidate.paragraph !== resolved.paragraph) {
        const opType = options.opType || 'operation';
        onInfo(`[Target] [P${parsedRef}] appears stale after prior edits; using strict text rematch for ${opType}.`);
        return { paragraph: bestCandidate.paragraph, resolvedBy: 'strict_text_after_ref_drift' };
    }

    throw createTargetNotFoundError(
        `Target paragraph [P${parsedRef}] no longer matches its batch-start anchor.`
    );

}

/**
 * Resolves a contiguous paragraph range using paragraph references.
 *
 * @param {Document} xmlDoc - XML document
 * @param {string|number|null} startRef - Start paragraph reference (e.g. P12)
 * @param {string|number|null} endRef - End paragraph reference (e.g. P15)
 * @param {Object} [options={}] - Resolution options
 * @param {string} [options.opType='redline'] - Operation type hint
 * @param {Array|null} [options.targetRefSnapshot=null] - Optional target snapshot
 * @param {(message: string) => void} [options.onInfo] - Optional info logger
 * @param {(message: string) => void} [options.onWarn] - Optional warn logger
 * @returns {Element[]|null}
 */
export function resolveParagraphRangeByRefs(xmlDoc, startRef, endRef, options = {}) {
    if (!xmlDoc || !startRef || !endRef) return null;

    const opType = options?.opType || 'redline';
    const targetRefSnapshot = options?.targetRefSnapshot || null;
    const onInfo = typeof options?.onInfo === 'function' ? options.onInfo : () => { };
    const onWarn = typeof options?.onWarn === 'function' ? options.onWarn : () => { };

    const start = resolveTargetParagraphWithSnapshot(xmlDoc, {
        targetRef: startRef,
        opType,
        targetRefSnapshot,
        onInfo,
        onWarn
    })?.paragraph;
    if (!start) return null;

    const end = resolveTargetParagraphWithSnapshot(xmlDoc, {
        targetRef: endRef,
        opType,
        targetRefSnapshot,
        onInfo,
        onWarn
    })?.paragraph;
    if (!end) return null;

    const allParagraphs = Array.from(xmlDoc.getElementsByTagNameNS('*', 'p'));
    const startIdx = allParagraphs.indexOf(start);
    const endIdx = allParagraphs.indexOf(end);
    if (startIdx < 0 || endIdx < startIdx) return null;

    const range = allParagraphs.slice(startIdx, endIdx + 1);
    if (range.length === 0) return null;

    const parent = range[0]?.parentNode || null;
    if (!parent) return null;
    if (!range.every(node => node && node.parentNode === parent)) return null;
    return range;
}
