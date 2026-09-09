/**
 * Surgical reconciliation mode.
 *
 * This mode performs in-place run-level edits and preserves existing structure,
 * making it safe for tables and other complex OOXML containers.
 */

import { getApplicableFormatHints } from '../pipeline/markdown-processor.js';
import { computeCharacterDiffs, computeInsertionOnlyDiffs, computeWordDiffs } from '../pipeline/diff-engine.js';
import { getDocumentParagraphs } from './format-extraction.js';
import { buildSpanIndex, buildSurgicalTextSpans, forEachOverlappingSpan } from './surgical-spans.js';
import {
    processDelete,
    processInsert,
    reconcileFormattingForTextSpan
} from './surgical-diff-application.js';
import { withOoxmlSourceType } from '../core/word-xml.js';
import { createReplacementRevisionEvent } from '../core/types.js';
import { extractCanonicalParagraphText } from '../core/paragraph-text.js';

function checkSafeAdjacencyForPairing(spanIndex, startPos, endPos, allowInsertionCarrier = false) {
    const spans = [];
    forEachOverlappingSpan(spanIndex, startPos, endPos, span => spans.push(span));
    if (spans.length === 0) return { safe: false };

    const firstRun = spans[0].runElement;
    const parent = firstRun?.parentNode;
    if (!parent) return { safe: false };

    // All overlapping runs must share the exact same parent element
    const sameParent = spans.every(s => s.runElement?.parentNode === parent);
    if (!sameParent) return { safe: false, structuralBoundary: true };

    // Parent container itself cannot be an existing revision or unsupported container
    const parentLocal = (parent.localName || parent.nodeName.replace(/^.*:/, ''));
    if (
        ['hyperlink', 'sdt', 'del', 'moveFrom', 'moveTo'].includes(parentLocal)
        || (parentLocal === 'ins' && !allowInsertionCarrier)
    ) {
        return { safe: false, structuralBoundary: true };
    }

    const structuralTags = new Set([
        'hyperlink', 'fldSimple', 'sdt',
        'commentRangeStart', 'commentRangeEnd', 'commentReference',
        'bookmarkStart', 'bookmarkEnd',
        'moveFrom', 'moveTo', 'ins', 'del'
    ]);

    for (const span of spans) {
        const run = span.runElement;
        for (const child of Array.from(run.childNodes || [])) {
            if (child.nodeType === 1) {
                const tag = child.localName || child.nodeName.replace(/^.*:/, '');
                if (structuralTags.has(tag) || tag === 'fldChar') {
                    return { safe: false, structuralBoundary: true };
                }
            }
        }
    }

    const lastRun = spans[spans.length - 1].runElement;
    let curr = firstRun;
    while (curr && curr !== lastRun) {
        if (curr !== firstRun) {
            const tag = curr.localName || curr.nodeName.replace(/^.*:/, '');
            if (structuralTags.has(tag)) {
                return { safe: false, structuralBoundary: true };
            }
        }
        curr = curr.nextSibling;
    }

    function hasStructuralDescendant(node) {
        if (!node || node.nodeType !== 1) return false;
        const tag = node.localName || node.nodeName.replace(/^.*:/, '');
        if (structuralTags.has(tag) || tag === 'fldChar') return true;
        for (const child of Array.from(node.childNodes || [])) {
            if (child.nodeType === 1 && hasStructuralDescendant(child)) return true;
        }
        return false;
    }

    // Inspect immediate adjacent siblings of the deleted range
    if (hasStructuralDescendant(firstRun.previousSibling) || hasStructuralDescendant(lastRun.nextSibling)) {
        return { safe: false, structuralBoundary: true };
    }

    return { safe: true };
}

/**
 * Applies surgical mode reconciliation.
 *
 * @param {Document} xmlDoc - XML document
 * @param {string} originalText - Original text
 * @param {string} modifiedText - Modified text
 * @param {XMLSerializer} serializer - Serializer instance
 * @param {string} author - Author name
 * @param {Array} formatHints - Format hints
 * @param {boolean} [generateRedlines=true] - Track change toggle
 * @param {Element|null} [targetParagraph=null] - Optional scope paragraph
 * @param {{ diffTimeoutSeconds?: number }} [diffOptions={}] - Diff configuration
 * @param {Object} [options={}] - Additional reconciliation options (e.g. pairReplacements)
 * @returns {{ oxml: string, hasChanges: boolean, warnings?: string[], sourceType?: 'package'|'document'|'fragment' }}
 */
export function applySurgicalMode(xmlDoc, originalText, modifiedText, serializer, author, formatHints, generateRedlines = true, targetParagraph = null, diffOptions = {}, options = {}) {
    void originalText;

    const allParagraphs = targetParagraph
        ? [targetParagraph]
        : getDocumentParagraphs(xmlDoc);

    const { fullText, textSpans } = buildSurgicalTextSpans(allParagraphs);
    const insertionOnlyDiffs = options.existingRevisions === 'slice-cross-author'
        ? computeInsertionOnlyDiffs(fullText, modifiedText)
        : null;
    const diffs = insertionOnlyDiffs || refineSpaceEquivalentReplacements(
        computeWordDiffs(fullText, modifiedText, diffOptions),
        diffOptions
    );
    const spanIndex = buildSpanIndex(textSpans);
    const pairReplacements = options.pairReplacements === true;
    const warnings = [];

    let originalPos = 0;
    let newPos = 0;
    let hasChanges = false;

    const insertionOperations = insertionOnlyDiffs
        ? collectInsertionOperations(insertionOnlyDiffs)
        : [];
    if (insertionOperations.length > 1 && formatHints.length === 0) {
        for (const operation of insertionOperations.slice().reverse()) {
            const liveSpans = buildSurgicalTextSpans(allParagraphs).textSpans;
            const liveSpanIndex = buildSpanIndex(liveSpans);
            const textWithoutNewlines = operation.text.replace(/\n/g, ' ');
            if (textWithoutNewlines.length === 0) continue;
            const insertResult = processInsert(
                xmlDoc,
                liveSpanIndex,
                operation.originalPos,
                textWithoutNewlines,
                author,
                formatHints,
                operation.newPos,
                generateRedlines,
                allParagraphs[0] || null,
                null,
                options?.insertionAffinity || null,
                options?.existingRevisions || 'merge-same-author'
            );
            if (insertResult && typeof insertResult === 'object' && insertResult.error) {
                return withOoxmlSourceType({
                    oxml: serializer.serializeToString(xmlDoc),
                    hasChanges: false,
                    status: 'error',
                    error: insertResult.error
                });
            }
            if (insertResult === true) hasChanges = true;
        }
    } else if (formatHints.length === 0) {
        const editOperations = collectTextEditOperations(diffs);
        for (const operation of editOperations.reverse()) {
            const liveSpanIndex = buildSpanIndex(buildSurgicalTextSpans(allParagraphs).textSpans);
            if (operation.type === 'insert') {
                const inserted = processInsert(
                    xmlDoc,
                    liveSpanIndex,
                    operation.start,
                    operation.text.replace(/\n/g, ' '),
                    author,
                    formatHints,
                    operation.newPos,
                    generateRedlines,
                    allParagraphs[0] || null,
                    null,
                    options?.insertionAffinity || null,
                    options?.existingRevisions || 'merge-same-author'
                );
                if (inserted && typeof inserted === 'object' && inserted.error) {
                    return withOoxmlSourceType({
                        oxml: serializer.serializeToString(xmlDoc),
                        hasChanges: false,
                        status: 'error',
                        error: inserted.error
                    });
                }
                if (inserted === true) hasChanges = true;
                continue;
            }

            let delMetadata = null;
            let insMetadata = null;
            if (operation.type === 'replace' && pairReplacements && generateRedlines && operation.text.replace(/\n/g, ' ')) {
                const checkResult = checkSafeAdjacencyForPairing(
                    liveSpanIndex,
                    operation.start,
                    operation.end,
                    options?.existingRevisions === 'slice-cross-author'
                );
                if (checkResult.safe) {
                    const event = createReplacementRevisionEvent(author, xmlDoc);
                    delMetadata = { id: event.deletionId, author: event.author, date: event.date };
                    insMetadata = { id: event.insertionId, author: event.author, date: event.date };
                } else if (checkResult.structuralBoundary) {
                    warnings.push('PAIRING_SKIPPED_STRUCTURAL_BOUNDARY');
                }
            }

            if (processDelete(xmlDoc, liveSpanIndex, operation.start, operation.end, author, generateRedlines, delMetadata)) {
                hasChanges = true;
            }
            if (operation.type === 'replace') {
                const inserted = processInsert(
                    xmlDoc,
                    liveSpanIndex,
                    operation.end,
                    operation.text.replace(/\n/g, ' '),
                    author,
                    formatHints,
                    operation.newPos,
                    generateRedlines,
                    allParagraphs[0] || null,
                    insMetadata,
                    options?.insertionAffinity || null,
                    options?.existingRevisions || 'merge-same-author'
                );
                if (inserted && typeof inserted === 'object' && inserted.error) {
                    return withOoxmlSourceType({
                        oxml: serializer.serializeToString(xmlDoc),
                        hasChanges: false,
                        status: 'error',
                        error: inserted.error
                    });
                }
                if (inserted === true) hasChanges = true;
            }
        }
    } else {

      for (let i = 0; i < diffs.length; i++) {
        const [op, text] = diffs[i];
        if (op === 0) {
            const len = text.length;
            const startPos = originalPos;
            const endPos = originalPos + len;

            forEachOverlappingSpan(spanIndex, startPos, endPos, span => {
                const overlapStartOriginal = Math.max(span.charStart, startPos);
                const overlapEndOriginal = Math.min(span.charEnd, endPos);
                const segmentLen = overlapEndOriginal - overlapStartOriginal;
                const relativeOffset = overlapStartOriginal - startPos;
                const overlapStartNew = newPos + relativeOffset;
                const overlapEndNew = overlapStartNew + segmentLen;
                const applicableHints = getApplicableFormatHints(formatHints, overlapStartNew, overlapEndNew);
                if (reconcileFormattingForTextSpan(xmlDoc, span, overlapStartOriginal, overlapEndOriginal, applicableHints, author, generateRedlines)) {
                    hasChanges = true;
                }
            });

            originalPos += len;
            newPos += len;
        } else if (op === -1) {
            const hasNextInsert = (i + 1 < diffs.length) && (diffs[i + 1][0] === 1);
            let paired = false;
            let delMetadata = null;
            let insMetadata = null;

            if (pairReplacements && generateRedlines && hasNextInsert) {
                const nextText = diffs[i + 1][1];
                const textWithoutNewlines = nextText.replace(/\n/g, ' ');
                if (textWithoutNewlines.length > 0) {
                    const checkResult = checkSafeAdjacencyForPairing(
                        spanIndex,
                        originalPos,
                        originalPos + text.length,
                        options?.existingRevisions === 'slice-cross-author'
                    );
                    if (checkResult.safe) {
                        const event = createReplacementRevisionEvent(author, xmlDoc);
                        delMetadata = { id: event.deletionId, author: event.author, date: event.date };
                        insMetadata = { id: event.insertionId, author: event.author, date: event.date };
                        paired = true;
                    } else if (checkResult.structuralBoundary) {
                        warnings.push('PAIRING_SKIPPED_STRUCTURAL_BOUNDARY');
                    }
                }
            }

            if (processDelete(xmlDoc, spanIndex, originalPos, originalPos + text.length, author, generateRedlines, delMetadata)) {
                hasChanges = true;
            }
            originalPos += text.length;

            if (paired) {
                i++;
                const [, nextText] = diffs[i];
                const textWithoutNewlines = nextText.replace(/\n/g, ' ');
                if (textWithoutNewlines.length > 0) {
                    const insertResult = processInsert(xmlDoc, spanIndex, originalPos, textWithoutNewlines, author, formatHints, newPos, generateRedlines, allParagraphs[0] || null, insMetadata, options?.insertionAffinity || null, options?.existingRevisions || 'merge-same-author');
                    if (insertResult && typeof insertResult === 'object' && insertResult.error) {
                        return withOoxmlSourceType({
                            oxml: serializer.serializeToString(xmlDoc),
                            hasChanges: false,
                            status: 'error',
                            error: insertResult.error
                        });
                    }
                    if (insertResult === true) {
                        hasChanges = true;
                    }
                }
                newPos += nextText.length;
            }
        } else if (op === 1) {
            const textWithoutNewlines = text.replace(/\n/g, ' ');
            if (textWithoutNewlines.length > 0) {
                const insertResult = processInsert(xmlDoc, spanIndex, originalPos, textWithoutNewlines, author, formatHints, newPos, generateRedlines, allParagraphs[0] || null, null, options?.insertionAffinity || null, options?.existingRevisions || 'merge-same-author');
                if (insertResult && typeof insertResult === 'object' && insertResult.error) {
                    return withOoxmlSourceType({
                        oxml: serializer.serializeToString(xmlDoc),
                        hasChanges: false,
                        status: 'error',
                        error: insertResult.error
                    });
                }
                if (insertResult === true) {
                    hasChanges = true;
                }
            }
            newPos += text.length;
        }
      }
    }

    const actualText = allParagraphs.map(paragraph => extractCanonicalParagraphText(paragraph)).join('\n');
    const expectedText = String(modifiedText).replace(/\r\n/g, '\n');
    if (options.existingRevisions === 'slice-cross-author' && actualText !== expectedText) {
        const mismatchOffset = firstMismatchOffset(expectedText, actualText);
        return withOoxmlSourceType({
            oxml: serializer.serializeToString(xmlDoc),
            hasChanges: false,
            status: 'error',
            error: {
                code: 'PATCH_ROUNDTRIP_MISMATCH',
                message: 'Generated OOXML accepted-view text does not match the requested modified text; the mutation was rejected.',
                mismatchOffset,
                expectedExcerpt: excerptAt(expectedText, mismatchOffset),
                actualExcerpt: excerptAt(actualText, mismatchOffset),
                expectedCodePoint: codePointAtOffset(expectedText, mismatchOffset),
                actualCodePoint: codePointAtOffset(actualText, mismatchOffset)
            },
            ...(warnings.length > 0 ? { warnings: [...new Set(warnings)] } : {})
        });
    }

    return withOoxmlSourceType({
        oxml: serializer.serializeToString(xmlDoc),
        hasChanges,
        ...(warnings.length > 0 ? { warnings: [...new Set(warnings)] } : {})
    });
}

function firstMismatchOffset(expected, actual) {
    const limit = Math.min(expected.length, actual.length);
    for (let index = 0; index < limit; index++) {
        if (expected[index] !== actual[index]) return index;
    }
    return limit;
}

function excerptAt(text, offset, radius = 40) {
    const start = Math.max(0, offset - radius);
    const end = Math.min(text.length, offset + radius);
    return text.slice(start, end);
}

function codePointAtOffset(text, offset) {
    if (offset >= text.length) return 'END';
    return `U+${text.codePointAt(offset).toString(16).toUpperCase().padStart(4, '0')}`;
}

function refineSpaceEquivalentReplacements(diffs, diffOptions) {
    const refined = [];
    for (let index = 0; index < diffs.length; index++) {
        const [op, text] = diffs[index];
        const next = diffs[index + 1];
        if (
            op === -1
            && next?.[0] === 1
            && text !== next[1]
            && text.length === next[1].length
            && text.replace(/\u00a0/g, ' ') === next[1].replace(/\u00a0/g, ' ')
        ) {
            refined.push(...computeCharacterDiffs(text, next[1], diffOptions));
            index++;
            continue;
        }
        refined.push([op, text]);
    }
    return refined;
}

function collectTextEditOperations(diffs) {
    const operations = [];
    let originalPos = 0;
    let newPos = 0;
    for (let index = 0; index < diffs.length; index++) {
        const [op, text] = diffs[index];
        if (op === 0) {
            originalPos += text.length;
            newPos += text.length;
            continue;
        }
        if (op === -1) {
            const next = diffs[index + 1];
            if (next?.[0] === 1) {
                operations.push({
                    type: 'replace',
                    start: originalPos,
                    end: originalPos + text.length,
                    newPos,
                    text: next[1]
                });
                originalPos += text.length;
                newPos += next[1].length;
                index++;
            } else {
                operations.push({ type: 'delete', start: originalPos, end: originalPos + text.length, newPos, text: '' });
                originalPos += text.length;
            }
            continue;
        }
        operations.push({ type: 'insert', start: originalPos, end: originalPos, newPos, text });
        newPos += text.length;
    }
    return operations;
}

function collectInsertionOperations(diffs) {
    const operations = [];
    let originalPos = 0;
    let newPos = 0;
    for (const [op, text] of diffs) {
        if (op === 0) {
            originalPos += text.length;
            newPos += text.length;
        } else if (op === -1) {
            originalPos += text.length;
        } else if (op === 1) {
            operations.push({ originalPos, newPos, text });
            newPos += text.length;
        }
    }
    return operations;
}
