/**
 * Surgical reconciliation mode.
 *
 * This mode performs in-place run-level edits and preserves existing structure,
 * making it safe for tables and other complex OOXML containers.
 */

import { getApplicableFormatHints } from '../pipeline/markdown-processor.js';
import { computeWordDiffs } from '../pipeline/diff-engine.js';
import { getDocumentParagraphs } from './format-extraction.js';
import { buildSpanIndex, buildSurgicalTextSpans, forEachOverlappingSpan } from './surgical-spans.js';
import {
    processDelete,
    processInsert,
    reconcileFormattingForTextSpan
} from './surgical-diff-application.js';
import { withOoxmlSourceType } from '../core/word-xml.js';
import { createReplacementRevisionEvent } from '../core/types.js';

function checkSafeAdjacencyForPairing(spanIndex, startPos, endPos) {
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
    if (['hyperlink', 'sdt', 'ins', 'del', 'moveFrom', 'moveTo'].includes(parentLocal)) {
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
    const diffs = computeWordDiffs(fullText, modifiedText, diffOptions);
    const spanIndex = buildSpanIndex(textSpans);
    const pairReplacements = options.pairReplacements === true;
    const warnings = [];

    let originalPos = 0;
    let newPos = 0;
    let hasChanges = false;

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

            if (pairReplacements && hasNextInsert) {
                const nextText = diffs[i + 1][1];
                const textWithoutNewlines = nextText.replace(/\n/g, ' ');
                if (textWithoutNewlines.trim().length > 0) {
                    const checkResult = checkSafeAdjacencyForPairing(spanIndex, originalPos, originalPos + text.length);
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
                if (textWithoutNewlines.trim().length > 0) {
                    const insertResult = processInsert(xmlDoc, spanIndex, originalPos, textWithoutNewlines, author, formatHints, newPos, generateRedlines, allParagraphs[0] || null, insMetadata, options?.insertionAffinity || null);
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
            if (textWithoutNewlines.trim().length > 0) {
                const insertResult = processInsert(xmlDoc, spanIndex, originalPos, textWithoutNewlines, author, formatHints, newPos, generateRedlines, allParagraphs[0] || null, null, options?.insertionAffinity || null);
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

    return withOoxmlSourceType({
        oxml: serializer.serializeToString(xmlDoc),
        hasChanges,
        ...(warnings.length > 0 ? { warnings: [...new Set(warnings)] } : {})
    });
}
