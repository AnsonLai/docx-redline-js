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
 * @returns {{ oxml: string, hasChanges: boolean, sourceType?: 'package'|'document'|'fragment' }}
 */
export function applySurgicalMode(xmlDoc, originalText, modifiedText, serializer, author, formatHints, generateRedlines = true, targetParagraph = null, diffOptions = {}) {
    void originalText;

    const allParagraphs = targetParagraph
        ? [targetParagraph]
        : getDocumentParagraphs(xmlDoc);

    const { fullText, textSpans } = buildSurgicalTextSpans(allParagraphs);
    const diffs = computeWordDiffs(fullText, modifiedText, diffOptions);
    const spanIndex = buildSpanIndex(textSpans);

    let originalPos = 0;
    let newPos = 0;
    let hasChanges = false;

    for (const [op, text] of diffs) {
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
            if (processDelete(xmlDoc, spanIndex, originalPos, originalPos + text.length, author, generateRedlines)) {
                hasChanges = true;
            }
            originalPos += text.length;
        } else if (op === 1) {
            const textWithoutNewlines = text.replace(/\n/g, ' ');
            if (textWithoutNewlines.trim().length > 0) {
                if (processInsert(xmlDoc, spanIndex, originalPos, textWithoutNewlines, author, formatHints, newPos, generateRedlines, allParagraphs[0] || null)) {
                    hasChanges = true;
                }
            }
            newPos += text.length;
        }
    }

    return withOoxmlSourceType({ oxml: serializer.serializeToString(xmlDoc), hasChanges });
}
