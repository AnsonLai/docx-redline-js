import { getApplicableFormatHints } from '../pipeline/markdown-processor.js';
import {
    createTrackChange,
    createTextRun,
    createFormattedRuns,
    createTextRunWithRPrElement,
    injectFormattingToRPr
} from './run-builders.js';
import { createRevisionMetadata } from '../core/types.js';
import {
    createRunFromPieces,
    getRunContentPieces,
    getRunTextLength,
    insertRunPiecesBefore,
    sliceRunPieces
} from './surgical-run-splitting.js';
import {
    findContainingSpan,
    findFirstSpanEndingAt,
    findLastSpanEndingBeforeOrAt,
    forEachOverlappingSpan,
    describeInsertionBoundary
} from './surgical-spans.js';
import { extractFormatFromRPr } from './rpr-helpers.js';
import { isWordElement } from '../core/word-xml.js';
import { NS_W } from '../core/types.js';

export function reconcileFormattingForTextSpan(xmlDoc, span, start, end, applicableHints, author, generateRedlines) {
    // Plain modified text carries no negative formatting instruction. Preserve
    // unchanged source formatting unless Markdown explicitly targets this span.
    if (applicableHints.length === 0) return false;

    const rPr = span.rPr;
    const existingFormat = extractFormatFromRPr(rPr);
    const desiredFormat = { ...existingFormat };
    applicableHints.forEach(h => Object.assign(desiredFormat, h.format));

    const formatsToCheck = ['bold', 'italic', 'underline', 'strikethrough'];
    const changesNeeded = formatsToCheck.some(f => !!desiredFormat[f] !== existingFormat[f]);

    if (!changesNeeded) return false;

    const parent = span.runElement.parentNode;
    if (!parent) return false;

    const fullText = span.textElement.textContent || '';
    const runStart = span.charStart;

    const localStart = start - runStart;
    const localEnd = end - runStart;

    const beforeText = fullText.substring(0, localStart);
    const affectedText = fullText.substring(localStart, localEnd);
    const afterText = fullText.substring(localEnd);

    if (beforeText.length > 0) {
        const beforeRun = createTextRun(xmlDoc, beforeText, rPr, false);
        parent.insertBefore(beforeRun, span.runElement);
    }

    const newRPr = injectFormattingToRPr(xmlDoc, rPr, desiredFormat, author, generateRedlines);
    const newRun = createTextRunWithRPrElement(xmlDoc, affectedText, newRPr, false);
    parent.insertBefore(newRun, span.runElement);

    if (afterText.length > 0) {
        const afterRun = createTextRun(xmlDoc, afterText, rPr, false);
        parent.insertBefore(afterRun, span.runElement);
    }

    parent.removeChild(span.runElement);
    return true;
}

export function processDelete(xmlDoc, spanIndex, startPos, endPos, author, generateRedlines, revisionMetadata = null) {
    const spans = [];
    forEachOverlappingSpan(spanIndex, startPos, endPos, span => {
        spans.push(span);
    });

    if (spans.length === 0) return false;

    const spansByRun = new Map();
    spans.forEach(span => {
        if (!span.runElement?.parentNode) return;
        if (!spansByRun.has(span.runElement)) spansByRun.set(span.runElement, []);
        spansByRun.get(span.runElement).push(span);
    });

    let changed = false;
    let usedDelMetadata = false;
    spansByRun.forEach((runSpans, runElement) => {
        const parent = runElement.parentNode;
        if (!parent) return;

        const pieces = getRunContentPieces(runElement);
        if (pieces.length === 0) return;

        let deleteStart = Infinity;
        let deleteEnd = -Infinity;
        runSpans.forEach(span => {
            const piece = pieces.find(candidate => candidate.node === span.textElement);
            if (!piece) return;

            const spanDeleteStart = Math.max(0, startPos - span.charStart);
            const spanDeleteEnd = Math.min(span.charEnd - span.charStart, endPos - span.charStart);
            if (spanDeleteEnd <= spanDeleteStart) return;

            deleteStart = Math.min(deleteStart, piece.start + spanDeleteStart);
            deleteEnd = Math.max(deleteEnd, piece.start + spanDeleteEnd);
        });

        if (!Number.isFinite(deleteStart) || deleteEnd <= deleteStart) return;

        const beforePieces = sliceRunPieces(xmlDoc, pieces, 0, deleteStart, false);
        const deletedPieces = sliceRunPieces(xmlDoc, pieces, deleteStart, deleteEnd, true);
        const afterPieces = sliceRunPieces(xmlDoc, pieces, deleteEnd, getRunTextLength(pieces), false);

        insertRunPiecesBefore(xmlDoc, parent, runElement, beforePieces, runSpans[0].rPr);

        if (generateRedlines && deletedPieces.length > 0) {
            const delRun = createRunFromPieces(xmlDoc, deletedPieces, runSpans[0].rPr);
            const metadata = revisionMetadata
                ? (usedDelMetadata ? { ...revisionMetadata, id: createRevisionMetadata(author, xmlDoc).id } : revisionMetadata)
                : null;
            usedDelMetadata = true;
            const delWrapper = createTrackChange(xmlDoc, 'del', delRun, author, metadata);
            parent.insertBefore(delWrapper, runElement);
        }

        insertRunPiecesBefore(xmlDoc, parent, runElement, afterPieces, runSpans[0].rPr);
        parent.removeChild(runElement);
        changed = true;
    });

    return changed;
}

export function processInsert(xmlDoc, spanIndex, pos, text, author, formatHints = [], insertOffset = 0, generateRedlines = true, fallbackParagraph = null, revisionMetadata = null, affinity = null) {
    if (!affinity) {
        let targetSpan = findContainingSpan(spanIndex, pos);

        if (!targetSpan && pos > 0) {
            targetSpan = findFirstSpanEndingAt(spanIndex, pos);
        }

        if (!targetSpan && pos > 0) {
            targetSpan = findLastSpanEndingBeforeOrAt(spanIndex, pos);
        }

        if (!targetSpan && spanIndex.spans.length > 0) {
            targetSpan = spanIndex.spans[spanIndex.spans.length - 1];
        }

        if (!targetSpan) {
            if (!fallbackParagraph) return false;
            insertTextRuns(xmlDoc, fallbackParagraph, null, text, null, author, formatHints, insertOffset, generateRedlines, revisionMetadata);
            return true;
        }

        const parent = targetSpan.runElement.parentNode;
        if (!parent) {
            if (!fallbackParagraph) return false;
            insertTextRuns(xmlDoc, fallbackParagraph, null, text, targetSpan.rPr, author, formatHints, insertOffset, generateRedlines, revisionMetadata);
            return true;
        }

        const pieces = getRunContentPieces(targetSpan.runElement);
        const targetPiece = pieces.find(piece => piece.node === targetSpan.textElement);
        const localInsertPos = targetPiece
            ? targetPiece.start + Math.max(0, Math.min(pos - targetSpan.charStart, targetSpan.charEnd - targetSpan.charStart))
            : (pos <= targetSpan.charStart ? 0 : getRunTextLength(pieces));

        if (localInsertPos > 0 && localInsertPos < getRunTextLength(pieces)) {
            const beforePieces = sliceRunPieces(xmlDoc, pieces, 0, localInsertPos, false);
            const afterPieces = sliceRunPieces(xmlDoc, pieces, localInsertPos, getRunTextLength(pieces), false);

            insertRunPiecesBefore(xmlDoc, parent, targetSpan.runElement, beforePieces, targetSpan.rPr);
            insertTextRuns(xmlDoc, parent, targetSpan.runElement, text, targetSpan.rPr, author, formatHints, insertOffset, generateRedlines, revisionMetadata);
            insertRunPiecesBefore(xmlDoc, parent, targetSpan.runElement, afterPieces, targetSpan.rPr);
            parent.removeChild(targetSpan.runElement);
            return true;
        }

        const referenceNode = pos <= targetSpan.charStart ? targetSpan.runElement : targetSpan.runElement.nextSibling;
        insertTextRuns(xmlDoc, parent, referenceNode, text, targetSpan.rPr, author, formatHints, insertOffset, generateRedlines, revisionMetadata);
        return true;
    }

    const boundary = describeInsertionBoundary(spanIndex, pos, fallbackParagraph);

    // Validate hyperlink affinity
    const isLeftInHyperlink = boundary.leftSpan && isWordElement(boundary.leftSpan.runElement?.parentNode, 'hyperlink');
    const isRightInHyperlink = boundary.rightSpan && isWordElement(boundary.rightSpan.runElement?.parentNode, 'hyperlink');
    const isContainingInHyperlink = boundary.containingSpan && isWordElement(boundary.containingSpan.runElement?.parentNode, 'hyperlink');

    if (affinity.hyperlink === 'outside') {
        if (boundary.isInterior && isContainingInHyperlink) {
            return {
                error: {
                    code: 'UNSUPPORTED_INSERTION_AFFINITY',
                    message: 'Cannot place insertion outside hyperlink from strictly interior position.'
                }
            };
        }
    } else if (affinity.hyperlink === 'inside') {
        if (!isLeftInHyperlink && !isRightInHyperlink && !isContainingInHyperlink) {
            return {
                error: {
                    code: 'UNSUPPORTED_INSERTION_AFFINITY',
                    message: 'Cannot place insertion inside hyperlink when no hyperlink is present at boundary.'
                }
            };
        }
    }

    // Determine formatting (baseRPr)
    let baseRPr = null;
    if (affinity.formatting === 'none') {
        baseRPr = null;
    } else if (affinity.formatting === 'right') {
        baseRPr = boundary.rightSpan?.rPr || null;
    } else if (affinity.formatting === 'left') {
        baseRPr = boundary.leftSpan?.rPr || null;
    } else {
        baseRPr = (boundary.containingSpan || boundary.leftSpan || boundary.rightSpan)?.rPr || null;
    }

    // Check interior of a run
    if (boundary.isInterior) {
        const targetSpan = boundary.containingSpan;
        const parent = targetSpan.runElement.parentNode || fallbackParagraph;
        if (!parent) return false;

        const pieces = getRunContentPieces(targetSpan.runElement);
        const targetPiece = pieces.find(piece => piece.node === targetSpan.textElement);
        const localInsertPos = targetPiece
            ? targetPiece.start + Math.max(0, Math.min(pos - targetSpan.charStart, targetSpan.charEnd - targetSpan.charStart))
            : (pos <= targetSpan.charStart ? 0 : getRunTextLength(pieces));

        if (localInsertPos > 0 && localInsertPos < getRunTextLength(pieces)) {
            const beforePieces = sliceRunPieces(xmlDoc, pieces, 0, localInsertPos, false);
            const afterPieces = sliceRunPieces(xmlDoc, pieces, localInsertPos, getRunTextLength(pieces), false);

            insertRunPiecesBefore(xmlDoc, parent, targetSpan.runElement, beforePieces, targetSpan.rPr);
            insertTextRuns(xmlDoc, parent, targetSpan.runElement, text, baseRPr, author, formatHints, insertOffset, generateRedlines, revisionMetadata);
            insertRunPiecesBefore(xmlDoc, parent, targetSpan.runElement, afterPieces, targetSpan.rPr);
            parent.removeChild(targetSpan.runElement);
            return true;
        }
    }

    // Boundary between runs or at start/end of paragraph
    let parent = null;
    let referenceNode = null;

    if (affinity.hyperlink === 'outside') {
        if (isRightInHyperlink) {
            const hyperlinkNode = boundary.rightSpan.runElement.parentNode;
            parent = hyperlinkNode.parentNode || fallbackParagraph;
            referenceNode = hyperlinkNode;
        } else if (isLeftInHyperlink) {
            const hyperlinkNode = boundary.leftSpan.runElement.parentNode;
            parent = hyperlinkNode.parentNode || fallbackParagraph;
            referenceNode = hyperlinkNode.nextSibling;
        }
    } else if (affinity.hyperlink === 'inside') {
        if (isRightInHyperlink) {
            parent = boundary.rightSpan.runElement.parentNode;
            referenceNode = boundary.rightSpan.runElement;
        } else if (isLeftInHyperlink) {
            parent = boundary.leftSpan.runElement.parentNode;
            referenceNode = boundary.leftSpan.runElement.nextSibling;
        }
    }

    if (!parent) {
        if (boundary.rightSpan) {
            parent = boundary.rightSpan.runElement.parentNode || fallbackParagraph;
            referenceNode = boundary.rightSpan.runElement;
        } else if (boundary.leftSpan) {
            parent = boundary.leftSpan.runElement.parentNode || fallbackParagraph;
            referenceNode = boundary.leftSpan.runElement.nextSibling;
        } else {
            parent = fallbackParagraph;
            referenceNode = null;
        }
    }

    // Check bookmark range affinity
    if (affinity.bookmark && parent) {
        if (affinity.bookmark === 'outside') {
            if (referenceNode && isWordElement(referenceNode.previousSibling, 'bookmarkStart')) {
                referenceNode = referenceNode.previousSibling;
            }
            if (boundary.leftSpan && isWordElement(boundary.leftSpan.runElement.nextSibling, 'bookmarkEnd')) {
                referenceNode = boundary.leftSpan.runElement.nextSibling.nextSibling;
            }
        } else if (affinity.bookmark === 'inside') {
            if (referenceNode && isWordElement(referenceNode, 'bookmarkStart')) {
                referenceNode = referenceNode.nextSibling;
            }
            if (boundary.leftSpan && isWordElement(boundary.leftSpan.runElement.nextSibling, 'bookmarkEnd')) {
                referenceNode = boundary.leftSpan.runElement.nextSibling;
            }
        }
    }

    // Check comment range affinity
    if (affinity.comment && parent) {
        if (affinity.comment === 'outside') {
            if (referenceNode && isWordElement(referenceNode.previousSibling, 'commentRangeStart')) {
                referenceNode = referenceNode.previousSibling;
            }
            if (boundary.leftSpan && isWordElement(boundary.leftSpan.runElement.nextSibling, 'commentRangeEnd')) {
                let afterComment = boundary.leftSpan.runElement.nextSibling.nextSibling;
                if (afterComment && (isWordElement(afterComment, 'commentReference') || isWordElement(afterComment, 'r'))) {
                    const hasCRef = Array.from(afterComment.childNodes || []).some(n => isWordElement(n, 'commentReference'));
                    if (hasCRef) afterComment = afterComment.nextSibling;
                }
                referenceNode = afterComment;
            }
        } else if (affinity.comment === 'inside') {
            if (referenceNode && isWordElement(referenceNode, 'commentRangeStart')) {
                referenceNode = referenceNode.nextSibling;
            }
            if (boundary.leftSpan && isWordElement(boundary.leftSpan.runElement.nextSibling, 'commentRangeEnd')) {
                referenceNode = boundary.leftSpan.runElement.nextSibling;
            }
        }
    }

    // Check revision affinity (coalesce_same_author)
    if (generateRedlines && affinity.revision === 'coalesce_same_author') {
        let insElem = null;
        let insRef = null;

        if (boundary.leftSpan && isWordElement(boundary.leftSpan.runElement.parentNode, 'ins')) {
            const candidate = boundary.leftSpan.runElement.parentNode;
            const candAuthor = candidate.getAttribute('w:author') || candidate.getAttributeNS(NS_W, 'author');
            if (candAuthor === author) {
                insElem = candidate;
                insRef = boundary.leftSpan.runElement.nextSibling;
            }
        } else if (boundary.rightSpan && isWordElement(boundary.rightSpan.runElement.parentNode, 'ins')) {
            const candidate = boundary.rightSpan.runElement.parentNode;
            const candAuthor = candidate.getAttribute('w:author') || candidate.getAttributeNS(NS_W, 'author');
            if (candAuthor === author) {
                insElem = candidate;
                insRef = boundary.rightSpan.runElement;
            }
        }

        if (insElem) {
            const insRun = createTextRun(xmlDoc, text, baseRPr, false);
            insElem.insertBefore(insRun, insRef);
            return true;
        }
    }

    insertTextRuns(xmlDoc, parent, referenceNode, text, baseRPr, author, formatHints, insertOffset, generateRedlines, revisionMetadata);
    return true;
}

function insertTextRuns(xmlDoc, parent, referenceNode, text, baseRPr, author, formatHints, insertOffset, generateRedlines, revisionMetadata = null) {
    const applicableHints = getApplicableFormatHints(formatHints, insertOffset, insertOffset + text.length);

    if (applicableHints.length === 0) {
        const insRun = createTextRun(xmlDoc, text, baseRPr, false);
        if (generateRedlines) {
            const insWrapper = createTrackChange(xmlDoc, 'ins', insRun, author, revisionMetadata);
            parent.insertBefore(insWrapper, referenceNode);
        } else {
            parent.insertBefore(insRun, referenceNode);
        }
        return;
    }

    const runs = createFormattedRuns(xmlDoc, text, baseRPr, applicableHints, insertOffset, author, generateRedlines);

    if (generateRedlines) {
        const insWrapper = createTrackChange(xmlDoc, 'ins', null, author, revisionMetadata);
        runs.forEach(run => insWrapper.appendChild(run));
        parent.insertBefore(insWrapper, referenceNode);
    } else {
        runs.forEach(run => parent.insertBefore(run, referenceNode));
    }
}
