import { getApplicableFormatHints } from '../pipeline/markdown-processor.js';
import {
    createTrackChange,
    createTextRun,
    createFormattedRuns,
    createTextRunWithRPrElement,
    injectFormattingToRPr
} from './run-builders.js';
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
    forEachOverlappingSpan
} from './surgical-spans.js';
import { extractFormatFromRPr } from './rpr-helpers.js';

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

export function processDelete(xmlDoc, spanIndex, startPos, endPos, author, generateRedlines) {
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
            const delWrapper = createTrackChange(xmlDoc, 'del', delRun, author);
            parent.insertBefore(delWrapper, runElement);
        }

        insertRunPiecesBefore(xmlDoc, parent, runElement, afterPieces, runSpans[0].rPr);
        parent.removeChild(runElement);
        changed = true;
    });

    return changed;
}

export function processInsert(xmlDoc, spanIndex, pos, text, author, formatHints = [], insertOffset = 0, generateRedlines = true, fallbackParagraph = null) {
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
        insertTextRuns(xmlDoc, fallbackParagraph, null, text, null, author, formatHints, insertOffset, generateRedlines);
        return true;
    }

    const parent = targetSpan.runElement.parentNode;
    if (!parent) {
        if (!fallbackParagraph) return false;
        insertTextRuns(xmlDoc, fallbackParagraph, null, text, targetSpan.rPr, author, formatHints, insertOffset, generateRedlines);
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
        insertTextRuns(xmlDoc, parent, targetSpan.runElement, text, targetSpan.rPr, author, formatHints, insertOffset, generateRedlines);
        insertRunPiecesBefore(xmlDoc, parent, targetSpan.runElement, afterPieces, targetSpan.rPr);
        parent.removeChild(targetSpan.runElement);
        return true;
    }

    const referenceNode = pos <= targetSpan.charStart ? targetSpan.runElement : targetSpan.runElement.nextSibling;
    insertTextRuns(xmlDoc, parent, referenceNode, text, targetSpan.rPr, author, formatHints, insertOffset, generateRedlines);
    return true;
}

function insertTextRuns(xmlDoc, parent, referenceNode, text, baseRPr, author, formatHints, insertOffset, generateRedlines) {
    const applicableHints = getApplicableFormatHints(formatHints, insertOffset, insertOffset + text.length);

    if (applicableHints.length === 0) {
        const insRun = createTextRun(xmlDoc, text, baseRPr, false);
        if (generateRedlines) {
            const insWrapper = createTrackChange(xmlDoc, 'ins', insRun, author);
            parent.insertBefore(insWrapper, referenceNode);
        } else {
            parent.insertBefore(insRun, referenceNode);
        }
        return;
    }

    const runs = createFormattedRuns(xmlDoc, text, baseRPr, applicableHints, insertOffset, author, generateRedlines);

    if (generateRedlines) {
        const insWrapper = createTrackChange(xmlDoc, 'ins', null, author);
        runs.forEach(run => insWrapper.appendChild(run));
        parent.insertBefore(insWrapper, referenceNode);
    } else {
        runs.forEach(run => parent.insertBefore(run, referenceNode));
    }
}
