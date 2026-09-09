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
    sliceRunPieces,
    splitTrackChangeCarrier
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

    const records = [];
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

        records.push({
            runElement,
            parent,
            rPr: runSpans[0].rPr,
            beforePieces: sliceRunPieces(xmlDoc, pieces, 0, deleteStart, false),
            deletedPieces: sliceRunPieces(xmlDoc, pieces, deleteStart, deleteEnd, true),
            afterPieces: sliceRunPieces(xmlDoc, pieces, deleteEnd, getRunTextLength(pieces), false),
            globalStart: Math.max(startPos, Math.min(...runSpans.map(span => span.charStart))),
            globalEnd: Math.min(endPos, Math.max(...runSpans.map(span => span.charEnd))),
            carrierGlobalStart: isWordElement(parent, 'ins') ? getCarrierGlobalStart(spanIndex, parent) : null
        });
    });

    const groups = [];
    for (const record of records) {
        const previousGroup = groups[groups.length - 1];
        const previousRecord = previousGroup?.[previousGroup.length - 1];
        if (
            previousRecord
            && previousRecord.parent === record.parent
            && nextElementSibling(previousRecord.runElement) === record.runElement
        ) {
            previousGroup.push(record);
        } else {
            groups.push([record]);
        }
    }

    let changed = false;
    let usedDelMetadata = false;
    for (const group of groups) {
        const firstRecord = group[0];
        let delWrapper = null;
        if (generateRedlines && group.some(record => record.deletedPieces.length > 0)) {
            const metadata = revisionMetadata
                ? (usedDelMetadata ? { ...revisionMetadata, id: createRevisionMetadata(author, xmlDoc, 'del').id } : revisionMetadata)
                : null;
            usedDelMetadata = true;
            delWrapper = createTrackChange(xmlDoc, 'del', null, author, metadata);
        }

        for (const record of group) {
            const { parent, runElement } = record;
            insertRunPiecesBefore(xmlDoc, parent, runElement, record.beforePieces, record.rPr);
            if (delWrapper && record === firstRecord) {
                parent.insertBefore(delWrapper, runElement);
            }
            if (delWrapper && record.deletedPieces.length > 0) {
                delWrapper.appendChild(createRunFromPieces(xmlDoc, record.deletedPieces, record.rPr));
            }
            const afterRun = insertRunPiecesBefore(xmlDoc, parent, runElement, record.afterPieces, record.rPr);
            if (
                record.globalEnd === endPos
                && !isWordElement(parent, 'ins')
            ) {
                if (!spanIndex.replacementInsertionAnchors) spanIndex.replacementInsertionAnchors = new Map();
                spanIndex.replacementInsertionAnchors.set(endPos, {
                    parent,
                    referenceNode: afterRun || runElement.nextSibling,
                    rPr: record.rPr
                });
            }
            parent.removeChild(runElement);
            changed = true;
        }

        const carrier = isWordElement(firstRecord.parent, 'ins') ? firstRecord.parent : null;
        const groupEnd = Math.max(...group.map(record => record.globalEnd));
        if (carrier && groupEnd === endPos) {
            const carrierStart = firstRecord.carrierGlobalStart;
            const deletedBeforeEnd = group
                .filter(record => record.globalStart < endPos)
                .reduce((sum, record) => sum + record.deletedPieces.reduce((n, piece) => n + (piece.textContent || '').length, 0), 0);
            const currentOffset = Math.max(0, endPos - carrierStart - deletedBeforeEnd);
            if (!spanIndex.revisionInsertionAnchors) spanIndex.revisionInsertionAnchors = new Map();
            spanIndex.revisionInsertionAnchors.set(endPos, {
                carrier,
                splitOffset: currentOffset,
                rPr: firstRecord.rPr
            });
        }
    }

    return changed;
}

export function processInsert(xmlDoc, spanIndex, pos, text, author, formatHints = [], insertOffset = 0, generateRedlines = true, fallbackParagraph = null, revisionMetadata = null, affinity = null, existingRevisions = 'merge-same-author') {
    const mutationAnchor = spanIndex.revisionInsertionAnchors?.get(pos) || null;
    if (
        mutationAnchor
        && existingRevisions === 'slice-cross-author'
        && isConnected(mutationAnchor.carrier)
        && isForeignInsertion(mutationAnchor.carrier, author)
    ) {
        spanIndex.revisionInsertionAnchors.delete(pos);
        return spliceInsertionAtCarrierOffset(
            xmlDoc,
            mutationAnchor.carrier,
            mutationAnchor.splitOffset,
            text,
            mutationAnchor.rPr,
            author,
            formatHints,
            insertOffset,
            generateRedlines,
            revisionMetadata
        );
    }

    const replacementAnchor = spanIndex.replacementInsertionAnchors?.get(pos) || null;
    if (
        replacementAnchor
        && !affinity
        && isConnected(replacementAnchor.parent)
        && (!replacementAnchor.referenceNode || replacementAnchor.referenceNode.parentNode === replacementAnchor.parent)
    ) {
        spanIndex.replacementInsertionAnchors.delete(pos);
        insertTextRuns(
            xmlDoc,
            replacementAnchor.parent,
            replacementAnchor.referenceNode,
            text,
            replacementAnchor.rPr,
            author,
            formatHints,
            insertOffset,
            generateRedlines,
            revisionMetadata
        );
        return true;
    }

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

        const generateNestedRevision = !(
            generateRedlines
            && existingRevisions === 'slice-cross-author'
            && isSameAuthorInsertion(parent, author)
        );

        if (
            generateRedlines
            && existingRevisions === 'slice-cross-author'
            && isForeignInsertion(parent, author)
        ) {
            return spliceInsertionAtCarrierOffset(
                xmlDoc,
                parent,
                getCarrierSplitOffset(spanIndex, parent, pos),
                text,
                targetSpan.rPr,
                author,
                formatHints,
                insertOffset,
                generateRedlines,
                revisionMetadata
            );
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
            insertTextRuns(xmlDoc, parent, targetSpan.runElement, text, targetSpan.rPr, author, formatHints, insertOffset, generateNestedRevision, revisionMetadata);
            insertRunPiecesBefore(xmlDoc, parent, targetSpan.runElement, afterPieces, targetSpan.rPr);
            parent.removeChild(targetSpan.runElement);
            return true;
        }

        const referenceNode = pos <= targetSpan.charStart ? targetSpan.runElement : targetSpan.runElement.nextSibling;
        insertTextRuns(xmlDoc, parent, referenceNode, text, targetSpan.rPr, author, formatHints, insertOffset, generateNestedRevision, revisionMetadata);
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

    if (
        generateRedlines
        && existingRevisions === 'slice-cross-author'
        && isForeignInsertion(parent, author)
    ) {
        return spliceInsertionAtCarrierOffset(
            xmlDoc,
            parent,
            getCarrierSplitOffset(spanIndex, parent, pos),
            text,
            baseRPr,
            author,
            formatHints,
            insertOffset,
            generateRedlines,
            revisionMetadata
        );
    }

    insertTextRuns(xmlDoc, parent, referenceNode, text, baseRPr, author, formatHints, insertOffset, generateRedlines, revisionMetadata);
    return true;
}

function spliceInsertionAtCarrierOffset(xmlDoc, carrier, splitOffset, text, baseRPr, author, formatHints, insertOffset, generateRedlines, revisionMetadata) {
    const parent = carrier.parentNode;
    if (!parent) return false;

    const { leftCarrier, rightCarrier } = splitTrackChangeCarrier(xmlDoc, carrier, splitOffset);
    if (leftCarrier) parent.insertBefore(leftCarrier, carrier);
    insertTextRuns(
        xmlDoc,
        parent,
        carrier,
        text,
        withoutRunPropertyChanges(baseRPr),
        author,
        formatHints,
        insertOffset,
        generateRedlines,
        revisionMetadata
    );
    if (rightCarrier) parent.insertBefore(rightCarrier, carrier);
    parent.removeChild(carrier);
    return true;
}

function withoutRunPropertyChanges(rPr) {
    if (!rPr) return null;
    const clone = rPr.cloneNode(true);
    const changes = Array.from(clone.getElementsByTagName?.('*') || [])
        .filter(node => isWordElement(node, 'rPrChange'));
    changes.forEach(node => node.parentNode?.removeChild(node));
    return clone;
}

function getCarrierSplitOffset(spanIndex, carrier, pos) {
    const carrierStart = getCarrierGlobalStart(spanIndex, carrier);
    const carrierLength = spanIndex.spans
        .filter(span => span.runElement?.parentNode === carrier)
        .reduce((length, span) => length + (span.charEnd - span.charStart), 0);
    return Math.max(0, Math.min(pos - carrierStart, carrierLength));
}

function getCarrierGlobalStart(spanIndex, carrier) {
    const carrierSpans = spanIndex.spans.filter(span => span.runElement?.parentNode === carrier);
    return carrierSpans.length > 0 ? Math.min(...carrierSpans.map(span => span.charStart)) : 0;
}

function isForeignInsertion(node, author) {
    if (!isWordElement(node, 'ins')) return false;
    const carrierAuthor = node.getAttribute('w:author') || node.getAttributeNS?.(NS_W, 'author') || '';
    return carrierAuthor.trim().toLowerCase() !== String(author || '').trim().toLowerCase();
}

function isSameAuthorInsertion(node, author) {
    if (!isWordElement(node, 'ins')) return false;
    const carrierAuthor = node.getAttribute('w:author') || node.getAttributeNS?.(NS_W, 'author') || '';
    return carrierAuthor.trim().toLowerCase() === String(author || '').trim().toLowerCase();
}

function nextElementSibling(node) {
    let sibling = node?.nextSibling || null;
    while (sibling && sibling.nodeType !== 1) sibling = sibling.nextSibling;
    return sibling;
}

function isConnected(node) {
    return !!node?.parentNode;
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
