/**
 * Reconstruction writer.
 *
 * Applies diff segments to mapped reconstruction context and writes updated DOM content.
 */

import { getApplicableFormatHints } from '../pipeline/markdown-processor.js';
import {
    createTrackChange,
    createFormattedRuns,
    markParagraphMarkDeleted,
    markParagraphMarkInserted
} from './run-builders.js';
import { getFirstElementByTagNSOrTag } from '../core/xml-query.js';
import { NS_W, createReplacementRevisionEvent, createRevisionMetadata } from '../core/types.js';
import { createWordElement, isWordElement } from '../core/word-xml.js';

/**
 * Applies diffs to reconstruction context and writes updated XML.
 *
 * @param {Document} xmlDoc - XML document
 * @param {Array<[number, string]>} diffs - Diff tuples from diff-match-patch
 * @param {ReturnType<import('./reconstruction-mapper.js').buildReconstructionMapping>} context - Reconstruction mapping
 * @param {XMLSerializer} serializer - Serializer instance
 * @param {string} author - Author name
 * @param {Array} formatHints - Format hints
 * @param {boolean} [generateRedlines=true] - Track change toggle
 * @returns {{ oxml: string, hasChanges: boolean }}
 */
export function applyReconstructionDiffs(xmlDoc, diffs, context, serializer, author, formatHints, generateRedlines = true, options = {}) {
    const {
        paragraphs,
        containerFragments,
        sentinelMapByStart,
        referenceMap,
        replacementContainers,
        getParagraphInfo,
        getRunProperties,
        getPropertySpanLength,
        isParagraphStart
    } = context;

    const createNewParagraph = (pPr) => {
        const newParagraph = createWordElement(xmlDoc, 'w:p');
        if (pPr) {
            const clonedPPr = pPr.cloneNode(true);
            // Safeguard: Never place w:sectPr on both paragraphs during split
            const sectPr = getFirstElementByTagNSOrTag(clonedPPr, NS_W, 'sectPr');
            if (sectPr) {
                clonedPPr.removeChild(sectPr);
            }
            newParagraph.appendChild(clonedPPr);
        }
        return newParagraph;
    };

    const startInfo = getParagraphInfo(0);
    let currentParagraph = createNewParagraph(startInfo.pPr);
    const initialFragment = containerFragments.get(startInfo.container);
    if (initialFragment) {
        initialFragment.appendChild(currentParagraph);
    }

    let currentOriginalIndex = 0;
    let currentInsertOffset = 0;
    const pairReplacements = options?.pairReplacements === true;
    let pendingReplacementStart = null;
    let pendingReplacementEvent = null;
    const emittedCommentMarkers = new WeakSet();

    for (let diffIndex = 0; diffIndex < diffs.length; diffIndex++) {
        const [op, text] = diffs[diffIndex];
        if (op === 0 || op === -1) {
            const type = op === 0 ? 'equal' : 'delete';
            if (op === 0) {
                pendingReplacementStart = null;
                pendingReplacementEvent = null;
            } else if (pendingReplacementStart === null) {
                pendingReplacementStart = currentOriginalIndex;
                let hasInsert = false;
                for (let k = diffIndex + 1; k < diffs.length; k++) {
                    if (diffs[k][0] === 1) { hasInsert = true; break; }
                    if (diffs[k][0] === 0) break;
                }
                if (pairReplacements && hasInsert) {
                    pendingReplacementEvent = createReplacementRevisionEvent(author, xmlDoc);
                }
            }
            let offset = 0;

            while (offset < text.length) {
                const chunkStart = currentOriginalIndex + offset;
                const properties = getRunProperties(chunkStart);
                const chunkLength = getPropertySpanLength(chunkStart, text.length - offset);
                const chunk = text.substring(offset, offset + chunkLength);

                const appendResult = appendTextToCurrent(
                    xmlDoc,
                    chunk,
                    type,
                    properties.rPr,
                    properties.wrapper,
                    chunkStart,
                    currentParagraph,
                    containerFragments,
                    sentinelMapByStart,
                    referenceMap,
                    replacementContainers,
                    getParagraphInfo,
                    createNewParagraph,
                    author,
                    formatHints,
                    currentInsertOffset,
                    generateRedlines,
                    emittedCommentMarkers,
                    pendingReplacementEvent
                );
                currentParagraph = appendResult.currentParagraph;

                if (op === 0) {
                    currentInsertOffset += chunkLength;
                }
                offset += chunkLength;
            }

            currentOriginalIndex += text.length;
            continue;
        }

        if (op === 1) {
            const propertyIndex = pendingReplacementStart !== null
                ? pendingReplacementStart
                : (currentOriginalIndex > 0 && !isParagraphStart(currentOriginalIndex)
                    ? currentOriginalIndex - 1
                    : currentOriginalIndex);
            const properties = getRunProperties(propertyIndex);

            const appendResult = appendTextToCurrent(
                xmlDoc,
                text,
                'insert',
                properties.rPr,
                properties.wrapper,
                currentOriginalIndex,
                currentParagraph,
                containerFragments,
                sentinelMapByStart,
                referenceMap,
                replacementContainers,
                getParagraphInfo,
                createNewParagraph,
                author,
                formatHints,
                currentInsertOffset,
                generateRedlines,
                emittedCommentMarkers,
                pendingReplacementEvent
            );
            currentParagraph = appendResult.currentParagraph;
            currentInsertOffset += text.length;
            pendingReplacementStart = null;
            pendingReplacementEvent = null;
        }
    }

    if (generateRedlines) {
        containerFragments.forEach(fragment => {
            Array.from(fragment.childNodes).forEach(node => {
                if (isWordElement(node, 'p')) {
                    const hasVisibleText = Array.from(node.getElementsByTagNameNS(NS_W, 't')).length > 0;
                    const hasDeletedText = Array.from(node.getElementsByTagNameNS(NS_W, 'delText')).length > 0;
                    if (paragraphs.length === 1 && !hasVisibleText && (hasDeletedText || context.originalFullText.trim() !== '')) {
                        markParagraphMarkDeleted(xmlDoc, node, author);
                    }
                }
            });
        });
    }

    containerFragments.forEach(fragment => {
        const createdParagraphs = Array.from(fragment.childNodes).filter(node => isWordElement(node, 'p'));
        if (createdParagraphs.length > 1) {
            const firstP = createdParagraphs[0];
            const lastP = createdParagraphs[createdParagraphs.length - 1];
            const firstPPr = getFirstElementByTagNSOrTag(firstP, NS_W, 'pPr');
            const sectPr = firstPPr ? getFirstElementByTagNSOrTag(firstPPr, NS_W, 'sectPr') : null;
            if (sectPr) {
                firstPPr.removeChild(sectPr);
                let lastPPr = getFirstElementByTagNSOrTag(lastP, NS_W, 'pPr');
                if (!lastPPr) {
                    lastPPr = createWordElement(xmlDoc, 'w:pPr');
                    lastP.insertBefore(lastPPr, lastP.firstChild || null);
                }
                lastPPr.appendChild(sectPr);
            }
        }
    });

    const paragraphSet = new Set(paragraphs);
    const insertionAnchors = new Map();
    paragraphs.forEach(paragraph => {
        const container = paragraph.parentNode;
        if (!container || insertionAnchors.has(container)) return;
        let anchor = paragraph.nextSibling;
        while (anchor && paragraphSet.has(anchor)) anchor = anchor.nextSibling;
        insertionAnchors.set(container, anchor);
    });

    paragraphs.forEach(paragraph => {
        if (paragraph.parentNode) {
            paragraph.parentNode.removeChild(paragraph);
        }
    });

    let hasDocumentTarget = false;
    let serializedDocumentOutput = '';

    containerFragments.forEach((fragment, container) => {
        const replacement = replacementContainers.get(container);
        const target = replacement || container;

        if (target.nodeType === 9) {
            hasDocumentTarget = true;
            if (fragment.childNodes.length === 1) {
                target.appendChild(fragment.firstChild);
            } else {
                serializedDocumentOutput = Array.from(fragment.childNodes)
                    .map(node => serializer.serializeToString(node))
                    .join('');
            }
            return;
        }

        const anchor = replacement ? null : insertionAnchors.get(container);
        if (anchor && anchor.parentNode === target) {
            target.insertBefore(fragment, anchor);
        } else {
            target.appendChild(fragment);
        }
    });

    const oxml = (hasDocumentTarget && serializedDocumentOutput)
        ? serializedDocumentOutput
        : serializer.serializeToString(xmlDoc);

    return { oxml, hasChanges: true };
}

function appendTextToCurrent(
    xmlDoc,
    text,
    type,
    rPr,
    wrapper,
    baseIndex,
    currentParagraphRef,
    containerFragments,
    sentinelMapByStart,
    referenceMap,
    replacementContainers,
    getParagraphInfo,
    createNewParagraph,
    author,
    formatHints = [],
    insertOffset = 0,
    generateRedlines = true,
    emittedCommentMarkers = new WeakSet(),
    replacementEvent = null
) {
    let localBaseIndex = baseIndex;
    let localInsertOffset = insertOffset;
    let localParagraph = currentParagraphRef;

    const parts = text.split(/([\n\uFFFC]|[\uE000-\uF8FF])/);

    parts.forEach(part => {
        const sentinelsAtOffset = sentinelMapByStart.get(localBaseIndex) || [];
        const commentMarkers = sentinelsAtOffset.filter(sentinel => sentinel.isCommentMarker && !emittedCommentMarkers.has(sentinel.node));

        commentMarkers.forEach(marker => {
            emittedCommentMarkers.add(marker.node);
            if (isWordElement(marker.node, 'commentReference')) {
                const run = createWordElement(xmlDoc, 'w:r');
                run.appendChild(marker.node.cloneNode(true));
                localParagraph.appendChild(run);
            } else {
                localParagraph.appendChild(marker.node.cloneNode(true));
            }
        });

        if (part === '\n') {
            const info = getParagraphInfo(localBaseIndex + 1);
            const nextParagraph = createNewParagraph(info.pPr);
            if (generateRedlines && type === 'insert') {
                markParagraphMarkInserted(xmlDoc, localParagraph, author);
            } else if (generateRedlines && type === 'delete') {
                markParagraphMarkDeleted(xmlDoc, localParagraph, author);
            }

            const fragment = containerFragments.get(info.container);
            if (fragment) {
                fragment.appendChild(nextParagraph);
                localParagraph = nextParagraph;
            }
            localBaseIndex++;
            if (type !== 'delete') localInsertOffset++;
            return;
        }

        if (part === '\uFFFC') {
            const sentinel = sentinelsAtOffset.find(entry => !entry.isCommentMarker) || sentinelsAtOffset[0];
            if (sentinel) {
                const clone = sentinel.node.cloneNode(true);
                if (sentinel.isTextBox && sentinel.originalContainer) {
                    const newContainer = getFirstElementByTagNSOrTag(clone, NS_W, 'txbxContent');
                    if (newContainer) {
                        while (newContainer.firstChild) newContainer.removeChild(newContainer.firstChild);
                        replacementContainers.set(sentinel.originalContainer, newContainer);
                    }
                }
                if (sentinel.wrapInRun) {
                    const run = createWordElement(xmlDoc, 'w:r');
                    if (sentinel.rPr) run.appendChild(sentinel.rPr.cloneNode(true));
                    run.appendChild(clone);
                    localParagraph.appendChild(run);
                } else {
                    localParagraph.appendChild(clone);
                }
            }
            localBaseIndex++;
            if (type !== 'delete') localInsertOffset++;
            return;
        }

        if (referenceMap.has(part)) {
            if (type !== 'delete') {
                const refNode = referenceMap.get(part);
                if (refNode) {
                    const clone = refNode.cloneNode(true);
                    const run = createWordElement(xmlDoc, 'w:r');
                    if (rPr) run.appendChild(rPr.cloneNode(true));
                    run.appendChild(clone);
                    localParagraph.appendChild(run);
                }
            }
            localBaseIndex++;
            if (type !== 'delete') localInsertOffset++;
            return;
        }

        if (part.length === 0) return;

        let parent = localParagraph;
        if (wrapper) {
            const wrapperClone = wrapper.cloneNode(false);
            parent = wrapperClone;
            localParagraph.appendChild(wrapperClone);
        }

        if (type === 'delete') {
            const run = createWordElement(xmlDoc, 'w:r');
            if (rPr) run.appendChild(rPr.cloneNode(true));
            const delText = createWordElement(xmlDoc, 'w:delText');
            delText.setAttribute('xml:space', 'preserve');
            delText.textContent = part;
            run.appendChild(delText);

            if (generateRedlines) {
                let metadata = null;
                if (replacementEvent) {
                    const id = replacementEvent.usedDeletionId
                        ? createRevisionMetadata(author, xmlDoc, 'del').id
                        : replacementEvent.deletionId;
                    replacementEvent.usedDeletionId = true;
                    metadata = {
                        id,
                        author: replacementEvent.author,
                        date: replacementEvent.date
                    };
                }
                const del = createTrackChange(xmlDoc, 'del', run, author, metadata);
                parent.appendChild(del);
            }
        } else {
            const applicableHints = getApplicableFormatHints(formatHints, localInsertOffset, localInsertOffset + part.length);
            const runs = createFormattedRuns(xmlDoc, part, rPr, applicableHints, localInsertOffset, author, generateRedlines);

            if (type === 'insert' && generateRedlines) {
                let metadata = null;
                if (replacementEvent) {
                    const id = replacementEvent.usedInsertionId
                        ? createRevisionMetadata(author, xmlDoc, 'ins').id
                        : replacementEvent.insertionId;
                    replacementEvent.usedInsertionId = true;
                    metadata = {
                        id,
                        author: replacementEvent.author,
                        date: replacementEvent.date
                    };
                }
                const ins = createTrackChange(xmlDoc, 'ins', null, author, metadata);
                runs.forEach(run => ins.appendChild(run));
                parent.appendChild(ins);
            } else {
                runs.forEach(run => parent.appendChild(run));
            }
        }

        if (type !== 'delete') {
            localInsertOffset += part.length;
        }
        localBaseIndex += part.length;
    });

    return { currentParagraph: localParagraph };
}
