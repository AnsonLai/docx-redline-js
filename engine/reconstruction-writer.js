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
import { NS_W } from '../core/types.js';
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
export function applyReconstructionDiffs(xmlDoc, diffs, context, serializer, author, formatHints, generateRedlines = true) {
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
        if (pPr) newParagraph.appendChild(pPr.cloneNode(true));
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
    const emittedCommentMarkers = new WeakSet();

    for (const [op, text] of diffs) {
        if (op === 0 || op === -1) {
            const type = op === 0 ? 'equal' : 'delete';
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
                    emittedCommentMarkers
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
            const properties = currentOriginalIndex > 0 && !isParagraphStart(currentOriginalIndex)
                ? getRunProperties(currentOriginalIndex - 1)
                : getRunProperties(currentOriginalIndex);

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
                emittedCommentMarkers
            );
            currentParagraph = appendResult.currentParagraph;
            currentInsertOffset += text.length;
        }
    }

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

    containerFragments.forEach((fragment, container) => {
        const replacement = replacementContainers.get(container);
        const target = replacement || container;

        if (target.nodeType === 9) {
            const firstChild = fragment.firstChild;
            if (firstChild) {
                target.appendChild(firstChild);
                while (fragment.firstChild) {
                    target.documentElement.appendChild(fragment.firstChild);
                }
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

    return { oxml: serializer.serializeToString(xmlDoc), hasChanges: true };
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
    emittedCommentMarkers = new WeakSet()
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
                localParagraph.appendChild(clone);
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
                const del = createTrackChange(xmlDoc, 'del', run, author);
                parent.appendChild(del);
            }
        } else {
            const applicableHints = getApplicableFormatHints(formatHints, localInsertOffset, localInsertOffset + part.length);
            const runs = createFormattedRuns(xmlDoc, part, rPr, applicableHints, localInsertOffset, author, generateRedlines);

            if (type === 'insert' && generateRedlines) {
                const ins = createTrackChange(xmlDoc, 'ins', null, author);
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
