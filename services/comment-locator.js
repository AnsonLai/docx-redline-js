/**
 * Comment text location and marker injection helpers.
 */

import { createWordElement } from '../core/word-xml.js';
import { getElementsByTag, getFirstElementByTag } from '../core/xml-query.js';
import { refreshRunPropertyChangeIds } from '../core/revision-cloning.js';
import { readCanonicalRunText, isNodeVisibleInRevisionView } from '../core/paragraph-text.js';

/**
 * Builds a paragraph text index in a single pass for repeated lookups.
 *
 * @param {Element} paragraph - w:p element
 * @param {Object} [options={}] - Options
 * @param {'accepted'|'rejected'} [options.revisionView='accepted'] - Revision view
 * @returns {{ fullText: string, runOffsets: Array<{run: Element, start: number, end: number}> }}
 */
export function createParagraphTextIndex(paragraph, options = {}) {
    const revisionView = options.revisionView === 'current' ? 'accepted' : (options.revisionView || 'accepted');
    const runs = getElementsByTag(paragraph, 'w:r');
    const runOffsets = [];
    let fullText = '';

    for (const run of runs) {
        if (!isNodeVisibleInRevisionView(run, paragraph, revisionView)) continue;
        const start = fullText.length;
        const text = readCanonicalRunText(run, { revisionView, boundary: paragraph });
        fullText += text;
        runOffsets.push({ run, start, end: fullText.length });
    }

    return { fullText, runOffsets };
}

function candidateOffsets(haystack, needle) {
    if (!needle) return [];
    const candidates = [];
    let offset = haystack.indexOf(needle);
    while (offset !== -1) {
        candidates.push({ start: offset, end: offset + needle.length });
        offset = haystack.indexOf(needle, offset + 1);
    }
    return candidates;
}

function spaceEquivalentText(value) {
    return String(value).replace(/[ \u00a0]/g, ' ');
}

function anchorError(code, message, candidates = []) {
    return {
        code,
        message,
        ...(candidates.length > 0 ? { candidates } : {})
    };
}

function attachRunOffsets(paragraphIndex, candidate, resolvedBy) {
    let startRun = null;
    let endRun = null;
    let startOffset = 0;
    let endOffset = 0;
    for (const { run, start, end } of paragraphIndex.runOffsets) {
        if (candidate.start >= start && candidate.start < end) {
            startRun = run;
            startOffset = candidate.start - start;
        }
        if (candidate.end > start && candidate.end <= end) {
            endRun = run;
            endOffset = candidate.end - start;
        }
    }
    if (!startRun || !endRun) return null;
    return { found: true, resolvedBy, ...candidate, startRun, startOffset, endRun, endOffset };
}

function findAncestorTag(node, tagNames, boundary = null) {
    let current = node?.parentNode;
    while (current && current.nodeType === 1 && current !== boundary && (current.localName || current.nodeName.replace(/^.*:/, '')) !== 'p') {
        const local = current.localName || current.nodeName.replace(/^.*:/, '');
        if (tagNames.includes(local)) return { node: current, tag: local };
        current = current.parentNode;
    }
    return null;
}

function validateAnchorLocation(paragraphIndex, location) {
    if (!location || !location.found) return location;

    const intersectingRuns = paragraphIndex.runOffsets
        .filter(entry => entry.end > location.start && entry.start < location.end)
        .map(entry => entry.run);

    for (const run of intersectingRuns) {
        if (findAncestorTag(run, ['del'])) {
            return {
                found: false,
                error: anchorError('UNSAFE_REVISION_NESTING', 'Refusing to attach comment to pending deletion.', [location])
            };
        }
        if (findAncestorTag(run, ['moveFrom', 'moveTo'])) {
            return {
                found: false,
                error: anchorError('UNSAFE_REVISION_NESTING', 'Refusing to comment on move revision until move lifecycle is designed.', [location])
            };
        }
    }

    return location;
}

/**
 * Resolves a unique comment anchor without mutating the paragraph DOM.
 * Exact matching wins. A one-to-one ASCII-space/NBSP comparison is used only
 * when no exact match exists, so raw offsets remain stable.
 */
export function resolveTextInParagraphIndex(paragraphIndex, searchText) {
    const needle = String(searchText ?? '');
    const exactCandidates = candidateOffsets(paragraphIndex.fullText, needle);
    if (exactCandidates.length > 1) {
        return {
            found: false,
            error: anchorError('AMBIGUOUS_ANCHOR', `Anchor text matched ${exactCandidates.length} locations in the target paragraph.`, exactCandidates)
        };
    }
    if (exactCandidates.length === 1) {
        const resolved = attachRunOffsets(paragraphIndex, exactCandidates[0], 'exact_anchor');
        if (resolved) return validateAnchorLocation(paragraphIndex, resolved);
    }

    const normalizedNeedle = spaceEquivalentText(needle);
    const normalizedParagraph = spaceEquivalentText(paragraphIndex.fullText);
    const equivalentCandidates = candidateOffsets(normalizedParagraph, normalizedNeedle);
    if (equivalentCandidates.length > 1) {
        return {
            found: false,
            error: anchorError('AMBIGUOUS_ANCHOR', `Space-equivalent anchor text matched ${equivalentCandidates.length} locations in the target paragraph.`, equivalentCandidates)
        };
    }
    if (equivalentCandidates.length === 1) {
        const resolved = attachRunOffsets(paragraphIndex, equivalentCandidates[0], 'space_equivalent_anchor');
        if (resolved) return validateAnchorLocation(paragraphIndex, resolved);
    }

    return {
        found: false,
        error: anchorError('ANCHOR_NOT_FOUND', `Could not find anchor text in the target paragraph: "${needle}".`)
    };
}

/**
 * Finds text within a prebuilt paragraph text index.
 *
 * @param {{ fullText: string, runOffsets: Array<{run: Element, start: number, end: number}> }} paragraphIndex - Prebuilt index
 * @param {string} searchText - Text to locate
 * @returns {{ found: boolean, startRun?: Element, startOffset?: number, endRun?: Element, endOffset?: number }}
 */
export function findTextInParagraphIndex(paragraphIndex, searchText) {
    return resolveTextInParagraphIndex(paragraphIndex, searchText);
}

function cloneRunWithText(xmlDoc, rPr, newText, revisionIdAllocator, preserveRevisionIds = false) {
    const newRun = createWordElement(xmlDoc, 'w:r');
    if (rPr) {
        const clonedRPr = rPr.cloneNode(true);
        if (!preserveRevisionIds) {
            refreshRunPropertyChangeIds(clonedRPr, revisionIdAllocator);
        }
        newRun.appendChild(clonedRPr);
    }

    const newTextNode = createWordElement(xmlDoc, 'w:t');
    newTextNode.setAttribute('xml:space', 'preserve');
    newTextNode.textContent = newText;
    newRun.appendChild(newTextNode);
    return newRun;
}

/**
 * Injects comment markers around text in a paragraph.
 *
 * @param {Document} xmlDoc - XML document
 * @param {Element} paragraph - w:p element
 * @param {string} textToFind - Target text
 * @param {number} commentId - Comment id
 * @param {{ fullText: string, runOffsets: Array<{run: Element, start: number, end: number}> }|null} [paragraphIndex=null] - Optional prebuilt index
 * @param {import('../core/types.js').RevisionIdAllocator|null} [revisionIdAllocator=null] - Document-scoped revision ID allocator
 * @returns {boolean}
 */
export function injectMarkersIntoParagraph(xmlDoc, paragraph, textToFind, commentId, paragraphIndex = null, revisionIdAllocator = null, resolvedLocation = null) {
    const activeIndex = paragraphIndex || createParagraphTextIndex(paragraph);
    const location = resolvedLocation || resolveTextInParagraphIndex(activeIndex, textToFind);
    if (!location.found || !location.startRun) {
        return false;
    }

    const startMarker = createWordElement(xmlDoc, 'w:commentRangeStart');
    startMarker.setAttribute('w:id', String(commentId));

    const endMarker = createWordElement(xmlDoc, 'w:commentRangeEnd');
    endMarker.setAttribute('w:id', String(commentId));

    const referenceRun = createWordElement(xmlDoc, 'w:r');
    const reference = createWordElement(xmlDoc, 'w:commentReference');
    reference.setAttribute('w:id', String(commentId));
    referenceRun.appendChild(reference);

    if (location.startRun === location.endRun) {
        const run = location.startRun;
        const textNode = getFirstElementByTag(run, 'w:t');
        if (!textNode) {
            run.parentNode.insertBefore(startMarker, run);
            if (run.nextSibling) {
                run.parentNode.insertBefore(endMarker, run.nextSibling);
                endMarker.parentNode.insertBefore(referenceRun, endMarker.nextSibling);
            } else {
                run.parentNode.appendChild(endMarker);
                run.parentNode.appendChild(referenceRun);
            }
            return true;
        }

        const fullText = textNode.textContent || '';
        const beforeText = fullText.substring(0, location.startOffset);
        const highlightedText = fullText.substring(location.startOffset, location.endOffset);
        const afterText = fullText.substring(location.endOffset);
        const rPr = getFirstElementByTag(run, 'w:rPr');
        const parent = run.parentNode;

        if (beforeText) {
            parent.insertBefore(cloneRunWithText(xmlDoc, rPr, beforeText, revisionIdAllocator, true), run);
            refreshRunPropertyChangeIds(rPr, revisionIdAllocator);
        }

        parent.insertBefore(startMarker, run);
        textNode.textContent = highlightedText;

        if (run.nextSibling) {
            parent.insertBefore(endMarker, run.nextSibling);
        } else {
            parent.appendChild(endMarker);
        }
        parent.insertBefore(referenceRun, endMarker.nextSibling || null);

        if (afterText) {
            parent.insertBefore(cloneRunWithText(xmlDoc, rPr, afterText, revisionIdAllocator), referenceRun.nextSibling || null);
        }

        return true;
    }

    const startTextNode = getFirstElementByTag(location.startRun, 'w:t');
    if (startTextNode && location.startOffset > 0) {
        const fullText = startTextNode.textContent || '';
        const beforeText = fullText.substring(0, location.startOffset);
        const highlightedStart = fullText.substring(location.startOffset);

        if (beforeText) {
            const rPr = getFirstElementByTag(location.startRun, 'w:rPr');
            location.startRun.parentNode.insertBefore(
                cloneRunWithText(xmlDoc, rPr, beforeText, revisionIdAllocator, true),
                location.startRun
            );
            refreshRunPropertyChangeIds(rPr, revisionIdAllocator);
        }
        startTextNode.textContent = highlightedStart;
    }

    location.startRun.parentNode.insertBefore(startMarker, location.startRun);

    const endRun = location.endRun || location.startRun;
    const endTextNode = getFirstElementByTag(endRun, 'w:t');
    if (endTextNode && location.endOffset < (endTextNode.textContent || '').length) {
        const fullText = endTextNode.textContent || '';
        const highlightedEnd = fullText.substring(0, location.endOffset);
        const afterText = fullText.substring(location.endOffset);

        endTextNode.textContent = highlightedEnd;

        if (afterText) {
            const rPr = getFirstElementByTag(endRun, 'w:rPr');
            if (endRun.nextSibling) {
                endRun.parentNode.insertBefore(
                    cloneRunWithText(xmlDoc, rPr, afterText, revisionIdAllocator),
                    endRun.nextSibling
                );
            } else {
                endRun.parentNode.appendChild(cloneRunWithText(xmlDoc, rPr, afterText, revisionIdAllocator));
            }
        }
    }

    if (endRun.nextSibling) {
        endRun.parentNode.insertBefore(endMarker, endRun.nextSibling);
        endMarker.parentNode.insertBefore(referenceRun, endMarker.nextSibling);
    } else {
        endRun.parentNode.appendChild(endMarker);
        endRun.parentNode.appendChild(referenceRun);
    }

    return true;
}
