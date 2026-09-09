import { createWordElement, isWordElement } from '../core/word-xml.js';
import {
    NS_W,
    RevisionIdAllocator,
    createRevisionIdAllocator,
    getRevisionIdAllocatorForDocument
} from '../core/types.js';
import { refreshRunPropertyChangeIds } from '../core/revision-cloning.js';
import { getRunChildText, isTextLikeRunChild } from './surgical-spans.js';

const TRACK_CHANGE_CARRIERS = new Set(['ins', 'del']);

export function getRunContentPieces(runElement) {
    const pieces = [];
    let offset = 0;

    for (const child of Array.from(runElement.childNodes || [])) {
        if (isWordElement(child, 'rPr')) continue;
        if (!isTextLikeRunChild(child)) continue;

        const text = getRunChildText(child);
        if (text.length === 0) continue;

        pieces.push({
            node: child,
            start: offset,
            end: offset + text.length,
            text
        });
        offset += text.length;
    }

    return pieces;
}

export function getRunTextLength(pieces) {
    if (pieces.length === 0) return 0;
    return pieces[pieces.length - 1].end;
}

export function sliceRunPieces(xmlDoc, pieces, start, end, asDeletedText) {
    const sliced = [];
    if (end <= start) return sliced;

    pieces.forEach(piece => {
        const overlapStart = Math.max(start, piece.start);
        const overlapEnd = Math.min(end, piece.end);
        if (overlapEnd <= overlapStart) return;

        const localStart = overlapStart - piece.start;
        const localEnd = overlapEnd - piece.start;
        const text = piece.text.slice(localStart, localEnd);

        sliced.push(cloneRunPiece(xmlDoc, piece.node, text, asDeletedText));
    });

    return sliced;
}

export function createRunFromPieces(xmlDoc, pieces, rPr) {
    const run = createWordElement(xmlDoc, 'w:r');
    if (rPr) run.appendChild(rPr.cloneNode(true));
    pieces.forEach(piece => run.appendChild(piece));
    return run;
}

export function insertRunPiecesBefore(xmlDoc, parent, referenceNode, pieces, rPr) {
    if (pieces.length === 0) return null;
    const run = createRunFromPieces(xmlDoc, pieces, rPr);
    parent.insertBefore(run, referenceNode);
    return run;
}

/**
 * Splits a run-level tracked-change carrier at its visible-view character
 * offset without mutating the source carrier. The original revision ID stays
 * with the leading fragment; an interior trailing fragment receives a fresh,
 * document-scoped ID while all other carrier metadata remains unchanged.
 *
 * @param {Document} xmlDoc
 * @param {Element} carrierElement
 * @param {number} splitOffset
 * @param {RevisionIdAllocator|null} [allocator=null]
 * @returns {{ leftCarrier: Element|null, rightCarrier: Element|null }}
 */
export function splitTrackChangeCarrier(xmlDoc, carrierElement, splitOffset, allocator = null) {
    const carrierName = getLocalName(carrierElement);
    if (!TRACK_CHANGE_CARRIERS.has(carrierName)) {
        throw new TypeError('splitTrackChangeCarrier requires a w:ins or w:del carrier.');
    }
    if (!Number.isInteger(splitOffset) || splitOffset < 0) {
        throw new RangeError('splitOffset must be a non-negative integer.');
    }

    const children = Array.from(carrierElement.childNodes || []);
    const totalLength = children.reduce((length, child) => {
        return length + (isWordElement(child, 'r') ? getRunTextLength(getRunContentPieces(child)) : 0);
    }, 0);
    if (splitOffset > totalLength) {
        throw new RangeError(`splitOffset ${splitOffset} exceeds carrier text length ${totalLength}.`);
    }

    if (splitOffset === 0) {
        return { leftCarrier: null, rightCarrier: carrierElement.cloneNode(true) };
    }
    if (splitOffset === totalLength) {
        return { leftCarrier: carrierElement.cloneNode(true), rightCarrier: null };
    }

    const leftCarrier = carrierElement.cloneNode(false);
    const rightCarrier = carrierElement.cloneNode(false);
    let offset = 0;

    for (const child of children) {
        if (!isWordElement(child, 'r')) {
            const destination = offset <= splitOffset ? leftCarrier : rightCarrier;
            destination.appendChild(child.cloneNode(true));
            continue;
        }

        const pieces = getRunContentPieces(child);
        const runLength = getRunTextLength(pieces);
        const runEnd = offset + runLength;

        if (runEnd <= splitOffset) {
            leftCarrier.appendChild(child.cloneNode(true));
        } else if (offset >= splitOffset) {
            rightCarrier.appendChild(child.cloneNode(true));
        } else {
            const localOffset = splitOffset - offset;
            const rPr = Array.from(child.childNodes || []).find(node => isWordElement(node, 'rPr')) || null;
            const asDeletedText = carrierName === 'del';
            const leftPieces = sliceRunPieces(xmlDoc, pieces, 0, localOffset, asDeletedText);
            const rightPieces = sliceRunPieces(xmlDoc, pieces, localOffset, runLength, asDeletedText);
            leftCarrier.appendChild(createRunFromPieces(xmlDoc, leftPieces, rPr));
            const rightRun = createRunFromPieces(xmlDoc, rightPieces, rPr);
            rightCarrier.appendChild(rightRun);
        }
        offset = runEnd;
    }

    const resolvedAllocator = resolveAllocator(xmlDoc, allocator);
    refreshRunPropertyChangeIds(rightCarrier, resolvedAllocator);
    const nextId = resolvedAllocator.next();
    setWordAttribute(rightCarrier, 'id', String(nextId));
    resolvedAllocator._receiptCollector?.recordRevision(nextId, carrierName);

    return { leftCarrier, rightCarrier };
}

function resolveAllocator(xmlDoc, allocator) {
    return allocator instanceof RevisionIdAllocator
        ? allocator
        : (getRevisionIdAllocatorForDocument(xmlDoc) || createRevisionIdAllocator(xmlDoc));
}

function setWordAttribute(element, localName, value) {
    if (typeof element.setAttributeNS === 'function') {
        element.setAttributeNS(NS_W, `w:${localName}`, value);
    } else {
        element.setAttribute(`w:${localName}`, value);
    }
}

function getLocalName(element) {
    return String(element?.localName || element?.nodeName || '').replace(/^.*:/, '');
}

function cloneRunPiece(xmlDoc, sourceNode, text, asDeletedText) {
    if (asDeletedText && (isWordElement(sourceNode, 'delText') || isWordElement(sourceNode, 't'))) {
        const delText = createWordElement(xmlDoc, 'w:delText');
        delText.setAttribute('xml:space', 'preserve');
        delText.textContent = text;
        return delText;
    }

    if (isWordElement(sourceNode, 't')) {
        const textNode = sourceNode.cloneNode(false);
        textNode.textContent = text;
        if (/^\s|\s$/.test(text)) {
            textNode.setAttribute('xml:space', 'preserve');
        }
        return textNode;
    }

    if (text === '\n' && (isWordElement(sourceNode, 'br') || isWordElement(sourceNode, 'cr'))) {
        return sourceNode.cloneNode(true);
    }
    if (text === '\t' && isWordElement(sourceNode, 'tab')) {
        return sourceNode.cloneNode(true);
    }
    if (text === '\u2011' && isWordElement(sourceNode, 'noBreakHyphen')) {
        return sourceNode.cloneNode(true);
    }

    if (text === '\u00ad' && isWordElement(sourceNode, 'softHyphen')) {
        return sourceNode.cloneNode(true);
    }

    if (asDeletedText) {
        const delText = createWordElement(xmlDoc, 'w:delText');
        delText.setAttribute('xml:space', 'preserve');
        delText.textContent = text;
        return delText;
    }

    const textNode = createWordElement(xmlDoc, 'w:t');
    textNode.setAttribute('xml:space', 'preserve');
    textNode.textContent = text;
    return textNode;
}
