import { createWordElement, isWordElement } from '../core/word-xml.js';
import { getRunChildText, isTextLikeRunChild } from './surgical-spans.js';

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

function cloneRunPiece(xmlDoc, sourceNode, text, asDeletedText) {
    if (asDeletedText) {
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

    const textNode = createWordElement(xmlDoc, 'w:t');
    textNode.setAttribute('xml:space', 'preserve');
    textNode.textContent = text;
    return textNode;
}
