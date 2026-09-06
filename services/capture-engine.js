/**
 * Capture entity management, resolution, and validation for batch operations.
 */

import { NS_W } from '../core/types.js';
import {
    createParagraphFingerprint,
    getDocumentParagraphNodes,
    getParagraphId,
    findContainingWordElement
} from '../core/paragraph-targeting.js';
import { extractCanonicalParagraphText } from '../core/paragraph-text.js';

function createCaptureError(code, message) {
    const error = new Error(message);
    error.code = code;
    return error;
}

export function ensureParagraphIdsOnImportedNode(node, session) {
    // Preserves existing paragraph attributes without injecting undeclared namespaces into the document.
}

/**
 * Derives a serializable CapturedEntity record from the live DOM nodes produced by an operation.
 *
 * @param {import('./document-operation-session.js').DocumentOperationSession} session
 * @param {object} operation
 * @param {Array<Element>} liveNodes
 * @returns {object} CapturedEntity
 */
export function deriveCapturedEntity(session, operation, liveNodes) {
    const paragraphs = [];
    let hasTable = false;
    let hasList = false;

    for (const node of (liveNodes || [])) {
        if (!node) continue;
        if (node.localName === 'tbl') {
            hasTable = true;
            paragraphs.push(...Array.from(node.getElementsByTagNameNS(NS_W, 'p')));
        } else if (node.localName === 'p') {
            if (node.getElementsByTagNameNS(NS_W, 'numPr').length > 0) {
                hasList = true;
            }
            paragraphs.push(node);
        }
    }

    const kind = hasTable
        ? 'table'
        : (hasList ? 'list' : (paragraphs.length > 1 ? 'range' : 'paragraph'));

    const generatedParagraphIds = [];
    const fingerprints = [];
    const expectedText = [];
    const structuralPathHints = [];

    const xmlDoc = session.document;
    const allDocParagraphs = getDocumentParagraphNodes(xmlDoc);

    for (const p of paragraphs) {
        const paraId = getParagraphId(p) || '';
        generatedParagraphIds.push(paraId);
        const text = extractCanonicalParagraphText(p);
        expectedText.push(text);
        const pIndex = allDocParagraphs.indexOf(p) + 1;
        fingerprints.push(createParagraphFingerprint(p, {
            text,
            paragraphId: paraId || null,
            index: pIndex > 0 ? pIndex : undefined,
            inTable: hasTable || !!findContainingWordElement(p, 'tbl')
        }));
        const parentTag = p.parentNode?.localName || 'body';
        structuralPathHints.push(`${parentTag}/${p.localName}`);
    }

    return {
        captureKey: operation.captureKey,
        operationIndex: typeof operation.index === 'number' ? operation.index : 0,
        kind,
        generatedParagraphIds,
        fingerprints,
        expectedText,
        structuralPathHints,
        stale: false
    };
}

/**
 * Invalidates any captures in the session capture table that referenced removed or replaced paragraphs.
 *
 * @param {Map<string, object>} captureTable
 * @param {Array<Element>} removedNodes
 */
export function invalidateAffectedCaptures(captureTable, removedNodes) {
    if (!captureTable || !(captureTable instanceof Map) || !Array.isArray(removedNodes) || removedNodes.length === 0) return;
    const removedIds = new Set();
    for (const node of removedNodes) {
        if (!node) continue;
        if (node.localName === 'p') {
            const id = getParagraphId(node);
            if (id) removedIds.add(id);
        } else if (typeof node.getElementsByTagNameNS === 'function') {
            for (const p of Array.from(node.getElementsByTagNameNS(NS_W, 'p'))) {
                const id = getParagraphId(p);
                if (id) removedIds.add(id);
            }
        }
    }
    if (removedIds.size === 0) return;
    for (const [, captured] of captureTable) {
        if (captured.stale) continue;
        if (captured.generatedParagraphIds.some(id => id && removedIds.has(id))) {
            captured.stale = true;
        }
    }
}

/**
 * Resolves a target paragraph from an existing capture in the session capture table.
 *
 * @param {Document} xmlDoc
 * @param {import('./document-operation-session.js').DocumentOperationSession} session
 * @param {object} targetDescriptor
 * @param {string} opType
 * @param {object} options
 * @returns {{ paragraph: Element, resolvedBy: 'capture' }}
 */
export function resolveTargetFromCapture(xmlDoc, session, targetDescriptor, opType = 'redline', options = {}) {
    const captureRef = targetDescriptor?.captureRef;
    if (!captureRef) {
        throw createCaptureError('INVALID_OPERATION', 'Operation target is missing captureRef.');
    }

    if (!session || !session.captureTable || !session.captureTable.has(captureRef)) {
        throw createCaptureError('CAPTURE_NOT_FOUND', `Capture "${captureRef}" was not found in the session capture table.`);
    }

    const captured = session.captureTable.get(captureRef);
    if (captured.stale) {
        throw createCaptureError('CAPTURE_STALE', `Capture "${captureRef}" is stale; the captured content was modified or removed.`);
    }

    const metadataIndex = session.getParagraphMetadataIndex();
    const liveParagraphs = [];

    for (let i = 0; i < captured.expectedText.length; i++) {
        const expectedParaId = captured.generatedParagraphIds[i];
        const expected = captured.expectedText[i];
        let foundNode = null;

        if (expectedParaId && metadataIndex.byId.has(expectedParaId)) {
            foundNode = metadataIndex.byId.get(expectedParaId).paragraph;
        }

        if (!foundNode) {
            const candidate = metadataIndex.entries.find(entry =>
                entry.text === expected || entry.fingerprint === captured.fingerprints[i]
            );
            if (candidate) foundNode = candidate.paragraph;
        }

        if (!foundNode || !foundNode.parentNode) {
            captured.stale = true;
            throw createCaptureError('CAPTURE_STALE', `Capture "${captureRef}" is stale; paragraph ${i + 1} was removed from the document.`);
        }

        const currentText = extractCanonicalParagraphText(foundNode);
        if (currentText !== expected) {
            captured.stale = true;
            throw createCaptureError('CAPTURE_STALE', `Capture "${captureRef}" is stale; paragraph ${i + 1} content drifted from "${expected}" to "${currentText}".`);
        }

        liveParagraphs.push(foundNode);
    }

    if (liveParagraphs.length === 0) {
        captured.stale = true;
        throw createCaptureError('CAPTURE_STALE', `Capture "${captureRef}" is stale; no live paragraphs found.`);
    }

    const select = targetDescriptor.select;
    if (select == null || select === '') {
        if (liveParagraphs.length === 1) {
            return { paragraph: liveParagraphs[0], resolvedBy: 'capture' };
        }
        throw createCaptureError(
            'AMBIGUOUS_CAPTURE_SELECTION',
            `Capture "${captureRef}" contains ${liveParagraphs.length} paragraphs; target.select is required to disambiguate.`
        );
    }

    let matched = [];
    const trimmedSelect = String(select).trim();
    if (trimmedSelect === ':first' || trimmedSelect === 'first') {
        matched = [liveParagraphs[0]];
    } else if (trimmedSelect === ':last' || trimmedSelect === 'last') {
        matched = [liveParagraphs[liveParagraphs.length - 1]];
    } else if (/^\d+$/.test(trimmedSelect)) {
        const index1Based = parseInt(trimmedSelect, 10);
        if (index1Based >= 1 && index1Based <= liveParagraphs.length) {
            matched = [liveParagraphs[index1Based - 1]];
        }
    } else {
        const exactMatches = liveParagraphs.filter(p => extractCanonicalParagraphText(p).trim() === trimmedSelect);
        if (exactMatches.length > 0) {
            matched = exactMatches;
        } else {
            matched = liveParagraphs.filter(p => extractCanonicalParagraphText(p).includes(trimmedSelect));
        }
    }

    if (matched.length === 0) {
        throw createCaptureError('TARGET_NOT_FOUND', `Target selection "${select}" was not found in capture "${captureRef}".`);
    }

    if (matched.length > 1) {
        throw createCaptureError(
            'AMBIGUOUS_CAPTURE_SELECTION',
            `Target selection "${select}" matched ${matched.length} paragraphs in capture "${captureRef}".`
        );
    }

    return { paragraph: matched[0], resolvedBy: 'capture' };
}
