/**
 * Revision/comment management utilities for OOXML payloads.
 */

import { NS_W } from '../core/types.js';
import { createParser, createSerializer } from '../adapters/xml-adapter.js';
import { getXmlParseError } from '../core/xml-query.js';

function getAttributeByLocalName(node, localName) {
    if (!node || !node.attributes) return '';
    for (const attr of Array.from(node.attributes)) {
        if ((attr.localName || '').toLowerCase() === localName.toLowerCase()) {
            return String(attr.value || '');
        }
    }
    return String(
        node.getAttribute?.(`w:${localName}`)
        || node.getAttribute?.(localName)
        || ''
    );
}

function normalizeAuthor(author) {
    return typeof author === 'string' ? author.trim().toLowerCase() : '';
}

function isElement(node) {
    return !!node && node.nodeType === 1;
}

function isWordElement(node, localName) {
    return isElement(node)
        && node.namespaceURI === NS_W
        && String(node.localName || '').toLowerCase() === localName.toLowerCase();
}

function getWordElementsByLocalName(xmlDoc, localName) {
    return Array.from(xmlDoc.getElementsByTagNameNS(NS_W, localName));
}

function resolveAuthorFilter(options = {}) {
    if (options?.allAuthors === true) {
        return { valid: true, allAuthors: true, normalizedAuthor: '' };
    }
    const normalizedAuthor = normalizeAuthor(options?.author);
    if (!normalizedAuthor) {
        return {
            valid: false,
            allAuthors: false,
            normalizedAuthor: '',
            warning: 'No author provided. Pass { author } or set { allAuthors: true }.'
        };
    }
    return { valid: true, allAuthors: false, normalizedAuthor };
}

function authorMatchesNode(node, filter) {
    if (filter.allAuthors) return true;
    const nodeAuthor = normalizeAuthor(getAttributeByLocalName(node, 'author'));
    return !!nodeAuthor && nodeAuthor === filter.normalizedAuthor;
}

function parseXmlWithWarnings(oxml, parseFailurePrefix) {
    const parser = createParser();
    const xmlDoc = parser.parseFromString(oxml, 'application/xml');
    const parseError = getXmlParseError(xmlDoc);
    if (parseError) {
        return {
            xmlDoc: null,
            serializer: null,
            warning: `${parseFailurePrefix}: ${parseError.textContent || 'parse error'}`
        };
    }
    return { xmlDoc, serializer: createSerializer(), warning: null };
}

function removeNode(node) {
    if (node?.parentNode) {
        node.parentNode.removeChild(node);
        return true;
    }
    return false;
}

function unwrapNode(node) {
    const parent = node?.parentNode;
    if (!parent) return false;
    while (node.firstChild) {
        parent.insertBefore(node.firstChild, node);
    }
    parent.removeChild(node);
    return true;
}

function isTableRowRevisionMarker(node) {
    const parent = node?.parentNode;
    return isWordElement(parent, 'trPr') && isWordElement(parent?.parentNode, 'tr');
}

/**
 * Accepts tracked changes (`w:ins`, `w:del`, and *PrChange tags) for one author
 * or all authors in the provided OOXML payload.
 *
 * @param {string} oxml
 * @param {{ author?: string, allAuthors?: boolean }} [options]
 * @returns {{ oxml: string, hasChanges: boolean, acceptedCount: number, warnings: string[] }}
 */
export function acceptTrackedChangesInOoxml(oxml, options = {}) {
    const warnings = [];
    const filter = resolveAuthorFilter(options);
    if (!filter.valid) {
        return { oxml, hasChanges: false, acceptedCount: 0, warnings: [filter.warning] };
    }

    const parseResult = parseXmlWithWarnings(oxml, 'Failed to parse OOXML');
    if (!parseResult.xmlDoc) {
        return { oxml, hasChanges: false, acceptedCount: 0, warnings: [parseResult.warning] };
    }

    const { xmlDoc, serializer } = parseResult;
    let acceptedCount = 0;

    for (const insNode of getWordElementsByLocalName(xmlDoc, 'ins')) {
        if (!insNode.parentNode || !authorMatchesNode(insNode, filter)) continue;
        if (isTableRowRevisionMarker(insNode)) {
            if (removeNode(insNode)) acceptedCount += 1;
            continue;
        }
        if (unwrapNode(insNode)) acceptedCount += 1;
    }

    for (const delNode of getWordElementsByLocalName(xmlDoc, 'del')) {
        if (!delNode.parentNode || !authorMatchesNode(delNode, filter)) continue;
        if (isTableRowRevisionMarker(delNode)) {
            const rowNode = delNode.parentNode?.parentNode;
            if (removeNode(rowNode)) acceptedCount += 1;
            continue;
        }
        if (removeNode(delNode)) acceptedCount += 1;
    }

    const changeTags = ['rPrChange', 'pPrChange', 'tblPrChange', 'trPrChange', 'tcPrChange'];
    for (const localName of changeTags) {
        for (const changeNode of getWordElementsByLocalName(xmlDoc, localName)) {
            if (!changeNode.parentNode || !authorMatchesNode(changeNode, filter)) continue;
            if (removeNode(changeNode)) acceptedCount += 1;
        }
    }

    return {
        oxml: serializer.serializeToString(xmlDoc),
        hasChanges: acceptedCount > 0,
        acceptedCount,
        warnings
    };
}

function convertDeletionTextNodes(xmlDoc, delNode) {
    for (const delTextNode of Array.from(delNode.getElementsByTagNameNS(NS_W, 'delText'))) {
        const normalText = xmlDoc.createElementNS(NS_W, 'w:t');
        const spaceValue = delTextNode.getAttribute('xml:space');
        if (spaceValue) {
            normalText.setAttribute('xml:space', spaceValue);
        }
        while (delTextNode.firstChild) {
            normalText.appendChild(delTextNode.firstChild);
        }
        delTextNode.parentNode?.replaceChild(normalText, delTextNode);
    }
}

function rejectPropertyChangeNode(changeNode, localName) {
    const parent = changeNode?.parentNode;
    if (!parent) return false;

    const baseLocalName = localName.endsWith('Change')
        ? localName.slice(0, -'Change'.length)
        : '';

    if (
        !baseLocalName
        || String(parent.localName || '').toLowerCase() !== baseLocalName.toLowerCase()
        || parent.namespaceURI !== NS_W
    ) {
        return removeNode(changeNode);
    }

    const historicalNode = Array.from(changeNode.childNodes || []).find(
        child => child.nodeType === 1 && child.namespaceURI === NS_W
            && String(child.localName || '').toLowerCase() === baseLocalName.toLowerCase()
    );

    if (!historicalNode) {
        return removeNode(changeNode);
    }

    const toAppend = Array.from(historicalNode.childNodes || []);
    while (parent.firstChild) {
        parent.removeChild(parent.firstChild);
    }
    for (const node of toAppend) {
        const clone = xmlDocImportNode(parent.ownerDocument, node);
        parent.appendChild(clone);
    }
    return true;
}

function xmlDocImportNode(xmlDoc, node) {
    if (xmlDoc && typeof xmlDoc.importNode === 'function') {
        return xmlDoc.importNode(node, true);
    }
    return node.cloneNode(true);
}

/**
 * Rejects tracked changes (`w:ins`, `w:del`, and *PrChange tags) for one author
 * or all authors in the provided OOXML payload.
 *
 * @param {string} oxml
 * @param {{ author?: string, allAuthors?: boolean }} [options]
 * @returns {{ oxml: string, hasChanges: boolean, rejectedCount: number, warnings: string[] }}
 */
export function rejectTrackedChangesInOoxml(oxml, options = {}) {
    const warnings = [];
    const filter = resolveAuthorFilter(options);
    if (!filter.valid) {
        return { oxml, hasChanges: false, rejectedCount: 0, warnings: [filter.warning] };
    }

    const parseResult = parseXmlWithWarnings(oxml, 'Failed to parse OOXML');
    if (!parseResult.xmlDoc) {
        return { oxml, hasChanges: false, rejectedCount: 0, warnings: [parseResult.warning] };
    }

    const { xmlDoc, serializer } = parseResult;
    let rejectedCount = 0;

    for (const insNode of getWordElementsByLocalName(xmlDoc, 'ins')) {
        if (!insNode.parentNode || !authorMatchesNode(insNode, filter)) continue;
        if (isTableRowRevisionMarker(insNode)) {
            const rowNode = insNode.parentNode?.parentNode;
            if (removeNode(rowNode)) rejectedCount += 1;
            continue;
        }
        if (removeNode(insNode)) rejectedCount += 1;
    }

    for (const delNode of getWordElementsByLocalName(xmlDoc, 'del')) {
        if (!delNode.parentNode || !authorMatchesNode(delNode, filter)) continue;
        if (isTableRowRevisionMarker(delNode)) {
            if (removeNode(delNode)) rejectedCount += 1;
            continue;
        }
        convertDeletionTextNodes(xmlDoc, delNode);
        if (unwrapNode(delNode)) rejectedCount += 1;
    }

    const changeTags = ['rPrChange', 'pPrChange', 'tblPrChange', 'trPrChange', 'tcPrChange'];
    for (const localName of changeTags) {
        for (const changeNode of getWordElementsByLocalName(xmlDoc, localName)) {
            if (!changeNode.parentNode || !authorMatchesNode(changeNode, filter)) continue;
            if (rejectPropertyChangeNode(changeNode, localName)) rejectedCount += 1;
        }
    }

    return {
        oxml: serializer.serializeToString(xmlDoc),
        hasChanges: rejectedCount > 0,
        rejectedCount,
        warnings
    };
}

function collectCommentTargetIds(xmlDoc, filter) {
    const targetIds = new Set();
    const commentNodes = getWordElementsByLocalName(xmlDoc, 'comment');

    for (const commentNode of commentNodes) {
        if (!authorMatchesNode(commentNode, filter)) continue;
        const id = getAttributeByLocalName(commentNode, 'id');
        if (id) targetIds.add(id);
    }

    return { targetIds, commentNodes };
}

function removeCommentNodesById(commentNodes, targetIds) {
    let removed = 0;
    for (const commentNode of commentNodes) {
        const id = getAttributeByLocalName(commentNode, 'id');
        if (!id || !targetIds.has(id)) continue;
        if (removeNode(commentNode)) removed += 1;
    }
    return removed;
}

function runIsOnlyCommentReference(runNode) {
    if (!isWordElement(runNode, 'r')) return false;
    const meaningfulChildren = Array.from(runNode.childNodes || []).filter(child => {
        if (child.nodeType === 3) return String(child.nodeValue || '').trim().length > 0;
        if (child.nodeType !== 1) return false;
        if (child.namespaceURI !== NS_W) return true;
        const local = String(child.localName || '').toLowerCase();
        return local !== 'rpr' && local !== 'commentreference';
    });
    return meaningfulChildren.length === 0;
}

function removeCommentAnchors(xmlDoc, targetIds) {
    let removed = 0;
    const anchorTags = ['commentRangeStart', 'commentRangeEnd', 'commentReference'];

    for (const localName of anchorTags) {
        for (const node of getWordElementsByLocalName(xmlDoc, localName)) {
            if (!node.parentNode) continue;
            const id = getAttributeByLocalName(node, 'id');
            if (!id || !targetIds.has(id)) continue;

            if (localName === 'commentReference' && runIsOnlyCommentReference(node.parentNode)) {
                if (removeNode(node.parentNode)) {
                    removed += 1;
                }
                continue;
            }
            if (removeNode(node)) {
                removed += 1;
            }
        }
    }
    return removed;
}

/**
 * Deletes comments authored by one user (or all users) and removes matching
 * comment anchors/references from the OOXML payload.
 *
 * @param {string} oxml
 * @param {{ author?: string, allAuthors?: boolean }} [options]
 * @returns {{ oxml: string, hasChanges: boolean, commentsRemoved: number, referencesRemoved: number, warnings: string[] }}
 */
export function deleteCommentsByAuthorInOoxml(oxml, options = {}) {
    const warnings = [];
    const filter = resolveAuthorFilter(options);
    if (!filter.valid) {
        return {
            oxml,
            hasChanges: false,
            commentsRemoved: 0,
            referencesRemoved: 0,
            warnings: [filter.warning]
        };
    }

    const parseResult = parseXmlWithWarnings(oxml, 'Failed to parse OOXML');
    if (!parseResult.xmlDoc) {
        return {
            oxml,
            hasChanges: false,
            commentsRemoved: 0,
            referencesRemoved: 0,
            warnings: [parseResult.warning]
        };
    }

    const { xmlDoc, serializer } = parseResult;
    const { targetIds, commentNodes } = collectCommentTargetIds(xmlDoc, filter);

    if (filter.allAuthors) {
        for (const localName of ['commentRangeStart', 'commentRangeEnd', 'commentReference']) {
            for (const node of getWordElementsByLocalName(xmlDoc, localName)) {
                const id = getAttributeByLocalName(node, 'id');
                if (id) targetIds.add(id);
            }
        }
    }

    const commentsRemoved = removeCommentNodesById(commentNodes, targetIds);
    const referencesRemoved = removeCommentAnchors(xmlDoc, targetIds);

    return {
        oxml: serializer.serializeToString(xmlDoc),
        hasChanges: commentsRemoved > 0 || referencesRemoved > 0,
        commentsRemoved,
        referencesRemoved,
        warnings
    };
}
