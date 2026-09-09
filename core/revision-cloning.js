import {
    NS_W,
    RevisionIdAllocator,
    createRevisionIdAllocator,
    getRevisionIdAllocatorForDocument
} from './types.js';
import { isWordElement } from './word-xml.js';

const REVISION_HISTORY_ELEMENTS = new Set([
    'ins', 'del', 'moveFrom', 'moveTo', 'pPrChange', 'rPrChange',
    'tblPrChange', 'trPrChange', 'tcPrChange', 'sectPrChange'
]);

/**
 * Clones effective paragraph/run properties without copying tracked history
 * into a newly created paragraph or run.
 */
export function clonePropertiesWithoutRevisionHistory(root) {
    if (!root) return null;
    const clone = root.cloneNode(true);
    const candidates = [clone, ...Array.from(clone.getElementsByTagName?.('*') || [])];
    for (const node of candidates.reverse()) {
        if (!REVISION_HISTORY_ELEMENTS.has(String(node.localName || node.nodeName || '').replace(/^.*:/, ''))) continue;
        node.parentNode?.removeChild(node);
    }
    return clone;
}

/**
 * Assigns fresh document-scoped IDs to w:rPrChange elements in a cloned
 * run-properties subtree. This preserves formatting-revision metadata while
 * preventing a DOM split from duplicating the original revision ID.
 *
 * @param {Element} root - Cloned subtree whose revision IDs should be refreshed
 * @param {RevisionIdAllocator|null} [allocator=null] - Document-scoped allocator
 * @returns {Element}
 */
export function refreshRunPropertyChangeIds(root, allocator = null) {
    if (!root) return root;

    const xmlDoc = root.nodeType === 9 ? root : root.ownerDocument;
    const resolvedAllocator = allocator instanceof RevisionIdAllocator
        ? allocator
        : (getRevisionIdAllocatorForDocument(xmlDoc) || createRevisionIdAllocator(xmlDoc));
    const candidates = [root, ...Array.from(root.getElementsByTagName?.('*') || [])];

    for (const node of candidates) {
        if (!isWordElement(node, 'rPrChange')) continue;
        const nextId = String(resolvedAllocator.next());
        if (typeof node.setAttributeNS === 'function') {
            node.setAttributeNS(NS_W, 'w:id', nextId);
        } else {
            node.setAttribute('w:id', nextId);
        }
        resolvedAllocator._receiptCollector?.recordRevision(Number(nextId), 'rPrChange');
    }

    return root;
}
