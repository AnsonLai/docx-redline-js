import {
    NS_W,
    RevisionIdAllocator,
    createRevisionIdAllocator,
    getRevisionIdAllocatorForDocument
} from './types.js';
import { isWordElement } from './word-xml.js';

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
    }

    return root;
}
