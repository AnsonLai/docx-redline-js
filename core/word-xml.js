import { NS_W } from './types.js';

/**
 * Returns true when a node is a WordprocessingML element with the given local name.
 *
 * @param {Node|null|undefined} node - Candidate node
 * @param {string} localName - Word local name, for example `r` or `tbl`
 * @returns {boolean}
 */
export function isWordElement(node, localName) {
    if (!node || node.nodeType !== 1) return false;
    if (node.namespaceURI === NS_W && node.localName === localName) return true;
    const nodeName = String(node.nodeName || '');
    return nodeName === `w:${localName}` || nodeName === localName;
}

/**
 * Creates a WordprocessingML element using namespace-aware DOM APIs when available.
 *
 * @param {Document} xmlDoc - Target document
 * @param {string} qualifiedName - Qualified name, for example `w:r`
 * @returns {Element}
 */
export function createWordElement(xmlDoc, qualifiedName) {
    return typeof xmlDoc.createElementNS === 'function'
        ? xmlDoc.createElementNS(NS_W, qualifiedName)
        : xmlDoc.createElement(qualifiedName);
}

function wordElementsByLocalName(xmlDoc, localName) {
    const namespaced = Array.from(xmlDoc?.getElementsByTagNameNS?.(NS_W, localName) || []);
    if (namespaced.length > 0) return namespaced;
    return Array.from(xmlDoc?.getElementsByTagName?.('*') || []).filter(node => isWordElement(node, localName));
}

// Keep revision discovery and author discovery on the same taxonomy. Cell
// markers are included for fail-closed detection even though selective
// acceptance/rejection of those structural revisions is not yet supported.
const TRACKED_CHANGE_NAMES = [
    'ins',
    'del',
    'moveFrom',
    'moveTo',
    'moveFromRangeStart',
    'moveFromRangeEnd',
    'moveToRangeStart',
    'moveToRangeEnd',
    'rPrChange',
    'pPrChange',
    'tblPrChange',
    'trPrChange',
    'tcPrChange',
    'cellIns',
    'cellDel'
];

/**
 * Returns true if a document or fragment contains Word tracked-change markup.
 *
 * @param {Document|Element} xmlDoc - Parsed OOXML document or element
 * @returns {boolean}
 */
export function containsTrackedChanges(xmlDoc) {
    return TRACKED_CHANGE_NAMES.some(localName => wordElementsByLocalName(xmlDoc, localName).length > 0);
}

/**
 * Collects all distinct author names from tracked changes in the given XML document or element.
 *
 * @param {Document|Element|null|undefined} xmlDocOrElement
 * @returns {string[]} Sorted unique list of author names.
 */
export function getTrackedChangeAuthors(xmlDocOrElement) {
    if (!xmlDocOrElement) return [];
    const authors = new Set();
    for (const localName of TRACKED_CHANGE_NAMES) {
        for (const node of wordElementsByLocalName(xmlDocOrElement, localName)) {
            const author = node.getAttribute?.('w:author')
                || node.getAttribute?.('author')
                || (typeof node.getAttributeNS === 'function' ? node.getAttributeNS(NS_W, 'author') : null);
            if (author && typeof author === 'string' && author.trim()) {
                authors.add(author.trim());
            }
        }
    }
    return [...authors].sort();
}

/**
 * Classifies the shape of an OOXML payload.
 *
 * @param {string} oxml - OOXML payload
 * @returns {'package'|'document'|'fragment'}
 */
export function classifyOoxmlSourceType(oxml) {
    const trimmed = String(oxml || '').trim();
    if (/^<\?xml\b[^>]*>\s*<pkg:package\b/i.test(trimmed) || /^<pkg:package\b/i.test(trimmed)) {
        return 'package';
    }
    if (/^<\?xml\b[^>]*>\s*<(?:w:)?document\b/i.test(trimmed) || /^<(?:w:)?document\b/i.test(trimmed)) {
        return 'document';
    }
    return 'fragment';
}

/**
 * Adds `sourceType` metadata to OOXML result objects without changing payloads.
 *
 * @template T
 * @param {T & { oxml?: string, sourceType?: 'package'|'document'|'fragment' }} result - Result object
 * @returns {T & { sourceType?: 'package'|'document'|'fragment' }}
 */
export function withOoxmlSourceType(result) {
    if (!result || typeof result !== 'object' || result.sourceType || typeof result.oxml !== 'string') {
        return result;
    }
    return { ...result, sourceType: classifyOoxmlSourceType(result.oxml) };
}
