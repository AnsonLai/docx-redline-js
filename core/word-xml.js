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

/**
 * Returns true if a document or fragment contains Word tracked-change markup.
 *
 * @param {Document|Element} xmlDoc - Parsed OOXML document or element
 * @returns {boolean}
 */
export function containsTrackedChanges(xmlDoc) {
    const trackedChangeNames = [
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
        'cellIns',
        'cellDel'
    ];

    return trackedChangeNames.some(localName => wordElementsByLocalName(xmlDoc, localName).length > 0);
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
