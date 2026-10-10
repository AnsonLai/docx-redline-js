/**
 * Single accessor for CURRENT paragraph/run properties.
 *
 * Descendant lookups (`getElementsByTagNameNS`, `getWordDescendants`, ...) cross `w:pPrChange` and
 * `w:rPrChange`, so a historical `numPr`/`pStyle`/`b` snapshot looks like a current value. Everything in
 * this module reads direct children of `w:pPr` / `w:rPr` only. The `*PrChange` snapshot is reachable
 * only through the explicit `getHistorical*` functions (reject simulation, restore).
 *
 * Depends only on `core/types.js`.
 */

import { NS_W } from './types.js';

function isWordElement(node, localName) {
    if (!node || node.nodeType !== 1) return false;
    const name = String(node.localName || node.nodeName || '').replace(/^.*:/, '');
    if (localName != null && name !== localName) return false;
    return !node.namespaceURI || node.namespaceURI === NS_W;
}

/** First direct child element with the given WordprocessingML local name. */
export function getDirectWordChild(node, localName) {
    if (!node) return null;
    for (const child of Array.from(node.childNodes || [])) {
        if (isWordElement(child, localName)) return child;
    }
    return null;
}

/** Direct child elements with the given local name (all direct Word elements when omitted). */
export function getDirectWordChildren(node, localName) {
    if (!node) return [];
    return Array.from(node.childNodes || []).filter(child => isWordElement(child, localName));
}

/** Direct child of a properties element; never descends into `*PrChange` snapshots. */
export function getCurrentChild(propsEl, localName) {
    return getDirectWordChild(propsEl, localName);
}

function readVal(el) {
    if (!el) return null;
    const value = (typeof el.getAttributeNS === 'function' && el.getAttributeNS(NS_W, 'val'))
        || el.getAttribute?.('w:val')
        || el.getAttribute?.('val')
        || null;
    return value === '' ? null : value;
}

function readNumPr(propsEl) {
    const numPr = getDirectWordChild(propsEl, 'numPr');
    if (!numPr) return null;
    return {
        element: numPr,
        numId: readVal(getDirectWordChild(numPr, 'numId')),
        ilvl: readVal(getDirectWordChild(numPr, 'ilvl')),
        numIdElement: getDirectWordChild(numPr, 'numId'),
        ilvlElement: getDirectWordChild(numPr, 'ilvl')
    };
}

function readParagraphProps(pPr) {
    const pStyleEl = getDirectWordChild(pPr, 'pStyle');
    const outlineEl = getDirectWordChild(pPr, 'outlineLvl');
    return {
        pPr: pPr || null,
        numPr: readNumPr(pPr),
        pStyle: readVal(pStyleEl),
        pStyleElement: pStyleEl,
        outlineLvl: readVal(outlineEl),
        outlineLvlElement: outlineEl
    };
}

/** Current paragraph properties: the paragraph's direct `w:pPr` only. */
export function getCurrentParagraphProperties(p) {
    return readParagraphProps(getDirectWordChild(p, 'pPr'));
}

/** Current run properties: the run's direct `w:rPr` (null when absent). */
export function getCurrentRunProperties(r) {
    return getDirectWordChild(r, 'rPr');
}

/** Paragraph-mark run properties (`w:pPr/w:rPr`), distinct from any run's `rPr`. */
export function getParagraphMarkProperties(p) {
    return getDirectWordChild(getDirectWordChild(p, 'pPr'), 'rPr');
}

/**
 * Historical paragraph properties: the `w:pPrChange/w:pPr` snapshot. Explicit opt-in; shape matches
 * `getCurrentParagraphProperties`. Returns null when the paragraph has no tracked property change.
 */
export function getHistoricalParagraphProperties(p) {
    const change = getDirectWordChild(getDirectWordChild(p, 'pPr'), 'pPrChange');
    const snapshot = getDirectWordChild(change, 'pPr');
    return snapshot ? readParagraphProps(snapshot) : null;
}

/** Historical run properties: the `w:rPrChange/w:rPr` snapshot, or null. */
export function getHistoricalRunProperties(r) {
    const change = getDirectWordChild(getDirectWordChild(r, 'rPr'), 'rPrChange');
    return getDirectWordChild(change, 'rPr');
}

/** All `w:p` elements at or under `root` (the paragraph itself when `root` is one). */
function paragraphsUnder(root) {
    if (!root || root.nodeType !== 1) return [];
    const out = isWordElement(root, 'p') ? [root] : [];
    if (typeof root.getElementsByTagNameNS === 'function') {
        out.push(...Array.from(root.getElementsByTagNameNS(NS_W, 'p')));
    }
    return out;
}

/**
 * Current `w:numId` elements (direct `pPr/numPr/numId`) of every paragraph at or under `root`.
 * Excludes `pPrChange` snapshots, so Reject still restores the snapshot numId.
 */
export function getCurrentNumIdElements(root) {
    const out = [];
    for (const p of paragraphsUnder(root)) {
        const el = getCurrentParagraphProperties(p).numPr?.numIdElement;
        if (el) out.push(el);
    }
    return out;
}

/**
 * Every `w:numId` element referenced anywhere under `root`, INCLUDING `pPrChange` snapshots. Explicitly
 * historical-inclusive: used when a numbering part must define every id the content can ever reference.
 */
export function getAllReferencedNumIdElements(root) {
    if (!root || typeof root.getElementsByTagNameNS !== 'function') return [];
    return Array.from(root.getElementsByTagNameNS(NS_W, 'numId'));
}
