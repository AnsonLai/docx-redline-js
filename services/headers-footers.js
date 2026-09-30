/**
 * Header and footer parts: discovery, targeting and safety checks.
 *
 * Word stores each header/footer as its own part (word/header1.xml ...), referenced from a section's
 * w:sectPr through w:headerReference / w:footerReference (type default | first | even) and the document
 * relationships. Behavior here is grounded in tests/fixtures/word-authored/header-footer.docx, saved by Word.
 */

import { createSerializer, parseOoxmlSafe } from '../adapters/xml-adapter.js';
import { findPartSpecByPartName } from './package-parts.js';

const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

const REVISION_ELEMENTS = new Set(['ins', 'del', 'moveFrom', 'moveTo', 'rPrChange', 'pPrChange', 'sectPrChange', 'tblPrChange', 'tblGridChange', 'tcPrChange', 'trPrChange', 'numberingChange']);
const REVISION_CONTENT_ELEMENTS = new Set(['ins', 'del', 'moveFrom', 'moveTo']);

const attrOf = (node, qualified, local) => node?.getAttribute?.(qualified) || node?.getAttribute?.(local) || '';

/** Resolves a relationship Target (relative to word/) to a package entry name such as word/header1.xml. */
function resolveTarget(target) {
    const clean = String(target || '');
    if (clean.startsWith('/')) return clean.replace(/^\/+/, '');
    const parts = ['word'];
    for (const segment of clean.split('/')) {
        if (segment === '..') parts.pop();
        else if (segment && segment !== '.') parts.push(segment);
    }
    return parts.join('/');
}

/**
 * Lists the header/footer parts a document uses.
 *
 * @param {{ documentXml: string, relsXml?: string|null, settingsXml?: string|null }} parts
 * `sections` lists sections that reference the part explicitly. A section with no reference of some type inherits
 * the previous section's part (Word omits the reference to mean "linked to previous"), so `appliesToSections` also
 * includes inherited sections.
 *
 * @returns {Array<{ path: string, kind: 'header'|'footer', type: 'default'|'first'|'even', sections: number[],
 *   appliesToSections: number[], sharedBySections: boolean, active: boolean }>}
 */
export function discoverHeaderFooterParts({ documentXml, relsXml = null, settingsXml = null }) {
    if (!documentXml || !relsXml) return [];
    const documentParsed = parseOoxmlSafe(documentXml, 'application/xml');
    const relsParsed = parseOoxmlSafe(relsXml, 'application/xml');
    if (!documentParsed.doc || documentParsed.error || !relsParsed.doc || relsParsed.error) return [];

    const targetByRelId = new Map();
    for (const rel of Array.from(relsParsed.doc.getElementsByTagNameNS('*', 'Relationship'))) {
        targetByRelId.set(rel.getAttribute('Id') || '', rel.getAttribute('Target') || '');
    }

    const evenAndOdd = !!settingsXml && /<w:evenAndOddHeaders\b(?![^>]*w:val="(?:0|false|off)")/.test(settingsXml);
    const byPath = new Map();
    const inForce = new Map(); // `${kind}:${type}` -> path currently in force
    const sectPrs = Array.from(documentParsed.doc.getElementsByTagNameNS(NS_W, 'sectPr'));
    sectPrs.forEach((sectPr, sectionIndex) => {
        const titlePg = Array.from(sectPr.childNodes || []).some(child => child.localName === 'titlePg' && !/^(0|false|off)$/i.test(attrOf(child, 'w:val', 'val')));
        for (const ref of Array.from(sectPr.childNodes || [])) {
            if (ref.nodeType !== 1 || (ref.localName !== 'headerReference' && ref.localName !== 'footerReference')) continue;
            const relId = ref.getAttributeNS?.(NS_R, 'id') || ref.getAttribute('r:id') || '';
            const target = targetByRelId.get(relId);
            if (!target) continue;
            const path = resolveTarget(target);
            const kind = ref.localName === 'headerReference' ? 'header' : 'footer';
            const type = attrOf(ref, 'w:type', 'type') || 'default';
            const entry = byPath.get(path) || { path, kind, type, sections: [], applies: [], titlePgSections: 0 };
            if (!entry.sections.includes(sectionIndex)) entry.sections.push(sectionIndex);
            if (titlePg) entry.titlePgSections += 1;
            byPath.set(path, entry);
            inForce.set(`${kind}:${type}`, path);
        }
        for (const path of inForce.values()) {
            const entry = byPath.get(path);
            if (entry && !entry.applies.includes(sectionIndex)) entry.applies.push(sectionIndex);
        }
    });

    return Array.from(byPath.values()).map(entry => ({
        path: entry.path,
        kind: entry.kind,
        type: entry.type,
        sections: entry.sections,
        appliesToSections: entry.applies,
        sharedBySections: entry.applies.length > 1,
        // first-page parts only show with w:titlePg; even parts only with w:evenAndOddHeaders in settings.xml.
        active: entry.type === 'first' ? entry.titlePgSections > 0 : entry.type === 'even' ? evenAndOdd : true
    }));
}

/**
 * Resolves an operation's `part` selector to one discovered part.
 *
 * @param {ReturnType<typeof discoverHeaderFooterParts>} parts
 * @param {string | { kind: 'header'|'footer', type?: string, section?: number }} selector
 * @returns {{ part: object } | { error: { code: string, message: string, candidates?: object[] } }}
 */
export function resolvePartSelector(parts, selector) {
    const describe = part => ({ path: part.path, kind: part.kind, type: part.type, sections: part.sections, appliesToSections: part.appliesToSections });
    if (typeof selector === 'string') {
        const wanted = selector.replace(/^\/+/, '').toLowerCase();
        const match = parts.find(part => part.path.toLowerCase() === wanted || part.path.toLowerCase() === `word/${wanted}`);
        if (match) return { part: match };
        return { error: { code: 'PART_NOT_FOUND', message: `No header/footer part '${selector}' is referenced by the document.`, candidates: parts.map(describe) } };
    }
    if (!selector || typeof selector !== 'object' || (selector.kind !== 'header' && selector.kind !== 'footer')) {
        return { error: { code: 'INVALID_OPERATION', message: 'part must be a part path or { kind: "header"|"footer", type?: "default"|"first"|"even", section?: number }.' } };
    }
    const type = selector.type || 'default';
    if (!['default', 'first', 'even'].includes(type)) {
        return { error: { code: 'INVALID_OPERATION', message: `part.type must be default, first or even (got '${type}').` } };
    }
    let matches = parts.filter(part => part.kind === selector.kind && part.type === type);
    if (Number.isInteger(selector.section)) matches = matches.filter(part => part.appliesToSections.includes(selector.section));
    if (matches.length === 1) return { part: matches[0] };
    if (matches.length === 0) {
        return { error: { code: 'PART_NOT_FOUND', message: `The document has no ${type} ${selector.kind}${Number.isInteger(selector.section) ? ` for section ${selector.section}` : ''}.`, candidates: parts.map(describe) } };
    }
    return { error: { code: 'PART_AMBIGUOUS', message: `Several ${type} ${selector.kind} parts exist; add "section" or use the part path.`, candidates: matches.map(describe) } };
}

/** Whether a package entry name is a header or footer part. */
export function isHeaderFooterPath(path) {
    const spec = findPartSpecByPartName(path);
    return spec?.id === 'header' || spec?.id === 'footer';
}

/**
 * Counts tracked-change elements that sit inside a field's instruction or cached result
 * (between w:fldChar begin and end, or inside w:fldSimple). A field is one atomic unit: PAGE, NUMPAGES, DATE and
 * friends are recomputed by Word, so a redline inside one is meaningless and can leave stray text after accept.
 *
 * @param {string} xml part XML
 * @returns {number}
 */
export function countRevisionsInsideFields(xml) {
    const parsed = parseOoxmlSafe(xml, 'application/xml');
    if (!parsed.doc || parsed.error) return 0;
    let depth = 0;
    let count = 0;
    const visit = node => {
        for (const child of Array.from(node.childNodes || [])) {
            if (child.nodeType !== 1) continue;
            const name = child.localName;
            if (name === 'fldChar') {
                const type = attrOf(child, 'w:fldCharType', 'fldCharType');
                if (type === 'begin') depth += 1;
                else if (type === 'end') depth = Math.max(0, depth - 1);
                continue;
            }
            if (REVISION_CONTENT_ELEMENTS.has(name) && depth > 0) count += 1;
            if (name === 'fldSimple') {
                depth += 1;
                visit(child);
                depth -= 1;
                continue;
            }
            visit(child);
        }
    };
    visit(parsed.doc.documentElement);
    return count;
}

/** All w:id values used by revision elements in an XML string. */
export function collectRevisionIds(xml) {
    const ids = new Set();
    for (const match of String(xml || '').matchAll(/<w:(?:ins|del|moveFrom|moveTo|rPrChange|pPrChange|sectPrChange|tblPrChange|tblGridChange|tcPrChange|trPrChange|numberingChange)\b[^>]*?\bw:id="(\d+)"/g)) ids.add(match[1]);
    return ids;
}

/**
 * Renumbers revision elements in `xml` whose w:id collides with an id in `taken`, so ids stay unique across
 * document.xml and header/footer parts. Ids that do not collide are left alone.
 *
 * @returns {{ xml: string, renumbered: number, idMap: Map<string, string> }} idMap maps each old id to its new id
 */
export function renumberCollidingRevisionIds(xml, taken) {
    const parsed = parseOoxmlSafe(xml, 'application/xml');
    if (!parsed.doc || parsed.error) return { xml, renumbered: 0, idMap: new Map() };
    const idMap = new Map();
    let next = 1 + Math.max(0, ...Array.from(taken, id => Number(id)).filter(Number.isFinite));
    const used = new Set(taken);
    let renumbered = 0;
    const visit = node => {
        for (const child of Array.from(node.childNodes || [])) {
            if (child.nodeType !== 1) continue;
            if (REVISION_ELEMENTS.has(child.localName)) {
                const id = attrOf(child, 'w:id', 'id');
                if (id && used.has(id)) {
                    while (used.has(String(next))) next += 1;
                    child.setAttributeNS(NS_W, 'w:id', String(next));
                    idMap.set(id, String(next));
                    used.add(String(next));
                    renumbered += 1;
                } else if (id) {
                    used.add(id);
                }
            }
            visit(child);
        }
    };
    visit(parsed.doc.documentElement);
    return renumbered > 0 ? { xml: createSerializer().serializeToString(parsed.doc), renumbered, idMap } : { xml, renumbered: 0, idMap };
}
