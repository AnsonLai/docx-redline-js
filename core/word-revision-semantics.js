import { NS_W } from './types.js';
import { readCanonicalRunText } from './paragraph-text.js';

/**
 * Word-semantics Reject All / Accept All projection over a WordprocessingML DOM.
 *
 * The projection is a list of paragraph records describing what Word shows after
 * rejecting or accepting every tracked revision. It never mutates the input and
 * never serializes. Rules (Word behavior):
 *   - a paragraph mark in w:pPr/w:rPr/w:ins (or moveTo) is removed on Reject; the
 *     paragraph's remaining content merges into the FOLLOWING paragraph of the same
 *     container, whose mark and properties survive. A deleted mark (w:del/moveFrom)
 *     is removed on Accept with the same merge rule.
 *   - at the end of a container: Reject keeps a table cell's last paragraph; otherwise a
 *     trailing paragraph left empty by the removed mark disappears.
 *   - w:ins/moveTo content is removed on Reject, w:del/moveFrom restored; Accept is
 *     the converse. Rows whose w:trPr carries w:ins (Reject) or w:del (Accept) vanish.
 *   - Reject restores the w:pPrChange snapshot as the paragraph's properties.
 *   - a w:body whose last block is a table gets one untracked empty paragraph from Word on
 *     open (a body must end with a paragraph). It is not a revision, so it survives Reject and
 *     Accept alike: such a body counts one extra paragraph in both views even when Reject
 *     removes the whole table (Word 16 COM evidence, R5).
 *
 * Shared by the post-apply fidelity verifier and the reject/accept tests.
 */

const CONTAINER_NAMES = new Set(['body', 'tc', 'txbxContent', 'footnote', 'endnote', 'comment', 'hdr', 'ftr']);
const WRAPPER_NAMES = new Set(['hyperlink', 'fldSimple', 'sdt', 'smartTag', 'customXml']);
const PROPERTY_SKIP = new Set(['rPr', 'pPrChange', 'sectPr']);

const localName = node => {
    const local = node?.localName;
    if (local) return local.indexOf(':') < 0 ? local : local.replace(/^.*:/, '');
    return String(node?.nodeName || '').replace(/^.*:/, '');
};
const isWordElement = node => node?.nodeType === 1 && (!node.namespaceURI || node.namespaceURI === NS_W);
const attr = (node, name) => node?.getAttribute?.(`w:${name}`) || node?.getAttribute?.(name) || '';

/** Document-order w:<name> descendants of `root`; a plain walk avoids the live node list's per-call cost. */
function descendantWordElements(root, name) {
    const found = [];
    if (!root) return found;
    const stack = [];
    for (let child = root.lastChild; child; child = child.previousSibling) stack.push(child);
    while (stack.length) {
        const node = stack.pop();
        if (node.nodeType !== 1) continue;
        if (node.namespaceURI === NS_W && node.localName === name) found.push(node);
        for (let child = node.lastChild; child; child = child.previousSibling) stack.push(child);
    }
    return found;
}

function directChild(node, name) {
    for (let child = node?.firstChild; child; child = child.nextSibling) {
        if (isWordElement(child) && localName(child) === name) return child;
    }
    return null;
}

/** Revision kind on the paragraph mark: 'ins' | 'del' | null (moveTo/moveFrom map to ins/del). */
export function paragraphMarkRevision(paragraph) {
    const rPr = directChild(directChild(paragraph, 'pPr'), 'rPr');
    if (!rPr) return null;
    for (let child = rPr.firstChild; child; child = child.nextSibling) {
        if (!isWordElement(child)) continue;
        const name = localName(child);
        if (name === 'ins' || name === 'moveTo') return 'ins';
        if (name === 'del' || name === 'moveFrom') return 'del';
    }
    return null;
}

function containerOf(paragraph) {
    for (let cur = paragraph.parentNode; cur; cur = cur.parentNode) {
        if (CONTAINER_NAMES.has(localName(cur))) return cur;
    }
    return null;
}

function hiddenByStructure(paragraph, view) {
    const revisionName = view === 'rejected' ? 'ins' : 'del';
    for (let cur = paragraph.parentNode; cur; cur = cur.parentNode) {
        const name = localName(cur);
        if (name === 'tr') {
            if (directChild(directChild(cur, 'trPr'), revisionName)) return true;
        } else if (name === 'tc') {
            const cellRevision = view === 'rejected' ? 'cellIns' : 'cellDel';
            if (directChild(directChild(cur, 'tcPr'), cellRevision)) return true;
        } else if (name === 'body') {
            break;
        }
    }
    return false;
}

function canonicalElement(node) {
    const attrs = Array.from(node.attributes || [])
        .filter(a => !/^xmlns(:|$)/.test(a.name) && !/(^|:)(rsid\w*|paraId|textId)$/.test(a.name))
        .map(a => `${a.name.replace(/^.*:/, '')}=${a.value}`)
        .sort();
    const children = Array.from(node.childNodes || [])
        .filter(isWordElement)
        .map(canonicalElement)
        .sort();
    return `${localName(node)}[${attrs.join(',')}]{${children.join('')}}`;
}

/** Properties a paragraph carries in a view (Reject restores the pPrChange snapshot). */
function viewProperties(paragraph, view) {
    const pPr = directChild(paragraph, 'pPr');
    let source = pPr;
    if (view === 'rejected') {
        const change = directChild(pPr, 'pPrChange');
        if (change) source = directChild(change, 'pPr');
    }
    let numPr = null;
    const children = [];
    for (const child of Array.from(source?.childNodes || [])) {
        if (!isWordElement(child) || PROPERTY_SKIP.has(localName(child))) continue;
        if (localName(child) === 'numPr') {
            numPr = {
                numId: attr(directChild(child, 'numId'), 'val') || null,
                ilvl: attr(directChild(child, 'ilvl'), 'val') || '0'
            };
            continue;
        }
        children.push(canonicalElement(child));
    }
    return { props: children.sort().join(''), numPr };
}

function wrapperKey(node) {
    const name = localName(node);
    if (name === 'hyperlink') {
        return `hyperlink:${node.getAttribute?.('r:id') || attr(node, 'anchor') || ''}`;
    }
    if (name === 'fldSimple') return `fldSimple:${attr(node, 'instr').trim()}`;
    if (name === 'sdt') {
        const sdtPr = directChild(node, 'sdtPr');
        return `sdt:${attr(directChild(sdtPr, 'id'), 'val')}:${attr(directChild(sdtPr, 'tag'), 'val')}`;
    }
    return `${name}:${attr(node, 'element') || attr(node, 'uri')}`;
}

/** Visible text and inline wrapper spans of one paragraph in a view. */
function paragraphContent(paragraph, view) {
    let text = '';
    const spans = [];
    const walk = (node, wrappers) => {
        for (let child = node.firstChild; child; child = child.nextSibling) {
            if (!isWordElement(child)) continue;
            const name = localName(child);
            if (name === 'pPr') continue;
            if (view === 'rejected' && (name === 'ins' || name === 'moveTo')) continue;
            if (view === 'accepted' && (name === 'del' || name === 'moveFrom')) continue;
            if (name === 'r') {
                const start = text.length;
                text += readCanonicalRunText(child, { revisionView: view, boundary: paragraph });
                if (text.length > start) {
                    for (const key of wrappers) spans.push({ key, start, end: text.length });
                }
                continue;
            }
            walk(child, WRAPPER_NAMES.has(name) ? [...wrappers, wrapperKey(child)] : wrappers);
        }
    };
    walk(paragraph, []);
    return { text, spans };
}

function mergeSpans(spans) {
    const byKey = new Map();
    for (const span of spans) {
        const list = byKey.get(span.key) || [];
        const last = list[list.length - 1];
        if (last && last.end >= span.start) last.end = Math.max(last.end, span.end);
        else list.push({ ...span });
        byKey.set(span.key, list);
    }
    return [...byKey.values()].flat()
        .sort((a, b) => a.start - b.start || a.end - b.end || (a.key < b.key ? -1 : 1))
        .map(span => `${span.key}@${span.start}-${span.end}`);
}

/**
 * Projects a document (or any element containing paragraphs) into the paragraph list Word
 * shows after Reject All (`view: 'rejected'`) or Accept All (`view: 'accepted'`).
 *
 * @returns {Array<{ text: string, props: string, numPr: {numId: string|null, ilvl: string}|null,
 *   wrappers: string[], container: string, sourceIndex: number }>}
 */
export function projectWordRevisionView(root, view = 'rejected') {
    if (view !== 'rejected' && view !== 'accepted') throw new Error(`Unknown revision view '${view}'.`);
    const paragraphs = descendantWordElements(root, 'p');
    const removedMark = view === 'rejected' ? 'ins' : 'del';
    const records = [];
    const pending = new Map();
    paragraphs.forEach((paragraph, sourceIndex) => {
        if (hiddenByStructure(paragraph, view)) return;
        const container = containerOf(paragraph);
        const carried = pending.get(container) || null;
        pending.delete(container);
        const { text, spans } = paragraphContent(paragraph, view);
        const offset = carried ? carried.text.length : 0;
        const merged = {
            text: (carried?.text || '') + text,
            spans: [...(carried?.spans || []), ...spans.map(s => ({ ...s, start: s.start + offset, end: s.end + offset }))]
        };
        if (paragraphMarkRevision(paragraph) === removedMark) {
            pending.set(container, { ...merged, paragraph, sourceIndex });
            return;
        }
        records.push(finishRecord(paragraph, merged, view, container, sourceIndex));
    });
    // A removed mark on the last paragraph of a container (Word 16 COM evidence,
    // tests/fixtures/word-revision-semantics/expectations.json): Reject keeps a cell's last
    // paragraph; otherwise an emptied trailing paragraph disappears and the previous mark becomes
    // the container's last. Non-empty trailing content is kept (not yet pinned by COM evidence).
    for (const [container, carried] of pending) {
        const keepCellParagraph = view === 'rejected' && localName(container) === 'tc';
        if (!keepCellParagraph && carried.text.length === 0) continue;
        records.push(finishRecord(carried.paragraph, carried, view, container, carried.sourceIndex));
    }
    records.sort((a, b) => a.sourceIndex - b.sourceIndex);
    appendImplicitBodyParagraphs(root, paragraphs, records);
    return records;
}

/** Word adds an untracked empty paragraph after a table that ends a body. */
function appendImplicitBodyParagraphs(root, paragraphs, records) {
    const bodies = localName(root) === 'body' ? [root] : descendantWordElements(root, 'body');
    for (const body of bodies) {
        const blocks = Array.from(body.childNodes || []).filter(node => isWordElement(node) && localName(node) !== 'sectPr');
        const last = blocks[blocks.length - 1];
        if (!last || localName(last) !== 'tbl') continue;
        records.push({ text: '', props: '', numPr: null, wrappers: [], container: 'body', sourceIndex: paragraphs.length, implicit: true });
    }
}

function finishRecord(paragraph, content, view, container, sourceIndex) {
    const { props, numPr } = viewProperties(paragraph, view);
    return {
        text: content.text,
        props,
        numPr,
        wrappers: mergeSpans(content.spans),
        container: container ? localName(container) : 'root',
        sourceIndex
    };
}

/** Plain text per paragraph after Word Reject All / Accept All. */
export function wordRevisionParagraphTexts(root, view = 'rejected') {
    return projectWordRevisionView(root, view).map(record => record.text);
}
