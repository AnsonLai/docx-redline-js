import { parseOoxmlSafe } from '../adapters/xml-adapter.js';
import { NS_W } from '../core/types.js';
import { extractCanonicalParagraphText, readCanonicalRunText, extractParagraphRevisionSegments } from '../core/paragraph-text.js';
import { createParagraphFingerprint, getDocumentParagraphNodes, getParagraphId } from '../core/paragraph-targeting.js';
import { extractDocumentPartsEntries, computeRevisionTokenSync } from './revision-token.js';

const attr = (node, name) => node?.getAttribute?.(`w:${name}`) || node?.getAttribute?.(name) || '';
const descendants = (node, name) => Array.from(node?.getElementsByTagNameNS?.(NS_W, name) || []);
const first = (node, name) => descendants(node, name)[0] || null;
function hasAncestor(node, localName) { let cursor = node?.parentNode; while (cursor) { if (cursor.localName === localName && (!cursor.namespaceURI || cursor.namespaceURI === NS_W)) return true; cursor = cursor.parentNode; } return false; }

function parseXml(xml, partName, required = false) {
    if (!xml) return required ? { error: { code: 'MISSING_PART', message: `Missing ${partName}.` } } : { doc: null };
    const parsed = parseOoxmlSafe(xml, 'application/xml');
    if (!parsed.doc || parsed.error) return { error: { code: 'PARSE_ERROR', message: `Could not parse ${partName}: ${parsed.error?.message || 'invalid XML'}` } };
    return { doc: parsed.doc, warnings: parsed.warnings || [] };
}

function paragraphListProperties(paragraph) {
    const numPr = first(first(paragraph, 'pPr'), 'numPr');
    if (!numPr) return null;
    const numId = attr(first(numPr, 'numId'), 'val');
    if (!numId || numId === '0') return null;
    return { numId, level: Number.parseInt(attr(first(numPr, 'ilvl'), 'val') || '0', 10) || 0 };
}

function parseNumbering(numberingDoc) {
    const abstracts = new Map();
    for (const abstract of descendants(numberingDoc, 'abstractNum')) {
        const levels = new Map();
        for (const level of descendants(abstract, 'lvl')) {
            const ilvl = Number.parseInt(attr(level, 'ilvl') || '0', 10) || 0;
            levels.set(ilvl, {
                start: Number.parseInt(attr(first(level, 'start'), 'val') || '1', 10) || 1,
                format: attr(first(level, 'numFmt'), 'val') || 'decimal',
                text: attr(first(level, 'lvlText'), 'val') || `%${ilvl + 1}.`
            });
        }
        abstracts.set(attr(abstract, 'abstractNumId'), levels);
    }
    const nums = new Map();
    for (const num of descendants(numberingDoc, 'num')) {
        const abstractId = attr(first(num, 'abstractNumId'), 'val');
        const levels = new Map(abstracts.get(abstractId) || []);
        for (const override of descendants(num, 'lvlOverride')) {
            const ilvl = Number.parseInt(attr(override, 'ilvl') || '0', 10) || 0;
            const embedded = first(override, 'lvl');
            const base = { ...(levels.get(ilvl) || { start: 1, format: 'decimal', text: `%${ilvl + 1}.` }) };
            if (embedded) {
                base.start = Number.parseInt(attr(first(embedded, 'start'), 'val') || String(base.start), 10) || base.start;
                base.format = attr(first(embedded, 'numFmt'), 'val') || base.format;
                base.text = attr(first(embedded, 'lvlText'), 'val') || base.text;
            }
            const startOverride = first(override, 'startOverride');
            if (startOverride) base.start = Number.parseInt(attr(startOverride, 'val') || String(base.start), 10) || base.start;
            levels.set(ilvl, base);
        }
        nums.set(attr(num, 'numId'), levels);
    }
    return nums;
}

function alpha(value, upper) {
    let n = Math.max(1, value); let out = '';
    while (n > 0) { n -= 1; out = String.fromCharCode(97 + (n % 26)) + out; n = Math.floor(n / 26); }
    return upper ? out.toUpperCase() : out;
}
function roman(value) {
    const pairs = [[1000,'M'],[900,'CM'],[500,'D'],[400,'CD'],[100,'C'],[90,'XC'],[50,'L'],[40,'XL'],[10,'X'],[9,'IX'],[5,'V'],[4,'IV'],[1,'I']];
    let n = value; let out = ''; for (const [amount, glyph] of pairs) while (n >= amount) { out += glyph; n -= amount; } return out;
}
function formatCounter(value, format) {
    if (format === 'lowerLetter') return alpha(value, false);
    if (format === 'upperLetter') return alpha(value, true);
    if (format === 'lowerRoman') return roman(value).toLowerCase();
    if (format === 'upperRoman') return roman(value);
    return String(value);
}

function createNumberingResolver(numberingDoc) {
    const nums = numberingDoc ? parseNumbering(numberingDoc) : new Map();
    const counters = new Map();
    return list => {
        if (!list?.numId) return null;
        const levels = nums.get(String(list.numId));
        if (!levels) return { ...list, label: null, format: null };
        const state = counters.get(list.numId) || [];
        const definition = levels.get(list.level) || { start: 1, format: 'decimal', text: `%${list.level + 1}.` };
        state[list.level] = state[list.level] == null ? definition.start : state[list.level] + 1;
        state.length = list.level + 1;
        counters.set(list.numId, state);
        const label = definition.text.replace(/%([1-9])/g, (_, raw) => {
            const level = Number(raw) - 1;
            const levelDef = levels.get(level) || definition;
            return formatCounter(state[level] ?? levelDef.start, levelDef.format);
        });
        return { ...list, label, format: definition.format };
    };
}

function headingLevel(paragraph) {
    const pPr = first(paragraph, 'pPr');
    const style = attr(first(pPr, 'pStyle'), 'val');
    const match = style.match(/^heading\s*([1-9])$/i);
    if (match) return Math.min(Number(match[1]), 6);
    const outline = Number.parseInt(attr(first(pPr, 'outlineLvl'), 'val'), 10);
    return Number.isInteger(outline) ? Math.min(outline + 1, 6) : null;
}

function structuralContext(paragraph, text) {
    const references = [];
    for (const [name, type] of [['footnoteReference', 'footnote'], ['endnoteReference', 'endnote'], ['commentReference', 'comment']]) {
        for (const node of descendants(paragraph, name)) references.push({ type, id: attr(node, 'id') || null });
    }
    let cell = paragraph.parentNode; while (cell && cell.localName !== 'tc') cell = cell.parentNode;
    let row = cell?.parentNode; while (row && row.localName !== 'tr') row = row.parentNode;
    let table = row?.parentNode; while (table && table.localName !== 'tbl') table = table.parentNode;
    const all = paragraph.ownerDocument;
    return {
        references,
        table: table ? {
            tableIndex: Array.from(all.getElementsByTagNameNS(NS_W, 'tbl')).indexOf(table) + 1,
            rowIndex: Array.from(table.getElementsByTagNameNS(NS_W, 'tr')).indexOf(row) + 1,
            cellIndex: Array.from(row.getElementsByTagNameNS(NS_W, 'tc')).indexOf(cell) + 1
        } : null,
        empty: text.length === 0
    };
}

function revisionAuthors(paragraph) {
    const authors = new Set();
    for (const name of ['ins', 'del', 'moveFrom', 'moveTo', 'rPrChange', 'pPrChange']) {
        for (const node of descendants(paragraph, name)) if (attr(node, 'author')) authors.add(attr(node, 'author'));
    }
    return [...authors].sort();
}

function readCommentDefinitions(commentsDoc) {
    const result = new Map();
    for (const comment of descendants(commentsDoc, 'comment')) {
        const paragraphs = descendants(comment, 'p');
        result.set(attr(comment, 'id'), {
            id: attr(comment, 'id'), author: attr(comment, 'author') || null, date: attr(comment, 'date') || null,
            text: paragraphs.map(p => extractCanonicalParagraphText(p)).join('\n'),
            paraId: paragraphs[0]?.getAttribute?.('w14:paraId') || paragraphs[0]?.getAttribute?.('paraId') || null
        });
    }
    return result;
}

function attachCommentThreadMetadata(comments, commentsExtendedDoc) {
    if (!commentsExtendedDoc) return;
    const byParaId = new Map([...comments.values()].filter(c => c.paraId).map(c => [c.paraId.toUpperCase(), c]));
    for (const entry of Array.from(commentsExtendedDoc.getElementsByTagNameNS('*', 'commentEx'))) {
        const paraId = entry.getAttribute('w15:paraId') || entry.getAttribute('paraId') || '';
        const parentParaId = entry.getAttribute('w15:paraIdParent') || entry.getAttribute('paraIdParent') || '';
        const comment = byParaId.get(paraId.toUpperCase());
        if (!comment) continue;
        comment.done = (entry.getAttribute('w15:done') || entry.getAttribute('done')) === '1';
        if (parentParaId) {
            comment.parentParaId = parentParaId;
            comment.parentCommentId = byParaId.get(parentParaId.toUpperCase())?.id || null;
        }
    }
}

function collectDocumentCommentAnchors(paragraphNodes, revisionView) {
    const active = new Map(); const anchors = new Map();
    let paragraphBoundary = null;
    const append = value => { for (const item of active.values()) item.text += value; };
    const visit = node => {
        for (const child of Array.from(node?.childNodes || [])) {
            if (child?.nodeType !== 1) continue;
            const name = child.localName;
            if ((revisionView === 'accepted' && (name === 'del' || name === 'moveFrom')) || (revisionView === 'rejected' && (name === 'ins' || name === 'moveTo'))) continue;
            if (name === 'commentRangeStart') active.set(attr(child, 'id'), { text: '' });
            else if (name === 'commentRangeEnd') { const id = attr(child, 'id'); if (active.has(id)) { anchors.set(id, active.get(id).text); active.delete(id); } }
            else if (name === 'r') append(readCanonicalRunText(child, { revisionView, boundary: paragraphBoundary }));
            else visit(child);
        }
    };
    paragraphNodes.forEach((paragraph, index) => { paragraphBoundary = paragraph; visit(paragraph); if (index < paragraphNodes.length - 1 && active.size) append('\n'); });
    return anchors;
}

/** Read-only, stable document-parts inspection for agents and package adapters. */
export function inspectDocumentParts(parts, options = {}) {
    const documentPart = parseXml(parts?.documentXml, 'word/document.xml', true);
    if (documentPart.error) return { status: 'error', error: documentPart.error, paragraphs: [], comments: [], warnings: [] };
    const commentsPart = parseXml(parts?.commentsXml, 'word/comments.xml');
    const commentsExtendedPart = parseXml(parts?.commentsExtendedXml, 'word/commentsExtended.xml');
    const numberingPart = parseXml(parts?.numberingXml, 'word/numbering.xml');
    const warnings = [...(documentPart.warnings || [])];
    if (commentsPart.error) warnings.push(commentsPart.error.message);
    if (commentsExtendedPart.error) warnings.push(commentsExtendedPart.error.message);
    if (numberingPart.error) warnings.push(numberingPart.error.message);
    const comments = readCommentDefinitions(commentsPart.doc);
    attachCommentThreadMetadata(comments, commentsExtendedPart.doc);
    const resolveNumbering = createNumberingResolver(numberingPart.doc);
    let nearestHeading = null;
    const paragraphNodes = getDocumentParagraphNodes(documentPart.doc);
    const commentAnchors = collectDocumentCommentAnchors(paragraphNodes, options.revisionView || 'accepted');
    let paragraphs = paragraphNodes.map((paragraph, zeroIndex) => {
        const text = extractCanonicalParagraphText(paragraph, { revisionView: options.revisionView || 'accepted' });
        const level = headingLevel(paragraph);
        if (level) nearestHeading = { level, text };
        const ids = [...new Set([...descendants(paragraph, 'commentRangeStart'), ...descendants(paragraph, 'commentReference')].map(node => attr(node, 'id')).filter(Boolean))];
        const authors = revisionAuthors(paragraph);
        const list = resolveNumbering(paragraphListProperties(paragraph));
        const styleId = attr(first(first(paragraph, 'pPr'), 'pStyle'), 'val') || null;
        const structure = structuralContext(paragraph, text);
        const index = zeroIndex + 1;
        const provision = list?.label && list.format !== 'bullet' ? list.label : null;
        const headingText = nearestHeading?.text || null;
        const humanReference = [provision, headingText, text.slice(0, options.excerptLength || 120)].filter(Boolean).join(' — ');
        const segments = extractParagraphRevisionSegments(paragraph);
        return {
            index, ref: `P${index}`, paragraphId: getParagraphId(paragraph), fingerprint: createParagraphFingerprint(paragraph),
            text, exactText: text, excerpt: text.slice(0, options.excerptLength || 120), humanReference, inTable: hasAncestor(paragraph, 'tc'), table: structure.table,
            styleId, headingLevel: level, nearestHeading, list, structuralReferences: structure.references, hasRevisions: authors.length > 0, revisionAuthors: authors, commentIds: ids,
            segments
        };
    });
    for (const paragraph of paragraphs) for (const id of paragraph.commentIds) {
        const definition = comments.get(id) || { id, author: null, date: null, text: '' };
        definition.paragraphIndex ??= paragraph.index; definition.targetRef ??= paragraph.ref;
        definition.anchoredText ??= commentAnchors.get(id) || paragraph.text;
        comments.set(id, definition);
    }
    if (options.revisedOnly) paragraphs = paragraphs.filter(item => item.hasRevisions);
    if (options.inTable != null) paragraphs = paragraphs.filter(item => item.inTable === !!options.inTable);
    if (options.skipEmpty) paragraphs = paragraphs.filter(item => item.text.length > 0);
    if (options.search) { const needle = String(options.search).toLowerCase(); paragraphs = paragraphs.filter(item => item.text.toLowerCase().includes(needle)); }
    if (Array.isArray(options.indexes)) { const indexes = new Set(options.indexes); paragraphs = paragraphs.filter(item => indexes.has(item.index)); }
    if (options.range) { const start = Number(options.range.start ?? options.range[0]); const end = Number(options.range.end ?? options.range[1]); paragraphs = paragraphs.filter(item => item.index >= start && item.index <= end); }
    const allRevisionAuthors = [...new Set(paragraphs.flatMap(item => item.revisionAuthors))].sort();
    const coveredEntries = extractDocumentPartsEntries(parts);
    const coveredParts = coveredEntries.map(e => e.name).sort();
    let revisionToken = null;
    if (typeof options.digestFn === 'function') {
        revisionToken = computeRevisionTokenSync({
            scope: 'document-parts',
            entries: coveredEntries,
            digestFn: options.digestFn
        });
    }
    return {
        status: 'ok',
        revisionToken,
        coveredParts,
        paragraphs,
        comments: [...comments.values()],
        revisionAuthors: allRevisionAuthors,
        commentAuthors: [...new Set([...comments.values()].map(item => item.author).filter(Boolean))].sort(),
        counts: { paragraphs: paragraphs.length, comments: comments.size, revisedParagraphs: paragraphs.filter(item => item.hasRevisions).length },
        warnings
    };
}
