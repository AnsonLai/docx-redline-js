import assert from 'assert/strict';

export function parseXml(xml) {
    const doc = new DOMParser().parseFromString(xml, 'text/xml');
    const parseError = doc.getElementsByTagName('parsererror')[0];
    assert(!parseError, parseError?.textContent || 'XML parse error');
    return doc;
}

export function parseXmlFragment(xml) {
    try {
        return parseXml(xml);
    } catch {
        return parseXml(`<w:root xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${xml}</w:root>`);
    }
}

export function serializeXml(doc) {
    return new XMLSerializer().serializeToString(doc);
}

export function localName(node) {
    return String(node?.localName || node?.nodeName || '').replace(/^.*:/, '');
}

export function elementsByLocalName(node, name) {
    return Array.from(node.getElementsByTagName('*')).filter(el => localName(el) === name);
}

export function directChildByLocalName(node, name) {
    return Array.from(node?.childNodes || []).find(child => child.nodeType === 1 && localName(child) === name) || null;
}

export function textContentByLocalName(node, name) {
    return elementsByLocalName(node, name).map(el => el.textContent || '').join('');
}

export function runText(run) {
    return elementsByLocalName(run, 't').map(t => t.textContent || '').join('');
}

export function findRunByText(xml, text) {
    const doc = parseXml(xml);
    return elementsByLocalName(doc, 'r').find(run => runText(run) === text) || null;
}

export function textIndex(xml, text) {
    const index = xml.indexOf(text);
    assert.notEqual(index, -1, `Expected output to contain "${text}"`);
    return index;
}

export function formatIsEnabled(rPr, name) {
    const el = directChildByLocalName(rPr, name);
    if (!el) return false;

    const value = String(el.getAttribute('w:val') || el.getAttribute('val') || '').toLowerCase();
    if (!value) return true;
    if (name === 'u') return !['none', '0', 'false', 'off'].includes(value);
    return !['0', 'false', 'off'].includes(value);
}

export function formatIsDisabled(rPr, name) {
    const el = directChildByLocalName(rPr, name);
    assert(el, `Expected <w:${name}> to be present`);

    const value = String(el.getAttribute('w:val') || el.getAttribute('val') || '').toLowerCase();
    if (name === 'u') return ['none', '0', 'false', 'off'].includes(value);
    return ['0', 'false', 'off'].includes(value);
}

export function assertRunFormat(xml, text, expected) {
    const run = findRunByText(xml, text);
    assert(run, `Expected run containing exact text "${text}"`);

    const rPr = directChildByLocalName(run, 'rPr');
    assert(rPr, `Expected run "${text}" to have run properties`);

    if ('bold' in expected) {
        assert.equal(formatIsEnabled(rPr, 'b'), expected.bold, `Unexpected bold state for "${text}"`);
    }
    if ('italic' in expected) {
        assert.equal(formatIsEnabled(rPr, 'i'), expected.italic, `Unexpected italic state for "${text}"`);
    }
    if ('underline' in expected) {
        assert.equal(formatIsEnabled(rPr, 'u'), expected.underline, `Unexpected underline state for "${text}"`);
    }
    if ('strikethrough' in expected) {
        assert.equal(formatIsEnabled(rPr, 'strike'), expected.strikethrough, `Unexpected strikethrough state for "${text}"`);
    }
}

export function assertRunFormatDisabled(xml, text, names) {
    const run = findRunByText(xml, text);
    assert(run, `Expected run containing exact text "${text}"`);

    const rPr = directChildByLocalName(run, 'rPr');
    assert(rPr, `Expected run "${text}" to have run properties`);

    names.forEach(name => {
        assert(formatIsDisabled(rPr, name), `Expected <w:${name}> to be explicitly disabled for "${text}"`);
    });
}

function attr(node, localName) {
    return node.getAttribute(`w:${localName}`) || node.getAttribute(localName) || '';
}

function xmlSpace(node) {
    return node.getAttribute('xml:space') ||
        node.getAttribute('space') ||
        node.getAttributeNS?.('http://www.w3.org/XML/1998/namespace', 'space') ||
        '';
}

export function assertNoNestedRevisions(xml) {
    const doc = parseXmlFragment(xml);
    const revisions = elementsByLocalName(doc, 'ins').concat(elementsByLocalName(doc, 'del'));

    for (const revision of revisions) {
        const nested = Array.from(revision.getElementsByTagName('*')).filter(el => {
            if (el === revision) return false;
            const name = localName(el);
            return name === 'ins' || name === 'del';
        });
        assert.equal(nested.length, 0, `Expected <${revision.nodeName}> not to contain nested revisions`);
    }
}

export function assertDelUsesDelText(xml) {
    const doc = parseXmlFragment(xml);
    for (const del of elementsByLocalName(doc, 'del')) {
        const normalTextNodes = elementsByLocalName(del, 't');
        assert.equal(normalTextNodes.length, 0, 'Expected <w:del> to contain <w:delText>, not <w:t>');

        for (const run of elementsByLocalName(del, 'r')) {
            const textNodes = elementsByLocalName(run, 't').concat(elementsByLocalName(run, 'delText'));
            if (textNodes.length === 0) continue;
            assert(
                textNodes.every(node => localName(node) === 'delText'),
                'Expected every text node in a deleted run to be <w:delText>'
            );
        }
    }
}

export function assertRevisionMetadata(xml) {
    const doc = parseXmlFragment(xml);
    const revisions = elementsByLocalName(doc, 'ins').concat(elementsByLocalName(doc, 'del'));

    for (const revision of revisions) {
        assert.notEqual(attr(revision, 'id'), '', `Expected ${revision.nodeName} to have w:id`);
        assert.notEqual(attr(revision, 'author'), '', `Expected ${revision.nodeName} to have w:author`);
        assert.match(attr(revision, 'date'), /^\d{4}-\d{2}-\d{2}T/, `Expected ${revision.nodeName} to have an ISO-ish w:date`);
    }
}

export function assertUniqueRevisionIds(xml) {
    const doc = parseXmlFragment(xml);
    const revisionNames = new Set(['ins', 'del', 'rPrChange', 'pPrChange']);
    const seen = new Map();

    for (const node of Array.from(doc.getElementsByTagName('*')).filter(el => revisionNames.has(localName(el)))) {
        const id = attr(node, 'id');
        if (!id) continue;
        assert(!seen.has(id), `Expected revision id ${id} to be unique`);
        seen.set(id, node);
    }
}

export function assertSpacePreserved(xml) {
    const doc = parseXmlFragment(xml);
    const textNodes = elementsByLocalName(doc, 't').concat(elementsByLocalName(doc, 'delText'));

    for (const node of textNodes) {
        const text = node.textContent || '';
        if (!/^\s|\s$/.test(text)) continue;
        assert.equal(xmlSpace(node), 'preserve', `Expected ${node.nodeName} with boundary whitespace to preserve space`);
    }
}

/**
 * Asserts no w:p is nested inside another w:p.
 *
 * CT_P has no paragraph child in WordprocessingML, so nested paragraphs are
 * schema-invalid and Word treats the file as corrupt. None of the text-level
 * assertions can see this, because the visible characters can still come out
 * in the right order.
 */
export function assertNoNestedParagraphs(xml) {
    const doc = parseXmlFragment(xml);

    for (const paragraph of elementsByLocalName(doc, 'p')) {
        const nested = Array.from(paragraph.getElementsByTagName('*'))
            .filter(el => el !== paragraph && localName(el) === 'p');
        assert.equal(nested.length, 0, 'Expected <w:p> not to contain a nested <w:p>');
    }
}

/**
 * Asserts w:sectPr is the last child of w:body when present.
 *
 * CT_Body puts section properties last. services/standalone-docx-plumbing.js
 * already rejects a misplaced sectPr ("Validation failed: w:sectPr not last"),
 * so engine output that violates it fails the package's own plumbing later.
 */
export function assertSectPrLast(xml) {
    const doc = parseXmlFragment(xml);
    const body = elementsByLocalName(doc, 'body')[0];
    if (!body) return;

    const children = Array.from(body.childNodes).filter(node => node.nodeType === 1);
    const index = children.findIndex(child => localName(child) === 'sectPr');
    if (index === -1) return;

    assert.equal(
        index,
        children.length - 1,
        `Expected <w:sectPr> to be the last child of <w:body>, found it at index ${index} of ${children.length}`
    );
}

/*
 * Lossless visible-text extraction for verification.
 *
 * ingestWordOoxmlToPlainText is deliberately lossy: normalizeInlineWhitespace in
 * pipeline/ingestion-export.js collapses runs of spaces/tabs and trims each line
 * so the output reads well as plain text. That is correct for a display reader
 * and wrong for a test oracle -- a redline that drops a w:tab, loses an
 * xml:space="preserve", or doubles a space at a splice point would compare equal.
 *
 * These helpers walk the DOM independently of the production reader so a bug in
 * ingestion cannot mask itself.
 */

// Container properties carry no visible text; w:instrText is field plumbing.
const NON_VISIBLE_CONTAINERS = new Set([
    'pPr', 'rPr', 'sectPr', 'tblPr', 'trPr', 'tcPr', 'tblGrid', 'numPr', 'instrText'
]);

// Content inside these is invisible in Word's accepted view.
const ACCEPTED_VIEW_HIDDEN = new Set(['del', 'moveFrom']);

function collectExactText(node, out) {
    for (const child of Array.from(node?.childNodes || [])) {
        if (child.nodeType !== 1) continue;

        const name = localName(child);
        if (ACCEPTED_VIEW_HIDDEN.has(name) || NON_VISIBLE_CONTAINERS.has(name)) continue;
        // A nested w:p is malformed OOXML, but extraction must stay well-defined:
        // every paragraph is enumerated separately, so never recurse into one.
        if (name === 'p') continue;

        if (name === 't') {
            out.push(child.textContent || '');
        } else if (name === 'tab') {
            out.push('\t');
        } else if (name === 'br' || name === 'cr') {
            out.push('\n');
        } else if (name === 'noBreakHyphen') {
            out.push('‑');
        } else if (name === 'delText' || name === 'softHyphen') {
            // delText only appears inside w:del (already skipped); softHyphen is invisible.
            continue;
        } else {
            collectExactText(child, out);
        }
    }
}

/** True when the paragraph's own mark is marked deleted (w:pPr/w:rPr/w:del). */
function paragraphMarkIsDeleted(paragraph) {
    const pPr = directChildByLocalName(paragraph, 'pPr');
    const rPr = pPr && directChildByLocalName(pPr, 'rPr');
    return Boolean(rPr && directChildByLocalName(rPr, 'del'));
}

/**
 * Extracts the exact visible text of an OOXML fragment with NO whitespace
 * normalization, modelling Word's *accepted* view: w:del and w:moveFrom content
 * is invisible, w:ins and w:moveTo content is visible, and a deleted paragraph
 * mark merges its paragraph into the next one.
 *
 * Mapping: w:t -> textContent, w:tab -> '\t', w:br|w:cr -> '\n',
 * w:noBreakHyphen -> U+2011. Paragraph boundaries emit '\n'.
 *
 * @param {string} xml - OOXML fragment, document, or package payload
 * @returns {string}
 */
export function extractExactVisibleText(xml) {
    const doc = parseXmlFragment(xml);
    const paragraphs = elementsByLocalName(doc, 'p');

    if (paragraphs.length === 0) {
        const out = [];
        collectExactText(doc, out);
        return out.join('');
    }

    let text = '';
    paragraphs.forEach((paragraph, index) => {
        const out = [];
        collectExactText(paragraph, out);
        text += out.join('');

        const isLast = index === paragraphs.length - 1;
        if (!isLast && !paragraphMarkIsDeleted(paragraph)) {
            text += '\n';
        }
    });
    return text;
}

/**
 * Normalizes only paragraph separators, leaving spaces and tabs byte-exact.
 *
 * Markdown treats a blank line as a paragraph separator while OOXML represents
 * one paragraph break as one boundary, so `\n\n` and `\n` are the same document.
 * Every other whitespace difference is a real regression and stays visible.
 *
 * @param {string} text
 * @returns {string}
 */
export function normalizeParagraphBreaks(text) {
    return String(text ?? '').replace(/\r\n?/g, '\n').replace(/\n{2,}/g, '\n');
}
