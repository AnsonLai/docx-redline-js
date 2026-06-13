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
