import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';

import './setup-xml-provider.mjs';
import { openDocx } from '../index.js';
import { parseOoxmlSafe } from '../adapters/xml-adapter.js';
import { unzipDocx, zipDocx } from '../document/zip-archive.js';
import { applyOperationsToDocumentXml } from '../services/standalone-operation-runner.js';

/**
 * Pre-0.9.0 Track A: (A1) a plain `redline` keeps the source paragraph's `w14:paraId` / `w14:textId` on the surviving
 * paragraph, with unique ids; (A2) a text-changed list-header conversion keeps the source `pPr` on the tracked-deleted
 * paragraph (Reject All restores it) and copies the direct properties onto the new list paragraph. Both entry paths run
 * with the default Reject All verifier (enforce).
 */

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const AUTHOR = 'Identity Author';
const dec = new TextDecoder();
const enc = new TextEncoder();
const TEMPLATE = new URL('./fixtures/agentic-lists/nested-lists-source.docx', import.meta.url);

const p = (text, id, pPr = '') => `<w:p w14:paraId="${id}" w14:textId="77777777">${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}<w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
const parse = xml => {
    const parsed = parseOoxmlSafe(xml);
    assert.ok(parsed.doc, parsed.error?.message || 'parseable OOXML');
    return parsed.doc;
};
const paragraphsOf = xml => Array.from(parse(xml).getElementsByTagNameNS(W, 'p'));
const idOf = paragraph => paragraph.getAttributeNS(W14, 'paraId') || null;
const visibleText = paragraph => Array.from(paragraph.getElementsByTagNameNS(W, 't')).map(node => node.textContent).join('');
const hasMark = (paragraph, kind) => Array.from(paragraph.getElementsByTagNameNS(W, kind))
    .some(node => node.parentNode?.localName === 'rPr' && node.parentNode.parentNode?.localName === 'pPr' && node.parentNode.parentNode.parentNode === paragraph);
const directChildren = paragraph => {
    const pPr = Array.from(paragraph.childNodes).find(node => node.localName === 'pPr');
    return pPr ? Array.from(pPr.childNodes).filter(node => node.nodeType === 1).map(node => node.localName) : [];
};
const assertUniqueIds = xml => {
    const ids = paragraphsOf(xml).map(idOf).filter(Boolean);
    assert.equal(new Set(ids).size, ids.length, 'paragraph ids are unique');
};

function buildDocx(bodyXml) {
    const entries = unzipDocx(readFileSync(TEMPLATE));
    const original = dec.decode(entries.get('word/document.xml'));
    const sectPr = original.match(/<w:sectPr[\s\S]*<\/w:sectPr>\s*<\/w:body>/)[0];
    const head = original.slice(0, original.indexOf('<w:body>') + '<w:body>'.length);
    entries.set('word/document.xml', enc.encode(`${head}${bodyXml}${sectPr}</w:document>`));
    return new Uint8Array(zipDocx(entries));
}
const bareDocument = body => `<w:document xmlns:w="${W}" xmlns:w14="${W14}"><w:body>${body}<w:sectPr/></w:body></w:document>`;

const quiet = console.log;
console.log = () => { };
try {
    // A1 -------------------------------------------------------------------------------------------------------
    const body = p('Hello old world', '1A2B3C4D') + p('Other paragraph', '1A2B3C4E');
    const replaceOp = { type: 'redline', target: { exactText: 'Hello old world' }, modified: 'Hello new world' };
    const splitOp = { type: 'redline', target: { exactText: 'Hello old world' }, modified: 'Hello new world\nAdded paragraph' };

    // Bare runner: a replacement keeps the identity.
    let bare = await applyOperationsToDocumentXml(bareDocument(body), [replaceOp], AUTHOR, {}, {});
    assert.ok(bare.hasChanges, JSON.stringify(bare.results?.[0]?.error || bare.error || null));
    let paragraphs = paragraphsOf(bare.documentXml);
    assert.equal(idOf(paragraphs[0]), '1A2B3C4D');
    assert.equal(paragraphs[0].getAttributeNS(W14, 'textId'), '77777777');
    assert.equal(idOf(paragraphs[1]), '1A2B3C4E');
    assertUniqueIds(bare.documentXml);

    // Bare runner: an inserted split gives the identity to exactly one paragraph (the one whose mark Reject All keeps).
    bare = await applyOperationsToDocumentXml(bareDocument(body), [splitOp], AUTHOR, {}, {});
    paragraphs = paragraphsOf(bare.documentXml);
    assert.equal(paragraphs.length, 3);
    assert.equal(paragraphs.filter(paragraph => idOf(paragraph) === '1A2B3C4D').length, 1, 'the source id survives exactly once');
    assert.equal(idOf(paragraphs.find(paragraph => visibleText(paragraph) === 'Added paragraph')), '1A2B3C4D');
    assert.equal(hasMark(paragraphs[0], 'ins'), true);
    assertUniqueIds(bare.documentXml);

    // openDocx: same, through the package facade; the id still addresses the edited paragraph afterwards.
    const result = await openDocx(buildDocx(body)).applyOperations([replaceOp], { author: AUTHOR });
    assert.equal(result.status, 'ok', JSON.stringify(result.results?.[0]?.error || result.error || null));
    const output = result.toUint8Array();
    const outputXml = dec.decode(unzipDocx(output).get('word/document.xml'));
    assert.equal(idOf(paragraphsOf(outputXml)[0]), '1A2B3C4D');
    assertUniqueIds(outputXml);
    const inspected = openDocx(output).inspect({}).paragraphs;
    assert.equal(inspected.filter(item => item.paragraphId === '1A2B3C4D').length, 1);

    // A2 -------------------------------------------------------------------------------------------------------
    const header = '<w:keepNext/><w:spacing w:before="240" w:after="60"/>';
    const headerBody = p('1. First Header', '1A000001', header) + p('Body one.', '1A000002') + p('2. Second Header', '1A000003', '<w:keepNext/>');
    const convert = { type: 'redline', target: { exactText: '1. First Header' }, modified: 'A. First Header' };
    const checkConversion = (xml, label) => {
        const list = paragraphsOf(xml);
        const deleted = list.find(paragraph => hasMark(paragraph, 'del'));
        const created = list.find(paragraph => hasMark(paragraph, 'ins'));
        assert.ok(deleted && created, `${label}: deleted and created paragraphs exist`);
        assert.deepEqual(directChildren(deleted).filter(name => name !== 'rPr'), ['keepNext', 'spacing'], `${label}: deleted paragraph keeps the source pPr`);
        const names = directChildren(created);
        assert.ok(names.includes('keepNext') && names.includes('spacing') && names.includes('numPr'), `${label}: new paragraph carries the intended properties`);
        assert.ok(names.indexOf('keepNext') < names.indexOf('numPr') && names.indexOf('numPr') < names.indexOf('spacing'), `${label}: schema order`);
        assert.equal(idOf(created), '1A000001', `${label}: the surviving paragraph keeps the source id`);
        assert.equal(idOf(deleted), null, `${label}: the deleted paragraph does not duplicate it`);
        assertUniqueIds(xml);
    };

    const converted = await openDocx(buildDocx(headerBody)).applyOperations([convert], { author: AUTHOR });
    assert.equal(converted.status, 'ok', JSON.stringify(converted.results?.[0]?.error || converted.error || null));
    checkConversion(dec.decode(unzipDocx(converted.toUint8Array()).get('word/document.xml')), 'openDocx');

    const entries = unzipDocx(buildDocx(headerBody));
    const bareConverted = await applyOperationsToDocumentXml(
        dec.decode(entries.get('word/document.xml')), [convert], AUTHOR,
        { numberingXml: dec.decode(entries.get('word/numbering.xml')), stylesXml: dec.decode(entries.get('word/styles.xml')) }, {}
    );
    assert.ok(bareConverted.documentXml && bareConverted.hasChanges, JSON.stringify(bareConverted.results?.[0]?.error || bareConverted.error || null));
    checkConversion(bareConverted.documentXml, 'bare');
} finally {
    console.log = quiet;
}
console.log('paragraph_identity_tests passed');
