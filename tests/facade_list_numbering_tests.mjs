import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';

import './setup-xml-provider.mjs';
import { openDocx, acceptTrackedChangesInOoxml, rejectTrackedChangesInOoxml } from '../index.js';
import { parseOoxmlSafe } from '../adapters/xml-adapter.js';
import { unzipDocx } from '../document/zip-archive.js';

// Regression: a same-kind list edit through the complete DOCX facade must not repoint source
// paragraphs at a new numbering definition, must continue the source list for the inserted
// items, and must leave Reject All referencing the original definitions.

const FIXTURE = new URL('./fixtures/agentic-lists/nested-lists-source.docx', import.meta.url);
const sourceBytes = new Uint8Array(readFileSync(FIXTURE));
const decoder = new TextDecoder();
const partOf = (bytes, name) => decoder.decode(unzipDocx(bytes).get(name));
const elements = (node, name) => Array.from(node.getElementsByTagNameNS('*', name));
const parse = xml => {
    const parsed = parseOoxmlSafe(xml);
    assert.ok(parsed.doc, parsed.error?.message || 'parseable OOXML');
    return parsed.doc;
};
const textOf = p => elements(p, 't').map(t => t.textContent).join('');
const numPrOf = p => {
    const numPr = elements(p, 'numPr').find(n => n.parentNode?.localName === 'pPr');
    if (!numPr) return null;
    const val = name => elements(numPr, name)[0]?.getAttribute('w:val') ?? null;
    return { numId: val('numId'), ilvl: val('ilvl') };
};
const listParagraphs = xml => elements(parse(xml), 'p').map(p => ({ text: textOf(p), numPr: numPrOf(p) }));

const definitionFor = (numberingXml, numId) => {
    const doc = parse(numberingXml);
    const num = elements(doc, 'num').find(n => n.getAttribute('w:numId') === String(numId));
    assert.ok(num, `numId ${numId} defined`);
    const absId = elements(num, 'abstractNumId')[0].getAttribute('w:val');
    const abstract = elements(doc, 'abstractNum').find(a => a.getAttribute('w:abstractNumId') === absId);
    assert.ok(abstract, `abstractNum ${absId} defined`);
    return {
        num: new XMLSerializer().serializeToString(num).replace(/\s+/g, ' '),
        abstract: new XMLSerializer().serializeToString(abstract).replace(/\s+/g, ' ')
    };
};

const doc = await openDocx(sourceBytes);
const result = await doc.applyOperations([{
    type: 'redline',
    target: { index: 3, exactText: 'Bullet Root A' },
    targetEnd: { index: 4, exactText: 'Bullet Insertion Anchor' },
    modified: '- Planner replacement parent\n    - Planner replacement child'
}], {
    author: 'Consumer Review', atomic: true, strictTargets: true, structuredContent: true, generateRedlines: true
});
assert.equal(result.status, 'ok');

const outBytes = result.uint8Array;
const sourceDocXml = partOf(sourceBytes, 'word/document.xml');
const sourceNumbering = partOf(sourceBytes, 'word/numbering.xml');
const outDocXml = partOf(outBytes, 'word/document.xml');
const outNumbering = partOf(outBytes, 'word/numbering.xml');

const source = listParagraphs(sourceDocXml);
const sourceByText = new Map(source.map(p => [p.text, p.numPr]));
const sourceRoot = sourceByText.get('Bullet Root A');
assert.ok(sourceRoot?.numId, 'fixture: Bullet Root A is a list item');

// 1. Every definition referenced by a source paragraph is semantically unchanged.
for (const numId of new Set(source.map(p => p.numPr?.numId).filter(Boolean))) {
    assert.deepEqual(definitionFor(outNumbering, numId), definitionFor(sourceNumbering, numId), `numId ${numId} definition unchanged`);
}

// 2. Untouched list paragraphs keep numId and ilvl.
const out = listParagraphs(outDocXml);
for (const p of source) {
    if (!p.numPr) continue;
    if (['Bullet Root A', 'Bullet Insertion Anchor'].includes(p.text)) continue;
    const match = out.find(o => o.text === p.text);
    assert.deepEqual(match?.numPr, p.numPr, `untouched '${p.text}' keeps numPr`);
}

// 3. Inserted paragraphs continue the source list (same numId, so Accept All keeps tail numbering).
const inserted = ['Planner replacement parent', 'Planner replacement child'].map(t => out.find(o => o.text === t));
assert.equal(inserted[0].numPr.numId, sourceRoot.numId, 'inserted parent reuses the source numId');
assert.equal(inserted[1].numPr.numId, sourceRoot.numId, 'inserted child reuses the source numId');
assert.equal(inserted[0].numPr.ilvl, '0');
assert.equal(inserted[1].numPr.ilvl, '1');

// 4. Reject All restores the source numPr per paragraph and numbering parts are not needed to change.
const rejected = listParagraphs(rejectTrackedChangesInOoxml(outDocXml, { allAuthors: true }).oxml);
assert.deepEqual(rejected, source, 'Reject All restores source paragraphs and numPr exactly');
assert.equal(outNumbering.match(/<w:abstractNum /g).length, sourceNumbering.match(/<w:abstractNum /g).length, 'no unused abstractNum definitions added');
assert.equal(outNumbering.match(/<w:num /g).length, sourceNumbering.match(/<w:num /g).length, 'no unused num definitions added');

// 5. Accept All: replacement items sit in the source list, tail untouched.
const accepted = listParagraphs(acceptTrackedChangesInOoxml(outDocXml, { allAuthors: true }).oxml);
const acceptedBullets = accepted.filter(p => p.numPr?.numId === sourceRoot.numId).map(p => p.text);
assert.deepEqual(acceptedBullets, ['Planner replacement parent', 'Planner replacement child', 'Bullet Untouched Tail', 'Bullet Root B']);

// 6. Format-changing edits stay on the allocation path: a numbered replacement of a bullet list keeps
// the source definitions intact and uses a fresh numId for the new items.
const changed = await (await openDocx(sourceBytes)).applyOperations([{
    type: 'redline',
    target: { index: 3, exactText: 'Bullet Root A' },
    targetEnd: { index: 4, exactText: 'Bullet Insertion Anchor' },
    modified: '1. Numbered replacement parent\n    1. Numbered replacement child'
}], { author: 'Consumer Review', atomic: true, strictTargets: true, structuredContent: true, generateRedlines: true });
assert.equal(changed.status, 'ok');
const changedOut = listParagraphs(partOf(changed.uint8Array, 'word/document.xml'));
const newItem = changedOut.find(p => p.text === 'Numbered replacement parent');
assert.notEqual(newItem.numPr.numId, sourceRoot.numId, 'format change gets its own numId');
assert.deepEqual(listParagraphs(rejectTrackedChangesInOoxml(partOf(changed.uint8Array, 'word/document.xml'), { allAuthors: true }).oxml), source, 'format-change Reject All restores source numPr');

console.log('facade_list_numbering_tests: all passed');
