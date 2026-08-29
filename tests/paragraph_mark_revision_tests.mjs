import './setup-xml-provider.mjs';

import assert from 'assert/strict';

import {
    acceptTrackedChangesInOoxml,
    applyRedlineToOxml,
    rejectTrackedChangesInOoxml
} from '../index.js';
import { assertRoundTripStructure } from './helpers/roundtrip.mjs';
import { directChildByLocalName, elementsByLocalName, parseXml } from './helpers/ooxml-assertions.mjs';

const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

function documentXml(paragraphs) {
    return `<w:document xmlns:w="${NS_W}"><w:body>${paragraphs.join('')}<w:sectPr/></w:body></w:document>`;
}

function paragraph(text) {
    return `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
}

function wordParagraphs(xml) {
    return Array.from(parseXml(xml).getElementsByTagNameNS(NS_W, 'p'));
}

function paragraphText(paragraphNode) {
    return elementsByLocalName(paragraphNode, 't').map(node => node.textContent || '').join('');
}

function paragraphMarkRevision(paragraphNode, type) {
    const pPr = directChildByLocalName(paragraphNode, 'pPr');
    const rPr = directChildByLocalName(pPr, 'rPr');
    return directChildByLocalName(rPr, type);
}

async function testInsertedParagraphMarkRoundTrip() {
    const source = documentXml([paragraph('one')]);
    const result = await applyRedlineToOxml(source, 'one', 'one\ntwo', {
        generateRedlines: true,
        author: 'Phase3'
    });

    assert.equal(result.hasChanges, true);
    assertRoundTripStructure(result.oxml);

    const redlinedParagraphs = wordParagraphs(result.oxml);
    assert.equal(redlinedParagraphs.length, 2, 'redlined insertion should contain the inserted paragraph node');
    assert.ok(paragraphMarkRevision(redlinedParagraphs[0], 'ins'), 'inserted boundary should carry a paragraph-mark insertion');

    const accepted = acceptTrackedChangesInOoxml(result.oxml, { author: 'Phase3' });
    const acceptedParagraphs = wordParagraphs(accepted.oxml);
    assert.equal(acceptedParagraphs.length, 2, 'accepting paragraph insertion should keep two paragraphs');
    assert.deepEqual(acceptedParagraphs.map(paragraphText), ['one', 'two']);

    const rejected = rejectTrackedChangesInOoxml(result.oxml, { author: 'Phase3' });
    const rejectedParagraphs = wordParagraphs(rejected.oxml);
    assert.equal(rejectedParagraphs.length, 1, 'rejecting paragraph insertion should restore one paragraph');
    assert.deepEqual(rejectedParagraphs.map(paragraphText), ['one']);
}

async function testDeletedParagraphMarkRoundTrip() {
    const source = documentXml([paragraph('one'), paragraph('two')]);
    const result = await applyRedlineToOxml(source, 'one\ntwo', 'one', {
        generateRedlines: true,
        author: 'Phase3'
    });

    assert.equal(result.hasChanges, true);
    assertRoundTripStructure(result.oxml);

    const redlinedParagraphs = wordParagraphs(result.oxml);
    assert.equal(redlinedParagraphs.length, 2, 'redlined deletion should retain the deleted paragraph node');
    assert.ok(paragraphMarkRevision(redlinedParagraphs[0], 'del'), 'deleted boundary should carry a paragraph-mark deletion');
    assert.ok(elementsByLocalName(redlinedParagraphs[1], 'delText').some(node => node.textContent === 'two'), 'deleted paragraph text should be in w:delText');

    const accepted = acceptTrackedChangesInOoxml(result.oxml, { author: 'Phase3' });
    const acceptedParagraphs = wordParagraphs(accepted.oxml);
    assert.equal(acceptedParagraphs.length, 1, 'accepting paragraph deletion should remove the deleted paragraph');
    assert.deepEqual(acceptedParagraphs.map(paragraphText), ['one']);

    const rejected = rejectTrackedChangesInOoxml(result.oxml, { author: 'Phase3' });
    const rejectedParagraphs = wordParagraphs(rejected.oxml);
    assert.equal(rejectedParagraphs.length, 2, 'rejecting paragraph deletion should restore both paragraphs');
    assert.deepEqual(rejectedParagraphs.map(paragraphText), ['one', 'two']);
}

async function run() {
    await testInsertedParagraphMarkRoundTrip();
    await testDeletedParagraphMarkRoundTrip();
    console.log('PASS: paragraph mark revision tests');
}

run().catch(err => {
    console.error('FAIL:', err.message);
    process.exit(1);
});
