/**
 * Paragraph boundary mutation implementation and safeguard tests.
 *
 * Validates:
 * 1. Split paragraph: retaining source text without false insertions.
 * 2. Boundary removal: merging paragraphs A and B with tracked deletion.
 * 3. Entire paragraph deletion: emitting w:pPr/w:rPr/w:del and verifying Accept All
 *    leaves no ghost paragraph (and Reject All restores cleanly).
 * 4. Blank paragraph insertion and deletion.
 * 5. List item boundary handling.
 * 6. Safeguards (UNSAFE_PARAGRAPH_BOUNDARY):
 *    - Refusing deletion of sole terminal paragraph in a table cell (violates OOXML schema).
 *    - Refusing boundary joins across different table cells / boundaries.
 *    - Refusing deletion of paragraph containing section break (w:sectPr).
 *    - Refusing split across field instructions (w:fldSimple or unclosed w:fldChar).
 *    - Refusing split across open bookmark ranges.
 */

import test from 'node:test';
import assert from 'node:assert/strict';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { configureXmlProvider } from '../adapters/xml-adapter.js';
import {
    applyOperationsToDocumentXml,
    preflightOperations
} from '../services/standalone-operation-runner.js';
import {
    acceptTrackedChangesInOoxml,
    rejectTrackedChangesInOoxml
} from '../services/revision-comment-management.js';
import { validateParagraphBoundaryMutation } from '../core/paragraph-targeting.js';

configureXmlProvider({ DOMParser, XMLSerializer });

const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

function parseXml(xmlString) {
    const cleanXml = String(xmlString).replace(/^\uFEFF/, '').trim();
    return new DOMParser().parseFromString(cleanXml, 'application/xml');
}

function getWordElements(docOrEl, localName) {
    const list = [];
    const elements = docOrEl.getElementsByTagNameNS ? docOrEl.getElementsByTagNameNS(NS_W, localName) : [];
    for (let i = 0; i < elements.length; i++) {
        list.push(elements[i]);
    }
    return list;
}

test('WP-10: Paragraph Split - splits paragraph retaining text and roundtrips Accept/Reject', async () => {
    const docXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${NS_W}">
  <w:body>
    <w:p>
      <w:r><w:t>The quick brown fox jumps over the lazy dog.</w:t></w:r>
    </w:p>
  </w:body>
</w:document>`;

    const result = await applyOperationsToDocumentXml(docXml, [{
        type: 'redline',
        target: 'The quick brown fox jumps over the lazy dog.',
        modified: 'The quick brown fox\njumps over the lazy dog.'
    }], 'TestAuthor');

    assert.equal(result.hasChanges, true);
    assert.equal(result.results[0].status, 'applied');

    const pendingDoc = parseXml(result.documentXml);
    const pendingParas = getWordElements(pendingDoc, 'p');
    assert.equal(pendingParas.length, 2, 'Should create 2 paragraphs in pending view');

    // Section 1: Check acceptTrackedChangesInOoxml
    const acceptedResult = acceptTrackedChangesInOoxml(result.documentXml, { allAuthors: true });
    const acceptedDoc = parseXml(acceptedResult.oxml);
    const acceptedParas = getWordElements(acceptedDoc, 'p');
    assert.equal(acceptedParas.length, 2, 'Accepted document must have 2 paragraphs');
    assert.equal(acceptedParas[0].textContent.trim(), 'The quick brown fox');
    assert.equal(acceptedParas[1].textContent.trim(), 'jumps over the lazy dog.');

    // Section 2: Check rejectTrackedChangesInOoxml
    const rejectedResult = rejectTrackedChangesInOoxml(result.documentXml, { allAuthors: true });
    const rejectedDoc = parseXml(rejectedResult.oxml);
    const rejectedParas = getWordElements(rejectedDoc, 'p');
    assert.equal(rejectedParas.length, 1, 'Rejected document must restore to 1 paragraph');
    assert.equal(rejectedParas[0].textContent.trim(), 'The quick brown fox jumps over the lazy dog.');
});

test('WP-10: Entire Paragraph Deletion - emits w:pPr/w:rPr/w:del and leaves NO ghost paragraph upon acceptance', async () => {
    const docXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${NS_W}">
  <w:body>
    <w:p><w:r><w:t>First paragraph.</w:t></w:r></w:p>
    <w:p><w:r><w:t>Second paragraph to delete.</w:t></w:r></w:p>
    <w:p><w:r><w:t>Third paragraph.</w:t></w:r></w:p>
  </w:body>
</w:document>`;

    const result = await applyOperationsToDocumentXml(docXml, [{
        type: 'redline',
        target: 'Second paragraph to delete.',
        modified: ''
    }], 'TestAuthor');

    assert.equal(result.hasChanges, true);
    assert.equal(result.results[0].status, 'applied');

    const pendingDoc = parseXml(result.documentXml);
    const pendingParas = getWordElements(pendingDoc, 'p');
    assert.equal(pendingParas.length, 3, 'Pending view retains 3 paragraphs with revision markup');

    // The deleted paragraph must have paragraph mark revision w:pPr/w:rPr/w:del
    const delP = pendingParas[1];
    const pPr = getWordElements(delP, 'pPr')[0];
    assert.ok(pPr, 'Deleted paragraph must have w:pPr');
    const rPr = getWordElements(pPr, 'rPr')[0];
    assert.ok(rPr, 'Deleted paragraph w:pPr must contain w:rPr');
    const pDel = getWordElements(rPr, 'del')[0];
    assert.ok(pDel, 'Deleted paragraph mark must contain w:del');

    // Acceptance must eliminate the paragraph completely (no ghost paragraph!)
    const acceptedResult = acceptTrackedChangesInOoxml(result.documentXml, { allAuthors: true });
    const acceptedDoc = parseXml(acceptedResult.oxml);
    const acceptedParas = getWordElements(acceptedDoc, 'p');
    assert.equal(acceptedParas.length, 2, 'Acceptance must remove the deleted paragraph completely');
    assert.equal(acceptedParas[0].textContent.trim(), 'First paragraph.');
    assert.equal(acceptedParas[1].textContent.trim(), 'Third paragraph.');

    // Rejection must restore all 3 paragraphs cleanly
    const rejectedResult = rejectTrackedChangesInOoxml(result.documentXml, { allAuthors: true });
    const rejectedDoc = parseXml(rejectedResult.oxml);
    const rejectedParas = getWordElements(rejectedDoc, 'p');
    assert.equal(rejectedParas.length, 3, 'Rejection must restore the original 3 paragraphs');
    assert.equal(rejectedParas[1].textContent.trim(), 'Second paragraph to delete.');
});

test('WP-10: Blank Paragraph Deletion - removes blank paragraph cleanly', async () => {
    const docXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${NS_W}">
  <w:body>
    <w:p><w:r><w:t>Heading</w:t></w:r></w:p>
    <w:p><w:pPr/></w:p>
    <w:p><w:r><w:t>Content</w:t></w:r></w:p>
  </w:body>
</w:document>`;

    const result = await applyOperationsToDocumentXml(docXml, [{
        type: 'redline',
        target: { index: 2 },
        modified: ''
    }], 'TestAuthor');

    assert.equal(result.hasChanges, true);
    assert.equal(result.results[0].status, 'applied');

    const acceptedResult = acceptTrackedChangesInOoxml(result.documentXml, { allAuthors: true });
    const acceptedDoc = parseXml(acceptedResult.oxml);
    const acceptedParas = getWordElements(acceptedDoc, 'p');
    assert.equal(acceptedParas.length, 2, 'Blank paragraph removed after acceptance');
});

test('WP-10: Safeguard - Refuse deletion of sole terminal paragraph in table cell', async () => {
    const docXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${NS_W}">
  <w:body>
    <w:tbl>
      <w:tr>
        <w:tc>
          <w:p><w:r><w:t>Sole cell paragraph.</w:t></w:r></w:p>
        </w:tc>
      </w:tr>
    </w:tbl>
  </w:body>
</w:document>`;

    const ops = [{
        type: 'redline',
        target: 'Sole cell paragraph.',
        modified: ''
    }];

    // 1. Preflight test
    const preflight = preflightOperations(docXml, ops, 'TestAuthor');
    assert.equal(preflight.valid, false);
    assert.equal(preflight.results[0].status, 'error');
    assert.equal(preflight.results[0].error.code, 'UNSAFE_PARAGRAPH_BOUNDARY');
    assert.match(preflight.results[0].error.message, /sole terminal paragraph.*table cell/i);

    // 2. Application test
    const result = await applyOperationsToDocumentXml(docXml, ops, 'TestAuthor');
    assert.equal(result.status, 'error');
    assert.equal(result.error.code, 'BATCH_OPERATION_FAILED');
    assert.equal(result.results[0].error.code, 'UNSAFE_PARAGRAPH_BOUNDARY');
    assert.equal(result.hasChanges, false);
});

test('WP-10: Safeguard - Refuse deletion of paragraph containing section break (w:sectPr)', async () => {
    const docXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${NS_W}">
  <w:body>
    <w:p>
      <w:pPr>
        <w:sectPr>
          <w:pgSz w:w="12240" w:h="15840"/>
        </w:sectPr>
      </w:pPr>
      <w:r><w:t>Paragraph with section break.</w:t></w:r>
    </w:p>
    <w:p><w:r><w:t>Following section paragraph.</w:t></w:r></w:p>
  </w:body>
</w:document>`;

    const ops = [{
        type: 'redline',
        target: 'Paragraph with section break.',
        modified: ''
    }];

    // 1. Preflight test
    const preflight = preflightOperations(docXml, ops, 'TestAuthor');
    assert.equal(preflight.valid, false);
    assert.equal(preflight.results[0].status, 'error');
    assert.equal(preflight.results[0].error.code, 'UNSAFE_PARAGRAPH_BOUNDARY');
    assert.match(preflight.results[0].error.message, /section properties/i);

    // 2. Application test
    const result = await applyOperationsToDocumentXml(docXml, ops, 'TestAuthor');
    assert.equal(result.status, 'error');
    assert.equal(result.error.code, 'BATCH_OPERATION_FAILED');
    assert.equal(result.results[0].error.code, 'UNSAFE_PARAGRAPH_BOUNDARY');
    assert.equal(result.hasChanges, false);
});

test('WP-10: Safeguard - Refuse split across simple field instruction (w:fldSimple)', async () => {
    const docXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${NS_W}">
  <w:body>
    <w:p>
      <w:r><w:t>See page </w:t></w:r>
      <w:fldSimple w:instr="PAGE"/>
      <w:r><w:t> for more information.</w:t></w:r>
    </w:p>
  </w:body>
</w:document>`;

    const ops = [{
        type: 'redline',
        target: 'See page for more information.',
        modified: 'See page\nfor more information.'
    }];

    const preflight = preflightOperations(docXml, ops, 'TestAuthor');
    assert.equal(preflight.valid, false);
    assert.equal(preflight.results[0].status, 'error');
    assert.equal(preflight.results[0].error.code, 'UNSAFE_PARAGRAPH_BOUNDARY');
    assert.match(preflight.results[0].error.message, /simple field instruction/i);

    const result = await applyOperationsToDocumentXml(docXml, ops, 'TestAuthor');
    assert.equal(result.status, 'error');
    assert.equal(result.error.code, 'BATCH_OPERATION_FAILED');
    assert.equal(result.results[0].error.code, 'UNSAFE_PARAGRAPH_BOUNDARY');
});

test('WP-10: Safeguard - Refuse split across unclosed field instruction (w:fldChar)', async () => {
    const docXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${NS_W}">
  <w:body>
    <w:p>
      <w:r><w:fldChar w:fldCharType="begin"/></w:r>
      <w:r><w:instrText>TOC \\o "1-3"</w:instrText></w:r>
      <w:r><w:t>Table of Contents</w:t></w:r>
    </w:p>
  </w:body>
</w:document>`;

    const ops = [{
        type: 'redline',
        target: 'Table of Contents',
        modified: 'Table\nof Contents'
    }];

    const preflight = preflightOperations(docXml, ops, 'TestAuthor');
    assert.equal(preflight.valid, false);
    assert.equal(preflight.results[0].status, 'error');
    assert.equal(preflight.results[0].error.code, 'UNSAFE_PARAGRAPH_BOUNDARY');
    assert.match(preflight.results[0].error.message, /unclosed field instruction/i);

    const result = await applyOperationsToDocumentXml(docXml, ops, 'TestAuthor');
    assert.equal(result.status, 'error');
    assert.equal(result.error.code, 'BATCH_OPERATION_FAILED');
    assert.equal(result.results[0].error.code, 'UNSAFE_PARAGRAPH_BOUNDARY');
});

test('WP-10: Safeguard - Refuse split across open bookmark range', async () => {
    const docXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${NS_W}">
  <w:body>
    <w:p>
      <w:bookmarkStart w:id="10" w:name="CrossDocBookmark"/>
      <w:r><w:t>Important bookmarked text spanning across paragraphs.</w:t></w:r>
    </w:p>
    <w:p>
      <w:r><w:t>Second paragraph.</w:t></w:r>
      <w:bookmarkEnd w:id="10"/>
    </w:p>
  </w:body>
</w:document>`;

    const ops = [{
        type: 'redline',
        target: 'Important bookmarked text spanning across paragraphs.',
        modified: 'Important bookmarked text\nspanning across paragraphs.'
    }];

    const preflight = preflightOperations(docXml, ops, 'TestAuthor');
    assert.equal(preflight.valid, false);
    assert.equal(preflight.results[0].status, 'error');
    assert.equal(preflight.results[0].error.code, 'UNSAFE_PARAGRAPH_BOUNDARY');
    assert.match(preflight.results[0].error.message, /open bookmark range/i);

    const result = await applyOperationsToDocumentXml(docXml, ops, 'TestAuthor');
    assert.equal(result.status, 'error');
    assert.equal(result.error.code, 'BATCH_OPERATION_FAILED');
    assert.equal(result.results[0].error.code, 'UNSAFE_PARAGRAPH_BOUNDARY');
});

test('WP-10: Safeguard - validateParagraphBoundaryMutation directly reports cross-cell join violations', () => {
    const docXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${NS_W}">
  <w:body>
    <w:tbl>
      <w:tr>
        <w:tc><w:p><w:r><w:t>Cell 1</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>Cell 2</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>
  </w:body>
</w:document>`;

    const doc = parseXml(docXml);
    const paras = getWordElements(doc, 'p');

    const check = validateParagraphBoundaryMutation(paras[0], 'Joined Text', {
        targetEndParagraph: paras[1]
    });

    assert.equal(check.valid, false);
    assert.equal(check.code, 'UNSAFE_PARAGRAPH_BOUNDARY');
    assert.match(check.message, /different table cells/i);
});

test('WP-10: Paragraph Boundary Removal (Join) - merges two paragraphs with tracked boundary deletion', async () => {
    const docXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${NS_W}">
  <w:body>
    <w:p><w:r><w:t>First sentence.</w:t></w:r></w:p>
    <w:p><w:r><w:t>Second sentence.</w:t></w:r></w:p>
  </w:body>
</w:document>`;

    const result = await applyOperationsToDocumentXml(docXml, [{
        type: 'redline',
        target: 'First sentence.\nSecond sentence.',
        modified: 'First sentence. Second sentence.'
    }], 'TestAuthor');

    assert.equal(result.hasChanges, true);
    assert.equal(result.results[0].status, 'applied');

    // Section 1: Check acceptTrackedChangesInOoxml
    const acceptedResult = acceptTrackedChangesInOoxml(result.documentXml, { allAuthors: true });
    const acceptedDoc = parseXml(acceptedResult.oxml);
    const acceptedParas = getWordElements(acceptedDoc, 'p');
    assert.equal(acceptedParas.length, 1, 'Acceptance must merge both paragraphs into 1');
    assert.equal(acceptedParas[0].textContent.trim(), 'First sentence. Second sentence.');

    // Section 2: Check rejectTrackedChangesInOoxml
    const rejectedResult = rejectTrackedChangesInOoxml(result.documentXml, { allAuthors: true });
    const rejectedDoc = parseXml(rejectedResult.oxml);
    const rejectedParas = getWordElements(rejectedDoc, 'p');
    assert.equal(rejectedParas.length, 2, 'Rejection must restore 2 separate paragraphs');
    assert.equal(rejectedParas[0].textContent.trim(), 'First sentence.');
    assert.equal(rejectedParas[1].textContent.trim(), 'Second sentence.');
});

test('WP-10: List Item Boundary - split preserves list numbering properties', async () => {
    const docXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${NS_W}">
  <w:body>
    <w:p>
      <w:pPr>
        <w:numPr>
          <w:ilvl w:val="0"/>
          <w:numId w:val="5"/>
        </w:numPr>
      </w:pPr>
      <w:r><w:t>List item alpha and beta.</w:t></w:r>
    </w:p>
  </w:body>
</w:document>`;

    const result = await applyOperationsToDocumentXml(docXml, [{
        type: 'redline',
        target: 'List item alpha and beta.',
        modified: 'List item alpha\nand beta.'
    }], 'TestAuthor');

    assert.equal(result.hasChanges, true);
    assert.equal(result.results[0].status, 'applied');

    const acceptedResult = acceptTrackedChangesInOoxml(result.documentXml, { allAuthors: true });
    const acceptedDoc = parseXml(acceptedResult.oxml);
    const acceptedParas = getWordElements(acceptedDoc, 'p');
    assert.equal(acceptedParas.length, 2, 'Accepted document must have 2 list paragraphs');

    for (const p of acceptedParas) {
        const numPr = getWordElements(p, 'numPr')[0];
        assert.ok(numPr, 'Each split paragraph must retain list properties (w:numPr)');
        const numId = getWordElements(numPr, 'numId')[0];
        assert.equal(numId.getAttribute('w:val'), '5', 'Must preserve numId=5');
    }
});

