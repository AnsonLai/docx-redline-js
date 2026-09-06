import assert from 'node:assert/strict';
import './setup-xml-provider.mjs';
import { applyRedlineToOxml, ingestWordOoxmlToPlainText } from '../index.js';
import { validateRedlineOoxml } from '../core/redline-validation.js';
import {
    acceptTrackedChangesInOoxml,
    rejectTrackedChangesInOoxml
} from '../services/revision-comment-management.js';
import { parseOoxmlSafe, createSerializer } from '../adapters/xml-adapter.js';
import { applySurgicalMode } from '../engine/surgical-mode.js';
import {
    applyOperationsToDocumentXml
} from '../services/standalone-operation-runner.js';
import { validateDocumentOperation } from '../services/document-operation-contract.js';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

function getElements(root, tagName) {
    return Array.from(root.getElementsByTagNameNS(W, tagName));
}

// Test 1: Plain run replacement receives identical author/date and distinct IDs
{
    const oxml = `<w:p xmlns:w="${W}"><w:r><w:t>The quick brown fox jumps.</w:t></w:r></w:p>`;
    const origText = 'The quick brown fox jumps.';
    const modText = 'The quick red fox jumps.';

    const result = await applyRedlineToOxml(oxml, origText, modText, {
        generateRedlines: true,
        author: 'Alice',
        pairReplacements: true
    });

    assert.equal(result.hasChanges, true);
    assert.equal(result.status, 'ok');

    const validation = validateRedlineOoxml(result.oxml);
    assert.equal(validation.valid, true, `Validation failed: ${JSON.stringify(validation.errors)}`);

    const { doc } = parseOoxmlSafe(result.oxml);
    const dels = getElements(doc, 'del');
    const inss = getElements(doc, 'ins');

    assert.equal(dels.length, 1, 'Expected exactly one w:del');
    assert.equal(inss.length, 1, 'Expected exactly one w:ins');

    const del = dels[0];
    const ins = inss[0];

    const delAuthor = del.getAttributeNS(W, 'author') || del.getAttribute('w:author');
    const insAuthor = ins.getAttributeNS(W, 'author') || ins.getAttribute('w:author');
    assert.equal(delAuthor, 'Alice');
    assert.equal(insAuthor, 'Alice');

    const delDate = del.getAttributeNS(W, 'date') || del.getAttribute('w:date');
    const insDate = ins.getAttributeNS(W, 'date') || ins.getAttribute('w:date');
    assert.ok(delDate, 'Expected del date to exist');
    assert.ok(insDate, 'Expected ins date to exist');
    assert.equal(delDate, insDate, 'Expected identical date timestamp on paired del/ins');

    const delId = del.getAttributeNS(W, 'id') || del.getAttribute('w:id');
    const insId = ins.getAttributeNS(W, 'id') || ins.getAttribute('w:id');
    assert.ok(delId, 'Expected del ID');
    assert.ok(insId, 'Expected ins ID');
    assert.notEqual(delId, insId, 'Paired del and ins must have distinct revision IDs');
}

// Test 2: Accepted/rejected output matches independent mode
{
    const oxml = `<w:p xmlns:w="${W}"><w:r><w:t>The quick brown fox jumps.</w:t></w:r></w:p>`;
    const origText = 'The quick brown fox jumps.';
    const modText = 'The quick red fox jumps.';

    const pairedResult = await applyRedlineToOxml(oxml, origText, modText, {
        generateRedlines: true,
        author: 'Alice',
        pairReplacements: true
    });
    const independentResult = await applyRedlineToOxml(oxml, origText, modText, {
        generateRedlines: true,
        author: 'Alice',
        pairReplacements: false
    });

    const acceptedPaired = ingestWordOoxmlToPlainText(acceptTrackedChangesInOoxml(pairedResult.oxml, { allAuthors: true }).oxml);
    const acceptedIndep = ingestWordOoxmlToPlainText(acceptTrackedChangesInOoxml(independentResult.oxml, { allAuthors: true }).oxml);
    assert.equal(acceptedPaired.trim(), 'The quick red fox jumps.');
    assert.equal(acceptedPaired, acceptedIndep);

    const rejectedPaired = ingestWordOoxmlToPlainText(rejectTrackedChangesInOoxml(pairedResult.oxml, { allAuthors: true }).oxml);
    const rejectedIndep = ingestWordOoxmlToPlainText(rejectTrackedChangesInOoxml(independentResult.oxml, { allAuthors: true }).oxml);
    assert.equal(rejectedPaired.trim(), 'The quick brown fox jumps.');
    assert.equal(rejectedPaired, rejectedIndep);
}

// Test 3: Replacement across two formatting runs retains both formats and shares author/date
{
    const oxml = `<w:p xmlns:w="${W}"><w:r><w:rPr><w:b/></w:rPr><w:t>bold </w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t>italic</w:t></w:r><w:r><w:t> tail</w:t></w:r></w:p>`;
    const origText = 'bold italic tail';
    const modText = 'combined tail';

    const result = await applyRedlineToOxml(oxml, origText, modText, {
        generateRedlines: true,
        author: 'Formatter',
        pairReplacements: true
    });

    assert.equal(result.hasChanges, true);
    const validation = validateRedlineOoxml(result.oxml);
    assert.equal(validation.valid, true, `Validation failed: ${JSON.stringify(validation.errors)}`);

    const { doc } = parseOoxmlSafe(result.oxml);
    const dels = getElements(doc, 'del');
    const inss = getElements(doc, 'ins');

    // Should have 2 del runs (one for bold, one for italic)
    assert.equal(dels.length, 2, 'Expected 2 del runs across formatting boundary');
    assert.equal(inss.length, 1, 'Expected 1 ins run');

    // Confirm distinct IDs across all revision elements
    const ids = new Set();
    for (const d of dels) {
        const id = d.getAttributeNS(W, 'id') || d.getAttribute('w:id');
        assert.ok(!ids.has(id), `Duplicate revision ID found: ${id}`);
        ids.add(id);
    }
    for (const i of inss) {
        const id = i.getAttributeNS(W, 'id') || i.getAttribute('w:id');
        assert.ok(!ids.has(id), `Duplicate revision ID found: ${id}`);
        ids.add(id);
    }

    // Confirm all dels and ins share the exact same author and date
    const insDate = inss[0].getAttributeNS(W, 'date') || inss[0].getAttribute('w:date');
    for (const d of dels) {
        const dAuthor = d.getAttributeNS(W, 'author') || d.getAttribute('w:author');
        const dDate = d.getAttributeNS(W, 'date') || d.getAttribute('w:date');
        assert.equal(dAuthor, 'Formatter');
        assert.equal(dDate, insDate, 'All del pieces in replacement should share date with ins');
    }

    // Verify format preservation on deleted runs
    const del1Runs = getElements(dels[0], 'r');
    const del2Runs = getElements(dels[1], 'r');
    assert.ok(getElements(del1Runs[0], 'b').length > 0, 'First del should retain bold');
    assert.ok(getElements(del2Runs[0], 'i').length > 0, 'Second del should retain italic');
}

// Test 4: Structural boundaries trigger safe fallback and emit PAIRING_SKIPPED_STRUCTURAL_BOUNDARY
{
    const serializer = createSerializer();

    // Test 4a: Hyperlink boundary (run inside hyperlink)
    const oxmlH = `<w:p xmlns:w="${W}"><w:hyperlink r:id="rId1" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:r><w:t>Link text</w:t></w:r></w:hyperlink><w:r><w:t xml:space="preserve"> after</w:t></w:r></w:p>`;
    const { doc: docH } = parseOoxmlSafe(oxmlH);
    const resultH = applySurgicalMode(docH, 'Link text after', 'New link after', serializer, 'Editor', [], true, null, {}, { pairReplacements: true });

    assert.equal(resultH.hasChanges, true);
    assert.ok(resultH.warnings?.includes('PAIRING_SKIPPED_STRUCTURAL_BOUNDARY'),
        `Expected PAIRING_SKIPPED_STRUCTURAL_BOUNDARY warning, got: ${JSON.stringify(resultH.warnings)}`);
    const valH = validateRedlineOoxml(resultH.oxml);
    assert.equal(valH.valid, true);

    // Test 4b: Bookmark boundary (run adjacent to bookmarkStart)
    const oxmlB = `<w:p xmlns:w="${W}"><w:bookmarkStart w:id="0" w:name="_bm1"/><w:r><w:t>Bookmarked text</w:t></w:r><w:bookmarkEnd w:id="0"/></w:p>`;
    const { doc: docB } = parseOoxmlSafe(oxmlB);
    const resultB = applySurgicalMode(docB, 'Bookmarked text', 'Changed bookmark', serializer, 'Editor', [], true, null, {}, { pairReplacements: true });

    assert.equal(resultB.hasChanges, true);
    assert.ok(resultB.warnings?.includes('PAIRING_SKIPPED_STRUCTURAL_BOUNDARY'),
        `Expected PAIRING_SKIPPED_STRUCTURAL_BOUNDARY warning, got: ${JSON.stringify(resultB.warnings)}`);
    const valB = validateRedlineOoxml(resultB.oxml);
    assert.equal(valB.valid, true);

    // Test 4c: Comment marker boundary (run adjacent to commentRangeStart)
    const oxmlC = `<w:p xmlns:w="${W}"><w:commentRangeStart w:id="1"/><w:r><w:t>Commented text</w:t></w:r><w:commentRangeEnd w:id="1"/></w:p>`;
    const { doc: docC } = parseOoxmlSafe(oxmlC);
    const resultC = applySurgicalMode(docC, 'Commented text', 'Replaced comment', serializer, 'Editor', [], true, null, {}, { pairReplacements: true });

    assert.equal(resultC.hasChanges, true);
    assert.ok(resultC.warnings?.includes('PAIRING_SKIPPED_STRUCTURAL_BOUNDARY'),
        `Expected PAIRING_SKIPPED_STRUCTURAL_BOUNDARY warning, got: ${JSON.stringify(resultC.warnings)}`);
    const valC = validateRedlineOoxml(resultC.oxml);
    assert.equal(valC.valid, true);

    // Test 4d: Field char boundary
    const oxmlF = `<w:p xmlns:w="${W}"><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:t>Field text</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>`;
    const { doc: docF } = parseOoxmlSafe(oxmlF);
    const resultF = applySurgicalMode(docF, 'Field text', 'Replaced field', serializer, 'Editor', [], true, null, {}, { pairReplacements: true });

    assert.equal(resultF.hasChanges, true);
    assert.ok(resultF.warnings?.includes('PAIRING_SKIPPED_STRUCTURAL_BOUNDARY'),
        `Expected PAIRING_SKIPPED_STRUCTURAL_BOUNDARY warning, got: ${JSON.stringify(resultF.warnings)}`);
    const valF = validateRedlineOoxml(resultF.oxml);
    assert.equal(valF.valid, true);
}

// Test 5: Multiple replacements in one operation receive separate event pairs
{
    const oxml = `<w:p xmlns:w="${W}"><w:r><w:t>The quick brown fox jumps over the lazy dog.</w:t></w:r></w:p>`;
    const origText = 'The quick brown fox jumps over the lazy dog.';
    const modText = 'A quick red fox jumps over the sleepy dog.';

    const result = await applyRedlineToOxml(oxml, origText, modText, {
        generateRedlines: true,
        author: 'MultiEditor',
        pairReplacements: true
    });

    assert.equal(result.hasChanges, true);
    const validation = validateRedlineOoxml(result.oxml);
    assert.equal(validation.valid, true);

    const { doc } = parseOoxmlSafe(result.oxml);
    const dels = getElements(doc, 'del');
    const inss = getElements(doc, 'ins');

    // 3 replacements: quick->fast, brown->red, lazy->sleepy
    assert.equal(dels.length, 3, 'Expected 3 del operations');
    assert.equal(inss.length, 3, 'Expected 3 ins operations');

    // Verify all revision IDs are unique
    const allIds = new Set();
    for (const el of [...dels, ...inss]) {
        const id = el.getAttributeNS(W, 'id') || el.getAttribute('w:id');
        assert.ok(!allIds.has(id), `Duplicate revision ID: ${id}`);
        allIds.add(id);
    }
    assert.equal(allIds.size, 6, 'Expected 6 distinct revision IDs for 3 pairs');
}

// Test 6: Validation rejects duplicate revision IDs
{
    const badOxml = `<w:p xmlns:w="${W}"><w:del w:id="42" w:author="Alice" w:date="2026-09-05T12:00:00Z"><w:r><w:delText>old</w:delText></w:r></w:del><w:ins w:id="42" w:author="Alice" w:date="2026-09-05T12:00:00Z"><w:r><w:t>new</w:t></w:r></w:ins></w:p>`;
    const validation = validateRedlineOoxml(badOxml);
    assert.equal(validation.valid, false);
    assert.ok(validation.issues.some(issue => issue.code === 'DUPLICATE_REVISION_ID' || issue.message.includes('Duplicate revision')));
}

// Test 7: Integration with Standalone Operation Runner and Contract Validation
{
    // Test contract validation
    const validOp = {
        type: 'replace',
        target: 'Target paragraph',
        modified: 'Replaced paragraph',
        pairReplacements: true
    };
    assert.equal(validateDocumentOperation(validOp).valid, true);

    const invalidOp = {
        type: 'replace',
        target: 'Target paragraph',
        modified: 'Replaced paragraph',
        pairReplacements: 'not-a-boolean'
    };
    assert.equal(validateDocumentOperation(invalidOp).valid, false);

    // Test runner execution with pairReplacements
    const docXml = `<w:document xmlns:w="${W}"><w:body><w:p><w:r><w:t>Alpha beta gamma</w:t></w:r></w:p></w:body></w:document>`;
    const runResult = await applyOperationsToDocumentXml(docXml, [
        {
            type: 'replace',
            target: 'Alpha beta gamma',
            modified: 'Alpha delta gamma',
            pairReplacements: true
        }
    ], 'BatchAuthor');

    assert.equal(runResult.hasChanges, true);
    assert.equal(runResult.results[0].status, 'applied');

    const { doc } = parseOoxmlSafe(runResult.documentXml);
    const dels = getElements(doc, 'del');
    const inss = getElements(doc, 'ins');

    assert.equal(dels.length, 1);
    assert.equal(inss.length, 1);

    const delDate = dels[0].getAttributeNS(W, 'date') || dels[0].getAttribute('w:date');
    const insDate = inss[0].getAttributeNS(W, 'date') || inss[0].getAttribute('w:date');
    assert.equal(delDate, insDate, 'Batch runner pairReplacements should share timestamp');
}

console.log('All WP-07 replacement event metadata tests passed successfully!');
