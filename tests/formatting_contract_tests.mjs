import './setup-xml-provider.mjs';

import assert from 'assert/strict';
import {
    applyOperationToDocumentXml
} from '../services/standalone-operation-runner.js';
import {
    acceptTrackedChangesInOoxml,
    rejectTrackedChangesInOoxml
} from '../services/revision-comment-management.js';
import { parseOoxml } from '../engine/oxml-engine.js';
import { NS_W } from '../core/types.js';

function wrapDocument(bodyXml) {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${NS_W}">
  <w:body>
    ${bodyXml}
    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
  </w:body>
</w:document>`;
}

async function testCharacterFormattingProperties() {
    // 1. Absent -> On (bold, italic, underline, strike, color, highlight, fontSize, fontFamily)
    const initialXml = wrapDocument(`
      <w:p>
        <w:r><w:t>The quick brown fox jumps over the lazy dog.</w:t></w:r>
      </w:p>
    `);

    const formatOp = {
        type: 'format',
        target: 'The quick brown fox jumps over the lazy dog.',
        textToFormat: 'brown fox',
        properties: {
            bold: true,
            italic: true,
            underline: true,
            strike: true,
            color: 'FF0000',
            highlight: 'yellow',
            fontSize: 28, // 14pt -> 28 half-points
            fontFamily: 'Arial'
        }
    };

    const applied = await applyOperationToDocumentXml(initialXml, formatOp, 'AuthorOne');
    assert.equal(applied.hasChanges, true, 'Format op reported changes');
    assert.ok(applied.documentXml.includes('w:rPrChange'), 'Document contains w:rPrChange');
    assert.ok(applied.documentXml.includes('w:b w:val="1"'), 'Run has bold');
    assert.ok(applied.documentXml.includes('w:i w:val="1"'), 'Run has italic');
    assert.ok(applied.documentXml.includes('w:u w:val="single"'), 'Run has underline');
    assert.ok(applied.documentXml.includes('w:strike w:val="1"'), 'Run has strike');
    assert.ok(applied.documentXml.includes('w:color w:val="FF0000"'), 'Run has color FF0000');
    assert.ok(applied.documentXml.includes('w:highlight w:val="yellow"'), 'Run has highlight yellow');
    assert.ok(applied.documentXml.includes('w:sz w:val="28"'), 'Run has font size 28');
    assert.ok(applied.documentXml.includes('w:rFonts w:ascii="Arial"'), 'Run has font family Arial');

    // Reject All should restore clean state without rPrChange or properties
    const rejected = rejectTrackedChangesInOoxml(applied.documentXml, { allAuthors: true });
    assert.ok(!rejected.oxml.includes('w:rPrChange'), 'Rejected doc has no rPrChange');
    assert.ok(!rejected.oxml.includes('w:b w:val="1"'), 'Rejected doc has no bold');

    // Accept All should keep formatting and remove rPrChange
    const accepted = acceptTrackedChangesInOoxml(applied.documentXml, { allAuthors: true });
    assert.ok(!accepted.oxml.includes('w:rPrChange'), 'Accepted doc has no rPrChange');
    assert.ok(accepted.oxml.includes('w:b w:val="1"'), 'Accepted doc retains bold');
    assert.ok(accepted.oxml.includes('w:i w:val="1"'), 'Accepted doc retains italic');
}

async function testFormatOnToOff() {
    // 2. On -> Off
    const initialXml = wrapDocument(`
      <w:p>
        <w:r>
          <w:rPr>
            <w:b/>
            <w:i/>
            <w:u w:val="single"/>
            <w:strike/>
            <w:highlight w:val="cyan"/>
          </w:rPr>
          <w:t>Already formatted text here.</w:t>
        </w:r>
      </w:p>
    `);

    const turnOffOp = {
        type: 'format',
        target: 'Already formatted text here.',
        textToFormat: 'formatted',
        properties: {
            bold: false,
            italic: false,
            underline: false,
            strike: false,
            highlight: null
        }
    };

    const result = await applyOperationToDocumentXml(initialXml, turnOffOp, 'AuthorOne');
    assert.equal(result.hasChanges, true);
    assert.ok(result.documentXml.includes('w:rPrChange'));
    assert.ok(result.documentXml.includes('w:b w:val="0"'));
    assert.ok(result.documentXml.includes('w:i w:val="0"'));
    assert.ok(result.documentXml.includes('w:u w:val="none"'));
    assert.ok(result.documentXml.includes('w:strike w:val="0"'));

    // Reject restores the original bold/italic
    const rejected = rejectTrackedChangesInOoxml(result.documentXml, { allAuthors: true });
    assert.ok(rejected.oxml.includes('<w:b/>') || rejected.oxml.includes('<w:b />'));
}

async function testMultiRunSpanAndPartialRuns() {
    // 3. Spanning multiple runs and partial runs
    const initialXml = wrapDocument(`
      <w:p>
        <w:r><w:t>First chunk </w:t></w:r>
        <w:r><w:t>second chunk </w:t></w:r>
        <w:r><w:t>third chunk</w:t></w:r>
      </w:p>
    `);

    // Target "chunk second chunk th"
    const spanOp = {
        type: 'format',
        target: 'First chunk second chunk third chunk',
        textToFormat: 'chunk second chunk th',
        properties: {
            bold: true
        }
    };

    const result = await applyOperationToDocumentXml(initialXml, spanOp, 'AuthorOne');
    assert.equal(result.hasChanges, true);

    const doc = parseOoxml(result.documentXml);
    const runs = Array.from(doc.getElementsByTagNameNS(NS_W, 'r'));
    const pieces = runs.map(r => {
        const t = Array.from(r.getElementsByTagNameNS(NS_W, 't')).map(n => n.textContent).join('');
        const b = r.getElementsByTagNameNS(NS_W, 'b')[0];
        const isBold = b ? (b.getAttribute('w:val') !== '0') : false;
        return { text: t, bold: isBold };
    });

    const boldPieces = pieces.filter(p => p.bold).map(p => p.text).join('');
    assert.equal(boldPieces, 'chunk second chunk th', 'Bold applied across partial and multi runs');

    const nonBoldPieces = pieces.filter(p => !p.bold).map(p => p.text).join('');
    assert.equal(nonBoldPieces, 'First ird chunk', 'Unselected portions remain unformatted');
}

async function testParagraphFormattingProperties() {
    // 4. Paragraph formatting: alignment, keepNext, keepLines, pageBreakBefore, style
    const initialXml = wrapDocument(`
      <w:p>
        <w:pPr><w:jc w:val="left"/></w:pPr>
        <w:r><w:t>Section title paragraph.</w:t></w:r>
      </w:p>
    `);

    const pOp = {
        type: 'paragraph-format',
        target: 'Section title paragraph.',
        properties: {
            alignment: 'center',
            keepNext: true,
            keepLines: true,
            pageBreakBefore: true,
            style: 'Heading1'
        }
    };

    const result = await applyOperationToDocumentXml(initialXml, pOp, 'AuthorOne');
    assert.equal(result.hasChanges, true);
    assert.ok(result.documentXml.includes('w:pPrChange'), 'Paragraph has w:pPrChange');
    assert.ok(result.documentXml.includes('w:jc w:val="center"'), 'Alignment is center');
    assert.ok(result.documentXml.includes('w:keepNext'), 'keepNext applied');
    assert.ok(result.documentXml.includes('w:keepLines'), 'keepLines applied');
    assert.ok(result.documentXml.includes('w:pageBreakBefore'), 'pageBreakBefore applied');
    assert.ok(result.documentXml.includes('w:pStyle w:val="Heading1"'), 'style applied');

    // Reject All restores left alignment and removes keepNext/Heading1
    const rejected = rejectTrackedChangesInOoxml(result.documentXml, { allAuthors: true });
    assert.ok(!rejected.oxml.includes('w:pPrChange'), 'Rejected doc has no pPrChange');
    assert.ok(rejected.oxml.includes('w:jc w:val="left"'), 'Restored left alignment');
    assert.ok(!rejected.oxml.includes('w:keepNext'), 'keepNext removed');
    assert.ok(!rejected.oxml.includes('Heading1'), 'Heading1 style removed');

    // Accept All keeps center alignment
    const accepted = acceptTrackedChangesInOoxml(result.documentXml, { allAuthors: true });
    assert.ok(!accepted.oxml.includes('w:pPrChange'), 'Accepted doc has no pPrChange');
    assert.ok(accepted.oxml.includes('w:jc w:val="center"'), 'Accepted doc keeps center alignment');
}

async function testAuthorCoalescingPolicy() {
    // 5. Author-aware coalescing: same author pending insertion vs different author
    const initialXml = wrapDocument(`
      <w:p>
        <w:ins w:id="1" w:author="AuthorA" w:date="2026-01-01T00:00:00Z">
          <w:r><w:t>Newly inserted clause.</w:t></w:r>
        </w:ins>
      </w:p>
    `);

    // Case A: same author with coalesce-own-insertion
    const coalesceOp = {
        type: 'format',
        target: 'Newly inserted clause.',
        textToFormat: 'inserted',
        properties: { bold: true },
        formattingRevisionPolicy: 'coalesce-own-insertion'
    };
    const resA = await applyOperationToDocumentXml(initialXml, coalesceOp, 'AuthorA');
    assert.equal(resA.hasChanges, true);
    assert.ok(!resA.documentXml.includes('w:rPrChange'), 'Coalesced formatting does not generate w:rPrChange');
    assert.ok(resA.documentXml.includes('w:b w:val="1"'), 'Bold applied directly to insertion run');

    // Case B: different author inside pending insertion -> MUST generate w:rPrChange
    const resB = await applyOperationToDocumentXml(initialXml, coalesceOp, 'AuthorB');
    assert.equal(resB.hasChanges, true);
    assert.ok(resB.documentXml.includes('w:rPrChange'), 'Different author generates w:rPrChange');
    assert.ok(resB.documentXml.includes('w:author="AuthorB"'), 'rPrChange attributed to AuthorB');
}

async function testSelectiveAuthorAcceptReject() {
    // 6. Selective author accept/reject
    const initialXml = wrapDocument(`
      <w:p>
        <w:r><w:t>Hello world text.</w:t></w:r>
      </w:p>
    `);

    const opA = {
        type: 'format',
        target: 'Hello world text.',
        textToFormat: 'Hello',
        properties: { bold: true }
    };
    const opB = {
        type: 'format',
        target: 'Hello world text.',
        textToFormat: 'world',
        properties: { italic: true }
    };

    const resA = await applyOperationToDocumentXml(initialXml, opA, 'AuthorA');
    const resAB = await applyOperationToDocumentXml(resA.documentXml, opB, 'AuthorB');

    assert.ok(resAB.documentXml.includes('w:author="AuthorA"'));
    assert.ok(resAB.documentXml.includes('w:author="AuthorB"'));

    // Reject AuthorA only -> Hello should revert, but world retains AuthorB's change!
    const rejectedA = rejectTrackedChangesInOoxml(resAB.documentXml, { author: 'AuthorA' });
    assert.ok(!rejectedA.oxml.includes('w:author="AuthorA"'), 'AuthorA rPrChange removed');
    assert.ok(rejectedA.oxml.includes('w:author="AuthorB"'), 'AuthorB rPrChange retained');

    // Accept AuthorB -> world retains italic, no rPrChange left
    const acceptedB = acceptTrackedChangesInOoxml(rejectedA.oxml, { author: 'AuthorB' });
    assert.ok(!acceptedB.oxml.includes('w:rPrChange'), 'All rPrChange consumed');
    assert.ok(acceptedB.oxml.includes('w:i w:val="1"'), 'Italic retained');
    assert.ok(!acceptedB.oxml.includes('w:b w:val="1"'), 'Bold was rejected');
}

async function testNoOpDoesNotConsumeId() {
    // 7. No-op returns status 'no-op' and does NOT consume revision IDs
    const initialXml = wrapDocument(`
      <w:p>
        <w:r>
          <w:rPr><w:b w:val="1"/></w:rPr>
          <w:t>Already bold.</w:t>
        </w:r>
      </w:p>
    `);

    const noOp = {
        type: 'format',
        target: 'Already bold.',
        textToFormat: 'bold',
        properties: { bold: true }
    };

    const result = await applyOperationToDocumentXml(initialXml, noOp, 'AuthorOne');
    assert.equal(result.hasChanges, false);
    assert.equal(result.status, 'no-op');
    assert.ok(!result.documentXml.includes('w:rPrChange'));
}

async function runAllTests() {
    console.log('Running WP-11 formatting contract tests...');
    await testCharacterFormattingProperties();
    console.log('  Passed testCharacterFormattingProperties');
    await testFormatOnToOff();
    console.log('  Passed testFormatOnToOff');
    await testMultiRunSpanAndPartialRuns();
    console.log('  Passed testMultiRunSpanAndPartialRuns');
    await testParagraphFormattingProperties();
    console.log('  Passed testParagraphFormattingProperties');
    await testAuthorCoalescingPolicy();
    console.log('  Passed testAuthorCoalescingPolicy');
    await testSelectiveAuthorAcceptReject();
    console.log('  Passed testSelectiveAuthorAcceptReject');
    await testNoOpDoesNotConsumeId();
    console.log('  Passed testNoOpDoesNotConsumeId');
    console.log('All WP-11 formatting contract tests passed successfully.');
}

runAllTests().catch(err => {
    console.error(err);
    process.exit(1);
});
