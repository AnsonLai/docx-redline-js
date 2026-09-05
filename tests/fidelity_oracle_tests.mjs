import assert from 'node:assert/strict';
import { buildZip } from '../scripts/lib/minimal-zip.mjs';
import { openDocx } from '../node/index.js';
import {
    MUTATION_ENVELOPES,
    verifyMutationFidelity
} from './helpers/mutation-envelopes.mjs';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';

// Base document components
const contentTypes = `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
  <Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
  <Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>
  <Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>`.trim();

const rels = `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="${R}/header" Target="header1.xml"/>
  <Relationship Id="rId2" Type="${R}/footer" Target="footer1.xml"/>
  <Relationship Id="rId3" Type="${R}/footnotes" Target="footnotes.xml"/>
  <Relationship Id="rId4" Type="${R}/endnotes" Target="endnotes.xml"/>
  <Relationship Id="rId5" Type="${R}/styles" Target="styles.xml"/>
  <Relationship Id="rId6" Type="${R}/image" Target="media/image1.png"/>
</Relationships>`.trim();

const headerXml = `<w:hdr xmlns:w="${W}"><w:p><w:r><w:t>Header Text</w:t></w:r></w:p></w:hdr>`;
const footerXml = `<w:ftr xmlns:w="${W}"><w:p><w:r><w:t>Footer Text</w:t></w:r></w:p></w:ftr>`;
const footnotesXml = `<w:footnotes xmlns:w="${W}"><w:footnote w:id="1"><w:p><w:r><w:t>Note 1</w:t></w:r></w:p></w:footnote></w:footnotes>`;
const endnotesXml = `<w:endnotes xmlns:w="${W}"><w:endnote w:id="1"><w:p><w:r><w:t>Endnote 1</w:t></w:r></w:p></w:endnote></w:endnotes>`;
const stylesXml = `<w:styles xmlns:w="${W}"><w:docDefaults><w:rPrDefault/></w:docDefaults></w:styles>`;
const binaryImage = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x01, 0x02, 0x03]);

// Multi-feature documentXml containing:
// P1: Normal intro paragraph
// P2: Paragraph with hyperlink and bookmark
// P3: Target paragraph for surgical text replacement
// P4: Paragraph with field codes and footnote reference
// P5: SDT containing content
// P6: Target paragraph for comment insertion
// P7: Existing table
// P8: Target paragraph for list insertion
// P9: Normal concluding paragraph
const documentXml = `
<w:document xmlns:w="${W}" xmlns:w14="${W14}" xmlns:r="${R}" xmlns:mc="${MC}" mc:Ignorable="w14">
  <w:body>
    <w:p w14:paraId="P0000001">
      <w:r><w:t>P1: Normal untouched introductory paragraph.</w:t></w:r>
    </w:p>
    <w:p w14:paraId="P0000002">
      <w:bookmarkStart w:id="10" w:name="BookmarkSection"/>
      <w:hyperlink r:id="rId1">
        <w:r><w:t>P2: Link text inside hyperlink</w:t></w:r>
      </w:hyperlink>
      <w:bookmarkEnd w:id="10"/>
    </w:p>
    <w:p w14:paraId="P0000003">
      <w:r><w:t>P3: Target for surgical replacement text.</w:t></w:r>
    </w:p>
    <w:p w14:paraId="P0000004">
      <w:fldSimple w:instr="PAGE"/>
      <w:r>
        <w:t>P4: Text with footnote </w:t>
        <w:footnoteReference w:id="1"/>
      </w:r>
    </w:p>
    <w:sdt>
      <w:sdtContent>
        <w:p w14:paraId="P0000005">
          <w:r><w:t>P5: Content inside Structured Document Tag.</w:t></w:r>
        </w:p>
      </w:sdtContent>
    </w:sdt>
    <w:p w14:paraId="P0000006">
      <w:r><w:t>P6: Target for comment injection.</w:t></w:r>
    </w:p>
    <w:tbl>
      <w:tr>
        <w:tc>
          <w:p w14:paraId="P0000007">
            <w:r><w:t>P7: Table cell text</w:t></w:r>
          </w:p>
        </w:tc>
      </w:tr>
    </w:tbl>
    <w:p w14:paraId="P0000008">
      <w:r><w:t>P8: Target to be converted to list.</w:t></w:r>
    </w:p>
    <w:p w14:paraId="P0000009">
      <w:r><w:t>P9: Concluding untouched paragraph.</w:t></w:r>
    </w:p>
    <w:sectPr/>
  </w:body>
</w:document>`.trim();

function createMasterPackage() {
    return buildZip([
        { name: '[Content_Types].xml', data: contentTypes },
        { name: 'word/_rels/document.xml.rels', data: rels },
        { name: 'word/document.xml', data: documentXml },
        { name: 'word/header1.xml', data: headerXml },
        { name: 'word/footer1.xml', data: footerXml },
        { name: 'word/footnotes.xml', data: footnotesXml },
        { name: 'word/endnotes.xml', data: endnotesXml },
        { name: 'word/styles.xml', data: stylesXml },
        { name: 'word/media/image1.png', data: binaryImage }
    ]);
}

// =========================================================================
// Test 1: Surgical text replacement fidelity
// =========================================================================
{
    const inputZip = createMasterPackage();
    const doc = openDocx(inputZip);

    const result = await doc.applyOperations([
        {
            type: 'redline',
            target: { paragraphId: 'P0000003' },
            modified: 'P3: Target for surgical replacement text with updated revision.',
            author: 'Editor'
        }
    ], { atomic: true });

    assert.equal(result.written, true);

    const fidelity = verifyMutationFidelity(inputZip, result.toBuffer(), {
        envelope: MUTATION_ENVELOPES.surgical_text,
        targetedParagraphIds: ['P0000003'],
        lifecycleCheck: true
    });

    // Verify untouched paragraphs, tables, SDTs, headers, footers, media
    assert.ok(fidelity.untouchedSubtreesChecked >= 6);
    console.log(`  ✓ Surgical text fidelity: ${fidelity.untouchedSubtreesChecked} untouched subtrees verified identical`);
}

// =========================================================================
// Test 2: Comment injection fidelity
// =========================================================================
{
    const inputZip = createMasterPackage();
    const doc = openDocx(inputZip);

    const result = await doc.applyOperations([
        {
            type: 'comment',
            target: { paragraphId: 'P0000006' },
            textToComment: 'comment injection',
            commentContent: 'Review note for legal compliance.',
            author: 'Reviewer'
        }
    ], { atomic: true });

    assert.equal(result.written, true);

    const fidelity = verifyMutationFidelity(inputZip, result.toBuffer(), {
        envelope: MUTATION_ENVELOPES.comment,
        targetedParagraphIds: ['P0000006'],
        lifecycleCheck: true
    });

    assert.ok(fidelity.untouchedSubtreesChecked >= 6);
    console.log(`  ✓ Comment injection fidelity: ${fidelity.untouchedSubtreesChecked} untouched subtrees verified identical`);
}

// =========================================================================
// Test 3: List generation fidelity (auxiliary numbering.xml created)
// =========================================================================
{
    const inputZip = createMasterPackage();
    const doc = openDocx(inputZip);

    const result = await doc.applyOperations([
        {
            type: 'redline',
            target: { paragraphId: 'P0000008' },
            modified: '1. First list item\n2. Second list item',
            author: 'Editor'
        }
    ], { atomic: true });

    assert.equal(result.written, true);

    const fidelity = verifyMutationFidelity(inputZip, result.toBuffer(), {
        envelope: MUTATION_ENVELOPES.list,
        targetedParagraphIds: ['P0000008'],
        lifecycleCheck: true
    });

    assert.ok(fidelity.untouchedSubtreesChecked >= 6);
    console.log(`  ✓ List generation fidelity: ${fidelity.untouchedSubtreesChecked} untouched subtrees verified identical`);
}

// =========================================================================
// Test 4: Hyperlink, Bookmark, and Field Preservation
// =========================================================================
{
    // Ensure P2 (hyperlink + bookmark) and P4 (field code + footnote) are 100% bit-identical
    // when an edit is applied to adjacent P3.
    const inputZip = createMasterPackage();
    const doc = openDocx(inputZip);

    const result = await doc.applyOperations([
        {
            type: 'redline',
            target: { paragraphId: 'P0000003' },
            modified: 'P3: Slightly changed text',
            author: 'Editor'
        }
    ], { atomic: true });

    const fidelity = verifyMutationFidelity(inputZip, result.toBuffer(), {
        envelope: MUTATION_ENVELOPES.surgical_text,
        targetedParagraphIds: ['P0000003'],
        lifecycleCheck: true
    });

    assert.ok(fidelity.untouchedSubtreesChecked >= 6);
    console.log('  ✓ Hyperlink, bookmark, field-code, and footnote preservation verified');
}

// =========================================================================
// Test 5: Table and SDT Preservation
// =========================================================================
{
    // Ensure P5 (inside SDT) and P7 (inside table) remain untouched during P1 edit
    const inputZip = createMasterPackage();
    const doc = openDocx(inputZip);

    const result = await doc.applyOperations([
        {
            type: 'redline',
            target: { paragraphId: 'P0000001' },
            modified: 'P1: Modified intro paragraph.',
            author: 'Editor'
        }
    ], { atomic: true });

    const fidelity = verifyMutationFidelity(inputZip, result.toBuffer(), {
        envelope: MUTATION_ENVELOPES.surgical_text,
        targetedParagraphIds: ['P0000001'],
        lifecycleCheck: true
    });

    assert.ok(fidelity.untouchedSubtreesChecked >= 6);
    console.log('  ✓ SDT container and Table container preservation verified');
}

// =========================================================================
// Test 6: Package parts immutability (headers, footers, media, styles)
// =========================================================================
{
    const inputZip = createMasterPackage();
    const doc = openDocx(inputZip);

    const result = await doc.applyOperations([
        {
            type: 'redline',
            target: { paragraphId: 'P0000009' },
            modified: 'P9: Modified conclusion.',
            author: 'Editor'
        }
    ], { atomic: true });

    const outputZip = result.toBuffer();

    // Verify binary media and auxiliary parts are byte-identical
    const fidelity = verifyMutationFidelity(inputZip, outputZip, {
        envelope: MUTATION_ENVELOPES.surgical_text,
        targetedParagraphIds: ['P0000009']
    });

    assert.equal(fidelity.packageComparison.unchanged.some(e => e.name === 'word/media/image1.png'), true);
    assert.equal(fidelity.packageComparison.unchanged.some(e => e.name === 'word/header1.xml'), true);
    assert.equal(fidelity.packageComparison.unchanged.some(e => e.name === 'word/footer1.xml'), true);
    assert.equal(fidelity.packageComparison.unchanged.some(e => e.name === 'word/footnotes.xml'), true);
    assert.equal(fidelity.packageComparison.unchanged.some(e => e.name === 'word/endnotes.xml'), true);
    assert.equal(fidelity.packageComparison.unchanged.some(e => e.name === 'word/styles.xml'), true);
    console.log('  ✓ Auxiliary parts (header, footer, footnotes, endnotes, media, styles) byte-identical');
}

console.log('fidelity oracle tests passed');
