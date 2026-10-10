import './setup-xml-provider.mjs';

import assert from 'assert';
import { parseOoxmlSafe } from '../adapters/xml-adapter.js';
import {
    getCurrentChild,
    getCurrentNumIdElements,
    getCurrentParagraphProperties,
    getCurrentRunProperties,
    getDirectWordChildren,
    getHistoricalParagraphProperties,
    getHistoricalRunProperties,
    getParagraphMarkProperties,
    getAllReferencedNumIdElements
} from '../core/current-properties.js';
import { ingestWordOoxmlToMarkdown } from '../index.js';
import { detectNumberingContext } from '../pipeline/ingestion-paragraph.js';
import { extractFirstParagraphNumId, overwriteParagraphNumIds } from '../services/numbering-helpers.js';
import { getParagraphListInfo } from '../core/list-targeting.js';

const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

// Current pPr: Heading1 / outline 0 / numId 7 ilvl 2. Snapshot (pPrChange): Heading3 / outline 2 / numId 9 ilvl 0.
const DIVERGENT_PARAGRAPH = `
<w:p xmlns:w="${NS_W}">
  <w:pPr>
    <w:pStyle w:val="Heading1"/>
    <w:numPr><w:ilvl w:val="2"/><w:numId w:val="7"/></w:numPr>
    <w:outlineLvl w:val="0"/>
    <w:rPr><w:b/></w:rPr>
    <w:pPrChange w:id="1" w:author="A" w:date="2026-01-01T00:00:00Z">
      <w:pPr>
        <w:pStyle w:val="Heading3"/>
        <w:numPr><w:ilvl w:val="0"/><w:numId w:val="9"/></w:numPr>
        <w:outlineLvl w:val="2"/>
      </w:pPr>
    </w:pPrChange>
  </w:pPr>
  <w:r>
    <w:rPr>
      <w:rFonts w:ascii="Arial"/>
      <w:rPrChange w:id="2" w:author="A" w:date="2026-01-01T00:00:00Z">
        <w:rPr><w:b/><w:i/><w:rFonts w:ascii="Courier New"/></w:rPr>
      </w:rPrChange>
    </w:rPr>
    <w:t>Body</w:t>
  </w:r>
</w:p>`;

function parseParagraph(xml) {
    const { doc } = parseOoxmlSafe(xml, 'application/xml');
    return doc.documentElement;
}

function testAccessorReturnsCurrentValues() {
    const p = parseParagraph(DIVERGENT_PARAGRAPH);
    const props = getCurrentParagraphProperties(p);
    assert.strictEqual(props.pStyle, 'Heading1');
    assert.strictEqual(props.outlineLvl, '0');
    assert.strictEqual(props.numPr.numId, '7');
    assert.strictEqual(props.numPr.ilvl, '2');
    assert.strictEqual(getCurrentChild(props.pPr, 'pStyle').getAttribute('w:val'), 'Heading1');
    assert.strictEqual(getCurrentChild(props.pPr, 'pPr'), null, 'snapshot pPr is not a direct child');

    const historical = getHistoricalParagraphProperties(p);
    assert.strictEqual(historical.pStyle, 'Heading3');
    assert.strictEqual(historical.outlineLvl, '2');
    assert.strictEqual(historical.numPr.numId, '9');
    assert.strictEqual(historical.numPr.ilvl, '0');

    assert.ok(getCurrentChild(getParagraphMarkProperties(p), 'b'), 'paragraph mark rPr is exposed explicitly');
    assert.strictEqual(getHistoricalParagraphProperties(parseParagraph(`<w:p xmlns:w="${NS_W}"/>`)), null);
}

function testRunAccessor() {
    const p = parseParagraph(DIVERGENT_PARAGRAPH);
    const run = getDirectWordChildren(p, 'r')[0];
    const rPr = getCurrentRunProperties(run);
    assert.ok(getCurrentChild(rPr, 'rFonts'));
    assert.strictEqual(getCurrentChild(rPr, 'rFonts').getAttribute('w:ascii'), 'Arial');
    assert.strictEqual(getCurrentChild(rPr, 'b'), null);
    assert.strictEqual(getCurrentChild(rPr, 'i'), null);
    const historical = getHistoricalRunProperties(run);
    assert.ok(getCurrentChild(historical, 'b') && getCurrentChild(historical, 'i'));
    assert.strictEqual(getCurrentChild(historical, 'rFonts').getAttribute('w:ascii'), 'Courier New');
}

function testNumIdEnumeration() {
    const p = parseParagraph(DIVERGENT_PARAGRAPH);
    assert.deepStrictEqual(getCurrentNumIdElements(p).map(el => el.getAttribute('w:val')), ['7']);
    assert.deepStrictEqual(getAllReferencedNumIdElements(p).map(el => el.getAttribute('w:val')), ['7', '9']);
}

function testMigratedReaders() {
    const p = parseParagraph(DIVERGENT_PARAGRAPH);

    const context = detectNumberingContext(p);
    assert.strictEqual(context.numId, '7');
    assert.strictEqual(context.ilvl, 2);

    const snapshotOnly = parseParagraph(`<w:p xmlns:w="${NS_W}"><w:pPr><w:pPrChange w:id="1" w:author="A"><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="9"/></w:numPr></w:pPr></w:pPrChange></w:pPr></w:p>`);
    assert.strictEqual(detectNumberingContext(snapshotOnly), null, 'snapshot-only numPr is not a current list binding');

    assert.strictEqual(extractFirstParagraphNumId([p]), '7');
    overwriteParagraphNumIds([p], 42);
    assert.deepStrictEqual(getAllReferencedNumIdElements(p).map(el => el.getAttribute('w:val')), ['42', '9'],
        'overwrite keeps the pPrChange snapshot numId so Reject restores it');

    const listInfo = getParagraphListInfo?.(parseParagraph(DIVERGENT_PARAGRAPH));
    if (listInfo) assert.notStrictEqual(String(listInfo.numId ?? ''), '9');
}

function testMarkdownExportUsesCurrentProperties() {
    const doc = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${NS_W}"><w:body>
<w:p><w:pPr><w:pStyle w:val="Heading2"/><w:pPrChange w:id="1" w:author="A"><w:pPr><w:pStyle w:val="Heading5"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr></w:pPr></w:pPrChange></w:pPr><w:r><w:rPr><w:rPrChange w:id="2" w:author="A"><w:rPr><w:b/><w:i/></w:rPr></w:rPrChange></w:rPr><w:t>Title</w:t></w:r></w:p>
</w:body></w:document>`;
    const markdown = ingestWordOoxmlToMarkdown(doc);
    assert.strictEqual(markdown.trim(), '## Title');
}

testAccessorReturnsCurrentValues();
testRunAccessor();
testNumIdEnumeration();
testMigratedReaders();
testMarkdownExportUsesCurrentProperties();
console.log('PASS: current properties tests');
