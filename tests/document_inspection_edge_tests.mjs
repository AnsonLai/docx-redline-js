import assert from 'node:assert/strict';
import './setup-xml-provider.mjs';
import { inspectDocumentParts, getParagraphListInfo } from '../index.js';
import { parseOoxmlSafe } from '../adapters/xml-adapter.js';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const documentXml = `<w:document xmlns:w="${W}"><w:body>
  <w:p><w:pPr><w:outlineLvl w:val="1"/></w:pPr><w:r><w:t>Scope</w:t></w:r></w:p>
  <w:p><w:commentRangeStart w:id="8"/><w:r><w:t>First half</w:t></w:r></w:p>
  <w:p><w:r><w:t>second half</w:t></w:r><w:commentRangeEnd w:id="8"/><w:r><w:commentReference w:id="8"/></w:r></w:p>
  <w:tbl><w:tr><w:tc><w:p><w:r><w:t>Cell one</w:t><w:footnoteReference w:id="9"/></w:r></w:p></w:tc><w:tc><w:p><w:r><w:endnoteReference w:id="10"/><w:t>Cell two</w:t></w:r></w:p></w:tc></w:tr></w:tbl>
  <w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="4"/></w:numPr></w:pPr><w:r><w:t>One</w:t></w:r></w:p>
  <w:p><w:pPr><w:numPr><w:ilvl w:val="1"/><w:numId w:val="4"/></w:numPr></w:pPr><w:r><w:t>Nested</w:t></w:r></w:p>
  <w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="4"/></w:numPr></w:pPr><w:r><w:t>Two</w:t></w:r></w:p>
  <w:p><w:r><w:t>Search Needle</w:t></w:r></w:p><w:p/>
  <w:sectPr/>
</w:body></w:document>`;
const commentsXml = `<w:comments xmlns:w="${W}"><w:comment w:id="8" w:author="A"><w:p><w:r><w:t>Across paragraphs</w:t></w:r></w:p></w:comment><w:comment w:id="99" w:author="Orphan"><w:p><w:r><w:t>Unanchored</w:t></w:r></w:p></w:comment></w:comments>`;
const numberingXml = `<w:numbering xmlns:w="${W}"><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="3"/><w:numFmt w:val="upperRoman"/><w:lvlText w:val="%1."/></w:lvl><w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%1.%2)"/></w:lvl></w:abstractNum><w:num w:numId="4"><w:abstractNumId w:val="1"/></w:num></w:numbering>`;

const result = inspectDocumentParts({ documentXml, commentsXml, numberingXml });
assert.equal(result.status, 'ok');
assert.equal(result.comments.find(item => item.id === '8').anchoredText, 'First half\nsecond half');
assert.equal(result.comments.find(item => item.id === '99').targetRef, undefined);
assert.deepEqual(result.commentAuthors, ['A', 'Orphan']);
assert.equal(result.paragraphs[0].headingLevel, 2);
assert.equal(result.paragraphs[3].table.tableIndex, 1);
assert.deepEqual(result.paragraphs[3].table, { tableIndex: 1, rowIndex: 1, cellIndex: 1 });
assert.deepEqual(result.paragraphs[3].structuralReferences, [{ type: 'footnote', id: '9' }]);
assert.deepEqual(result.paragraphs[4].structuralReferences, [{ type: 'endnote', id: '10' }]);
assert.deepEqual(result.paragraphs.slice(5, 8).map(item => item.list.label), ['III.', 'III.a)', 'IV.']);
assert.equal(result.paragraphs[5].humanReference.startsWith('III.'), true);
assert.deepEqual(inspectDocumentParts({ documentXml }, { inTable: true }).paragraphs.map(item => item.text), ['Cell one', 'Cell two']);
assert.deepEqual(inspectDocumentParts({ documentXml }, { search: 'needle' }).paragraphs.map(item => item.text), ['Search Needle']);
assert.deepEqual(inspectDocumentParts({ documentXml }, { indexes: [1, 8] }).paragraphs.map(item => item.index), [1, 8]);
assert.equal(inspectDocumentParts({ documentXml }, { skipEmpty: true }).paragraphs.some(item => item.text === ''), false);

const missing = inspectDocumentParts({});
assert.equal(missing.status, 'error'); assert.equal(missing.error.code, 'MISSING_PART');
const optionalMalformed = inspectDocumentParts({ documentXml, commentsXml: '<bad' });
assert.equal(optionalMalformed.status, 'ok'); assert.equal(optionalMalformed.warnings.length > 0, true);
console.log('document inspection edge tests passed');

// Historical (w:pPrChange) paragraph properties must never be reported as current.
{
    const hist = (inner) => `<w:pPrChange w:id="7" w:author="Prior Editor"><w:pPr>${inner}</w:pPr></w:pPrChange>`;
    const num = (id) => `<w:numPr><w:ilvl w:val="0"/><w:numId w:val="${id}"/></w:numPr>`;
    const histDoc = `<w:document xmlns:w="${W}"><w:body>
      <w:p><w:pPr>${hist(num(1))}</w:pPr><w:r><w:t>Historical only</w:t></w:r></w:p>
      <w:p><w:pPr>${num(1)}</w:pPr><w:r><w:t>Active control</w:t></w:r></w:p>
      <w:p><w:pPr>${num(2)}${hist(num(1))}</w:pPr><w:r><w:t>Both</w:t></w:r></w:p>
      <w:p><w:pPr><w:pStyle w:val="Heading2"/><w:outlineLvl w:val="2"/>${hist('<w:pStyle w:val="Heading1"/><w:outlineLvl w:val="0"/>')}</w:pPr><w:r><w:t>Style current</w:t></w:r></w:p>
      <w:p><w:pPr>${hist('<w:pStyle w:val="Heading1"/><w:outlineLvl w:val="0"/>')}</w:pPr><w:r><w:t>Style historical only</w:t></w:r></w:p>
      <w:sectPr/></w:body></w:document>`;
    const out = inspectDocumentParts({ documentXml: histDoc }).paragraphs;
    assert.equal(out[0].list, null, 'numPr only under pPrChange is not a current list');
    assert.deepEqual([out[1].list.numId, out[1].list.level], ['1', 0]);
    assert.equal(out[2].list.numId, '2', 'current numPr wins over historical');
    assert.equal(out[3].styleId, 'Heading2');
    assert.equal(out[3].headingLevel, 2);
    assert.equal(out[4].styleId, null, 'historical pStyle is not current');
    assert.equal(out[4].headingLevel, null, 'historical heading is not current');
}

// Shared list-targeting helper must also ignore historical numPr.
{
    const num = id => `<w:numPr><w:ilvl w:val="1"/><w:numId w:val="${id}"/></w:numPr>`;
    const histOnly = `<w:pPrChange w:id="7" w:author="P"><w:pPr>${num(1)}</w:pPr></w:pPrChange>`;
    const doc = parseOoxmlSafe(`<w:document xmlns:w="${W}"><w:body>
      <w:p><w:pPr>${histOnly}</w:pPr></w:p>
      <w:p><w:pPr>${num(2)}${histOnly}</w:pPr></w:p>
      <w:sectPr/></w:body></w:document>`, 'application/xml').doc;
    const [historicalOnly, both] = Array.from(doc.getElementsByTagNameNS(W, 'p'));
    assert.equal(getParagraphListInfo(historicalOnly), null);
    assert.deepEqual(getParagraphListInfo(both), { numId: '2', ilvl: 1 });
}
