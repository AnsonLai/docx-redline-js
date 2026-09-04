import assert from 'node:assert/strict';
import './setup-xml-provider.mjs';
import { inspectDocumentParts } from '../index.js';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const documentXml = `<w:document xmlns:w="${W}"><w:body>
<w:p w:paraId="A1"><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Terms</w:t></w:r></w:p>
<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="7"/></w:numPr></w:pPr><w:commentRangeStart w:id="5"/><w:r><w:t xml:space="preserve">Alpha </w:t><w:tab/><w:br/><w:noBreakHyphen/></w:r><w:commentRangeEnd w:id="5"/><w:del w:id="2" w:author="Old"><w:r><w:delText>gone</w:delText></w:r></w:del><w:ins w:id="3" w:author="New"><w:r><w:t>kept</w:t></w:r></w:ins><w:r><w:commentReference w:id="5"/></w:r></w:p>
<w:sectPr/></w:body></w:document>`;
const numberingXml = `<w:numbering xmlns:w="${W}"><w:abstractNum w:abstractNumId="2"><w:lvl w:ilvl="0"><w:start w:val="4"/><w:numFmt w:val="upperLetter"/><w:lvlText w:val="%1)"/></w:lvl></w:abstractNum><w:num w:numId="7"><w:abstractNumId w:val="2"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="6"/></w:lvlOverride></w:num></w:numbering>`;
const commentsXml = `<w:comments xmlns:w="${W}"><w:comment w:id="5" w:author="Reviewer" w:date="2026-09-03T00:00:00Z"><w:p><w:r><w:t>Check it</w:t></w:r></w:p></w:comment></w:comments>`;

const result = inspectDocumentParts({ documentXml, numberingXml, commentsXml });
assert.equal(result.status, 'ok');
assert.equal(result.paragraphs[1].exactText, 'Alpha \t\n\u2011kept');
assert.equal(result.paragraphs[1].list.label, 'F)');
assert.deepEqual(result.paragraphs[1].revisionAuthors, ['New', 'Old']);
assert.equal(result.comments[0].author, 'Reviewer');
assert.equal(result.comments[0].targetRef, 'P2');
assert.equal(result.comments[0].anchoredText, 'Alpha \t\n\u2011');
assert.equal(result.paragraphs[1].nearestHeading.text, 'Terms');
assert.equal(result.paragraphs[0].styleId, 'Heading1');
assert.match(result.paragraphs[1].humanReference, /^F\)/);
assert.equal(inspectDocumentParts({ documentXml, numberingXml }, { revisedOnly: true }).paragraphs.length, 1);
assert.equal(inspectDocumentParts({ documentXml }, { range: { start: 2, end: 2 } }).paragraphs[0].index, 2);
console.log('document inspection tests passed');
