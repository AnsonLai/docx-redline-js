import assert from 'node:assert/strict';
import './setup-xml-provider.mjs';
import { parseOoxmlSafe } from '../adapters/xml-adapter.js';
import { extractCanonicalParagraphText, getParagraphText, readCanonicalRunText } from '../index.js';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const xml = `<w:p xmlns:w="${W}">
  <w:r><w:t>A</w:t></w:r>
  <w:del w:id="1" w:author="Old"><w:r><w:delText>old</w:delText></w:r></w:del>
  <w:ins w:id="2" w:author="New"><w:r><w:t>new</w:t></w:r></w:ins>
  <w:moveFrom w:id="3"><w:r><w:delText>gone</w:delText></w:r></w:moveFrom>
  <w:moveTo w:id="4"><w:r><w:t>moved</w:t></w:r></w:moveTo>
  <w:r><w:tab/><w:br/><w:cr/><w:noBreakHyphen/><w:softHyphen/><w:instrText> HIDDEN </w:instrText><w:t>shown</w:t><w:footnoteReference w:id="7"/></w:r>
</w:p>`;
const parsed = parseOoxmlSafe(xml, 'application/xml');
assert.equal(parsed.error, null);
const paragraph = parsed.doc.documentElement;

assert.equal(extractCanonicalParagraphText(paragraph), 'Anewmoved\t\n\n\u2011\u00adshown');
assert.equal(extractCanonicalParagraphText(paragraph, { revisionView: 'current' }), 'Anewmoved\t\n\n\u2011\u00adshown');
assert.equal(extractCanonicalParagraphText(paragraph, { revisionView: 'rejected' }), 'Aoldgone\t\n\n\u2011\u00adshown');
assert.equal(getParagraphText(paragraph), extractCanonicalParagraphText(paragraph));

const insertedRun = paragraph.getElementsByTagNameNS(W, 'ins')[0].getElementsByTagNameNS(W, 'r')[0];
assert.equal(readCanonicalRunText(insertedRun, { boundary: paragraph, revisionView: 'accepted' }), 'new');
assert.equal(readCanonicalRunText(insertedRun, { boundary: paragraph, revisionView: 'rejected' }), '');
assert.equal(extractCanonicalParagraphText(null), '');
console.log('canonical paragraph text tests passed');
