import assert from 'node:assert/strict';
import { canonicalizeOoxml } from './helpers/canonical-ooxml.mjs';

// 1. Different harmless prefixes compare equal
const xmlWithStandardPrefixes = `
<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:r>
    <w:t>Hello world</w:t>
  </w:r>
</w:p>`.trim();

const xmlWithCustomPrefixes = `
<myw:p xmlns:myw="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <myw:r>
    <myw:t>Hello world</myw:t>
  </myw:r>
</myw:p>`.trim();

const canonA = canonicalizeOoxml(xmlWithStandardPrefixes);
const canonB = canonicalizeOoxml(xmlWithCustomPrefixes);
assert.equal(canonA.canonicalXml, canonB.canonicalXml);
assert.equal(canonA.sha256, canonB.sha256);

// 2. Different attribute orders compare equal
const xmlAttrOrder1 = `
<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
     w:rsidR="00A1B2C3" w:rsidRDefault="00D4E5F6">
  <w:r><w:t>Text</w:t></w:r>
</w:p>`.trim();

const xmlAttrOrder2 = `
<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
     w:rsidRDefault="00D4E5F6" w:rsidR="00A1B2C3">
  <w:r><w:t>Text</w:t></w:r>
</w:p>`.trim();

const canonAttr1 = canonicalizeOoxml(xmlAttrOrder1);
const canonAttr2 = canonicalizeOoxml(xmlAttrOrder2);
assert.equal(canonAttr1.canonicalXml, canonAttr2.canonicalXml);
assert.equal(canonAttr1.sha256, canonAttr2.sha256);

// 3. Different child orders compare unequal
const xmlChildOrder1 = `
<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:r><w:t>Alpha</w:t></w:r>
  <w:r><w:t>Beta</w:t></w:r>
</w:p>`.trim();

const xmlChildOrder2 = `
<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:r><w:t>Beta</w:t></w:r>
  <w:r><w:t>Alpha</w:t></w:r>
</w:p>`.trim();

const canonChild1 = canonicalizeOoxml(xmlChildOrder1);
const canonChild2 = canonicalizeOoxml(xmlChildOrder2);
assert.notEqual(canonChild1.canonicalXml, canonChild2.canonicalXml);
assert.notEqual(canonChild1.sha256, canonChild2.sha256);

// 4. w:b absent and w:b w:val="0" compare unequal
const xmlNoBold = `
<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:rPr/>
  <w:t>Clause</w:t>
</w:r>`.trim();

const xmlBoldOff = `
<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:rPr>
    <w:b w:val="0"/>
  </w:rPr>
  <w:t>Clause</w:t>
</w:r>`.trim();

const xmlBoldOn = `
<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:rPr>
    <w:b/>
  </w:rPr>
  <w:t>Clause</w:t>
</w:r>`.trim();

const canonNoBold = canonicalizeOoxml(xmlNoBold);
const canonBoldOff = canonicalizeOoxml(xmlBoldOff);
const canonBoldOn = canonicalizeOoxml(xmlBoldOn);

assert.notEqual(canonNoBold.canonicalXml, canonBoldOff.canonicalXml);
assert.notEqual(canonNoBold.sha256, canonBoldOff.sha256);
assert.notEqual(canonBoldOff.canonicalXml, canonBoldOn.canonicalXml);
assert.notEqual(canonBoldOff.sha256, canonBoldOn.sha256);

// 5. Leading/trailing spaces and xml:space differences compare unequal when changing semantics
const xmlPreserve = `
<w:t xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xml:space="preserve"> test </w:t>`.trim();

const xmlNoPreserve = `
<w:t xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"> test </w:t>`.trim();

const canonPreserve = canonicalizeOoxml(xmlPreserve);
const canonNoPreserve = canonicalizeOoxml(xmlNoPreserve);
assert.notEqual(canonPreserve.canonicalXml, canonNoPreserve.canonicalXml);
assert.notEqual(canonPreserve.sha256, canonNoPreserve.sha256);

// 6. mc:Ignorable remains bound to the intended namespace prefixes
const xmlMc1 = `
<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
     xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
     xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
     mc:Ignorable="w14">
  <w:r><w:t>Content</w:t></w:r>
</w:p>`.trim();

const xmlMc2 = `
<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
     xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
     xmlns:word2010="http://schemas.microsoft.com/office/word/2010/wordml"
     mc:Ignorable="word2010">
  <w:r><w:t>Content</w:t></w:r>
</w:p>`.trim();

const xmlMcDifferentNs = `
<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
     xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
     xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"
     mc:Ignorable="w15">
  <w:r><w:t>Content</w:t></w:r>
</w:p>`.trim();

const canonMc1 = canonicalizeOoxml(xmlMc1);
const canonMc2 = canonicalizeOoxml(xmlMc2);
const canonMcDiff = canonicalizeOoxml(xmlMcDifferentNs);

assert.equal(canonMc1.canonicalXml, canonMc2.canonicalXml);
assert.equal(canonMc1.sha256, canonMc2.sha256);
assert.notEqual(canonMc1.canonicalXml, canonMcDiff.canonicalXml);
assert.notEqual(canonMc1.sha256, canonMcDiff.sha256);

console.log('canonical ooxml tests passed');
