import assert from 'node:assert/strict';

import './setup-xml-provider.mjs';
import {
    acceptTrackedChangesInOoxml,
    ingestWordOoxmlToPlainText,
    rejectTrackedChangesInOoxml,
    validateRedlineOoxml
} from '../index.js';
import { parseOoxmlSafe } from '../adapters/xml-adapter.js';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const DATE_A = '2026-09-08T16:00:00Z';
const DATE_B = '2026-09-08T17:00:00Z';

function localName(node) {
    return (node?.localName || node?.nodeName || '').replace(/^.*:/, '');
}

function descendants(root, name) {
    const all = Array.from(root.getElementsByTagName('*'));
    return name === '*' ? all : all.filter(node => localName(node) === name);
}

function elementChildren(root) {
    return Array.from(root.childNodes || []).filter(node => node.nodeType === 1);
}

function revisionElements(doc) {
    const names = new Set(['ins', 'del', 'moveFrom', 'moveTo', 'pPrChange', 'rPrChange', 'tblPrChange', 'tcPrChange']);
    return descendants(doc, '*').filter(node => names.has(localName(node)) && node.hasAttribute('w:id'));
}

function parse(xml) {
    const parsed = parseOoxmlSafe(xml, 'application/xml');
    assert(parsed.doc, parsed.error?.message || 'Fixture must parse');
    return parsed.doc;
}

function plainText(xml) {
    return ingestWordOoxmlToPlainText(xml).replace(/\r/g, '').trim();
}

function assertUniqueRevisionIds(xml, label) {
    const revisions = revisionElements(parse(xml));
    const ids = revisions.map(node => node.getAttribute('w:id'));
    assert.equal(new Set(ids).size, ids.length, `${label}: revision IDs must be globally unique`);
}

function assertValidOracle(xml, label) {
    const validation = validateRedlineOoxml(xml);
    const errors = (validation.issues || []).filter(issue => issue.severity === 'error');
    assert.deepEqual(errors, [], `${label}: ${JSON.stringify(errors)}`);
}

const inlineInsertionOracle = `
<w:document xmlns:w="${W}" xmlns:w14="${W14}">
  <w:body>
    <w:p w14:paraId="10000001">
      <w:pPr><w:rPr><w:del w:id="700" w:author="Reviewer A" w:date="${DATE_A}"/></w:rPr></w:pPr>
      <w:del w:id="701" w:author="Reviewer A" w:date="${DATE_A}"><w:r><w:delText>8.1 Standard Charges</w:delText></w:r></w:del>
    </w:p>
    <w:p w14:paraId="10000002">
      <w:pPr><w:rPr><w:del w:id="710" w:author="Reviewer A" w:date="${DATE_A}"/></w:rPr></w:pPr>
      <w:del w:id="711" w:author="Reviewer A" w:date="${DATE_A}"><w:r><w:delText>The account holder must pa</w:delText></w:r></w:del>
      <w:ins w:id="712" w:author="Reviewer B" w:date="${DATE_B}"><w:r><w:t>REVIEWER B INSERTION</w:t></w:r></w:ins>
      <w:del w:id="713" w:author="Reviewer A" w:date="${DATE_A}"><w:r><w:delText>y each undisputed invoice.</w:delText></w:r></w:del>
    </w:p>
    <w:p w14:paraId="10000003"><w:r><w:t>Following retained paragraph.</w:t></w:r></w:p>
    <w:sectPr/>
  </w:body>
</w:document>`.trim();

const paragraphRestorationOracle = `
<w:document xmlns:w="${W}" xmlns:w14="${W14}">
  <w:body>
    <w:p w14:paraId="20000001">
      <w:pPr><w:rPr><w:del w:id="714" w:author="Reviewer A" w:date="${DATE_A}"/></w:rPr></w:pPr>
      <w:del w:id="715" w:author="Reviewer A" w:date="${DATE_A}"><w:r><w:delText>8.2 Usage Adjustments</w:delText></w:r></w:del>
    </w:p>
    <w:p w14:paraId="20000002">
      <w:pPr><w:rPr><w:del w:id="716" w:author="Reviewer A" w:date="${DATE_A}"/></w:rPr></w:pPr>
      <w:del w:id="717" w:author="Reviewer A" w:date="${DATE_A}"><w:r><w:delText>Usage above the stated threshold may be billed at the next tier.</w:delText></w:r></w:del>
    </w:p>
    <w:p w14:paraId="30000001">
      <w:pPr><w:pStyle w:val="Heading2"/><w:rPr>
        <w:ins w:id="720" w:author="Reviewer B" w:date="${DATE_B}"/>
        <w:b/>
        <w:rPrChange w:id="721" w:author="Reviewer B" w:date="${DATE_B}"><w:rPr><w:ins w:id="722" w:author="Reviewer B" w:date="${DATE_B}"/></w:rPr></w:rPrChange>
      </w:rPr></w:pPr>
      <w:ins w:id="723" w:author="Reviewer B" w:date="${DATE_B}"><w:r><w:rPr><w:b/><w:rPrChange w:id="724" w:author="Reviewer B" w:date="${DATE_B}"><w:rPr/></w:rPrChange></w:rPr><w:t>8.2 Usage Adjustments</w:t></w:r></w:ins>
    </w:p>
    <w:p w14:paraId="30000002">
      <w:pPr><w:rPr><w:ins w:id="725" w:author="Reviewer B" w:date="${DATE_B}"/></w:rPr></w:pPr>
      <w:ins w:id="726" w:author="Reviewer B" w:date="${DATE_B}"><w:r><w:t>Usage above the stated threshold may be billed at the next tier.</w:t></w:r></w:ins>
    </w:p>
    <w:p w14:paraId="20000003">
      <w:pPr><w:rPr><w:del w:id="727" w:author="Reviewer A" w:date="${DATE_A}"/></w:rPr></w:pPr>
      <w:del w:id="728" w:author="Reviewer A" w:date="${DATE_A}"><w:r><w:delText>8.3 Following Deleted Heading</w:delText></w:r></w:del>
    </w:p>
    <w:p w14:paraId="20000004"><w:r><w:t>Following retained paragraph.</w:t></w:r></w:p>
    <w:sectPr/>
  </w:body>
</w:document>`.trim();

console.log('--- Word deleted-section edit oracle characterization ---');

{
    const doc = parse(inlineInsertionOracle);
    const paragraphs = descendants(doc, 'p');
    const editedParagraph = paragraphs[1];
    const childOrder = elementChildren(editedParagraph).map(localName);
    assert.deepEqual(childOrder, ['pPr', 'del', 'ins', 'del']);

    const topLevelRevisions = elementChildren(editedParagraph).filter(node => ['del', 'ins'].includes(localName(node)));
    assert.deepEqual(topLevelRevisions.map(node => node.getAttribute('w:author')), ['Reviewer A', 'Reviewer B', 'Reviewer A']);
    assert.deepEqual(topLevelRevisions.map(node => node.getAttribute('w:id')), ['711', '712', '713']);
    assert.equal(descendants(topLevelRevisions[1], 'del').length, 0, 'Reviewer B insertion must not be nested in a deletion');
    assert.equal(descendants(editedParagraph, 'ins').filter(node => node.parentNode && localName(node.parentNode) === 'rPr').length, 0,
        'Inline insertion must not create an inserted paragraph mark');
    assertUniqueRevisionIds(inlineInsertionOracle, 'inline insertion oracle');
    assertValidOracle(inlineInsertionOracle, 'inline insertion oracle');

    const rejectB = rejectTrackedChangesInOoxml(inlineInsertionOracle, { author: 'Reviewer B' });
    assert(!plainText(rejectB.oxml).includes('REVIEWER B INSERTION'));
    const remainingDeletions = descendants(parse(rejectB.oxml), 'del');
    assert.equal(remainingDeletions.length, 5, 'Rejecting Reviewer B must preserve Reviewer A content and paragraph-mark deletions');
    assert(remainingDeletions.every(node => node.getAttribute('w:author') === 'Reviewer A'));

    const rejectA = rejectTrackedChangesInOoxml(inlineInsertionOracle, { author: 'Reviewer A' });
    assert(plainText(rejectA.oxml).includes('The account holder must paREVIEWER B INSERTIONy each undisputed invoice.'));

    const rejectAll = rejectTrackedChangesInOoxml(inlineInsertionOracle, { allAuthors: true });
    assert(plainText(rejectAll.oxml).includes('The account holder must pay each undisputed invoice.'));
    assert(!plainText(rejectAll.oxml).includes('REVIEWER B INSERTION'));

    const acceptAll = acceptTrackedChangesInOoxml(inlineInsertionOracle, { allAuthors: true });
    assert(plainText(acceptAll.oxml).includes('REVIEWER B INSERTION'));
    assert(!plainText(acceptAll.oxml).includes('account holder'));
}

{
    const doc = parse(paragraphRestorationOracle);
    const paragraphs = descendants(doc, 'p');
    const body = descendants(doc, 'body')[0];
    const bodyParagraphs = elementChildren(body).filter(node => localName(node) === 'p');
    assert.deepEqual(bodyParagraphs.map(node => node.getAttribute('w14:paraId')), [
        '20000001', '20000002', '30000001', '30000002', '20000003', '20000004'
    ], 'Restored paragraphs must follow the deleted source block and precede the following source paragraph');

    for (const sourceParagraph of paragraphs.slice(0, 2)) {
        assert.equal(descendants(sourceParagraph, 'ins').length, 0, 'Deleted source paragraphs must remain untouched');
    }
    for (const restoredParagraph of paragraphs.slice(2, 4)) {
        const pMarkInsertions = descendants(restoredParagraph, 'rPr')
            .flatMap(rPr => elementChildren(rPr).filter(node => localName(node) === 'ins'));
        assert(pMarkInsertions.some(node => node.getAttribute('w:author') === 'Reviewer B'),
            'Each restored paragraph needs a Reviewer B paragraph-mark insertion');
        assert(elementChildren(restoredParagraph).some(node => localName(node) === 'ins' && node.getAttribute('w:author') === 'Reviewer B'),
            'Each restored paragraph needs a Reviewer B content insertion');
    }
    assertUniqueRevisionIds(paragraphRestorationOracle, 'paragraph restoration oracle');
    assertValidOracle(paragraphRestorationOracle, 'paragraph restoration oracle');

    const rejectB = rejectTrackedChangesInOoxml(paragraphRestorationOracle, { author: 'Reviewer B' });
    assert(!plainText(rejectB.oxml).includes('Usage Adjustments'));
    assert(!plainText(rejectB.oxml).includes('next tier'));

    const rejectA = rejectTrackedChangesInOoxml(paragraphRestorationOracle, { author: 'Reviewer A' });
    const rejectAText = plainText(rejectA.oxml);
    assert(rejectAText.includes('8.2 Usage Adjustments'));
    assert(rejectAText.includes('Usage above the stated threshold may be billed at the next tier.'));
    assert(rejectAText.includes('8.3 Following Deleted Heading'));

    const rejectAll = rejectTrackedChangesInOoxml(paragraphRestorationOracle, { allAuthors: true });
    const rejectAllText = plainText(rejectAll.oxml);
    assert(rejectAllText.includes('8.2 Usage Adjustments'));
    assert(rejectAllText.includes('8.3 Following Deleted Heading'));

    const acceptAll = acceptTrackedChangesInOoxml(paragraphRestorationOracle, { allAuthors: true });
    const acceptAllText = plainText(acceptAll.oxml);
    assert(acceptAllText.includes('8.2 Usage Adjustments'));
    assert(acceptAllText.includes('Usage above the stated threshold may be billed at the next tier.'));
    assert(!acceptAllText.includes('8.3 Following Deleted Heading'));
}

console.log('PASS: sanitized inline and paragraph-range Word oracles satisfy structure, identity, and lifecycle expectations');
