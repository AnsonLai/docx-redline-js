import './setup-xml-provider.mjs';

import assert from 'assert/strict';

import {
    acceptTrackedChangesInOoxml,
    containsTrackedChanges,
    ingestWordOoxmlToPlainText,
    rejectTrackedChangesInOoxml
} from '../index.js';
import { elementsByLocalName, parseXml } from './helpers/ooxml-assertions.mjs';

const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

function moveFixture() {
    return `<w:document xmlns:w="${NS_W}">
  <w:body>
    <w:p>
      <w:r><w:t xml:space="preserve">Before </w:t></w:r>
      <w:moveFromRangeStart w:id="77" w:author="Mover" w:date="2026-01-01T00:00:00Z" w:name="move1"/>
      <w:moveFrom w:id="78" w:author="Mover" w:date="2026-01-01T00:00:00Z">
        <w:r><w:delText>moved sentence</w:delText></w:r>
      </w:moveFrom>
      <w:moveFromRangeEnd w:id="77"/>
      <w:r><w:t xml:space="preserve"> after.</w:t></w:r>
    </w:p>
    <w:p>
      <w:r><w:t xml:space="preserve">Destination: </w:t></w:r>
      <w:moveToRangeStart w:id="79" w:author="Mover" w:date="2026-01-01T00:00:00Z" w:name="move1"/>
      <w:moveTo w:id="80" w:author="Mover" w:date="2026-01-01T00:00:00Z">
        <w:r><w:t>moved sentence</w:t></w:r>
      </w:moveTo>
      <w:moveToRangeEnd w:id="79"/>
    </w:p>
  </w:body>
</w:document>`;
}

function paragraphTexts(xml) {
    const doc = parseXml(xml);
    return Array.from(doc.getElementsByTagNameNS(NS_W, 'p')).map(p =>
        elementsByLocalName(p, 't').map(node => node.textContent || '').join('')
    );
}

function assertNoMoveMarkup(xml) {
    const doc = parseXml(xml);
    for (const localName of ['moveFrom', 'moveTo', 'moveFromRangeStart', 'moveFromRangeEnd', 'moveToRangeStart', 'moveToRangeEnd']) {
        assert.equal(doc.getElementsByTagNameNS(NS_W, localName).length, 0, `Expected no ${localName} markup`);
    }
}

function testAcceptAllConsumesMove() {
    const result = acceptTrackedChangesInOoxml(moveFixture(), { allAuthors: true });
    assert.equal(result.hasChanges, true);
    assertNoMoveMarkup(result.oxml);
    assert.deepEqual(paragraphTexts(result.oxml), ['Before  after.', 'Destination: moved sentence']);
}

function testRejectAllConsumesMove() {
    const result = rejectTrackedChangesInOoxml(moveFixture(), { allAuthors: true });
    assert.equal(result.hasChanges, true);
    assertNoMoveMarkup(result.oxml);
    assert.deepEqual(paragraphTexts(result.oxml), ['Before moved sentence after.', 'Destination: ']);
}

function testAuthorFilteredAcceptLeavesNonMatchingMoveUntouched() {
    const source = moveFixture();
    const result = acceptTrackedChangesInOoxml(source, { author: 'Someone Else' });
    assert.equal(result.hasChanges, false);
    assert.equal(result.oxml, source);
}

function testIngestionShowsMovedSentenceOnce() {
    const plain = ingestWordOoxmlToPlainText(moveFixture());
    assert.equal((plain.match(/moved sentence/g) || []).length, 1);
    assert.ok(plain.includes('Destination: moved sentence'));
}

function testContainsTrackedChangesSeesMoves() {
    assert.equal(containsTrackedChanges(parseXml(moveFixture())), true);
}

testAcceptAllConsumesMove();
testRejectAllConsumesMove();
testAuthorFilteredAcceptLeavesNonMatchingMoveUntouched();
testIngestionShowsMovedSentenceOnce();
testContainsTrackedChangesSeesMoves();

console.log('PASS: move revision tests');
