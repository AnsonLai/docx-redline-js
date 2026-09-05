import assert from 'assert/strict';

import './setup-xml-provider.mjs';
import {
    acceptTrackedChangesInOoxml,
    applyRedlineToOxml,
    rejectTrackedChangesInOoxml,
    validateRedlineOoxml
} from '../index.js';
import { parseOoxmlSafe } from '../adapters/xml-adapter.js';
import { detectNumberingContext } from '../pipeline/ingestion.js';
import { NumberingService } from '../services/numbering-service.js';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const SOURCE = `<w:p xmlns:w="${W}">
  <w:pPr>
    <w:pStyle w:val="Level1"/>
    <w:numPr><w:ilvl w:val="0"/><w:numId w:val="0"/></w:numPr>
    <w:jc w:val="both"/>
    <w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="24"/><w:u w:val="single"/></w:rPr>
  </w:pPr>
  <w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="24"/></w:rPr><w:t>A.</w:t></w:r>
  <w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="24"/></w:rPr><w:tab/></w:r>
  <w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="24"/><w:u w:val="single"/></w:rPr><w:t>PURPOSE</w:t></w:r>
</w:p>`;
const MODIFIED = '* Article A. Purpose and Interagency Alignment\n* Key Focus: Joint Street Outreach & Medical Triage';
const AUTHOR = 'Structural List Test';

function elements(node, localName) {
    return Array.from(node.getElementsByTagNameNS('*', localName));
}

function paragraphs(oxml) {
    const parsed = parseOoxmlSafe(oxml);
    assert.ok(parsed.doc, parsed.error?.message || 'Expected parseable OOXML');
    return elements(parsed.doc, 'p');
}

function paragraphText(paragraph, includeDeleted = false) {
    const pieces = [];
    const visit = node => {
        if (node.nodeType === 1) {
            if (node.localName === 't' || (includeDeleted && node.localName === 'delText')) {
                pieces.push(node.textContent || '');
                return;
            }
            if (node.localName === 'tab') {
                pieces.push('\t');
                return;
            }
        }
        for (const child of Array.from(node.childNodes || [])) visit(child);
    };
    visit(paragraph);
    return pieces.join('');
}

function hasParagraphMark(paragraph, type) {
    return elements(paragraph, type).some(node => (
        node.parentNode?.localName === 'rPr'
        && node.parentNode?.parentNode?.localName === 'pPr'
    ));
}

const sourceDoc = parseOoxmlSafe(SOURCE).doc;
assert.equal(detectNumberingContext(elements(sourceDoc, 'p')[0]), null,
    'numId="0" must be treated as numbering suppression, not list context');

const numberingService = new NumberingService();
numberingService.registerExistingNumId('bullet', '0');
assert.equal(numberingService.getOrCreateNumId(
    { type: 'bullet' },
    { numId: '0', type: 'unknown', ilvl: 0 }
), '1', 'NumberingService must never reuse numId="0"');

const redlined = await applyRedlineToOxml(SOURCE, 'A.\tPURPOSE', MODIFIED, {
    author: AUTHOR,
    generateRedlines: true
});
assert.equal(redlined.hasChanges, true);
assert.equal(validateRedlineOoxml(redlined.oxml).valid, true);

const markedParagraphs = paragraphs(redlined.oxml);
assert.equal(markedParagraphs.length, 4, 'packaged fragment must retain its trailing insertion sentinel');
assert.equal(paragraphText(markedParagraphs[0], true), 'A.\tPURPOSE');
assert.equal(paragraphText(markedParagraphs[1]), 'Article A. Purpose and Interagency Alignment');
assert.equal(paragraphText(markedParagraphs[2]), 'Key Focus: Joint Street Outreach & Medical Triage');
assert.ok(hasParagraphMark(markedParagraphs[0], 'del'), 'source paragraph mark must be deleted');
assert.ok(hasParagraphMark(markedParagraphs[1], 'ins'), 'first list paragraph mark must be inserted');
assert.ok(hasParagraphMark(markedParagraphs[2], 'ins'), 'second list paragraph mark must be inserted');

for (const paragraph of markedParagraphs.slice(1, 3)) {
    const numId = elements(paragraph, 'numId')[0]?.getAttribute('w:val');
    assert.ok(Number.parseInt(numId, 10) > 0, 'inserted list paragraph must use a positive numId');
    const insertion = elements(paragraph, 'ins').find(node => node.parentNode?.localName === 'p');
    const rFonts = elements(insertion, 'rFonts')[0];
    const size = elements(insertion, 'sz')[0];
    assert.equal(rFonts?.getAttribute('w:ascii'), 'Times New Roman');
    assert.equal(size?.getAttribute('w:val'), '24');
    assert.equal(elements(insertion, 'b').length, 0, 'heading emphasis must not leak into list body text');
    assert.equal(elements(insertion, 'u').length, 0, 'heading underline must not leak into list body text');
}

const accepted = acceptTrackedChangesInOoxml(redlined.oxml, { author: AUTHOR });
assert.deepEqual(paragraphs(accepted.oxml).map(paragraph => paragraphText(paragraph)), [
    'Article A. Purpose and Interagency Alignment',
    'Key Focus: Joint Street Outreach & Medical Triage',
    ''
]);

const rejected = rejectTrackedChangesInOoxml(redlined.oxml, { author: AUTHOR });
const rejectedParagraphs = paragraphs(rejected.oxml);
assert.equal(rejectedParagraphs.length, 2, 'Reject must remove inserted list paragraphs and retain only the packaging sentinel');
assert.equal(paragraphText(rejectedParagraphs[0]), 'A.\tPURPOSE');
assert.equal(paragraphText(rejectedParagraphs[1]), '');
assert.equal(elements(rejectedParagraphs[0], 'numId')[0]?.getAttribute('w:val'), '0',
    'Reject must restore the source paragraph properties exactly');

console.log('PASS: structural list replacement regression tests');
