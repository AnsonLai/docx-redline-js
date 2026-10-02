import './setup-xml-provider.mjs';

import assert from 'assert/strict';
import { inspectDocumentParts } from '../index.js';
import { applyOperationsToDocumentXml } from '../services/standalone-operation-runner.js';
import { validateDocumentOperation } from '../services/document-operation-contract.js';
import { acceptTrackedChangesInOoxml, rejectTrackedChangesInOoxml } from '../services/revision-comment-management.js';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const xml = `<w:document ${W}><w:body>`
  + '<w:p><w:r><w:t>Other paragraph.</w:t></w:r></w:p>'
  + '<w:p><w:r><w:t xml:space="preserve">laws of British Columbia, courts of </w:t></w:r>'
  + '<w:r><w:rPr><w:b/></w:rPr><w:t>British Columbia</w:t></w:r><w:r><w:t>.</w:t></w:r></w:p>'
  + '<w:sectPr/></w:body></w:document>';
const paragraph = inspectDocumentParts({ documentXml: xml }).paragraphs[1];
const target = { index: 2, exactText: paragraph.exactText, paragraphId: paragraph.paragraphId, fingerprint: paragraph.fingerprint };
const run = (op) => applyOperationsToDocumentXml(xml, [op], 'Tester', null, { atomic: true, strictTargets: true, generateRedlines: true });
const base = { type: 'format', target, textToFormat: 'British Columbia', properties: { bold: false } };

// Second occurrence (the bold one) is unbolded; the first is untouched.
const second = await run({ ...base, textOccurrence: 2 });
assert.equal(second.status, 'ok', JSON.stringify(second.error || second.results));
assert.equal(second.hasChanges, true);
const accepted = acceptTrackedChangesInOoxml(second.documentXml, { allAuthors: true }).oxml;
assert(accepted.includes('<w:b w:val="0"/>') && !accepted.includes('<w:b/>'), 'accepted: second occurrence unbolded');
assert(accepted.includes('<w:r><w:t xml:space="preserve">laws of British Columbia, courts of </w:t></w:r>'), 'first occurrence untouched');
assert(!accepted.includes('w:rPrChange'));
const rejected = rejectTrackedChangesInOoxml(second.documentXml, { allAuthors: true }).oxml;
assert(rejected.includes('<w:b/>'), 'rejected: bold restored');
assert(!rejected.includes('w:rPrChange'));

// First occurrence (default and explicit 1) targets the non-bold match: bolding it.
const first = await run({ ...base, properties: { bold: true }, textOccurrence: 1 });
assert.equal(first.status, 'ok');
const firstAccepted = acceptTrackedChangesInOoxml(first.documentXml, { allAuthors: true }).oxml;
assert(firstAccepted.indexOf('<w:b') > -1 && firstAccepted.indexOf('<w:b') < firstAccepted.indexOf('British Columbia', firstAccepted.indexOf('laws of') + 10) + 1 && /<w:rPr><w:b[^>]*\/><w:bCs[^>]*\/><\/w:rPr><w:t[^>]*>British Columbia<\/w:t>/.test(firstAccepted), 'first occurrence bolded');
const dflt = await run({ ...base, properties: { bold: true } });
const nd = (x) => x.replace(/w:date="[^"]*"/g, '');
assert.equal(nd(dflt.documentXml), nd(first.documentXml), 'default textOccurrence equals 1');

// Out of range fails closed.
const out = await run({ ...base, textOccurrence: 3 });
assert.equal(out.status, 'error');
assert.equal(out.hasChanges, false);
assert.equal(out.results?.[0]?.error?.code ?? out.error?.code, 'PATCH_SOURCE_NOT_FOUND');

// Invalid values are rejected by the contract.
for (const bad of [0, -1, 1.5, '2']) {
    assert.equal(validateDocumentOperation({ ...base, textOccurrence: bad }).valid, false, `textOccurrence ${bad}`);
}
assert.equal(validateDocumentOperation({ ...base, textOccurrence: 2 }).valid, true);

console.log('format text occurrence tests passed');
