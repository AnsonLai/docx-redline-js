import assert from 'assert/strict';

import './setup-xml-provider.mjs';
import {
    acceptTrackedChangesInOoxml,
    applyRedlineToOxml
} from '../index.js';
import { extractCanonicalParagraphText } from '../core/paragraph-text.js';
import { parseOoxmlSafe } from '../adapters/xml-adapter.js';
import { applyOperationsToDocumentXml } from '../services/standalone-operation-runner.js';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const nbsp = '\u00a0';

function run(text) {
    const preserve = /^\s|\s$/.test(text) ? ' xml:space="preserve"' : '';
    return `<w:r><w:t${preserve}>${text}</w:t></w:r>`;
}

function insertion(id, text) {
    return `<w:ins w:id="${id}" w:author="Prior Reviewer" w:date="2026-09-08T16:00:00Z">${run(text)}</w:ins>`;
}

function hyperlink(id, text) {
    return `<w:hyperlink r:id="${id}">${run(text)}</w:hyperlink>`;
}

function acceptedParagraphText(xml) {
    const parsed = parseOoxmlSafe(xml, 'application/xml');
    assert.equal(parsed.error, null);
    const paragraph = parsed.doc.getElementsByTagNameNS(W, 'p')[0];
    return extractCanonicalParagraphText(paragraph);
}

const original = `Acme will take reasonable measures to safeguard Data. This Agreement incorporates by reference the Processing Addendum located at${nbsp}example.com/legal/dpa, which sets forth obligations of the Parties. In addition, Acme will use Data only in accordance with the Agreement and, to the extent Data includes Personal Data, Acme's Widget Policy located at${nbsp}example.com/legal/wp/${nbsp}(the "Widget Policy").`;
const modified = `Acme will take reasonable measures to safeguard Data. This Agreement incorporates by reference the Processing Addendum located at example.com/legal/dpa, which sets forth obligations of the Parties. In addition, Acme will use Data only in accordance with the Agreement and, to the extent Data includes Personal Data, Acme's Widget Policy, as in effect as of the date this Agreement is executed, located at example.com/legal/wp/ (the "Widget Policy").`;

const source = `<w:p xmlns:w="${W}" xmlns:r="${R}">`
    + insertion(1, 'Acme will take reasonable measures to safeguard Data. This Agreement incorporates by reference the Processing Addendum located at')
    + insertion(2, nbsp)
    + hyperlink('rIdDpa', 'example.com/legal/dpa')
    + insertion(3, ", which sets forth obligations of the Parties. In addition, Acme will use Data only in accordance with the Agreement and, to the extent Data includes Personal Data, Acme's Widget Policy located at")
    + insertion(4, nbsp)
    + hyperlink('rIdWidget', 'example.com/legal/wp/')
    + insertion(5, `${nbsp}(the "Widget Policy").`)
    + '</w:p>';

const result = await applyRedlineToOxml(source, original, modified, {
    author: 'Reviewer',
    existingRevisions: 'slice-cross-author',
    pairReplacements: true
});

assert.equal(result.status, 'ok', JSON.stringify(result.error));
assert.equal(result.hasChanges, true);
assert.equal(acceptedParagraphText(result.oxml), modified,
    'accepted-view text must exactly equal the requested modified text');

const accepted = acceptTrackedChangesInOoxml(result.oxml, { allAuthors: true });
assert.equal(acceptedParagraphText(accepted.oxml), modified,
    'Accept All must exactly equal the requested modified text');

const parsedResult = parseOoxmlSafe(result.oxml, 'application/xml');
assert.equal(parsedResult.doc.getElementsByTagNameNS(W, 'hyperlink').length, 2,
    'both hyperlink relationship containers must survive');

// The document operation path used by CLI apply must preserve the same invariant.
{
    const documentXml = `<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>${source}<w:sectPr/></w:body></w:document>`;
    const batch = await applyOperationsToDocumentXml(documentXml, [{
        type: 'redline',
        target: { exactText: original },
        modified,
        author: 'Reviewer',
        existingRevisions: 'slice-cross-author'
    }], 'Reviewer', null, { atomic: true, strictTargets: true });
    assert.equal(batch.status, 'ok', JSON.stringify(batch.error));
    assert.equal(batch.hasChanges, true);
    assert.equal(batch.results[0]?.status, 'applied');
    const batchAccepted = acceptTrackedChangesInOoxml(batch.documentXml, { allAuthors: true });
    assert.equal(acceptedParagraphText(batchAccepted.oxml), modified);
}

// A structural mutation that cannot reconstruct the requested exact text must fail closed.
{
    const newlineSource = `<w:p xmlns:w="${W}">${insertion(10, 'Alpha Beta')}</w:p>`;
    const mismatch = await applyRedlineToOxml(newlineSource, 'Alpha Beta', 'Alpha\nBeta', {
        author: 'Reviewer',
        existingRevisions: 'slice-cross-author',
        pairReplacements: true,
        structuredContent: false
    });
    assert.equal(mismatch.status, 'error');
    assert.equal(mismatch.hasChanges, false);
    assert.equal(mismatch.error?.code, 'PATCH_ROUNDTRIP_MISMATCH');
    assert.equal(mismatch.oxml, newlineSource, 'round-trip mismatch must return the exact input OOXML');
    assert.equal(mismatch.error?.mismatchOffset, 5);
}

console.log('PASS: cross-author slicing hyperlink round-trip regression tests');
