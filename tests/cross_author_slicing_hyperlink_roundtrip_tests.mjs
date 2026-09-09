import assert from 'assert/strict';

import './setup-xml-provider.mjs';
import {
    acceptTrackedChangesInOoxml,
    applyRedlineToOxml,
    rejectTrackedChangesInOoxml
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
    return `<w:hyperlink r:id="${id}" w:history="1">${run(text)}</w:hyperlink>`;
}

function acceptedParagraphText(xml) {
    const parsed = parseOoxmlSafe(xml, 'application/xml');
    assert.equal(parsed.error, null);
    const paragraph = parsed.doc.getElementsByTagNameNS(W, 'p')[0];
    return extractCanonicalParagraphText(paragraph);
}

const original = `Example Co. will take reasonable measures to safeguard records. This agreement incorporates the Processing Schedule located at${nbsp}example.invalid/schedule, which states the parties' obligations. Example Co. will also use records under the agreement and the Service Policy located at${nbsp}example.invalid/policy/${nbsp}(the "Service Policy").`;
const modified = `Example Co. will take reasonable measures to safeguard records. This agreement incorporates the Processing Schedule located at example.invalid/schedule, which states the parties' obligations. Example Co. will also use records under the agreement and the Service Policy, as in effect on execution, located at example.invalid/policy/ (the "Service Policy").`;

const source = `<w:p xmlns:w="${W}" xmlns:r="${R}">`
    + insertion(1, 'Example Co. will take reasonable measures to safeguard records. This agreement incorporates the Processing Schedule located at')
    + insertion(2, nbsp)
    + hyperlink('rIdSchedule', 'example.invalid/schedule')
    + insertion(3, ", which states the parties' obligations. Example Co. will also use records under the agreement and the Service Policy located at")
    + insertion(4, nbsp)
    + hyperlink('rIdPolicy', 'example.invalid/policy/')
    + insertion(5, `${nbsp}(the "Service Policy").`)
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

// Replacing the space immediately after a hyperlink must keep the paired
// insertion at that boundary rather than appending it after the following run.
{
    const policyOriginal = `Example Co.’s Service Policy located at${nbsp}example.invalid/policy/${nbsp}(the “Service Policy”).`;
    const policyCallerTarget = 'Example Co.’s Service Policy located at example.invalid/policy/ (the “Service Policy”).';
    const policyModified = 'Example Co.’s Service Policy, as in effect on execution, located at example.invalid/policy/ (the “Service Policy”).';
    const policySource = `<w:p xmlns:w="${W}" xmlns:r="${R}">`
        + run('Example Co.’s Service Policy located at')
        + run(nbsp)
        + hyperlink('rIdPolicy', 'example.invalid/policy/')
        + run(`${nbsp}(the “`)
        + '<w:r><w:rPr><w:b/><w:bCs/><w:u w:val="single"/></w:rPr><w:t>Service Policy</w:t></w:r>'
        + run('”).')
        + '</w:p>';
    const policyResult = await applyRedlineToOxml(policySource, policyOriginal, policyModified, {
        author: 'Reviewer',
        existingRevisions: 'slice-cross-author',
        pairReplacements: true,
        structuredContent: false
    });
    assert.equal(policyResult.status, 'ok', JSON.stringify(policyResult.error));
    assert.equal(acceptedParagraphText(policyResult.oxml), policyModified);
    const policyParsed = parseOoxmlSafe(policyResult.oxml, 'application/xml');
    const policyLinks = policyParsed.doc.getElementsByTagNameNS(W, 'hyperlink');
    assert.equal(policyLinks.length, 1);
    assert.equal(policyLinks[0].getAttribute('r:id'), 'rIdPolicy');
    assert.equal(policyLinks[0].getAttribute('w:history'), '1');

    const policyAccepted = acceptTrackedChangesInOoxml(policyResult.oxml, { allAuthors: true });
    assert.equal(acceptedParagraphText(policyAccepted.oxml), policyModified);

    const policyDocument = `<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>${policySource}<w:sectPr/></w:body></w:document>`;
    const policyBatch = await applyOperationsToDocumentXml(policyDocument, [{
        type: 'redline',
        target: { exactText: policyCallerTarget },
        modified: policyModified,
        author: 'Reviewer',
        existingRevisions: 'slice-cross-author'
    }], 'Reviewer', null, { atomic: true, strictTargets: true });
    assert.equal(policyBatch.status, 'ok', JSON.stringify(policyBatch.error));
    assert.equal(policyBatch.results[0]?.status, 'applied');
    assert.equal(policyBatch.results[0]?.resolvedTarget?.targetTextMatch?.mode, 'space_equivalent');
    assert.deepEqual(
        policyBatch.results[0]?.resolvedTarget?.targetTextMatch?.differences.map(item => [item.sourceCodePoint, item.requestedCodePoint]),
        [['U+00A0', 'U+0020'], ['U+00A0', 'U+0020']]
    );
    const policyBatchAccepted = acceptTrackedChangesInOoxml(policyBatch.documentXml, { allAuthors: true });
    assert.equal(acceptedParagraphText(policyBatchAccepted.oxml), policyModified);
    const policyBatchRejected = rejectTrackedChangesInOoxml(policyBatch.documentXml, { author: 'Reviewer' });
    assert.equal(acceptedParagraphText(policyBatchRejected.oxml), policyOriginal,
        'Rejecting the current reviewer must restore the exact NBSP-bearing source text');
}

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
