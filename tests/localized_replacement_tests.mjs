import './setup-xml-provider.mjs';

import assert from 'node:assert/strict';
import { parseOoxmlSafe } from '../adapters/xml-adapter.js';
import { getParagraphText } from '../core/paragraph-targeting.js';
import { validateDocumentOperation } from '../services/document-operation-contract.js';
import { compileExactReplacements } from '../services/localized-replacement-compiler.js';
import {
    applyOperationToDocumentXml,
    applyOperationsToDocumentXml,
    preflightOperations
} from '../services/standalone-operation-runner.js';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

function documentXml(text) {
    return `<w:document xmlns:w="${W}"><w:body><w:p w:paraId="A1"><w:r><w:t>${text}</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`;
}

function multiParagraphDocumentXml(paragraphs) {
    const body = paragraphs.map(({ id, text }) => (
        `<w:p w:paraId="${id}"><w:r><w:t>${text}</w:t></w:r></w:p>`
    )).join('');
    return `<w:document xmlns:w="${W}"><w:body>${body}<w:sectPr/></w:body></w:document>`;
}

function acceptedText(xml) {
    const parsed = parseOoxmlSafe(xml, 'application/xml');
    assert.equal(parsed.error, null);
    return getParagraphText(parsed.doc.getElementsByTagNameNS(W, 'p')[0]);
}

const sourceText = 'Alpha owes five dollars to Alpha.';
const desiredText = 'Alpha owes ten dollars to Beta.';
const replacements = [
    { find: 'five dollars', replace: 'ten dollars' },
    { find: 'Alpha', occurrence: 2, replace: 'Beta' }
];

const compiled = compileExactReplacements(sourceText, replacements);
assert.equal(compiled.ok, true);
assert.equal(compiled.desiredText, desiredText);
assert.deepEqual(compiled.replacements.map(item => item.find), ['five dollars', 'Alpha']);

const duplicate = compileExactReplacements('pay promptly', [
    { find: 'promptly', replace: 'within five days' },
    { find: 'promptly', replace: 'within five days' }
]);
assert.equal(duplicate.ok, true);
assert.equal(duplicate.replacements.length, 1);

const ambiguous = compileExactReplacements('Party pays Party.', [
    { find: 'Party', replace: 'Buyer' }
]);
assert.equal(ambiguous.ok, false);
assert.equal(ambiguous.error.code, 'AMBIGUOUS_PATCH_SOURCE');
assert.equal(ambiguous.error.candidates.length, 2);

const missing = compileExactReplacements('Party pays Party.', [
    { find: 'Seller', replace: 'Buyer' }
]);
assert.equal(missing.ok, false);
assert.equal(missing.error.code, 'PATCH_SOURCE_NOT_FOUND');
assert.equal(missing.error.matchCount, 0);

const overlap = compileExactReplacements('All notices must be written.', [
    { find: 'All notices', replace: 'Every notice' },
    { find: 'notices must', replace: 'notice shall' }
]);
assert.equal(overlap.ok, false);
assert.equal(overlap.error.code, 'OVERLAPPING_PATCHES');

for (const invalid of [
    [],
    [{ find: '', replace: 'value' }],
    [{ find: 'same', replace: 'same' }],
    [{ find: 'value', replace: 3 }]
]) {
    const validation = compileExactReplacements('same value', invalid);
    assert.equal(validation.ok, false);
    assert.equal(validation.error.code, 'INVALID_OPERATION');
}

const source = documentXml(sourceText);
const operation = {
    type: 'redline',
    target: { paragraphId: 'A1' },
    replacements
};
const validation = validateDocumentOperation(operation);
assert.equal(validation.valid, true);
assert.equal(validation.operation.modified, undefined);

const preflight = preflightOperations(source, [operation], 'Localized Editor', {
    strictTargets: true,
    generateRedlines: false
});
assert.equal(preflight.valid, true, JSON.stringify(preflight));
assert.equal(preflight.results[0].status, 'ready');

const result = await applyOperationsToDocumentXml(source, [operation], 'Localized Editor', null, {
    atomic: true,
    strictTargets: true,
    generateRedlines: false
});
assert.equal(result.status, 'ok', JSON.stringify(result));
assert.equal(result.results[0].status, 'applied');
assert.equal(acceptedText(result.documentXml), desiredText);
assert.equal(result.results[0].change.kind, 'localized_replacement');
assert.deepEqual(
    result.results[0].change.replacements.map(item => item.find),
    ['five dollars', 'Alpha']
);
assert.equal(result.results[0].change.verification.acceptedViewMatchesCompiledText, true);

const directResult = await applyOperationToDocumentXml(source, operation, 'Localized Editor', null, {
    strictTargets: true,
    generateRedlines: false
});
assert.equal(directResult.status, 'ok', JSON.stringify(directResult));
assert.equal(directResult.hasChanges, true);
assert.equal(acceptedText(directResult.documentXml), desiredText);

const canonical = await applyOperationsToDocumentXml(source, [{
    type: 'redline',
    target: { paragraphId: 'A1' },
    modified: desiredText
}], 'Localized Editor', null, {
    atomic: true,
    strictTargets: true,
    generateRedlines: false
});
assert.equal(acceptedText(canonical.documentXml), acceptedText(result.documentXml));

const deletion = await applyOperationsToDocumentXml(documentXml('Pay the administrative fee.'), [{
    type: 'redline',
    target: { paragraphId: 'A1' },
    replacements: [{ find: ' administrative', replace: '' }]
}], 'Localized Editor', null, { generateRedlines: false });
assert.equal(acceptedText(deletion.documentXml), 'Pay the fee.');

for (const invalidOperation of [
    { ...operation, modified: desiredText },
    { ...operation, target: { paragraphId: 'A1', revisionView: 'rejected' } },
    { ...operation, target: { captureRef: 'created' } },
    { ...operation, captureKey: 'patched' },
    { ...operation, target: { exactText: sourceText, occurrence: 1 } },
    { ...operation, targetEnd: { paragraphId: 'A1' } },
    { ...operation, replacements: [{ find: 'five dollars', replace: 'ten\ndollars' }] }
]) {
    assert.equal(validateDocumentOperation(invalidOperation).valid, false);
}

const failed = await applyOperationsToDocumentXml(source, [{
    type: 'redline',
    target: { paragraphId: 'A1' },
    replacements: [{ find: 'not present', replace: 'replacement' }]
}], 'Localized Editor', null, {
    atomic: true,
    strictTargets: true,
    generateRedlines: false
});
assert.equal(failed.status, 'error');
assert.equal(failed.error.code, 'BATCH_OPERATION_FAILED');
assert.equal(failed.results[0].error.code, 'PATCH_SOURCE_NOT_FOUND');
assert.equal(failed.results[0].error.recovery.action, 'change_patch');
assert.equal(failed.documentXml, source);

const longPrefix = `  ${'context '.repeat(20)}`;
const longSuffix = `${' trailing'.repeat(20)}  `;
const longSource = `${longPrefix}old first${longSuffix}old second${longSuffix}`;
const longSummary = await applyOperationsToDocumentXml(documentXml(longSource), [{
    type: 'redline',
    target: { paragraphId: 'A1' },
    replacements: [
        { find: 'old first', replace: 'new first' },
        { find: 'old second', replace: 'new second' }
    ]
}], 'Localized Editor', null, {
    atomic: true,
    strictTargets: true,
    generateRedlines: false
});
assert.equal(longSummary.status, 'ok', JSON.stringify(longSummary));
assert.deepEqual(
    longSummary.results[0].change.replacements.map(item => item.find),
    ['old first', 'old second']
);
for (const item of longSummary.results[0].change.replacements) {
    assert(item.beforeExcerpt.length <= 120);
    assert(item.afterExcerpt.length <= 120);
}
assert.match(longSummary.results[0].change.replacements[0].beforeExcerpt, /old first/);
assert.match(longSummary.results[0].change.replacements[1].afterExcerpt, /new second/);

// A later independent failure may execute before the patch because comments
// receive scheduling priority. Progressive execution still attempts the patch,
// and atomic rollback must make the change evidence unambiguously non-committed.
const rollbackSource = multiParagraphDocumentXml([
    { id: 'R1', text: 'The old value applies.' },
    { id: 'R2', text: 'Comment target paragraph.' }
]);
const rolledBack = await applyOperationsToDocumentXml(rollbackSource, [
    {
        type: 'redline',
        target: { paragraphId: 'R1' },
        replacements: [{ find: 'old value', replace: 'new value' }]
    },
    {
        type: 'comment',
        target: { paragraphId: 'R2' },
        textToComment: 'not present',
        commentContent: 'This must fail.'
    }
], 'Localized Editor', null, {
    atomic: true,
    continueOnError: true,
    strictTargets: true,
    generateRedlines: false
});
assert.equal(rolledBack.status, 'error', JSON.stringify(rolledBack));
assert.equal(rolledBack.rolledBack, true);
assert.equal(rolledBack.documentXml, rollbackSource);
assert.equal(rolledBack.results[0].change.committed, false);
assert.equal(rolledBack.results[0].change.finalDisposition, 'rolled_back');

console.log('localized replacement tests passed');
