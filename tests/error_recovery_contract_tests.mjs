import assert from 'node:assert/strict';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { normalizeErrorWithRecovery } from '../services/error-recovery.js';
import { openDocx } from '../node/index.js';
import { executeCli, runCli } from '../node/cli.js';
import { buildZip } from '../scripts/lib/minimal-zip.mjs';

const invalid = normalizeErrorWithRecovery({
    code: 'INVALID_OPERATION',
    message: 'modified must be a string.',
    field: 'operations[0].modified'
});
assert.equal(invalid.recoveryVersion, 1);
assert.equal(invalid.category, 'request_fixable');
assert.equal(invalid.recovery.action, 'change_request');
assert.equal(invalid.recovery.sameArgumentsSafe, false);

const ambiguous = normalizeErrorWithRecovery({
    code: 'AMBIGUOUS_TARGET',
    message: 'Two paragraphs matched.',
    candidates: [{ index: 1, text: 'First candidate.' }, { index: 7, text: 'Second candidate.' }]
});
assert.equal(ambiguous.category, 'candidate_selection_required');
assert.equal(ambiguous.recovery.action, 'choose_candidate');
assert.equal(ambiguous.recovery.requiresReinspection, false);
assert.equal(ambiguous.candidates.length, 2);

const revisions = normalizeErrorWithRecovery({
    code: 'EXISTING_REVISIONS',
    message: 'Reviewer changes are present.',
    revisionAuthors: ['Reviewer A'],
    currentPolicy: 'merge-same-author'
}, { operationIndex: 3, operationId: 'mutuality-3' });
assert.equal(revisions.category, 'policy_choice_required');
assert.equal(revisions.recovery.action, 'set_option');
assert.equal(revisions.recovery.recommendedValue, 'slice-cross-author');
assert.equal(revisions.recovery.requiresUserAuthorization, false);
assert.deepEqual(revisions.context.revisionAuthors, ['Reviewer A']);
assert.equal(revisions.context.currentPolicy, 'merge-same-author');

const commentProtection = normalizeErrorWithRecovery({
    code: 'COMMENTED_CONTENT_DELETE',
    message: 'Commented content cannot be deleted.',
    commentIds: ['4']
});
assert.equal(commentProtection.category, 'user_authorization_required');
assert.equal(commentProtection.recovery.requiresUserAuthorization, true);

const generated = normalizeErrorWithRecovery({
    code: 'GENERATED_OOXML_INVALID',
    message: 'Invalid output.',
    generatedIssues: [{ code: 'A' }, { code: 'A' }, { code: 'B' }]
});
assert.equal(generated.category, 'library_or_builder_failure');
assert.deepEqual(generated.issueSummary, {
    total: 3,
    byCode: [{ code: 'A', count: 2 }, { code: 'B', count: 1 }]
});

const patchMismatch = normalizeErrorWithRecovery({
    code: 'PATCH_ROUNDTRIP_MISMATCH',
    message: 'The requested accepted view could not be reconstructed exactly.'
});
assert.equal(patchMismatch.category, 'request_fixable');
assert.equal(patchMismatch.recovery.action, 'reinspect_and_narrow');
assert.equal(patchMismatch.recovery.requiresReinspection, true);

const unsafeBoundary = normalizeErrorWithRecovery({
    code: 'SECTION_BREAK_PARAGRAPH',
    message: 'The target carries a section boundary.'
});
assert.equal(unsafeBoundary.category, 'manual_document_resolution');
assert.equal(unsafeBoundary.recovery.action, 'manual_resolution');

const fixturePath = new URL('./fixtures/sample_doc_test.docx', import.meta.url);
const fixtureFile = fileURLToPath(fixturePath);
const fixture = await readFile(fixturePath);
const document = openDocx(fixture);
const currentRevision = document.getRevisionToken();
const staleRevision = { ...currentRevision, value: '0'.repeat(currentRevision.value.length) };
const staleResult = await document.applyOperations([], { expectedRevision: staleRevision });
assert.equal(staleResult.status, 'error');
assert.equal(staleResult.error.code, 'REVISION_MISMATCH');
assert.deepEqual(staleResult.error.expectedRevision, staleRevision);
assert.deepEqual(staleResult.error.currentRevision, currentRevision);
assert.equal(staleResult.error.category, 'target_refresh_required');
assert.equal(staleResult.error.recovery.action, 'reinspect');
assert.equal(staleResult.retryPlan.base, 'original');

const conflictingResult = await openDocx(fixture).applyOperations([
    {
        type: 'redline',
        target: 'Notices. All notices must be in writing and sent to the addresses listed above.',
        modified: 'Notices state A.'
    },
    {
        type: 'redline',
        target: 'Notices. All notices must be in writing and sent to the addresses listed above.',
        modified: 'Notices state B.'
    }
], { atomic: false });
assert.equal(conflictingResult.status, 'error');
assert.equal(conflictingResult.written, false);
assert.equal(conflictingResult.rolledBack, true);
assert.equal(conflictingResult.retryPlan.base, 'original');
assert.equal(conflictingResult.retryPlan.replayWholeBatch, true);

const version = await executeCli(['version']);
assert.equal(version.contractVersion, 8);
assert(version.capabilities.includes('batch-start-source-binding'));
assert(version.capabilities.includes('recovery-envelope-v1'));
assert(version.capabilities.includes('require-complete-exit'));
assert(version.capabilities.includes('operations-stdin'));
assert(version.capabilities.includes('agent-safety-profile-v2'));
assert(version.capabilities.includes('deduplicated-cli-receipts'));
assert(version.capabilities.includes('compact-cli-json-v1'));
assert(version.capabilities.includes('command-help-v1'));
assert(version.capabilities.includes('inspection-context-v1'));
assert(version.capabilities.includes('bounded-inspection-v1'));
assert(version.capabilities.includes('human-document-references-v1'));
assert(version.capabilities.includes('localized-replacements-v1'));
assert(version.capabilities.includes('speculative-search-apply-v1'));
assert(version.capabilities.includes('localized-change-summary-v1'));

const temp = await mkdtemp(path.join(tmpdir(), 'docx-recovery-contract-'));
try {
    const operationsPath = path.join(temp, 'operations.json');
    const completeOutput = path.join(temp, 'require-complete.docx');
    const legacyOutput = path.join(temp, 'legacy-partial.docx');
    await writeFile(operationsPath, JSON.stringify([
        {
            operationId: 'valid-notice-edit',
            type: 'redline',
            target: 'Notices. All notices must be in writing and sent to the addresses listed above.',
            modified: 'Notices. All notices must be in writing and sent to the addresses listed above by email.'
        },
        {
            operationId: 'missing-target',
            type: 'redline',
            target: 'This target is not present.',
            modified: 'This operation must fail.'
        }
    ]), 'utf8');

    let strictStdout = '';
    const strictExit = await runCli([
        'apply', fixtureFile, '--operations', operationsPath,
        '--output', completeOutput, '--require-complete'
    ], { stdout: { write: value => { strictStdout += value; } } });
    const strictResult = JSON.parse(strictStdout);
    assert.equal(strictExit, 3);
    assert.equal(strictResult.status, 'partial');
    assert.equal(strictResult.completion, false);
    assert.equal(strictResult.retryPlan.base, 'output');
    assert.deepEqual(strictResult.retryPlan.committedIndexes, [1]);
    assert.deepEqual(strictResult.retryPlan.failedIndexes, [2]);
    assert.equal(strictResult.results[1].error.recoveryVersion, 1);
    assert.equal(strictResult.results[1].error.recovery.action, 'reinspect');

    let legacyStdout = '';
    const legacyExit = await runCli([
        'apply', fixtureFile, '--operations', operationsPath,
        '--output', legacyOutput
    ], { stdout: { write: value => { legacyStdout += value; } } });
    assert.equal(legacyExit, 0, 'partial status retains legacy exit behavior without --require-complete');
    assert.equal(JSON.parse(legacyStdout).status, 'partial');

    const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
    const mismatchInput = path.join(temp, 'mismatch.docx');
    const mismatchOperations = path.join(temp, 'mismatch-operations.json');
    const mismatchDocumentXml = `<w:document xmlns:w="${W}"><w:body><w:p w:paraId="A1"><w:ins w:id="10" w:author="Prior" w:date="2026-01-01T00:00:00Z"><w:r><w:t>Alpha Beta</w:t></w:r></w:ins></w:p><w:sectPr/></w:body></w:document>`;
    await writeFile(mismatchInput, buildZip([
        {
            name: '[Content_Types].xml',
            data: '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
        },
        { name: 'word/document.xml', data: mismatchDocumentXml },
        { name: 'word/_rels/document.xml.rels', data: '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>' }
    ]));
    await writeFile(mismatchOperations, JSON.stringify([{
        type: 'redline',
        target: { exactText: 'Alpha Beta', paragraphId: 'A1' },
        modified: 'Alpha\nBeta',
        author: 'Reviewer',
        existingRevisions: 'slice-cross-author'
    }]));
    const mismatchResult = await executeCli([
        'apply', mismatchInput,
        '--operations', mismatchOperations,
        '--atomic'
    ]);
    const mismatchError = mismatchResult.results[0].error;
    assert.equal(mismatchError.code, 'PATCH_ROUNDTRIP_MISMATCH');
    assert.equal(mismatchError.mismatchOffset, 5);
    assert.equal(mismatchError.expectedCodePoint, 'U+000A');
    assert.equal(mismatchError.actualCodePoint, 'U+0020');
    assert.match(mismatchError.expectedExcerpt, /Alpha\nBeta/);
    assert.equal(mismatchResult.results[0].receipt, undefined);
    assert.equal(mismatchResult.receipts[0].finalDisposition, 'refused');
} finally {
    await rm(temp, { recursive: true, force: true });
}

console.log('error recovery contract tests passed');
