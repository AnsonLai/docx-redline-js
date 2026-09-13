import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { configureLogger } from '../adapters/logger.js';
import {
    compileExactReplacements,
    createExampleAgentSession
} from '../examples/agent-session-wrapper.mjs';
import { openDocx } from '../node/index.js';

configureLogger({ info() {}, warn() {}, error() {} });

const simultaneous = compileExactReplacements(
    'Alpha owes five dollars to Alpha.',
    [
        { find: 'five dollars', replace: 'ten United States dollars' },
        { find: 'Alpha', occurrence: 2, replace: 'Beta' }
    ]
);
assert.equal(simultaneous.ok, true);
assert.equal(simultaneous.desiredText, 'Alpha owes ten United States dollars to Beta.');
assert.equal(simultaneous.replacements.length, 2);

const deduplicated = compileExactReplacements('pay promptly', [
    { find: 'promptly', replace: 'within five days' },
    { find: 'promptly', replace: 'within five days' }
]);
assert.equal(deduplicated.ok, true);
assert.equal(deduplicated.desiredText, 'pay within five days');
assert.equal(deduplicated.replacements.length, 1);

const ambiguous = compileExactReplacements('Party pays Party.', [
    { find: 'Party', replace: 'Buyer' }
]);
assert.equal(ambiguous.ok, false);
assert.equal(ambiguous.error.code, 'AMBIGUOUS_PATCH_SOURCE');
assert.equal(ambiguous.error.recovery.action, 'choose_occurrence');
assert.equal(ambiguous.error.candidates.length, 2);

const missing = compileExactReplacements('Party pays Party.', [
    { find: 'Seller', replace: 'Buyer' }
]);
assert.equal(missing.ok, false);
assert.equal(missing.error.code, 'PATCH_SOURCE_NOT_FOUND');
assert.equal(missing.error.recovery.action, 'change_patch');

const outOfRangeOccurrence = compileExactReplacements('Party pays Party.', [
    { find: 'Party', occurrence: 3, replace: 'Buyer' }
]);
assert.equal(outOfRangeOccurrence.ok, false);
assert.equal(outOfRangeOccurrence.error.code, 'PATCH_SOURCE_NOT_FOUND');
assert.equal(outOfRangeOccurrence.error.matchCount, 2);

const overlapping = compileExactReplacements('All notices must be written.', [
    { find: 'All notices', replace: 'Every notice' },
    { find: 'notices must', replace: 'notice shall' }
]);
assert.equal(overlapping.ok, false);
assert.equal(overlapping.error.code, 'OVERLAPPING_PATCHES');
assert.equal(overlapping.error.recovery.action, 'combine_patches');

const conflicting = compileExactReplacements('All notices.', [
    { find: 'notices', replace: 'communications' },
    { find: 'notices', replace: 'notices and demands' }
]);
assert.equal(conflicting.ok, false);
assert.equal(conflicting.error.code, 'CONFLICTING_PATCHES');

const invalid = compileExactReplacements('All notices.', []);
assert.equal(invalid.ok, false);
assert.equal(invalid.error.code, 'INVALID_AGENT_REQUEST');

const fixture = await readFile(new URL('./fixtures/sample_doc_test.docx', import.meta.url));
const session = createExampleAgentSession(fixture, { sessionId: 'localized-edit-session' });
const term = session.inspect({ indexes: [32] }).targets[0];
const originalTerm = term.exactText;
const desiredTerm = originalTerm.replace('five (5) years', 'three (3) years');
const termResult = await session.applyEdits([{
    operationId: 'term-duration',
    target: term.handle,
    replacements: [{ find: 'five (5) years', replace: 'three (3) years' }]
}]);
assert.equal(termResult.ok, true);
assert.equal(termResult.compiledEdits.length, 1);
assert.equal(termResult.compiledEdits[0].operationId, 'term-duration');
assert.equal(termResult.compiledEdits[0].replacements.length, 1);
assert.equal(termResult.refreshedTargets.length, 1);
assert.notEqual(termResult.refreshedTargets[0].target.handle, term.handle);
assert.equal(termResult.refreshedTargets[0].target.exactText, desiredTerm);

const termOutput = openDocx(termResult.outputBytes);
assert(termOutput.inspect().paragraphs.some(paragraph => paragraph.exactText === desiredTerm));
assert(termOutput.inspect({ revisionView: 'rejected' }).paragraphs.some(paragraph => paragraph.exactText === originalTerm));

const refreshedTerm = termResult.refreshedTargets[0].target;
const chainedTermResult = await session.applyEdits([{
    operationId: 'term-duration-follow-up',
    target: refreshedTerm.handle,
    replacements: [{ find: 'three (3) years', replace: 'four (4) years' }]
}]);
assert.equal(chainedTermResult.ok, true, 'a refreshed handle should support a follow-up edit without reinspection');
assert.equal(chainedTermResult.refreshedTargets.length, 1);
assert.notEqual(chainedTermResult.refreshedTargets[0].target.handle, refreshedTerm.handle);
assert(openDocx(chainedTermResult.outputBytes).inspect().paragraphs.some(
    paragraph => paragraph.exactText === desiredTerm.replace('three (3) years', 'four (4) years')
));

const mutualitySession = createExampleAgentSession(fixture, { sessionId: 'mutuality-session' });
const remedy = mutualitySession.inspect({ indexes: [46] }).targets[0];
const desiredRemedy = remedy.exactText
    .replaceAll('The Receiving Party', 'Each Party')
    .replaceAll('the Disclosing Party', 'the other Party');
const mutualityResult = await mutualitySession.applyEdits([{
    operationId: 'mutual-remedy',
    target: remedy.handle,
    replacements: [
        { find: 'The Receiving Party', replace: 'Each Party' },
        { find: 'the Disclosing Party', occurrence: 1, replace: 'the other Party' },
        { find: 'the Disclosing Party', occurrence: 2, replace: 'the other Party' }
    ]
}]);
assert.equal(mutualityResult.ok, true);
assert(openDocx(mutualityResult.outputBytes).inspect().paragraphs.some(paragraph => paragraph.exactText === desiredRemedy));

const failedSession = createExampleAgentSession(fixture, { sessionId: 'failed-localized-edit' });
const notice = failedSession.inspect({ indexes: [53] }).targets[0];
const beforeFailure = failedSession.toBuffer();
const failed = await failedSession.applyEdits([{
    target: notice.handle,
    replacements: [{ find: 'not present', replace: 'replacement' }]
}]);
assert.equal(failed.ok, false);
assert.equal(failed.error.code, 'PATCH_SOURCE_NOT_FOUND');
assert.equal(failed.outputBytes, null);
assert.deepEqual(failedSession.toBuffer(), beforeFailure);

const crossAuthorFixture = await readFile(new URL(
    './fixtures/cross-author-slicing/insert-interior-pending.docx',
    import.meta.url
));
const revisionSession = createExampleAgentSession(crossAuthorFixture, { sessionId: 'revision-policy-session' });
const revisedTarget = revisionSession.inspect({ revisedOnly: true }).targets[0];
const refusedRevisionEdit = await revisionSession.applyEdits([{
    target: revisedTarget.handle,
    replacements: [{ find: 'amended', replace: 'revised' }]
}]);
assert.equal(refusedRevisionEdit.ok, false);
assert(refusedRevisionEdit.results.some(result => result.error?.code === 'EXISTING_REVISIONS'));

const permittedRevisionEdit = await revisionSession.applyEdits([{
    target: revisedTarget.handle,
    replacements: [{ find: 'amended', replace: 'revised' }],
    existingRevisions: 'slice-cross-author'
}]);
assert.equal(permittedRevisionEdit.ok, true);
assert(openDocx(permittedRevisionEdit.outputBytes).inspect().paragraphs.some(
    paragraph => paragraph.exactText === 'Contract terms revised by this MASTER Agreement.'
));

console.log('localized edit example tests passed');
