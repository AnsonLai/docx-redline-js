import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { configureLogger } from '../adapters/logger.js';
import * as nodeFacade from '../node/index.js';
import {
    DEFAULT_AGENT_PROFILE,
    ExampleAgentDocumentSession,
    createExampleAgentSession
} from '../examples/agent-session-wrapper.mjs';

configureLogger({ info() {}, warn() {}, error() {} });

const fixture = await readFile(new URL('./fixtures/sample_doc_test.docx', import.meta.url));
const packageConfig = JSON.parse(await readFile(new URL('../package.json', import.meta.url), 'utf8'));
const sourceText = 'Notices. All notices must be in writing and sent to the addresses listed above.';
const desiredText = 'Notices. All notices must be in writing and sent to the addresses listed above by email.';

assert.equal(
    'createExampleAgentSession' in nodeFacade,
    false,
    'the demonstration agent must not become part of the published Node facade'
);
assert.equal(packageConfig.files.includes('examples/'), false);
assert(packageConfig.files.includes('!scripts/benchmark-agent-workflow.mjs'));
assert(packageConfig.files.includes('!scripts/lib/agent-performance-cases.mjs'));
assert.equal(DEFAULT_AGENT_PROFILE.atomic, true);
assert.equal(DEFAULT_AGENT_PROFILE.strictTargets, true);
assert.equal(DEFAULT_AGENT_PROFILE.validate, true);

const session = createExampleAgentSession(fixture, { sessionId: 'test-session' });
assert(session instanceof ExampleAgentDocumentSession);
const found = session.inspect({ search: 'All notices must', around: 1 });
assert.equal(found.ok, true);
assert.equal(found.sessionId, 'test-session');
assert.equal(found.packageRevision.scope, 'package');
assert.equal(found.counts.matches, 1);
assert(found.counts.context >= 1, 'context-window inspection should include a neighboring paragraph');
assert.equal(found.targets.filter(target => target.role === 'match').length, 1);
assert(found.targets.some(target => target.role === 'context'));
assert(found.targets.every(target => /^T\d+$/.test(target.handle)));
assert(found.targets.every(target => !('paragraphId' in target)), 'opaque handles should hide strict descriptors');

const match = found.targets.find(target => target.role === 'match');
assert.equal(match.exactText, sourceText);
const firstRevision = found.packageRevision.value;
const applied = await session.applyEdits([
    { operationId: 'notice-email', target: match.handle, desiredText }
]);
assert.equal(applied.ok, true);
assert.equal(applied.complete, true);
assert.equal(applied.changed, true);
assert.equal(applied.status, 'ok');
assert.equal(applied.effectiveProfile.atomic, true);
assert.equal(applied.effectiveProfile.strictTargets, true);
assert.equal(applied.results[0].status, 'applied');
assert(applied.outputBytes instanceof Buffer);
assert.notEqual(applied.packageRevision.value, firstRevision);
assert.equal(applied.refreshedTargets.length, 1);
assert.equal(applied.refreshedTargets[0].operationId, 'notice-email');
assert.equal(applied.refreshedTargets[0].previousHandle, match.handle);
assert.notEqual(applied.refreshedTargets[0].target.handle, match.handle);
assert.equal(applied.refreshedTargets[0].target.exactText, desiredText);

const outputDocument = nodeFacade.openDocx(applied.outputBytes);
assert(outputDocument.inspect().paragraphs.some(paragraph => paragraph.exactText === desiredText));
assert(outputDocument.inspect({ revisionView: 'rejected' }).paragraphs.some(paragraph => paragraph.exactText === sourceText));

const stale = await session.applyEdits([
    { target: match.handle, desiredText: `${desiredText} Stale retry.` }
]);
assert.equal(stale.ok, false);
assert.equal(stale.error.code, 'STALE_TARGET_HANDLE');
assert.equal(stale.error.recovery.action, 'reinspect');
assert.equal(stale.outputBytes, null);

const refreshed = session.inspect({ search: 'by email' });
assert.equal(refreshed.counts.matches, 1);
assert.notEqual(refreshed.targets[0].handle, match.handle);
assert.equal(refreshed.targets[0].handle, applied.refreshedTargets[0].target.handle);
const accepted = await session.resolveReview('accept', { allAuthors: true });
assert.equal(accepted.ok, true);
assert.equal(accepted.changed, true);
const acceptedDocument = nodeFacade.openDocx(accepted.outputBytes);
assert(acceptedDocument.inspect().paragraphs.some(paragraph => paragraph.exactText === desiredText));
assert(acceptedDocument.inspect({ revisionView: 'rejected' }).paragraphs.some(paragraph => paragraph.exactText === desiredText));

const invalidSession = createExampleAgentSession(fixture, { sessionId: 'invalid-session' });
const invalidTarget = invalidSession.inspect({ search: 'All notices must' }).targets[0];
const beforeInvalid = invalidSession.toBuffer();
const invalid = await invalidSession.applyEdits([
    { target: invalidTarget.handle, desiredText },
    { target: 'T999', desiredText: 'Never applied.' }
]);
assert.equal(invalid.ok, false);
assert.equal(invalid.error.code, 'TARGET_HANDLE_NOT_FOUND');
assert.deepEqual(invalid.retryPlan.failedIndexes, [2]);
assert.equal(invalid.retryPlan.base, 'original');
assert.deepEqual(invalidSession.toBuffer(), beforeInvalid);

const rollbackSession = createExampleAgentSession(fixture, { sessionId: 'rollback-session' });
const rollbackTarget = rollbackSession.inspect({ search: 'All notices must' }).targets[0];
const beforeRollback = rollbackSession.toBuffer();
const rolledBack = await rollbackSession.applyEdits([
    { target: rollbackTarget.handle, desiredText },
    { target: rollbackTarget.handle, commentContent: 'Cannot anchor this.', textToComment: 'missing anchor text' }
]);
assert.equal(rolledBack.ok, false);
assert.equal(rolledBack.status, 'error');
assert.equal(rolledBack.changed, false);
assert.equal(rolledBack.outputBytes, null);
assert.equal(rolledBack.partialOutputBytes, null);
assert.equal(rolledBack.retryPlan.base, 'original');
assert.equal(rolledBack.retryPlan.replayWholeBatch, true);
assert(rolledBack.results.some(result => result.error?.code === 'ANCHOR_NOT_FOUND'));
assert.deepEqual(rollbackSession.toBuffer(), beforeRollback);

const invalidInspection = createExampleAgentSession(fixture).inspect({ around: 21 });
assert.equal(invalidInspection.ok, false);
assert.equal(invalidInspection.error.code, 'INVALID_AGENT_REQUEST');

console.log('agent session example tests passed');
