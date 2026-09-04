import './setup-xml-provider.mjs';

import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import {
    applyOperationToDocumentXml as applyFromFacade,
    applyOperationsToDocumentXml as applyBatchFromFacade,
    orderOperationsForStableTargets as orderFromFacade
} from '../services/standalone-operation-runner.js';
import { applyOperationToDocumentXml as applyFromApplier } from '../services/document-operation-applier.js';
import {
    applyOperationsToDocumentXml as applyBatchFromOrchestrator,
    orderOperationsForStableTargets as orderFromOrchestrator
} from '../services/batch-operation-orchestrator.js';
import {
    cloneBatchRuntimeContext,
    commitBatchRuntimeContext,
    DocumentOperationSession
} from '../services/document-operation-session.js';

const DOCUMENT_XML = '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Original</w:t></w:r></w:p></w:body></w:document>';

assert.equal(applyFromFacade, applyFromApplier, 'facade must re-export the canonical applier');
assert.equal(applyBatchFromFacade, applyBatchFromOrchestrator, 'facade must re-export the canonical orchestrator');
assert.equal(orderFromFacade, orderFromOrchestrator, 'facade must re-export stable scheduling');

const facadeSource = await fs.readFile(new URL('../services/standalone-operation-runner.js', import.meta.url), 'utf8');
const mutationSource = await fs.readFile(new URL('../services/document-operation-mutations.js', import.meta.url), 'utf8');
assert.ok(facadeSource.split(/\r?\n/).length < 250, 'compatibility facade should stay below 250 lines');
assert.doesNotMatch(mutationSource, /from\s+['"]\.\.\/index\.js['"]/, 'operation internals must use leaf imports');

const session = new DocumentOperationSession(DOCUMENT_XML);
assert.equal(session.valid, true);
assert.match(session.serialize(), /<w:t>Original<\/w:t>/);
session.paragraphIndex = new Map([['Original', {}]]);
session.invalidateParagraphIndex();
assert.equal(session.paragraphIndex, null);
session.setDocumentXml('<changed/>');
assert.equal(session.rollback(), DOCUMENT_XML, 'rollback must return the exact original string');

const malformedSession = new DocumentOperationSession('<broken');
assert.equal(malformedSession.valid, false);
assert.equal(malformedSession.serialize(), '<broken');

const originalContext = {
    listFallbackSharedNumIdByKey: new Map([['list', '7']]),
    tableStructuralRedlineKeys: new Set(['table:1']),
    numberingIdState: {
        usedNumIds: new Set([7]),
        usedAbstractNumIds: new Set([8]),
        nextNumId: 9
    },
    listFallbackSequenceState: {
        explicitByNumberingKey: new Map([['list', { next: 2 }]])
    }
};
const context = cloneBatchRuntimeContext(originalContext);
context.listFallbackSharedNumIdByKey.set('new', '10');
context.tableStructuralRedlineKeys.add('table:2');
context.numberingIdState.usedNumIds.add(10);
assert.equal(originalContext.listFallbackSharedNumIdByKey.has('new'), false, 'working context must be isolated');
assert.equal(originalContext.tableStructuralRedlineKeys.has('table:2'), false, 'working sets must be isolated');
assert.equal(originalContext.numberingIdState.usedNumIds.has(10), false, 'numbering state must be isolated');

const originalMap = originalContext.listFallbackSharedNumIdByKey;
const originalSet = originalContext.numberingIdState.usedNumIds;
commitBatchRuntimeContext(originalContext, context);
assert.equal(originalContext.listFallbackSharedNumIdByKey, originalMap, 'commit must preserve caller map identity');
assert.equal(originalContext.numberingIdState.usedNumIds, originalSet, 'commit must preserve caller set identity');
assert.equal(originalMap.get('new'), '10');
assert.equal(originalSet.has(10), true);

assert.deepEqual(
    orderFromFacade([{ type: 'replace' }, { type: 'comment' }, { type: 'highlight' }]).map(op => op.type),
    ['comment', 'replace', 'highlight']
);

console.log('PASS: performance Phase 2 module boundaries');
