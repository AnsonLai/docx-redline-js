import './setup-xml-provider.mjs';
import assert from 'node:assert/strict';
import { parseOoxmlSafe } from '../adapters/xml-adapter.js';
import { getParagraphText } from '../core/paragraph-targeting.js';
import {
    applyOperationsToDocumentXml,
    preflightOperations
} from '../services/standalone-operation-runner.js';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

function documentXml(texts) {
    return `<w:document xmlns:w="${W}"><w:body>${texts.map(text => (
        `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`
    )).join('')}<w:sectPr/></w:body></w:document>`;
}

function paragraphTexts(xml) {
    const parsed = parseOoxmlSafe(xml, 'application/xml');
    assert.equal(parsed.error, null);
    return Array.from(parsed.doc.getElementsByTagNameNS(W, 'p')).map(getParagraphText);
}

function descriptor(resolvedTarget) {
    return {
        exactText: resolvedTarget.text,
        index: resolvedTarget.index,
        ...(resolvedTarget.paragraphId ? { paragraphId: resolvedTarget.paragraphId } : {}),
        fingerprint: resolvedTarget.fingerprint,
        inTable: resolvedTarget.inTable,
        revisionView: resolvedTarget.revisionView
    };
}

const source = documentXml([
    'Opening source paragraph.',
    'Independent middle paragraph.',
    'Independent final paragraph.'
]);
const inspected = preflightOperations(source, [
    { type: 'redline', target: { index: 1 }, modified: 'unused' },
    { type: 'redline', target: { index: 3 }, modified: 'unused' }
], 'Compiler Test');
assert.equal(inspected.valid, true);
const openingTarget = descriptor(inspected.results[0].resolvedTarget);
const finalTarget = descriptor(inspected.results[1].resolvedTarget);
const split = {
    operationId: 'split-opening',
    type: 'redline',
    target: openingTarget,
    modified: 'Opening source paragraph revised.\nInserted paragraph.'
};
const editFinal = {
    operationId: 'edit-final',
    type: 'redline',
    target: finalTarget,
    modified: 'Independent final paragraph revised.'
};

for (const operations of [[split, editFinal], [editFinal, split]]) {
    const result = await applyOperationsToDocumentXml(source, operations, 'Compiler Test', null, {
        atomic: true,
        continueOnError: true,
        strictTargets: true,
        generateRedlines: false
    });
    assert.equal(result.status, 'ok');
    assert.equal(result.hasChanges, true);
    assert.deepEqual(result.results.map(item => item.status), ['applied', 'applied']);
    assert.deepEqual(paragraphTexts(result.documentXml), [
        'Opening source paragraph revised.',
        'Inserted paragraph.',
        'Independent middle paragraph.',
        'Independent final paragraph revised.'
    ]);
}

const conflicting = await applyOperationsToDocumentXml(source, [
    { operationId: 'state-a', type: 'redline', target: finalTarget, modified: 'Final state A.' },
    { operationId: 'state-b', type: 'redline', target: finalTarget, modified: 'Final state B.' }
], 'Compiler Test', null, { atomic: false, strictTargets: true, generateRedlines: false });
assert.equal(conflicting.status, 'error');
assert.equal(conflicting.hasChanges, false);
assert.equal(conflicting.documentXml, source);
assert.equal(conflicting.error.code, 'OVERLAPPING_SOURCE_TARGETS');
assert.equal(conflicting.error.category, 'source_conflict');
assert.deepEqual(conflicting.error.operationIndexes, [1, 2]);
assert.deepEqual(conflicting.results.map(item => item.status), ['error', 'error']);
assert.deepEqual(conflicting.executionOrder, []);
assert.equal(conflicting.retryPlan.base, 'original');
assert.equal(conflicting.retryPlan.replayWholeBatch, true);

const conflictPreflight = preflightOperations(source, [
    { type: 'redline', target: finalTarget, modified: 'Final state A.' },
    { type: 'redline', target: finalTarget, modified: 'Final state B.' }
], 'Compiler Test');
assert.equal(conflictPreflight.valid, false);
assert.equal(conflictPreflight.conflicts[0].code, 'OVERLAPPING_SOURCE_TARGETS');
assert.equal(conflictPreflight.conflicts[0].category, 'source_conflict');

const createdParagraphText = 'Paragraph created by the structural rewrite.';
const createdDependencyOperations = [
    {
        operationId: 'format-created-paragraph',
        type: 'paragraph-format',
        target: { exactText: createdParagraphText },
        properties: { alignment: 'center' }
    },
    {
        operationId: 'create-paragraph',
        type: 'redline',
        target: openingTarget,
        modified: `Opening source paragraph revised.\n${createdParagraphText}`
    }
];
const createdDependencyPreflight = preflightOperations(
    source,
    createdDependencyOperations,
    'Compiler Test'
);
assert.equal(createdDependencyPreflight.valid, true);
assert.equal(createdDependencyPreflight.results[0].status, 'deferred');
assert.equal(createdDependencyPreflight.results[0].createdByOperation, 2);

const createdDependencyResult = await applyOperationsToDocumentXml(
    source,
    createdDependencyOperations,
    'Compiler Test',
    null,
    { atomic: true, strictTargets: true, generateRedlines: false }
);
assert.equal(createdDependencyResult.status, 'ok');
assert.deepEqual(createdDependencyResult.executionOrder, [2, 1]);
assert.deepEqual(createdDependencyResult.results.map(item => item.status), ['applied', 'applied']);
assert.deepEqual(paragraphTexts(createdDependencyResult.documentXml), [
    'Opening source paragraph revised.',
    createdParagraphText,
    'Independent middle paragraph.',
    'Independent final paragraph.'
]);

const createdFanout = [
    createdDependencyOperations[1],
    createdDependencyOperations[0],
    {
        operationId: 'highlight-same-created-paragraph',
        type: 'highlight',
        target: { exactText: createdParagraphText },
        textToHighlight: 'created',
        color: 'yellow'
    }
];
const createdFanoutPreflight = preflightOperations(source, createdFanout, 'Compiler Test');
assert.equal(createdFanoutPreflight.valid, false);
assert.equal(createdFanoutPreflight.error.code, 'CAPTURE_FANOUT_CONFLICT');
const createdFanoutResult = await applyOperationsToDocumentXml(
    source,
    createdFanout,
    'Compiler Test',
    null,
    { atomic: true, strictTargets: true, generateRedlines: false }
);
assert.equal(createdFanoutResult.status, 'error');
assert.equal(createdFanoutResult.error.code, 'CAPTURE_FANOUT_CONFLICT');
assert.equal(createdFanoutResult.hasChanges, false);

const savepointResult = await applyOperationsToDocumentXml(source, [
    { type: 'redline', target: openingTarget, modified: 'Opening source paragraph revised.' },
    {
        type: 'comment',
        target: { exactText: 'Independent middle paragraph.', index: 2 },
        textToComment: 'missing anchor',
        commentContent: 'This operation should fail and restore its savepoint.'
    },
    { type: 'redline', target: finalTarget, modified: 'Independent final paragraph revised.' }
], 'Compiler Test', null, {
    atomic: false,
    continueOnError: true,
    strictTargets: true,
    generateRedlines: false
});
assert.equal(savepointResult.status, 'partial');
assert.deepEqual(savepointResult.executionOrder, [2, 1, 3]);
assert.deepEqual(savepointResult.results.map(item => item.status), ['applied', 'error', 'applied']);
assert.equal(savepointResult.results[1].error.code, 'ANCHOR_NOT_FOUND');
assert.deepEqual(paragraphTexts(savepointResult.documentXml), [
    'Opening source paragraph revised.',
    'Independent middle paragraph.',
    'Independent final paragraph revised.'
]);
assert.equal(savepointResult.retryPlan.base, 'output');
assert.deepEqual(savepointResult.retryPlan.committedIndexes, [1, 3]);
assert.deepEqual(savepointResult.retryPlan.failedIndexes, [2]);

console.log('batch source binding tests passed');
