import './setup-xml-provider.mjs';

import assert from 'node:assert/strict';
import { parseOoxmlSafe } from '../adapters/xml-adapter.js';
import { getParagraphText } from '../core/paragraph-targeting.js';
import {
    applyOperationsToDocumentXml,
    applyOperationToDocumentXml,
    preflightOperations
} from '../services/standalone-operation-runner.js';

const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const NS_W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';

function documentXml(paragraphs) {
    const body = paragraphs.map(({ text, id, wrapper = null }) => {
        const run = `<w:r><w:t>${text}</w:t></w:r>`;
        return `<w:p w14:paraId="${id}">${wrapper ? wrapper(run) : run}</w:p>`;
    }).join('');
    return `<w:document xmlns:w="${NS_W}" xmlns:w14="${NS_W14}"><w:body>${body}</w:body></w:document>`;
}

function paragraphTexts(xml) {
    const parsed = parseOoxmlSafe(xml, 'application/xml');
    assert.equal(parsed.error, null);
    return Array.from(parsed.doc.getElementsByTagNameNS(NS_W, 'p')).map(getParagraphText);
}

async function testPerOperationAuthorsAndMetadata() {
    const input = documentXml([
        { id: 'AAA00001', text: 'Comment target paragraph.' },
        { id: 'AAA00002', text: 'Old replacement paragraph.' }
    ]);
    const result = await applyOperationsToDocumentXml(input, [
        {
            type: 'comment',
            target: { paragraphId: 'AAA00001', exactText: 'Comment target paragraph.' },
            commentContent: 'Review the entire paragraph.',
            author: 'Comment Author'
        },
        {
            type: 'replace',
            target: { paragraphId: 'AAA00002', exactText: 'Old replacement paragraph.' },
            modified: 'New replacement paragraph.',
            author: 'Edit Author'
        }
    ], 'Batch Fallback', null, { generateRedlines: true });

    assert.equal(result.hasChanges, true);
    assert.deepEqual(result.authorsUsed, ['Comment Author', 'Edit Author']);
    assert.deepEqual(result.results.map(item => item.authorUsed), ['Comment Author', 'Edit Author']);
    assert.deepEqual(result.results.map(item => item.resolvedBy), ['paragraph_id', 'paragraph_id']);
    assert.deepEqual(result.results.map(item => item.resolvedTarget.paragraphId), ['AAA00001', 'AAA00002']);
    assert.match(result.commentsXml, /w:author="Comment Author"/);
    assert.match(result.documentXml, /w:author="Edit Author"/);
}

async function testRuntimeOperationValidation() {
    const input = documentXml([{ id: 'BBB00001', text: 'Untouched.' }]);
    const single = await applyOperationToDocumentXml(
        input,
        { type: 'unsupported', target: 'Untouched.' },
        'Validator'
    );
    assert.equal(single.status, 'error');
    assert.equal(single.error.code, 'INVALID_OPERATION');
    assert.equal(single.documentXml, input);

    const batch = await applyOperationsToDocumentXml(
        input,
        [{ type: 'comment', target: 'Untouched.', commentContent: '' }],
        'Validator'
    );
    assert.equal(batch.rolledBack, true);
    assert.equal(batch.results[0].error.code, 'INVALID_OPERATION');
    assert.equal(batch.documentXml, input);

    const wholeParagraphComment = await applyOperationToDocumentXml(
        input,
        {
            type: 'comment',
            target: { paragraphId: 'BBB00001' },
            commentContent: 'Whole-paragraph comment.'
        },
        'Commenter'
    );
    assert.equal(wholeParagraphComment.hasChanges, true);
    assert.match(wholeParagraphComment.commentsXml, /Whole-paragraph comment\./);
}

async function testStrictAmbiguityAndDescriptors() {
    const input = documentXml([
        { id: 'CCC00001', text: 'Repeated boilerplate.' },
        { id: 'CCC00002', text: 'Repeated boilerplate.' }
    ]);
    const ambiguous = await applyOperationToDocumentXml(
        input,
        { type: 'replace', target: 'Repeated boilerplate.', modified: 'Changed.' },
        'Strict Editor',
        null,
        { generateRedlines: false, strictTargets: true }
    );
    assert.equal(ambiguous.status, 'error');
    assert.equal(ambiguous.error.code, 'AMBIGUOUS_TARGET');
    assert.equal(ambiguous.error.candidates.length, 2);
    assert.equal(ambiguous.documentXml, input);

    const occurrence = await applyOperationToDocumentXml(
        input,
        {
            type: 'replace',
            target: { exactText: 'Repeated boilerplate.', occurrence: 2 },
            modified: 'Second changed.'
        },
        'Strict Editor',
        null,
        { generateRedlines: false, strictTargets: true }
    );
    assert.equal(occurrence.status, 'ok');
    assert.equal(occurrence.resolvedBy, 'occurrence');
    assert.deepEqual(paragraphTexts(occurrence.documentXml), ['Repeated boilerplate.', 'Second changed.']);

    assert.equal(occurrence.resolvedTarget.fingerprint.startsWith('fnv1a32:'), true);

    const sourcePreflight = preflightOperations(input, [{
        type: 'replace',
        target: { exactText: 'Repeated boilerplate.', occurrence: 2 },
        modified: 'Fingerprint changed.'
    }], 'Strict Editor');
    const sourceFingerprint = sourcePreflight.results[0].resolvedTarget.fingerprint;
    const byFingerprint = await applyOperationToDocumentXml(
        input,
        {
            type: 'replace',
            target: { exactText: 'Repeated boilerplate.', fingerprint: sourceFingerprint },
            modified: 'Fingerprint changed.'
        },
        'Strict Editor',
        null,
        { generateRedlines: false, strictTargets: true }
    );
    assert.equal(byFingerprint.resolvedBy, 'fingerprint');
    assert.deepEqual(paragraphTexts(byFingerprint.documentXml), ['Repeated boilerplate.', 'Fingerprint changed.']);

    const staleFingerprint = await applyOperationToDocumentXml(
        input,
        {
            type: 'replace',
            target: { exactText: 'Repeated boilerplate.', fingerprint: 'fnv1a32:00000000' },
            modified: 'Must not apply.'
        },
        'Strict Editor',
        null,
        { generateRedlines: false, strictTargets: true }
    );
    assert.equal(staleFingerprint.error.code, 'TARGET_FINGERPRINT_MISMATCH');
    assert.equal(staleFingerprint.documentXml, input);
}

function testPreflightDiagnosticsAndConflicts() {
    const revision = attributes => run => `<w:ins ${attributes}>${run}</w:ins>`;
    const input = documentXml([
        { id: 'DDD00001', text: 'Repeated clause.' },
        { id: 'DDD00002', text: 'Repeated clause.' },
        {
            id: 'DDD00003',
            text: 'Revised clause.',
            wrapper: revision('w:id="71" w:author="Prior" w:date="2026-01-01T00:00:00Z"')
        },
        { id: 'DDD00004', text: 'List source.' }
    ]);
    const operations = [
        { type: 'replace', target: 'Repeated clause.', modified: 'Ambiguous edit.', author: 'Editor A' },
        { type: 'comment', target: { paragraphId: 'DDD00003' }, textToComment: 'missing', commentContent: 'Check.', author: 'Reviewer' },
        { type: 'replace', target: { paragraphId: 'DDD00003' }, modified: 'Changed revised clause.', author: 'Editor B' },
        { type: 'replace', target: { paragraphId: 'DDD00004' }, modified: '1. First item', author: 'Editor B' },
        { type: 'highlight', target: { paragraphId: 'DDD00004' }, textToHighlight: 'List source.', color: 'yellow', author: 'Reviewer' }
    ];
    const before = input;
    const result = preflightOperations(input, operations, 'Fallback');

    assert.equal(input, before, 'preflight must not mutate the source string');
    assert.equal(result.valid, false);
    assert.equal(result.results[0].error.code, 'AMBIGUOUS_TARGET');
    assert.equal(result.results[0].error.candidates.length, 2);
    assert.equal(result.results[1].error.code, 'ANCHOR_NOT_FOUND');
    assert.equal(result.results[2].error.code, 'EXISTING_REVISIONS');
    assert.equal(result.results[2].hasRevisions, true);
    assert.deepEqual(result.requiredArtifacts, { comments: true, numbering: true });
    assert.deepEqual(result.authorsUsed, ['Editor A', 'Reviewer', 'Editor B']);
    assert.equal(result.conflicts.some(conflict => conflict.code === 'REVISION_ORDER_CONFLICT'), true);
    assert.equal(result.results[3].resolvedTarget.fingerprint.startsWith('fnv1a32:'), true);
}

await testPerOperationAuthorsAndMetadata();
await testRuntimeOperationValidation();
await testStrictAmbiguityAndDescriptors();
testPreflightDiagnosticsAndConflicts();

console.log('agent_operation_contract_tests.mjs ... PASS');
