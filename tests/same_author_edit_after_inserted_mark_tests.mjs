/**
 * A paragraph whose mark is a same-author pending insertion (Word marks the ORIGINAL paragraph's
 * mark as inserted when paragraphs are added after it) keeps its own text in the merge baseline,
 * so a later same-author delete or edit of that paragraph tracks the text instead of dropping it.
 */
import './setup-xml-provider.mjs';
import assert from 'node:assert/strict';
import {
    acceptTrackedChangesInOoxml,
    parseOoxmlSafe,
    rejectTrackedChangesInOoxml
} from '../index.js';
import { applyOperationsToDocumentXml } from '../services/standalone-operation-runner.js';
import { projectWordRevisionView } from '../core/word-revision-semantics.js';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const p = text => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
const source = `<w:document xmlns:w="${W}"><w:body>${p('Before.')}${p('Old provision.')}${p('After.')}</w:body></w:document>`;
const options = {
    atomic: true,
    structuredContent: true,
    pairReplacements: true,
    generateRedlines: true,
    strictTargets: true,
    onInfo: () => {},
    onWarn: () => {}
};
const NEW_PARAGRAPHS = ['New opening:', 'Point one.', 'Point two.', 'Point three.'];

function viewTexts(documentXml, view) {
    const { doc } = parseOoxmlSafe(documentXml, 'text/xml');
    return projectWordRevisionView(doc, view).map(record => record.text);
}

function resolvedTexts(documentXml, resolve) {
    const resolved = resolve(documentXml, { allAuthors: true });
    assert.notEqual(resolved.status, 'error');
    const { doc } = parseOoxmlSafe(resolved.oxml, 'text/xml');
    return projectWordRevisionView(doc, 'accepted').map(record => record.text);
}

async function insertAfterOldProvision() {
    const result = await applyOperationsToDocumentXml(source, [{
        type: 'redline',
        target: { exactText: 'Old provision.', index: 2 },
        modified: ['Old provision.', ...NEW_PARAGRAPHS].join('\n'),
        structuredContent: false
    }], 'A', {}, options);
    assert.equal(result.error, undefined, 'stage 1 inserts the new paragraphs');
    assert.match(
        result.documentXml,
        /<w:pPr><w:rPr><w:ins [^>]*\/><\/w:rPr><\/w:pPr><w:r><w:t>Old provision\.<\/w:t>/,
        'the original paragraph carries the inserted mark'
    );
    return result.documentXml;
}

// Same-author delete of the paragraph whose mark is a pending insertion.
{
    const staged = await insertAfterOldProvision();
    const result = await applyOperationsToDocumentXml(staged, [{
        type: 'delete',
        target: { exactText: 'Old provision.', index: 2 }
    }], 'A', {}, options);
    assert.equal(result.error, undefined, JSON.stringify(result.results?.[0]?.error?.message));
    assert.equal(result.results[0].status, 'applied');

    assert.deepEqual(viewTexts(result.documentXml, 'rejected'), ['Before.', 'Old provision.', 'After.']);
    assert.deepEqual(viewTexts(result.documentXml, 'accepted'), ['Before.', ...NEW_PARAGRAPHS, 'After.']);
    assert.deepEqual(resolvedTexts(result.documentXml, rejectTrackedChangesInOoxml), ['Before.', 'Old provision.', 'After.']);
    assert.deepEqual(resolvedTexts(result.documentXml, acceptTrackedChangesInOoxml), ['Before.', ...NEW_PARAGRAPHS, 'After.']);

    // The deleted mark keeps its pending insertion, in schema order (ins before del).
    assert.match(result.documentXml, /<w:rPr><w:ins [^>]*w:author="A"[^>]*\/><w:del [^>]*w:author="A"[^>]*\/><\/w:rPr>/);
    assert.match(result.documentXml, /<w:delText[^>]*>Old provision\.<\/w:delText>/);
}

// Same-author text edit of that paragraph keeps the inserted mark and tracks the change.
{
    const staged = await insertAfterOldProvision();
    const result = await applyOperationsToDocumentXml(staged, [{
        type: 'redline',
        target: { exactText: 'Old provision.', index: 2 },
        modified: 'Old clause.',
        structuredContent: false
    }], 'A', {}, options);
    assert.equal(result.error, undefined, JSON.stringify(result.results?.[0]?.error?.message));

    assert.deepEqual(viewTexts(result.documentXml, 'rejected'), ['Before.', 'Old provision.', 'After.']);
    assert.deepEqual(viewTexts(result.documentXml, 'accepted'), ['Before.', 'Old clause.', ...NEW_PARAGRAPHS, 'After.']);
}

// Another author is still refused by the existing-revisions policy.
{
    const staged = await insertAfterOldProvision();
    const result = await applyOperationsToDocumentXml(staged, [{
        type: 'delete',
        target: { exactText: 'Old provision.', index: 2 }
    }], 'B', {}, options);
    assert.equal(result.error?.code, 'BATCH_OPERATION_FAILED');
    assert.equal(result.results[0].error?.code, 'EXISTING_REVISIONS');
    assert.equal(result.documentXml, staged);
}

// Deleting one of the author's own pending inserted paragraphs removes it (no empty paragraph is
// left for Accept All), whether it carries an inserted mark or holds the original paragraph's mark.
for (const [deleted, kept] of [
    ['New opening:', ['Point one.', 'Point two.', 'Point three.']],
    ['Point two.', ['New opening:', 'Point one.', 'Point three.']],
    ['Point three.', ['New opening:', 'Point one.', 'Point two.']]
]) {
    const staged = await insertAfterOldProvision();
    const result = await applyOperationsToDocumentXml(staged, [{
        type: 'delete',
        target: { exactText: deleted }
    }], 'A', {}, options);
    assert.equal(result.error, undefined, `${deleted}: ${JSON.stringify(result.results?.[0]?.error?.message)}`);
    const accepted = ['Before.', 'Old provision.', ...kept, 'After.'];
    assert.deepEqual(viewTexts(result.documentXml, 'accepted'), accepted, deleted);
    assert.deepEqual(viewTexts(result.documentXml, 'rejected'), ['Before.', 'Old provision.', 'After.'], deleted);
    assert.deepEqual(resolvedTexts(result.documentXml, acceptTrackedChangesInOoxml), accepted, deleted);
    assert.deepEqual(resolvedTexts(result.documentXml, rejectTrackedChangesInOoxml), ['Before.', 'Old provision.', 'After.'], deleted);
    assert.doesNotMatch(result.documentXml, /delText/, 'an own pending insertion is taken back, not tracked as deleted');
}

// Deleting every inserted paragraph returns the document to its source.
{
    let xml = await insertAfterOldProvision();
    for (const text of NEW_PARAGRAPHS) {
        const result = await applyOperationsToDocumentXml(xml, [{ type: 'delete', target: { exactText: text } }], 'A', {}, options);
        assert.equal(result.error, undefined, `${text}: ${JSON.stringify(result.results?.[0]?.error?.message)}`);
        xml = result.documentXml;
    }
    assert.doesNotMatch(xml, /<w:(ins|del)\b/);
    assert.deepEqual(viewTexts(xml, 'accepted'), ['Before.', 'Old provision.', 'After.']);
}

// Two edits of one paragraph in one batch: the second merges the first one's revisions, and the
// receipts describe the final document instead of failing reconciliation.
{
    const result = await applyOperationsToDocumentXml(source, [
        { type: 'redline', target: { exactText: 'Old provision.' }, modified: 'Old clause.', structuredContent: false },
        { type: 'redline', target: { exactText: 'Old clause.' }, modified: 'Older clause.', structuredContent: false }
    ], 'A', {}, options);
    assert.equal(result.error, undefined, `${result.error?.code}: ${result.error?.message}`);
    assert.deepEqual(viewTexts(result.documentXml, 'accepted'), ['Before.', 'Older clause.', 'After.']);
    assert.deepEqual(viewTexts(result.documentXml, 'rejected'), ['Before.', 'Old provision.', 'After.']);
    const liveIds = new Set([...result.documentXml.matchAll(/<w:(?:ins|del)\b[^>]*\bw:id="(\d+)"/g)].map(match => match[1]));
    const receipts = result.receipts || result.results.map(item => item.receipt).filter(Boolean);
    assert.ok(receipts.length === 2, 'both operations report a receipt');
    for (const receipt of receipts) {
        for (const item of receipt.revisionItems) assert.ok(liveIds.has(item.id), `receipt revision ${item.id} exists in the output`);
    }
    assert.equal(receipts[0].revisionItems.length, 0);
    assert.ok(receipts[1].revisionItems.length > 0);
}

console.log('PASS: same-author edit after inserted paragraph mark');
