import assert from 'node:assert/strict';

import './setup-xml-provider.mjs';
import { parseOoxmlSafe } from '../adapters/xml-adapter.js';
import { projectWordRevisionView } from '../core/word-revision-semantics.js';
import {
    applyOperationToDocumentXml,
    applyOperationsToDocumentXml
} from '../services/standalone-operation-runner.js';
import {
    createFidelityViews,
    verifyNoCreatedRevisions,
    verifyNumberingReferences,
    verifyUntrackedLocality
} from '../services/revision-fidelity-verifier.js';

/**
 * The session reuses each verified operation's after-state projections as the next operation's baseline.
 * These cases pin that the reuse is invalidated whenever the live DOM is not the one that was projected:
 * an untracked edit between tracked ones, a refused operation (restored savepoint), and an unverified commit.
 * A stale baseline would show up as a false REVISION_FIDELITY_VIOLATION or a missed one, so the batch outcome
 * is compared with the same operations applied one at a time, each in a fresh session with no cache.
 */

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const paragraph = (id, text) => `<w:p w14:paraId="${id}"><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
const SOURCE = `<w:document xmlns:w="${W}" xmlns:w14="${W14}" xmlns:r="${R}"><w:body>`
    + paragraph('0A000001', 'Alpha clause stays here.')
    + paragraph('0A000002', 'Bravo clause is edited without tracking.')
    + paragraph('0A000003', 'Charlie clause is edited with tracking.')
    + '<w:p w14:paraId="0A000004"><w:hyperlink r:id="rId9"><w:r><w:t xml:space="preserve">B. </w:t></w:r></w:hyperlink><w:r><w:t>Header</w:t></w:r></w:p>'
    + paragraph('0A000005', 'Echo clause is edited after the refusal.')
    + paragraph('0A000006', 'Foxtrot clause is edited last.')
    + paragraph('0A000007', 'Golf clause closes the batch.')
    + '<w:sectPr/></w:body></w:document>';

const target = (id, text) => ({ paragraphId: id, exactText: text });
const operations = [
    { type: 'redline', target: target('0A000001', 'Alpha clause stays here.'), replacements: [{ find: 'stays', replace: 'changes' }] },
    { type: 'redline', target: target('0A000002', 'Bravo clause is edited without tracking.'), replacements: [{ find: 'edited', replace: 'revised' }], generateRedlines: false },
    { type: 'redline', target: target('0A000003', 'Charlie clause is edited with tracking.'), replacements: [{ find: 'edited', replace: 'amended' }] },
    { type: 'redline', target: { exactText: 'B. Header' }, modified: 'B. Header' },
    { type: 'redline', target: target('0A000005', 'Echo clause is edited after the refusal.'), replacements: [{ find: 'edited', replace: 'handled' }] },
    { type: 'redline', target: target('0A000006', 'Foxtrot clause is edited last.'), replacements: [{ find: 'last', replace: 'finally' }], generateRedlines: false },
    { type: 'redline', target: target('0A000007', 'Golf clause closes the batch.'), replacements: [{ find: 'closes', replace: 'ends' }] }
];

const views = xml => {
    const parsed = parseOoxmlSafe(xml);
    assert.ok(parsed.doc, parsed.error?.message);
    return ['accepted', 'rejected'].map(view => projectWordRevisionView(parsed.doc, view).map(record => record.text));
};
const statusOf = result => (result.status === 'error' ? (result.error?.code || 'error') : 'applied');

// Batch with a shared session (cached views) versus fresh sessions (no cache).
{
    const options = { author: 'Cache Probe' };
    const batch = await applyOperationsToDocumentXml(SOURCE, operations, 'Cache Probe', null, { ...options });
    const batchStatuses = batch.results.map(statusOf);
    assert.equal(batchStatuses[3], 'REVISION_FIDELITY_VIOLATION', JSON.stringify(batchStatuses));

    let documentXml = SOURCE;
    const sequentialStatuses = [];
    for (const operation of operations) {
        const result = await applyOperationToDocumentXml(documentXml, operation, 'Cache Probe', null, { ...options });
        sequentialStatuses.push(statusOf(result));
        if (result.status !== 'error' && result.hasChanges) documentXml = result.documentXml;
    }
    assert.deepEqual(batchStatuses, sequentialStatuses, 'cached batch verifies exactly like uncached single operations');
    assert.deepEqual(views(batch.documentXml), views(documentXml), 'cached batch output matches uncached output in both revision views');
    // The operation after the refusal and the untracked edits all landed.
    const [accepted, rejected] = views(batch.documentXml);
    assert.ok(accepted.includes('Alpha clause changes here.') && accepted.includes('Golf clause ends the batch.'), JSON.stringify(accepted));
    assert.ok(accepted.includes('Bravo clause is revised without tracking.'));
    assert.ok(accepted.includes('Echo clause is handled after the refusal.'));
    assert.ok(accepted.includes('Foxtrot clause is edited finally.'));
    assert.ok(rejected.includes('Charlie clause is edited with tracking.'));
    assert.ok(rejected.includes('Alpha clause stays here.') && rejected.includes('Golf clause closes the batch.'));
}

// Verifier-level memo: views fill lazily, are reused, and hold plain data only.
{
    const parse = xml => parseOoxmlSafe(xml).doc;
    const base = parse(SOURCE);
    const edited = parse(SOURCE.replace('Bravo clause is edited', 'Bravo clause is changed'));
    const beforeViews = createFidelityViews(base);
    const afterViews = createFidelityViews(edited);
    const first = verifyUntrackedLocality({ before: base, after: edited, targetSourceIndex: 1, beforeViews, afterViews });
    assert.equal(first.ok, true);
    assert.ok(beforeViews.accepted && beforeViews.rejected && afterViews.accepted && afterViews.rejected);
    const memo = afterViews.accepted;
    // The after-state becomes the next baseline and verifies an edit elsewhere without being recomputed.
    const next = parse(SOURCE.replace('Bravo clause is edited', 'Bravo clause is changed').replace('Echo clause', 'Echo text'));
    assert.equal(verifyUntrackedLocality({ before: edited, after: next, targetSourceIndex: 4, beforeViews: afterViews, afterViews: createFidelityViews(next) }).ok, true);
    assert.equal(afterViews.accepted, memo, 'the baseline projection is reused, not recomputed');
    // The same baseline against an edit outside the target is a violation, so reuse cannot hide one.
    const neighbor = verifyUntrackedLocality({ before: edited, after: next, targetSourceIndex: 0, beforeViews: afterViews, afterViews: createFidelityViews(next) });
    assert.equal(neighbor.ok, false);
    assert.equal(neighbor.violations[0].invariant, 'ACCEPT_NEIGHBORS');

    // Revision-id signatures are plain strings (stable after the DOM they came from changes).
    const ins = id => `<w:ins w:id="${id}" w:author="${id === 5 ? 'Ann' : 'Bob'}" w:date="2026-01-01T00:00:00Z"><w:r><w:t>x</w:t></w:r></w:ins>`;
    const withRevision = parse(SOURCE.replace('<w:sectPr/>', `<w:p>${ins(5)}</w:p><w:sectPr/>`));
    const revisionViews = createFidelityViews(withRevision);
    assert.equal(verifyNoCreatedRevisions({ before: withRevision, after: withRevision, beforeViews: revisionViews, afterViews: createFidelityViews(withRevision) }).ok, true);
    assert.ok(revisionViews.revisionSignatures instanceof Map);
    assert.ok([...revisionViews.revisionSignatures.values()].every(value => typeof value === 'string'));
    const created = parse(SOURCE.replace('<w:sectPr/>', `<w:p>${ins(5)}${ins(6).replace('Bob', 'Cy').replace('<w:ins ', '<w:del ').replace('</w:ins>', '</w:del>')}</w:p><w:sectPr/>`));
    assert.equal(verifyNoCreatedRevisions({ before: withRevision, after: created, beforeViews: revisionViews, afterViews: createFidelityViews(created) }).ok, false);

    // Numbering: ids the source already used are not re-checked, and a reused numbering part gives the same answer.
    const withList = parse(SOURCE.replace('<w:p w14:paraId="0A000001">', '<w:p w14:paraId="0A000001"><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="3"/></w:numPr></w:pPr>'));
    const numbering = `<w:numbering xmlns:w="${W}"><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"/></w:abstractNum><w:num w:numId="3"><w:abstractNumId w:val="1"/></w:num></w:numbering>`;
    assert.equal(verifyNumberingReferences({ before: base, after: withList, numberingXmls: [numbering] }).ok, true);
    assert.equal(verifyNumberingReferences({ before: base, after: withList, numberingXmls: [numbering] }).ok, true);
    assert.equal(verifyNumberingReferences({ before: base, after: withList, numberingXmls: [null] }).ok, false);
    assert.equal(verifyNumberingReferences({ before: withList, after: withList, numberingXmls: [null], beforeViews: createFidelityViews(withList), afterViews: createFidelityViews(withList) }).ok, true);
}

console.log('PASS: revision fidelity view cache tests');
