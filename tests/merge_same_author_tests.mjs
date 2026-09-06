import './setup-xml-provider.mjs';

import assert from 'node:assert/strict';
import { DOMParser } from '@xmldom/xmldom';
import {
    applyRedlineToOxml,
    acceptTrackedChangesInOoxml,
    rejectTrackedChangesInOoxml,
    ingestWordOoxmlToPlainText,
    containsTrackedChanges,
    getTrackedChangeAuthors
} from '../index.js';
import {
    applyOperationToDocumentXml,
    applyOperationsToDocumentXml
} from '../services/standalone-operation-runner.js';
import { preflightOperations } from '../services/operation-preflight.js';
import { validateRedlineOoxml } from '../core/redline-validation.js';

const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

function createDocumentXml(bodyInner) {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${NS_W}">
  <w:body>
    ${bodyInner}
    <w:sectPr/>
  </w:body>
</w:document>`;
}

// ---------------------------------------------------------------------------
// 1. Successive editing turns by same author merge against baseline
// ---------------------------------------------------------------------------
{
    const initialXml = `<w:p xmlns:w="${NS_W}"><w:r><w:t>The quick brown fox jumps over the lazy dog.</w:t></w:r></w:p>`;

    // Turn 1: Alice changes "brown fox" to "red fox"
    const turn1 = await applyRedlineToOxml(
        initialXml,
        'The quick brown fox jumps over the lazy dog.',
        'The quick red fox jumps over the lazy dog.',
        { author: 'Alice' }
    );
    assert.equal(turn1.status, 'ok');
    assert.equal(turn1.hasChanges, true);
    assert.ok(turn1.oxml.includes('w:author="Alice"'));

    // Verify Turn 1 accept/reject
    const turn1Accepted = acceptTrackedChangesInOoxml(turn1.oxml, { author: 'Alice' });
    const turn1Rejected = rejectTrackedChangesInOoxml(turn1.oxml, { author: 'Alice' });
    assert.equal(ingestWordOoxmlToPlainText(turn1Accepted.oxml), 'The quick red fox jumps over the lazy dog.');
    assert.equal(ingestWordOoxmlToPlainText(turn1Rejected.oxml), 'The quick brown fox jumps over the lazy dog.');

    // Turn 2: Alice changes "red fox" to "green fox" and adds "sleeping " before "dog"
    const turn2 = await applyRedlineToOxml(
        turn1.oxml,
        'The quick red fox jumps over the lazy dog.',
        'The quick green fox jumps over the lazy sleeping dog.',
        { author: 'Alice' }
    );
    assert.equal(turn2.status, 'ok');
    assert.equal(turn2.hasChanges, true);

    // Intermediate "red" should NOT be present anywhere in the OOXML!
    assert.ok(!turn2.oxml.includes('red'), 'Intermediate draft text "red" must not remain in revision markup');

    // Rejection MUST revert all the way to the original baseline "brown", NOT the intermediate "red"!
    const turn2Rejected = rejectTrackedChangesInOoxml(turn2.oxml, { author: 'Alice' });
    assert.equal(
        ingestWordOoxmlToPlainText(turn2Rejected.oxml),
        'The quick brown fox jumps over the lazy dog.',
        'Rejecting Turn 2 revisions must restore original pre-revision baseline'
    );

    // Acceptance MUST yield Turn 2 desired text
    const turn2Accepted = acceptTrackedChangesInOoxml(turn2.oxml, { author: 'Alice' });
    assert.equal(
        ingestWordOoxmlToPlainText(turn2Accepted.oxml),
        'The quick green fox jumps over the lazy sleeping dog.',
        'Accepting Turn 2 revisions must yield final modified text'
    );

    // Turn 3: Alice changes her mind and reverts back to the original baseline
    const turn3 = await applyRedlineToOxml(
        turn2.oxml,
        'The quick green fox jumps over the lazy sleeping dog.',
        'The quick brown fox jumps over the lazy dog.',
        { author: 'Alice' }
    );
    assert.equal(turn3.hasChanges, true);
    // Turn 3 output has no tracked changes because it reverted cleanly to baseline
    assert.ok(!turn3.oxml.includes('<w:ins'), 'Reverting to baseline should leave no ins tags');
    assert.ok(!turn3.oxml.includes('<w:del'), 'Reverting to baseline should leave no del tags');
    assert.equal(ingestWordOoxmlToPlainText(turn3.oxml), 'The quick brown fox jumps over the lazy dog.');
}

// ---------------------------------------------------------------------------
// 2. Full document multi-turn operations by same author
// ---------------------------------------------------------------------------
{
    const initialDoc = createDocumentXml(`
        <w:p w:paraId="P1">
            <w:r><w:t>Clause 1: Initial terms and conditions.</w:t></w:r>
        </w:p>
        <w:p w:paraId="P2">
            <w:r><w:t>Clause 2: Payment within 30 days.</w:t></w:r>
        </w:p>
    `);

    // Turn 1
    const res1 = await applyOperationToDocumentXml(initialDoc, {
        type: 'replace',
        target: 'Clause 2: Payment within 30 days.',
        modified: 'Clause 2: Payment within 45 days.'
    }, 'Editor');
    assert.equal(res1.hasChanges, true);
    assert.ok(res1.documentXml.includes('w:author="Editor"'));

    // Turn 2: Editor edits the same clause again
    const res2 = await applyOperationToDocumentXml(res1.documentXml, {
        type: 'replace',
        target: 'Clause 2: Payment within 45 days.',
        modified: 'Clause 2: Payment within 60 calendar days.'
    }, 'Editor');
    assert.equal(res2.hasChanges, true);
    assert.ok(!res2.documentXml.includes('45'), 'Intermediate "45 days" should not appear in document markup');

    // Rejecting the document restores "30 days"
    const rejectedDoc = rejectTrackedChangesInOoxml(res2.documentXml, { author: 'Editor' });
    assert.ok(ingestWordOoxmlToPlainText(rejectedDoc.oxml).includes('Clause 2: Payment within 30 days.'));
    assert.ok(!ingestWordOoxmlToPlainText(rejectedDoc.oxml).includes('60 calendar days.'));

    // Accepting the document yields "60 calendar days"
    const acceptedDoc = acceptTrackedChangesInOoxml(res2.documentXml, { author: 'Editor' });
    assert.ok(ingestWordOoxmlToPlainText(acceptedDoc.oxml).includes('Clause 2: Payment within 60 calendar days.'));
    assert.ok(!ingestWordOoxmlToPlainText(acceptedDoc.oxml).includes('within 30 days'));
}

// ---------------------------------------------------------------------------
// 3. Different author protection under default merge-same-author
// ---------------------------------------------------------------------------
{
    const docWithAliceEdit = createDocumentXml(`
        <w:p w:paraId="P1">
            <w:r><w:t xml:space="preserve">Standard </w:t></w:r>
            <w:del w:id="1" w:author="Alice" w:date="2026-01-01T00:00:00Z">
                <w:r><w:delText>old</w:delText></w:r>
            </w:del>
            <w:ins w:id="2" w:author="Alice" w:date="2026-01-01T00:00:00Z">
                <w:r><w:t>new</w:t></w:r>
            </w:ins>
            <w:r><w:t xml:space="preserve"> clause.</w:t></w:r>
        </w:p>
    `);

    // Preflight by Bob: fails with EXISTING_REVISIONS
    const preflightBob = preflightOperations(docWithAliceEdit, [{
        type: 'replace',
        target: 'Standard new clause.',
        modified: 'Standard updated clause.'
    }], 'Bob');
    assert.equal(preflightBob.valid, false);
    assert.equal(preflightBob.results[0].error.code, 'EXISTING_REVISIONS');
    assert.ok(preflightBob.results[0].error.message.includes('Alice'));

    // Apply by Bob: fails with EXISTING_REVISIONS
    const applyBob = await applyOperationToDocumentXml(docWithAliceEdit, {
        type: 'replace',
        target: 'Standard new clause.',
        modified: 'Standard updated clause.'
    }, 'Bob');
    assert.equal(applyBob.hasChanges, false);
    assert.equal(applyBob.status, 'error');
    assert.equal(applyBob.error.code, 'EXISTING_REVISIONS');
    assert.equal(applyBob.documentXml, docWithAliceEdit);

    // Apply by Bob with explicit existingRevisions: 'accept-all-first': succeeds!
    const applyBobAccept = await applyOperationToDocumentXml(docWithAliceEdit, {
        type: 'replace',
        target: 'Standard new clause.',
        modified: 'Standard updated clause.',
        existingRevisions: 'accept-all-first'
    }, 'Bob');
    assert.equal(applyBobAccept.hasChanges, true);
    assert.ok(applyBobAccept.documentXml.includes('w:author="Bob"'));
    assert.ok(!applyBobAccept.documentXml.includes('w:author="Alice"'));
}

// ---------------------------------------------------------------------------
// 4. Mixed authors in target paragraph protection
// ---------------------------------------------------------------------------
{
    const docWithMixedAuthors = createDocumentXml(`
        <w:p w:paraId="P1">
            <w:r><w:t xml:space="preserve">Clause with </w:t></w:r>
            <w:ins w:id="1" w:author="Alice" w:date="2026-01-01T00:00:00Z">
                <w:r><w:t>Alice edit </w:t></w:r>
            </w:ins>
            <w:ins w:id="2" w:author="Bob" w:date="2026-01-01T00:00:00Z">
                <w:r><w:t>and Bob edit.</w:t></w:r>
            </w:ins>
        </w:p>
    `);

    // Neither Alice nor Bob can merge when multiple authors' revisions are present
    const preflightAlice = preflightOperations(docWithMixedAuthors, [{
        type: 'replace',
        target: 'Clause with Alice edit and Bob edit.',
        modified: 'Clause with further edit.'
    }], 'Alice');
    assert.equal(preflightAlice.valid, false);
    assert.equal(preflightAlice.results[0].error.code, 'EXISTING_REVISIONS');

    const applyAlice = await applyOperationToDocumentXml(docWithMixedAuthors, {
        type: 'replace',
        target: 'Clause with Alice edit and Bob edit.',
        modified: 'Clause with further edit.'
    }, 'Alice');
    assert.equal(applyAlice.hasChanges, false);
    assert.equal(applyAlice.status, 'error');
    assert.equal(applyAlice.error.code, 'EXISTING_REVISIONS');
}

// ---------------------------------------------------------------------------
// 5. Batch operation with multiple edits to same paragraph in single run
// ---------------------------------------------------------------------------
{
    const initialDoc = createDocumentXml(`
        <w:p w:paraId="P1">
            <w:r><w:t>The product warranty shall last 12 months from purchase.</w:t></w:r>
        </w:p>
    `);

    const batch = [
        {
            type: 'replace',
            target: 'The product warranty shall last 12 months from purchase.',
            modified: 'The product warranty shall last 24 months from purchase.',
            author: 'Reviewer'
        },
        {
            type: 'replace',
            target: 'The product warranty shall last 24 months from purchase.',
            modified: 'The product warranty shall last 36 months from original purchase date.',
            author: 'Reviewer'
        }
    ];

    const result = await applyOperationsToDocumentXml(initialDoc, batch, 'Reviewer');
    const accepted = acceptTrackedChangesInOoxml(result.documentXml, { author: 'Reviewer' });
    const rejected = rejectTrackedChangesInOoxml(result.documentXml, { author: 'Reviewer' });
    assert.ok(ingestWordOoxmlToPlainText(accepted.oxml).includes('The product warranty shall last 36 months from original purchase date.'));
    assert.ok(ingestWordOoxmlToPlainText(rejected.oxml).includes('The product warranty shall last 12 months from purchase.'));
    assert.ok(!result.documentXml.includes('24 months'), 'Intermediate step in batch must be merged');

    const validation = validateRedlineOoxml(result.documentXml);
    assert.equal(validation.valid, true);
}

// ---------------------------------------------------------------------------
// 6. Same-author merging refuses commented revision content
// ---------------------------------------------------------------------------
{
    const commentedRevision = `<w:p xmlns:w="${NS_W}">
        <w:r><w:t xml:space="preserve">Base </w:t></w:r>
        <w:ins w:id="1" w:author="Alice">
            <w:commentRangeStart w:id="7"/>
            <w:r><w:t>inserted</w:t></w:r>
            <w:commentRangeEnd w:id="7"/>
            <w:r><w:commentReference w:id="7"/></w:r>
        </w:ins>
        <w:r><w:t xml:space="preserve"> end</w:t></w:r>
    </w:p>`;

    const direct = await applyRedlineToOxml(
        commentedRevision,
        'Base inserted end',
        'Base revised end',
        { author: 'Alice' }
    );
    assert.equal(direct.status, 'error');
    assert.equal(direct.error?.code, 'COMMENTED_CONTENT_MERGE');
    assert.deepEqual(direct.error?.commentIds, ['7']);
    assert.equal(direct.oxml, commentedRevision);

    const documentXml = createDocumentXml(commentedRevision);
    const operation = {
        type: 'replace',
        target: 'Base inserted end',
        modified: 'Base revised end',
        author: 'Alice'
    };
    const preflight = preflightOperations(documentXml, [operation], 'Alice');
    assert.equal(preflight.valid, false);
    assert.equal(preflight.results[0].error.code, 'COMMENTED_CONTENT_MERGE');

    const applied = await applyOperationToDocumentXml(documentXml, operation, 'Alice');
    assert.equal(applied.status, 'error');
    assert.equal(applied.error?.code, 'COMMENTED_CONTENT_MERGE');
    assert.equal(applied.documentXml, documentXml);
}

// ---------------------------------------------------------------------------
// 7. Empty accepted-view paragraphs cannot bypass third-party protection
// ---------------------------------------------------------------------------
{
    const documentXml = createDocumentXml(`
        <w:p w:paraId="P1">
            <w:del w:id="1" w:author="Alice">
                <w:r><w:delText>Pending deletion.</w:delText></w:r>
            </w:del>
        </w:p>
    `);
    const operation = {
        type: 'replace',
        target: { paragraphId: 'P1' },
        modified: '',
        author: 'Bob'
    };

    const preflight = preflightOperations(documentXml, [operation], 'Bob');
    assert.equal(preflight.valid, false);
    assert.equal(preflight.results[0].error.code, 'EXISTING_REVISIONS');

    const applied = await applyOperationToDocumentXml(documentXml, operation, 'Bob');
    assert.equal(applied.status, 'error');
    assert.equal(applied.error?.code, 'EXISTING_REVISIONS');
    assert.equal(applied.hasChanges, false);
    assert.equal(applied.documentXml, documentXml);

    const sameAuthorOperation = { ...operation, author: 'Alice' };
    const sameAuthorApplied = await applyOperationToDocumentXml(documentXml, sameAuthorOperation, 'Alice');
    assert.equal(sameAuthorApplied.hasChanges, true);
    assert.ok(!sameAuthorApplied.documentXml.includes('w:id="1"'), 'prior deletion must be replaced, not layered');
    const rejected = rejectTrackedChangesInOoxml(sameAuthorApplied.documentXml, { author: 'Alice' });
    assert.ok(ingestWordOoxmlToPlainText(rejected.oxml).includes('Pending deletion.'));
}

// ---------------------------------------------------------------------------
// 8. Property revisions participate in author detection and protection
// ---------------------------------------------------------------------------
{
    const tableCellXml = `<w:tc xmlns:w="${NS_W}">
        <w:tcPr>
            <w:tcPrChange w:id="2" w:author="Bob"><w:tcPr/></w:tcPrChange>
        </w:tcPr>
        <w:p><w:r><w:t>Hello</w:t></w:r></w:p>
    </w:tc>`;
    const parsed = new DOMParser().parseFromString(tableCellXml, 'text/xml');
    assert.equal(containsTrackedChanges(parsed), true);
    assert.deepEqual(getTrackedChangeAuthors(parsed), ['Bob']);

    const applied = await applyRedlineToOxml(tableCellXml, 'Hello', 'Goodbye', { author: 'Alice' });
    assert.equal(applied.status, 'error');
    assert.equal(applied.error?.code, 'EXISTING_REVISIONS');
    assert.equal(applied.oxml, tableCellXml);

    const unsupportedCellRevision = `<w:tc xmlns:w="${NS_W}">
        <w:tcPr><w:cellIns w:id="3" w:author="Alice"/></w:tcPr>
        <w:p><w:r><w:t>Hello</w:t></w:r></w:p>
    </w:tc>`;
    const sameAuthor = await applyRedlineToOxml(
        unsupportedCellRevision,
        'Hello',
        'Goodbye',
        { author: 'Alice' }
    );
    assert.equal(sameAuthor.status, 'error');
    assert.equal(sameAuthor.error?.code, 'UNSAFE_REVISION_NESTING');
    assert.equal(sameAuthor.oxml, unsupportedCellRevision);
}

console.log('PASS: merge_same_author_tests.mjs');
