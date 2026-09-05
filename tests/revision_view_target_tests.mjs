import assert from 'node:assert/strict';
import './setup-xml-provider.mjs';
import {
    preflightOperations,
    applyOperationToDocumentXml
} from '../services/standalone-operation-runner.js';
import { validateDocumentOperation } from '../services/document-operation-contract.js';
import {
    resolveTargetParagraph,
    buildParagraphMetadataIndex
} from '../core/paragraph-targeting.js';
import { parseOoxmlSafe } from '../adapters/xml-adapter.js';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

function createDocumentXml(paragraphsXml) {
    return `<w:document xmlns:w="${W}"><w:body>${paragraphsXml}</w:body></w:document>`;
}

// Test 1: Same phrase at different accepted/rejected locations
{
    // Paragraph 1 has an insertion containing "Target phrase"
    // Paragraph 2 has a deletion containing "Target phrase"
    const docXml = createDocumentXml(`
        <w:p w:paraId="P1">
            <w:r><w:t>P1 preamble </w:t></w:r>
            <w:ins w:id="1" w:author="Alice"><w:r><w:t>Target phrase</w:t></w:r></w:ins>
        </w:p>
        <w:p w:paraId="P2">
            <w:r><w:t>P2 preamble </w:t></w:r>
            <w:del w:id="2" w:author="Bob"><w:r><w:delText>Target phrase</w:delText></w:r></w:del>
        </w:p>
    `);

    // In accepted view, "Target phrase" is in P1
    const acceptedPreflight = preflightOperations(docXml, [
        {
            type: 'comment',
            target: { exactText: 'P1 preamble Target phrase', revisionView: 'accepted' },
            commentContent: 'Accepted note'
        }
    ], 'Reviewer');
    assert.equal(acceptedPreflight.valid, true);
    assert.equal(acceptedPreflight.results[0].status, 'ready');
    assert.equal(acceptedPreflight.results[0].resolvedTarget.index, 1);
    assert.equal(acceptedPreflight.results[0].resolvedTarget.paragraphId, 'P1');
    assert.equal(acceptedPreflight.results[0].resolvedTarget.revisionView, 'accepted');

    // In rejected view, "Target phrase" is in P2
    const rejectedPreflight = preflightOperations(docXml, [
        {
            type: 'comment',
            target: { exactText: 'P2 preamble Target phrase', revisionView: 'rejected' },
            commentContent: 'Rejected note'
        }
    ], 'Reviewer');
    assert.equal(rejectedPreflight.valid, true);
    assert.equal(rejectedPreflight.results[0].status, 'ready');
    assert.equal(rejectedPreflight.results[0].resolvedTarget.index, 2);
    assert.equal(rejectedPreflight.results[0].resolvedTarget.paragraphId, 'P2');
    assert.equal(rejectedPreflight.results[0].resolvedTarget.revisionView, 'rejected');
}

// Test 2: Phrase exists only in an insertion
{
    const docXml = createDocumentXml(`
        <w:p w:paraId="P1">
            <w:r><w:t>Baseline intro. </w:t></w:r>
            <w:ins w:id="3" w:author="Alice"><w:r><w:t>Newly added text.</w:t></w:r></w:ins>
        </w:p>
    `);

    // Accepted view finds it
    const accepted = preflightOperations(docXml, [
        {
            type: 'comment',
            target: { exactText: 'Baseline intro. Newly added text.', revisionView: 'accepted' },
            commentContent: 'Comment on new text'
        }
    ], 'Reviewer');
    assert.equal(accepted.valid, true);
    assert.equal(accepted.results[0].status, 'ready');

    // Rejected view does NOT find it
    const rejected = preflightOperations(docXml, [
        {
            type: 'comment',
            target: { exactText: 'Baseline intro. Newly added text.', revisionView: 'rejected' },
            commentContent: 'Comment on new text'
        }
    ], 'Reviewer');
    assert.equal(rejected.valid, false);
    assert.equal(rejected.results[0].status, 'error');
    assert.equal(rejected.results[0].error.code, 'TARGET_NOT_FOUND');
}

// Test 3: Phrase exists only in a deletion
{
    const docXml = createDocumentXml(`
        <w:p w:paraId="P1">
            <w:r><w:t>Baseline intro. </w:t></w:r>
            <w:del w:id="4" w:author="Bob"><w:r><w:delText>Historical deleted text.</w:delText></w:r></w:del>
        </w:p>
    `);

    // Rejected view finds it
    const rejected = preflightOperations(docXml, [
        {
            type: 'comment',
            target: { exactText: 'Baseline intro. Historical deleted text.', revisionView: 'rejected' },
            commentContent: 'Comment on deleted text'
        }
    ], 'Reviewer');
    assert.equal(rejected.valid, true);
    assert.equal(rejected.results[0].status, 'ready');
    assert.equal(rejected.results[0].resolvedTarget.paragraphId, 'P1');

    // Accepted view does NOT find it
    const accepted = preflightOperations(docXml, [
        {
            type: 'comment',
            target: { exactText: 'Baseline intro. Historical deleted text.', revisionView: 'accepted' },
            commentContent: 'Comment on deleted text'
        }
    ], 'Reviewer');
    assert.equal(accepted.valid, false);
    assert.equal(accepted.results[0].status, 'error');
    assert.equal(accepted.results[0].error.code, 'TARGET_NOT_FOUND');
}

// Test 4: Duplicate phrase in one view but unique in the other
{
    // In accepted view:
    // P1: "Duplicate section" (baseline)
    // P2: "Duplicate section" (inserted by Alice: <w:ins><w:t>Duplicate section</w:t></w:ins>)
    // In rejected view:
    // P1: "Duplicate section"
    // P2: "" (insertion is omitted)
    const docXml = createDocumentXml(`
        <w:p w:paraId="P1">
            <w:r><w:t>Duplicate section</w:t></w:r>
        </w:p>
        <w:p w:paraId="P2">
            <w:ins w:id="5" w:author="Alice"><w:r><w:t>Duplicate section</w:t></w:r></w:ins>
        </w:p>
    `);

    // In accepted view: AMBIGUOUS_TARGET without occurrence
    const acceptedAmbiguous = preflightOperations(docXml, [
        {
            type: 'comment',
            target: { exactText: 'Duplicate section', revisionView: 'accepted' },
            commentContent: 'Note'
        }
    ], 'Reviewer');
    assert.equal(acceptedAmbiguous.valid, false);
    assert.equal(acceptedAmbiguous.results[0].status, 'error');
    assert.equal(acceptedAmbiguous.results[0].error.code, 'AMBIGUOUS_TARGET');

    // In rejected view: Unique! (P2 has no text in rejected view)
    const rejectedUnique = preflightOperations(docXml, [
        {
            type: 'comment',
            target: { exactText: 'Duplicate section', revisionView: 'rejected' },
            commentContent: 'Note'
        }
    ], 'Reviewer');
    assert.equal(rejectedUnique.valid, true);
    assert.equal(rejectedUnique.results[0].status, 'ready');
    assert.equal(rejectedUnique.results[0].resolvedTarget.index, 1);
    assert.equal(rejectedUnique.results[0].resolvedTarget.paragraphId, 'P1');

    // Occurrence in accepted view: occurrence 2 resolves to P2
    const acceptedOccur2 = preflightOperations(docXml, [
        {
            type: 'comment',
            target: { exactText: 'Duplicate section', occurrence: 2, revisionView: 'accepted' },
            commentContent: 'Note'
        }
    ], 'Reviewer');
    assert.equal(acceptedOccur2.valid, true);
    assert.equal(acceptedOccur2.results[0].resolvedTarget.index, 2);

    // Occurrence in rejected view: occurrence 2 does not exist!
    const rejectedOccur2 = preflightOperations(docXml, [
        {
            type: 'comment',
            target: { exactText: 'Duplicate section', occurrence: 2, revisionView: 'rejected' },
            commentContent: 'Note'
        }
    ], 'Reviewer');
    assert.equal(rejectedOccur2.valid, false);
    assert.equal(rejectedOccur2.results[0].error.code, 'TARGET_NOT_FOUND');
}

// Test 5: Descriptor fields disagree despite one individually matching
{
    const docXml = createDocumentXml(`
        <w:p w:paraId="Alpha1">
            <w:r><w:t>First section content.</w:t></w:r>
        </w:p>
        <w:p w:paraId="Beta2">
            <w:r><w:t>Second section content.</w:t></w:r>
        </w:p>
    `);

    // Disagreement 1: paragraphId matches Alpha1, but index is 2 (Beta2's index)
    const disagreeIndex = preflightOperations(docXml, [
        {
            type: 'comment',
            target: { paragraphId: 'Alpha1', index: 2 },
            commentContent: 'Note'
        }
    ], 'Reviewer');
    assert.equal(disagreeIndex.valid, false);
    assert.equal(disagreeIndex.results[0].error.code, 'TARGET_INDEX_MISMATCH');

    // Disagreement 2: paragraphId matches Alpha1, but exactText is Beta2's text
    const disagreeText = preflightOperations(docXml, [
        {
            type: 'comment',
            target: { paragraphId: 'Alpha1', exactText: 'Second section content.' },
            commentContent: 'Note'
        }
    ], 'Reviewer');
    assert.equal(disagreeText.valid, false);
    assert.equal(disagreeText.results[0].error.code, 'TARGET_TEXT_MISMATCH');

    // Disagreement 3: paragraphId matches Alpha1, but inTable is true
    const disagreeContext = preflightOperations(docXml, [
        {
            type: 'comment',
            target: { paragraphId: 'Alpha1', inTable: true },
            commentContent: 'Note'
        }
    ], 'Reviewer');
    assert.equal(disagreeContext.valid, false);
    assert.equal(disagreeContext.results[0].error.code, 'TARGET_CONTEXT_MISMATCH');
}

// Test 6: Invalid view is rejected by runtime contract and JSON schema
{
    const invalidViewOp = {
        type: 'comment',
        target: { exactText: 'Something', revisionView: 'original' }, // only 'accepted' or 'rejected' allowed
        commentContent: 'Note'
    };
    const validation = validateDocumentOperation(invalidViewOp);
    assert.equal(validation.valid, false);
    assert.equal(validation.error.code, 'INVALID_OPERATION');
    assert.match(validation.error.message, /must be "accepted" or "rejected"/i);

    // Verify JSON schema definition
    const { readFile } = await import('node:fs/promises');
    const schemaText = await readFile(new URL('../docs/schemas/document-operations.schema.json', import.meta.url), 'utf8');
    const schema = JSON.parse(schemaText);
    const targetProps = schema.$defs.target.oneOf[1].properties;
    assert(targetProps.revisionView, 'schema must define revisionView in target properties');
    assert.deepEqual(targetProps.revisionView.enum, ['accepted', 'rejected']);
}

// Test 7: apply refuses revisionView: 'rejected' with UNSUPPORTED_REVISION_VIEW_MUTATION
{
    const docXml = createDocumentXml(`
        <w:p w:paraId="P1">
            <w:del w:id="1" w:author="Old"><w:r><w:delText>Obsolete text</w:delText></w:r></w:del>
        </w:p>
    `);

    const result = await applyOperationToDocumentXml(docXml, {
        type: 'redline',
        target: { exactText: 'Obsolete text', revisionView: 'rejected' },
        modified: 'Revived text'
    }, 'Editor');

    assert.equal(result.hasChanges, false);
    assert.equal(result.status, 'error');
    assert.equal(result.error.code, 'UNSUPPORTED_REVISION_VIEW_MUTATION');
    assert.match(result.error.message, /not supported yet/i);
}

console.log('revision view target tests passed');
