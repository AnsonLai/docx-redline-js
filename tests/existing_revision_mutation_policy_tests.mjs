import './setup-xml-provider.mjs';

import assert from 'assert/strict';
import {
    applyOperationToDocumentXml,
    applyOperationsToDocumentXml
} from '../services/standalone-operation-runner.js';
import { preflightOperations } from '../services/operation-preflight.js';
import { acceptTrackedChangesInOoxml, rejectTrackedChangesInOoxml } from '../index.js';
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
// 1. Baseline Text: Normal tracked mutations and comments
// ---------------------------------------------------------------------------
{
    const docXml = createDocumentXml(`
        <w:p w:paraId="P1">
            <w:r><w:t>Baseline sentence for editing.</w:t></w:r>
        </w:p>
    `);

    // Comment on baseline text
    const commentRes = await applyOperationToDocumentXml(docXml, {
        type: 'comment',
        target: 'Baseline sentence for editing.',
        textToComment: 'sentence',
        commentContent: 'Review baseline wording.'
    }, 'Reviewer');
    assert.equal(commentRes.hasChanges, true);
    assert.ok(commentRes.documentXml.includes('<w:commentRangeStart'));
    assert.ok(commentRes.commentsXml.includes('Review baseline wording.'));

    // Redline replacement on baseline text
    const redlineRes = await applyOperationToDocumentXml(docXml, {
        type: 'replace',
        target: 'Baseline sentence for editing.',
        modified: 'Baseline sentence for revision.'
    }, 'Editor');
    assert.equal(redlineRes.hasChanges, true);
    assert.ok(redlineRes.documentXml.includes('<w:ins'));
    assert.ok(redlineRes.documentXml.includes('<w:del'));
}

// ---------------------------------------------------------------------------
// 2. Pending insertion, comment only: Same author and Different author
// ---------------------------------------------------------------------------
{
    // Document with pending insertion by Alice
    const docXml = createDocumentXml(`
        <w:p w:paraId="P1">
            <w:r><w:t xml:space="preserve">Standard prefix </w:t></w:r>
            <w:ins w:id="101" w:author="Alice" w:date="2026-01-01T00:00:00Z">
                <w:r><w:t>freshly inserted clause</w:t></w:r>
            </w:ins>
            <w:r><w:t xml:space="preserve"> standard suffix.</w:t></w:r>
        </w:p>
    `);

    // Preflight check: comment entirely inside insertion is allowed
    const preflightSame = preflightOperations(docXml, [{
        type: 'comment',
        target: 'Standard prefix freshly inserted clause standard suffix.',
        textToComment: 'freshly inserted clause',
        commentContent: 'Same-author note',
        author: 'Alice'
    }], 'Alice');
    assert.equal(preflightSame.valid, true);
    assert.equal(preflightSame.results[0].status, 'ready');

    // 2a. Same author: Alice comments on her own insertion
    const sameAuthorRes = await applyOperationToDocumentXml(docXml, {
        type: 'comment',
        target: 'Standard prefix freshly inserted clause standard suffix.',
        textToComment: 'freshly inserted clause',
        commentContent: 'Review my own insertion.',
        author: 'Alice'
    }, 'Alice');
    assert.equal(sameAuthorRes.hasChanges, true);
    assert.ok(sameAuthorRes.documentXml.includes('<w:commentRangeStart'));
    assert.ok(sameAuthorRes.commentsXml.includes('Review my own insertion.'));
    // Structural validity: comment markers must be nested inside <w:ins>
    assert.match(
        sameAuthorRes.documentXml,
        /<w:ins[^>]*>[\s\S]*<w:commentRangeStart[\s\S]*<\/w:ins>/,
        'comment markers must be contained within the <w:ins> element'
    );
    const validSame = validateRedlineOoxml(sameAuthorRes.documentXml);
    assert.equal(validSame.valid, true);

    // 2b. Different author: Bob comments on Alice's insertion
    const diffAuthorRes = await applyOperationToDocumentXml(docXml, {
        type: 'comment',
        target: 'Standard prefix freshly inserted clause standard suffix.',
        textToComment: 'inserted clause',
        commentContent: 'Bob reviews Alice insertion.',
        author: 'Bob'
    }, 'Bob');
    assert.equal(diffAuthorRes.hasChanges, true);
    assert.ok(diffAuthorRes.commentsXml.includes('author="Bob"'));
    assert.match(
        diffAuthorRes.documentXml,
        /<w:ins[^>]*w:author="Alice"[^>]*>[\s\S]*<w:commentRangeStart[\s\S]*<\/w:ins>/,
        'comment markers must be contained within Alice <w:ins>'
    );
    const validDiff = validateRedlineOoxml(diffAuthorRes.documentXml);
    assert.equal(validDiff.valid, true);

    // 2c. Selective Accept: When Alice's insertion is accepted, the comment survives intact
    const acceptedXml = acceptTrackedChangesInOoxml(diffAuthorRes.documentXml, { author: 'Alice' });
    assert.ok(!acceptedXml.oxml.includes('w:author="Alice"'), 'Alice insertion must be unwrapped');
    assert.ok(acceptedXml.oxml.includes('<w:commentRangeStart'), 'Bob comment must remain anchored');
    assert.ok(acceptedXml.oxml.includes('inserted clause'), 'Accepted text must be preserved');

    // 2d. Selective Reject: When Alice's insertion is rejected, the insertion and nested comment markers are removed
    const rejectedXml = rejectTrackedChangesInOoxml(diffAuthorRes.documentXml, { author: 'Alice' });
    assert.ok(!rejectedXml.oxml.includes('freshly inserted clause'), 'Rejected text must be removed');
    assert.ok(!rejectedXml.oxml.includes('<w:commentRangeStart'), 'Comment markers on rejected text must be removed');
}

// ---------------------------------------------------------------------------
// 2e. Pending insertion: Comment spanning baseline and inserted text is ALLOWED
// ---------------------------------------------------------------------------
{
    const docXml = createDocumentXml(`
        <w:p w:paraId="P1">
            <w:r><w:t xml:space="preserve">Prefix </w:t></w:r>
            <w:ins w:id="102" w:author="Alice" w:date="2026-01-01T00:00:00Z">
                <w:r><w:t>clause</w:t></w:r>
            </w:ins>
            <w:r><w:t xml:space="preserve"> suffix.</w:t></w:r>
        </w:p>
    `);

    // Comment spanning from baseline 'Prefix ' into inserted 'clause'
    const crossingPreflight = preflightOperations(docXml, [{
        type: 'comment',
        target: 'Prefix clause suffix.',
        textToComment: 'Prefix clause',
        commentContent: 'Cross-boundary comment.'
    }], 'Bob');
    assert.equal(crossingPreflight.valid, true);

    const crossingApply = await applyOperationToDocumentXml(docXml, {
        type: 'comment',
        target: 'Prefix clause suffix.',
        textToComment: 'Prefix clause',
        commentContent: 'Cross-boundary comment.'
    }, 'Bob');
    assert.equal(crossingApply.hasChanges, true);
    assert.ok(crossingApply.documentXml.includes('<w:commentRangeStart'));
    assert.ok(crossingApply.commentsXml.includes('Cross-boundary comment.'));
    const validation = validateRedlineOoxml(crossingApply.documentXml);
    assert.equal(validation.valid, true);
}

// ---------------------------------------------------------------------------
// 3. Pending insertion, text replacement: Refused under reject-input
// ---------------------------------------------------------------------------
{
    const docXml = createDocumentXml(`
        <w:p w:paraId="P1">
            <w:ins w:id="103" w:author="Alice" w:date="2026-01-01T00:00:00Z">
                <w:r><w:t>Pending inserted text.</w:t></w:r>
            </w:ins>
        </w:p>
    `);

    // Under default 'reject-input': refused with EXISTING_REVISIONS
    const preflight = preflightOperations(docXml, [{
        type: 'replace',
        target: 'Pending inserted text.',
        modified: 'Attempted replacement text.'
    }], 'Bob');
    assert.equal(preflight.valid, false);
    assert.equal(preflight.results[0].error.code, 'EXISTING_REVISIONS');

    const applyDefault = await applyOperationToDocumentXml(docXml, {
        type: 'replace',
        target: 'Pending inserted text.',
        modified: 'Attempted replacement text.'
    }, 'Bob');
    assert.equal(applyDefault.hasChanges, false);
    assert.equal(applyDefault.status, 'error');
    assert.equal(applyDefault.error.code, 'EXISTING_REVISIONS');

    // Opt-in 'accept-all-first': normalizes prior insertion and redlines
    const applyOptIn = await applyOperationToDocumentXml(docXml, {
        type: 'replace',
        target: 'Pending inserted text.',
        modified: 'Attempted replacement text.',
        existingRevisions: 'accept-all-first'
    }, 'Bob');
    assert.equal(applyOptIn.hasChanges, true);
    assert.ok(applyOptIn.documentXml.includes('w:author="Bob"'));
}

// ---------------------------------------------------------------------------
// 4. Pending deletion, comment only: Refused with UNSAFE_REVISION_NESTING
// ---------------------------------------------------------------------------
{
    const docXml = createDocumentXml(`
        <w:p w:paraId="P1">
            <w:r><w:t>Live text and </w:t></w:r>
            <w:del w:id="104" w:author="Alice" w:date="2026-01-01T00:00:00Z">
                <w:r><w:delText>deleted text</w:delText></w:r>
            </w:del>
        </w:p>
    `);

    // Targeting deleted text in rejected view
    const preflight = preflightOperations(docXml, [{
        type: 'comment',
        target: { exactText: 'Live text and deleted text', revisionView: 'rejected' },
        textToComment: 'deleted text',
        commentContent: 'Comment on deleted text.'
    }], 'Bob');
    assert.equal(preflight.valid, false);
    assert.equal(preflight.results[0].error.code, 'UNSAFE_REVISION_NESTING');

    const apply = await applyOperationToDocumentXml(docXml, {
        type: 'comment',
        target: { exactText: 'Live text and deleted text', revisionView: 'rejected' },
        textToComment: 'deleted text',
        commentContent: 'Comment on deleted text.'
    }, 'Bob');
    assert.equal(apply.hasChanges, false);
    assert.equal(apply.status, 'error');
    // Rejected view mutation refusal or unsafe nesting
    assert.ok(
        ['UNSAFE_REVISION_NESTING', 'UNSUPPORTED_REVISION_VIEW_MUTATION'].includes(apply.error.code),
        `Expected UNSAFE_REVISION_NESTING or UNSUPPORTED_REVISION_VIEW_MUTATION, got ${apply.error.code}`
    );
}

// ---------------------------------------------------------------------------
// 5. Pending deletion, text replacement: Refused with UNSAFE_REVISION_NESTING
// ---------------------------------------------------------------------------
{
    const docXml = createDocumentXml(`
        <w:p w:paraId="P1">
            <w:r><w:t>Surviving text. </w:t></w:r>
            <w:del w:id="105" w:author="Alice" w:date="2026-01-01T00:00:00Z">
                <w:r><w:delText>Old deleted text.</w:delText></w:r>
            </w:del>
        </w:p>
    `);

    const preflight = preflightOperations(docXml, [{
        type: 'replace',
        target: 'Surviving text.',
        modified: 'Replaced text.'
    }], 'Bob');
    assert.equal(preflight.valid, false);
    assert.equal(preflight.results[0].error.code, 'UNSAFE_REVISION_NESTING');

    const apply = await applyOperationToDocumentXml(docXml, {
        type: 'replace',
        target: 'Surviving text.',
        modified: 'Replaced text.'
    }, 'Bob');
    assert.equal(apply.hasChanges, false);
    assert.equal(apply.status, 'error');
    assert.equal(apply.error.code, 'UNSAFE_REVISION_NESTING');
}

// ---------------------------------------------------------------------------
// 6. Move source/destination: Refused with UNSAFE_REVISION_NESTING
// ---------------------------------------------------------------------------
{
    const moveSourceXml = createDocumentXml(`
        <w:p w:paraId="P1">
            <w:moveFrom w:id="106" w:author="Alice" w:date="2026-01-01T00:00:00Z">
                <w:r><w:delText>Moved source text.</w:delText></w:r>
            </w:moveFrom>
        </w:p>
    `);

    // Move source: comment refused
    const commentMoveFrom = preflightOperations(moveSourceXml, [{
        type: 'comment',
        target: { exactText: 'Moved source text.', revisionView: 'rejected' },
        textToComment: 'Moved source text.',
        commentContent: 'Note.'
    }], 'Bob');
    assert.equal(commentMoveFrom.valid, false);
    assert.equal(commentMoveFrom.results[0].error.code, 'UNSAFE_REVISION_NESTING');

    // Move source: replace refused
    const replaceMoveFrom = preflightOperations(moveSourceXml, [{
        type: 'replace',
        target: { exactText: 'Moved source text.', revisionView: 'rejected' },
        modified: 'Replacement.'
    }], 'Bob');
    assert.equal(replaceMoveFrom.valid, false);
    assert.equal(replaceMoveFrom.results[0].error.code, 'UNSAFE_REVISION_NESTING');

    const moveDestXml = createDocumentXml(`
        <w:p w:paraId="P2">
            <w:moveTo w:id="107" w:author="Alice" w:date="2026-01-01T00:00:00Z">
                <w:r><w:t>Moved destination text.</w:t></w:r>
            </w:moveTo>
        </w:p>
    `);

    // Move destination: comment refused
    const commentMoveTo = preflightOperations(moveDestXml, [{
        type: 'comment',
        target: 'Moved destination text.',
        textToComment: 'destination',
        commentContent: 'Note.'
    }], 'Bob');
    assert.equal(commentMoveTo.valid, false);
    assert.equal(commentMoveTo.results[0].error.code, 'UNSAFE_REVISION_NESTING');

    // Move destination: replace refused
    const replaceMoveTo = preflightOperations(moveDestXml, [{
        type: 'replace',
        target: 'Moved destination text.',
        modified: 'Replacement.'
    }], 'Bob');
    assert.equal(replaceMoveTo.valid, false);
    assert.equal(replaceMoveTo.results[0].error.code, 'UNSAFE_REVISION_NESTING');

    const applyMoveTo = await applyOperationToDocumentXml(moveDestXml, {
        type: 'replace',
        target: 'Moved destination text.',
        modified: 'Replacement.'
    }, 'Bob');
    assert.equal(applyMoveTo.hasChanges, false);
    assert.equal(applyMoveTo.status, 'error');
    assert.equal(applyMoveTo.error.code, 'UNSAFE_REVISION_NESTING');
}

console.log('existing revision mutation policy tests passed');
