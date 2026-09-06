import './setup-xml-provider.mjs';
import assert from 'node:assert/strict';
import {
    applyOperationsToDocumentXml,
    applyOperationToDocumentXml
} from '../services/standalone-operation-runner.js';
import {
    reconcileReceiptsAgainstOutput,
    createEmptyReceipt
} from '../services/receipt-collector.js';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const sampleDocXml = `<w:document xmlns:w="${W}"><w:body><w:p><w:r><w:t>First paragraph.</w:t></w:r></w:p><w:p><w:r><w:t>Second paragraph.</w:t></w:r></w:p><w:p><w:r><w:t>Third paragraph.</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`;

// 1. Successful atomic batch
{
    const ops = [
        {
            operationId: 'op-1',
            type: 'redline',
            target: 'First paragraph.',
            modified: 'First paragraph updated.',
            generateRedlines: true
        },
        {
            operationId: 'op-2',
            type: 'format',
            target: 'Second paragraph.',
            textToFormat: 'Second',
            properties: { bold: true },
            generateRedlines: true
        }
    ];

    const result = await applyOperationsToDocumentXml(sampleDocXml, ops, 'Author', null, { atomic: true });
    assert.equal(result.hasChanges, true);
    assert.equal(result.rolledBack, undefined);
    assert.ok(Array.isArray(result.receipts), 'result.receipts must be an array');
    assert.equal(result.receipts.length, 2);

    const r1 = result.receipts[0];
    assert.equal(r1.operationIndex, 1);
    assert.equal(r1.operationId, 'op-1');
    assert.equal(r1.attemptedDisposition, 'applied');
    assert.equal(r1.finalDisposition, 'applied');
    assert.equal(r1.committed, true);
    assert.ok(r1.revisionItems.length > 0);

    const r2 = result.receipts[1];
    assert.equal(r2.operationIndex, 2);
    assert.equal(r2.operationId, 'op-2');
    assert.equal(r2.attemptedDisposition, 'applied');
    assert.equal(r2.finalDisposition, 'applied');
    assert.equal(r2.committed, true);

    // Each result item also has receipt attached
    assert.equal(result.results[0].receipt.committed, true);
    assert.equal(result.results[1].receipt.committed, true);
}

// 2. Successful non-atomic batch
{
    const ops = [
        {
            operationId: 'na-1',
            type: 'redline',
            target: 'First paragraph.',
            modified: 'First non-atomic.',
            generateRedlines: true
        },
        {
            operationId: 'na-2',
            type: 'redline',
            target: 'Second paragraph.',
            modified: 'Second non-atomic.',
            generateRedlines: true
        }
    ];

    const result = await applyOperationsToDocumentXml(sampleDocXml, ops, 'Author', null, { atomic: false });
    assert.equal(result.hasChanges, true);
    assert.equal(result.receipts.length, 2);
    assert.equal(result.receipts[0].committed, true);
    assert.equal(result.receipts[1].committed, true);
}

// 3. Early stop and continue-on-error in atomic mode
{
    const ops = [
        {
            operationId: 'step-1',
            type: 'redline',
            target: 'First paragraph.',
            modified: 'Will be rolled back.',
            generateRedlines: true
        },
        {
            operationId: 'step-2',
            type: 'redline',
            target: 'Non-existent paragraph target.',
            modified: 'Will fail.'
        },
        {
            operationId: 'step-3',
            type: 'redline',
            target: 'Third paragraph.',
            modified: 'Will not run.'
        }
    ];

    const result = await applyOperationsToDocumentXml(sampleDocXml, ops, 'Author', null, {
        atomic: true,
        continueOnError: false
    });

    assert.equal(result.rolledBack, true);
    assert.equal(result.hasChanges, false);
    assert.equal(result.receipts.length, 3);

    // Step 1: attempted and succeeded, but rolled back
    const r1 = result.receipts.find(r => r.operationIndex === 1);
    assert.equal(r1.attemptedDisposition, 'applied');
    assert.equal(r1.finalDisposition, 'rolled_back');
    assert.equal(r1.committed, false);

    // Step 2: failed
    const r2 = result.receipts.find(r => r.operationIndex === 2);
    assert.equal(r2.attemptedDisposition, 'refused');
    assert.equal(r2.finalDisposition, 'refused');
    assert.equal(r2.committed, false);

    // Step 3: not attempted
    const r3 = result.receipts.find(r => r.operationIndex === 3);
    assert.equal(r3.attemptedDisposition, 'not_attempted');
    assert.equal(r3.finalDisposition, 'not_attempted');
    assert.equal(r3.committed, false);
}

// 4. Early stop and continue-on-error in non-atomic mode
{
    const ops = [
        {
            operationId: 'step-1',
            type: 'redline',
            target: 'First paragraph.',
            modified: 'First preserved.',
            generateRedlines: true
        },
        {
            operationId: 'step-2',
            type: 'redline',
            target: 'Non-existent paragraph target.',
            modified: 'Will fail.'
        },
        {
            operationId: 'step-3',
            type: 'redline',
            target: 'Third paragraph.',
            modified: 'Will not run.'
        }
    ];

    const result = await applyOperationsToDocumentXml(sampleDocXml, ops, 'Author', null, {
        atomic: false,
        continueOnError: false
    });

    assert.equal(result.hasChanges, true);
    assert.equal(result.receipts.length, 3);

    // Step 1: applied and committed
    const r1 = result.receipts.find(r => r.operationIndex === 1);
    assert.equal(r1.attemptedDisposition, 'applied');
    assert.equal(r1.finalDisposition, 'applied');
    assert.equal(r1.committed, true);

    // Step 2: failed / refused
    const r2 = result.receipts.find(r => r.operationIndex === 2);
    assert.equal(r2.attemptedDisposition, 'refused');
    assert.equal(r2.finalDisposition, 'refused');
    assert.equal(r2.committed, false);

    // Step 3: not attempted
    const r3 = result.receipts.find(r => r.operationIndex === 3);
    assert.equal(r3.attemptedDisposition, 'not_attempted');
    assert.equal(r3.finalDisposition, 'not_attempted');
    assert.equal(r3.committed, false);
}

// 5. Continue on error in non-atomic mode
{
    const ops = [
        {
            operationId: 'c-1',
            type: 'redline',
            target: 'First paragraph.',
            modified: 'First applied.',
            generateRedlines: true
        },
        {
            operationId: 'c-2',
            type: 'redline',
            target: 'Non-existent paragraph target.',
            modified: 'Fails.'
        },
        {
            operationId: 'c-3',
            type: 'redline',
            target: 'Third paragraph.',
            modified: 'Third applied.',
            generateRedlines: true
        }
    ];

    const result = await applyOperationsToDocumentXml(sampleDocXml, ops, 'Author', null, {
        atomic: false,
        continueOnError: true
    });

    assert.equal(result.hasChanges, true);
    assert.equal(result.receipts.length, 3);
    assert.equal(result.receipts[0].committed, true);
    assert.equal(result.receipts[0].finalDisposition, 'applied');
    assert.equal(result.receipts[1].committed, false);
    assert.equal(result.receipts[1].finalDisposition, 'refused');
    assert.equal(result.receipts[2].committed, true);
    assert.equal(result.receipts[2].finalDisposition, 'applied');
}

// 6. Later failure rolling back earlier success in atomic batch
{
    const ops = [
        {
            operationId: 'later-fail-1',
            type: 'redline',
            target: 'First paragraph.',
            modified: 'First paragraph changed.',
            generateRedlines: true
        },
        {
            operationId: 'later-fail-2',
            type: 'comment',
            target: 'Missing paragraph.',
            textToComment: 'Missing',
            commentContent: 'Cannot place'
        }
    ];

    const result = await applyOperationsToDocumentXml(sampleDocXml, ops, 'Author', null, { atomic: true });
    assert.equal(result.rolledBack, true);
    assert.equal(result.hasChanges, false);
    assert.equal(result.documentXml, sampleDocXml);
    assert.equal(result.receipts[0].finalDisposition, 'rolled_back');
    assert.equal(result.receipts[0].committed, false);
    assert.equal(result.receipts[1].finalDisposition, 'refused');
    assert.equal(result.receipts[1].committed, false);
}

// 7. Mixed text/comment/list/format operations
{
    const ops = [
        {
            operationId: 'mix-comment',
            type: 'comment',
            target: 'First paragraph.',
            textToComment: 'First',
            commentContent: 'Review intro'
        },
        {
            operationId: 'mix-format',
            type: 'format',
            target: 'Second paragraph.',
            textToFormat: 'Second',
            properties: { italic: true },
            generateRedlines: true
        },
        {
            operationId: 'mix-list',
            type: 'redline',
            target: 'Third paragraph.',
            modified: '1. First item\n2. Second item',
            generateRedlines: true
        }
    ];

    const result = await applyOperationsToDocumentXml(sampleDocXml, ops, 'Author', null, { atomic: true });
    assert.equal(result.hasChanges, true);
    assert.equal(result.receipts.length, 3);

    const commentReceipt = result.receipts.find(r => r.operationId === 'mix-comment');
    assert.ok(commentReceipt.commentIds.length > 0, 'Comment receipt must record comment ID');
    assert.equal(commentReceipt.committed, true);

    const formatReceipt = result.receipts.find(r => r.operationId === 'mix-format');
    assert.ok(formatReceipt.revisionItems.some(i => i.kind === 'rPrChange'), 'Format receipt must record rPrChange');
    assert.equal(formatReceipt.committed, true);

    const listReceipt = result.receipts.find(r => r.operationId === 'mix-list');
    assert.ok(listReceipt.numberingIds.length > 0, 'List receipt must record numbering ID');
    assert.equal(listReceipt.committed, true);
}

// 8. Existing IDs are never incorrectly reported as newly allocated
{
    const docWithExistingRevision = `<w:document xmlns:w="${W}"><w:body><w:p><w:ins w:id="99" w:author="PriorAuthor" w:date="2026-01-01T00:00:00Z"><w:r><w:t>Existing revision text.</w:t></w:r></w:ins></w:p><w:p><w:r><w:t>Target paragraph.</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`;

    const result = await applyOperationToDocumentXml(
        docWithExistingRevision,
        {
            type: 'redline',
            target: 'Target paragraph.',
            modified: 'Target paragraph modified.',
            generateRedlines: true
        },
        'NewAuthor'
    );

    assert.equal(result.hasChanges, true);
    assert.ok(result.receipt);
    assert.equal(result.receipt.committed, true);
    const hasExistingId = result.receipt.revisionItems.some(item => item.id === '99');
    assert.equal(hasExistingId, false, 'Pre-existing revision ID 99 must not appear in receipt');
}

// 9. Reconciliation failure detection
{
    const fakeReceipt = {
        operationIndex: 1,
        attemptedDisposition: 'applied',
        finalDisposition: 'applied',
        committed: true,
        revisionItems: [{ id: '999999', kind: 'ins', partName: 'word/document.xml' }],
        commentIds: [],
        numberingIds: [],
        relationshipIds: [],
        affectedTargets: [],
        warnings: []
    };

    const reconciliation = reconcileReceiptsAgainstOutput({
        documentXml: sampleDocXml
    }, [fakeReceipt]);

    assert.equal(reconciliation.valid, false);
    assert.equal(reconciliation.error.code, 'RECEIPT_RECONCILIATION_FAILED');
    assert.ok(reconciliation.error.message.includes('999999'));
}

console.log('All MutationReceipt tests passed successfully.');
