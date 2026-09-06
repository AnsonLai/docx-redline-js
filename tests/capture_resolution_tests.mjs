import './setup-xml-provider.mjs';
import assert from 'node:assert/strict';
import {
    applyOperationsToDocumentXml
} from '../services/standalone-operation-runner.js';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const sampleDocXml = `<w:document xmlns:w="${W}"><w:body><w:p><w:r><w:t>First paragraph.</w:t></w:r></w:p><w:p><w:r><w:t>Second paragraph.</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`;

// 1. Plain paragraph insertion -> comment
{
    const ops = [
        {
            operationId: 'insert-p',
            captureKey: 'new-para',
            type: 'redline',
            target: 'First paragraph.',
            modified: 'First paragraph.\n\nBrand new paragraph.'
        },
        {
            operationId: 'comment-p',
            type: 'comment',
            target: { captureRef: 'new-para', select: 'Brand new paragraph.' },
            textToComment: 'Brand new paragraph.',
            commentContent: 'Review this new paragraph'
        }
    ];

    const result = await applyOperationsToDocumentXml(sampleDocXml, ops, 'Author', null, { generateRedlines: false });
    assert.equal(result.hasChanges, true);
    assert.equal(result.rolledBack, undefined);
    assert.ok(result.commentsXml, 'Comments XML should be present');
    assert.ok(result.commentsXml.includes('Review this new paragraph'));
    assert.equal(result.results[1].status, 'applied');
    assert.equal(result.results[1].resolvedBy, 'capture');
}

// 2. Structured heading/paragraph insertion -> comments targeting :first and :last
{
    const ops = [
        {
            operationId: 'insert-struct',
            captureKey: 'section-cap',
            type: 'redline',
            target: 'Second paragraph.',
            modified: '# New Section Title\n\nNew section body text.',
            structuredContent: true
        },
        {
            operationId: 'comment-heading',
            type: 'comment',
            target: { captureRef: 'section-cap', select: ':first' },
            textToComment: 'New Section Title',
            commentContent: 'Check title wording'
        },
        {
            operationId: 'comment-body',
            type: 'comment',
            target: { captureRef: 'section-cap', select: ':last' },
            textToComment: 'New section body text.',
            commentContent: 'Check body accuracy'
        }
    ];

    const result = await applyOperationsToDocumentXml(sampleDocXml, ops, 'Author', null, { generateRedlines: false });
    assert.equal(result.hasChanges, true);
    assert.equal(result.rolledBack, undefined);
    assert.ok(result.commentsXml.includes('Check title wording'));
    assert.ok(result.commentsXml.includes('Check body accuracy'));
    assert.equal(result.results[1].resolvedBy, 'capture');
    assert.equal(result.results[2].resolvedBy, 'capture');
}

// 3. List insertion -> format one item by 1-based index
{
    const ops = [
        {
            operationId: 'insert-list',
            captureKey: 'list-cap',
            type: 'redline',
            target: 'Second paragraph.',
            modified: '1. First item\n2. Second item\n3. Third item'
        },
        {
            operationId: 'format-item-2',
            type: 'format',
            target: { captureRef: 'list-cap', select: '2' },
            textToFormat: 'Second item',
            properties: { bold: true }
        }
    ];

    const result = await applyOperationsToDocumentXml(sampleDocXml, ops, 'Author', null, { generateRedlines: false });
    assert.equal(result.hasChanges, true);
    assert.equal(result.rolledBack, undefined);
    assert.equal(result.results[1].status, 'applied');
    assert.equal(result.results[1].resolvedBy, 'capture');
    assert.ok(result.documentXml.includes('<w:b/>') || result.documentXml.includes('<w:b '), 'Bold format should be applied');
}

// 4. Table insertion -> target cell by text
{
    const ops = [
        {
            operationId: 'insert-table',
            captureKey: 'table-cap',
            type: 'redline',
            target: 'Second paragraph.',
            modified: '| HeaderA | HeaderB |\n| --- | --- |\n| CellAlpha | CellBeta |',
            structuredContent: true
        },
        {
            operationId: 'comment-cell',
            type: 'comment',
            target: { captureRef: 'table-cap', select: 'CellBeta' },
            textToComment: 'CellBeta',
            commentContent: 'Verify Beta value'
        }
    ];

    const result = await applyOperationsToDocumentXml(sampleDocXml, ops, 'Author', null, { generateRedlines: false });
    assert.equal(result.hasChanges, true);
    assert.equal(result.rolledBack, undefined);
    assert.ok(result.commentsXml.includes('Verify Beta value'));
    assert.equal(result.results[1].resolvedBy, 'capture');
}

// 5. Producer no-op -> CAPTURE_NOT_FOUND
{
    const ops = [
        {
            operationId: 'noop-op',
            captureKey: 'noop-cap',
            type: 'redline',
            target: 'First paragraph.',
            modified: 'First paragraph.'
        },
        {
            operationId: 'consumer-op',
            type: 'comment',
            target: { captureRef: 'noop-cap' },
            textToComment: 'First',
            commentContent: 'Should not run'
        }
    ];

    const result = await applyOperationsToDocumentXml(sampleDocXml, ops, 'Author', null, { generateRedlines: false, atomic: true });
    assert.equal(result.rolledBack, true);
    assert.equal(result.hasChanges, false);
    assert.equal(result.results[0].status, 'no_change');
    assert.equal(result.results[1].status, 'error');
    assert.equal(result.results[1].error.code, 'CAPTURE_NOT_FOUND');
}

// 6. Producer failure -> batch rolls back
{
    const ops = [
        {
            operationId: 'fail-producer',
            captureKey: 'fail-cap',
            type: 'redline',
            target: 'Text not in document',
            modified: 'New text'
        },
        {
            operationId: 'consumer-op',
            type: 'comment',
            target: { captureRef: 'fail-cap' },
            textToComment: 'New text',
            commentContent: 'Note'
        }
    ];

    const result = await applyOperationsToDocumentXml(sampleDocXml, ops, 'Author', null, { generateRedlines: false, atomic: true });
    assert.equal(result.rolledBack, true);
    assert.equal(result.hasChanges, false);
    assert.equal(result.results[0].status, 'error');
    assert.equal(result.results[0].error.code, 'TARGET_NOT_FOUND');
    assert.equal(result.results[1].status, 'error');
    assert.equal(result.results[1].error.code, 'CAPTURE_NOT_FOUND');
}

// 7. Consumer failure rolling back atomic batch
{
    const ops = [
        {
            operationId: 'success-producer',
            captureKey: 'succ-cap',
            type: 'redline',
            target: 'First paragraph.',
            modified: 'First paragraph.\n\nAdded paragraph.'
        },
        {
            operationId: 'fail-consumer',
            type: 'comment',
            target: { captureRef: 'succ-cap', select: 'Added paragraph.' },
            textToComment: 'Non-existent anchor text',
            commentContent: 'Failed anchor'
        }
    ];

    const result = await applyOperationsToDocumentXml(sampleDocXml, ops, 'Author', null, { generateRedlines: false, atomic: true });
    assert.equal(result.rolledBack, true);
    assert.equal(result.hasChanges, false);
    assert.equal(result.documentXml, sampleDocXml);
    assert.equal(result.commentsXml, null);
    assert.equal(result.results[0].status, 'applied');
    assert.equal(result.results[1].status, 'error');
    assert.equal(result.results[1].error.code, 'ANCHOR_NOT_FOUND');
}

// 8. Stale capture detection (CAPTURE_STALE) when intermediate op replaces captured paragraph
{
    const ops = [
        {
            operationId: 'producer',
            captureKey: 'cap-to-replace',
            type: 'redline',
            target: 'First paragraph.',
            modified: 'Temporary text that gets replaced.'
        },
        {
            operationId: 'destroyer',
            type: 'redline',
            target: 'Temporary text that gets replaced.',
            modified: 'Completely different replacement.'
        },
        {
            operationId: 'stale-consumer',
            type: 'redline',
            target: { captureRef: 'cap-to-replace' },
            modified: 'Should fail with CAPTURE_STALE'
        }
    ];

    const result = await applyOperationsToDocumentXml(sampleDocXml, ops, 'Author', null, { generateRedlines: false, atomic: true });
    assert.equal(result.rolledBack, true);
    assert.equal(result.hasChanges, false);
    assert.equal(result.results[2].status, 'error');
    assert.equal(result.results[2].error.code, 'CAPTURE_STALE');
}

// 9. Ambiguous selection in multi-paragraph capture
{
    // 9a. Multi-paragraph capture without select
    {
        const ops = [
            {
                operationId: 'prod-multi',
                captureKey: 'multi-cap',
                type: 'redline',
                target: 'First paragraph.',
                modified: 'Para Alpha.\n\nPara Beta.'
            },
            {
                operationId: 'consumer-no-select',
                type: 'comment',
                target: { captureRef: 'multi-cap' },
                textToComment: 'Para',
                commentContent: 'Should fail'
            }
        ];

        const result = await applyOperationsToDocumentXml(sampleDocXml, ops, 'Author', null, { generateRedlines: false, atomic: true });
        assert.equal(result.rolledBack, true);
        assert.equal(result.results[1].status, 'error');
        assert.equal(result.results[1].error.code, 'AMBIGUOUS_CAPTURE_SELECTION');
    }

    // 9b. Multi-paragraph capture with ambiguous select query
    {
        const ops = [
            {
                operationId: 'prod-multi-2',
                captureKey: 'multi-cap-2',
                type: 'redline',
                target: 'First paragraph.',
                modified: 'Same Text.\n\nSame Text.'
            },
            {
                operationId: 'consumer-ambiguous-select',
                type: 'comment',
                target: { captureRef: 'multi-cap-2', select: 'Same Text' },
                textToComment: 'Same Text',
                commentContent: 'Should fail with AMBIGUOUS_CAPTURE_SELECTION'
            }
        ];

        const result = await applyOperationsToDocumentXml(sampleDocXml, ops, 'Author', null, { generateRedlines: false, atomic: true });
        assert.equal(result.rolledBack, true);
        assert.equal(result.results[1].status, 'error');
        assert.equal(result.results[1].error.code, 'AMBIGUOUS_CAPTURE_SELECTION');
    }
}

// 10. Non-atomic batch (atomic: false) preserves successful producer when later op fails
{
    const ops = [
        {
            operationId: 'surviving-producer',
            captureKey: 'surv-cap',
            type: 'redline',
            target: 'First paragraph.',
            modified: 'Modified First Paragraph.'
        },
        {
            operationId: 'failing-consumer',
            type: 'comment',
            target: { captureRef: 'surv-cap' },
            textToComment: 'Non-existent anchor',
            commentContent: 'Failed comment'
        }
    ];

    const result = await applyOperationsToDocumentXml(sampleDocXml, ops, 'Author', null, {
        generateRedlines: false,
        atomic: false
    });

    assert.equal(result.rolledBack, undefined);
    assert.equal(result.hasChanges, true);
    assert.ok(result.documentXml.includes('Modified') && result.documentXml.includes('Paragraph.'));
    assert.equal(result.results[0].status, 'applied');
    assert.equal(result.results[1].status, 'error');
}

console.log('PASS: capture resolution tests passed');
