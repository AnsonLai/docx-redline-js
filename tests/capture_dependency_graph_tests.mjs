import './setup-xml-provider.mjs';
import assert from 'node:assert/strict';
import {
    buildOperationDependencyPlan,
    orderOperationsForStableTargets,
    preflightOperations,
    applyOperationsToDocumentXml
} from '../services/standalone-operation-runner.js';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const sampleDocXml = `<w:document xmlns:w="${W}"><w:body><w:p><w:r><w:t>First paragraph.</w:t></w:r></w:p><w:p><w:r><w:t>Second paragraph.</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`;

// 1. Linear dependency (Op1 -> Op2 -> Op3)
{
    const ops = [
        { operationId: 'op1', captureKey: 'cap1', type: 'redline', target: 'First paragraph.', modified: 'Updated first.' },
        { operationId: 'op2', captureKey: 'cap2', type: 'redline', target: { captureRef: 'cap1' }, modified: 'Updated first again.' },
        { operationId: 'op3', type: 'redline', target: { captureRef: 'cap2' }, modified: 'Final first.' }
    ];

    const plan = buildOperationDependencyPlan(ops);
    assert.equal(plan.valid, true);
    assert.deepEqual(plan.scheduled.map(s => s.index), [0, 1, 2]);
}

// 2. Diamond dependency (Op1 -> Op2, Op3; Op2, Op3 -> Op4)
{
    const ops = [
        { operationId: 'op1', captureKey: 'cap1', type: 'redline', target: 'First paragraph.', modified: 'A' },
        { operationId: 'op2', captureKey: 'cap2', type: 'redline', target: { captureRef: 'cap1' }, modified: 'B' },
        { operationId: 'op3', captureKey: 'cap3', type: 'redline', target: { captureRef: 'cap1' }, modified: 'C' },
        { operationId: 'op4', type: 'redline', target: { captureRef: 'cap2' }, modified: 'D' }
    ];

    const plan = buildOperationDependencyPlan(ops);
    assert.equal(plan.valid, true);
    const order = plan.scheduled.map(s => s.index);
    assert.ok(order.indexOf(0) < order.indexOf(1));
    assert.ok(order.indexOf(0) < order.indexOf(2));
    assert.ok(order.indexOf(1) < order.indexOf(3));
}

// 3. Independent comments run first before producer, but dependent comments wait for producer
{
    const ops = [
        { operationId: 'producer', captureKey: 'cap1', type: 'redline', target: 'First paragraph.', modified: 'Updated' },
        { operationId: 'consumer-comment', type: 'comment', target: { captureRef: 'cap1' }, textToComment: 'Updated', commentContent: 'Note on update' },
        { operationId: 'independent-comment', type: 'comment', target: 'Second paragraph.', textToComment: 'Second', commentContent: 'Static note' }
    ];

    const plan = buildOperationDependencyPlan(ops);
    assert.equal(plan.valid, true);
    // Independent comment (index 2) has in-degree 0 and comment priority -> runs first
    // Producer (index 0) runs second
    // Consumer comment (index 1) depended on producer -> runs third
    assert.deepEqual(plan.scheduled.map(s => s.index), [2, 0, 1]);
}

// 4. Duplicate capture key detection
{
    const ops = [
        { captureKey: 'duplicate_key', type: 'redline', target: 'First paragraph.', modified: 'A' },
        { captureKey: 'duplicate_key', type: 'redline', target: 'Second paragraph.', modified: 'B' }
    ];

    const plan = buildOperationDependencyPlan(ops);
    assert.equal(plan.valid, false);
    assert.equal(plan.error.code, 'DUPLICATE_CAPTURE_KEY');
    assert.ok(plan.error.message.includes('duplicate_key'));
}

// 5. Missing producer detection
{
    const ops = [
        { type: 'redline', target: { captureRef: 'non_existent_key' }, modified: 'A' }
    ];

    const plan = buildOperationDependencyPlan(ops);
    assert.equal(plan.valid, false);
    assert.equal(plan.error.code, 'CAPTURE_NOT_FOUND');
    assert.ok(plan.error.message.includes('non_existent_key'));
}

// 6. Self-cycle detection
{
    const ops = [
        { captureKey: 'self_key', type: 'redline', target: { captureRef: 'self_key' }, modified: 'A' }
    ];

    const plan = buildOperationDependencyPlan(ops);
    assert.equal(plan.valid, false);
    assert.equal(plan.error.code, 'CAPTURE_DEPENDENCY_CYCLE');
}

// 7. Multi-node cycle detection
{
    const ops = [
        { captureKey: 'key_a', type: 'redline', target: { captureRef: 'key_b' }, modified: 'A' },
        { captureKey: 'key_b', type: 'redline', target: { captureRef: 'key_a' }, modified: 'B' }
    ];

    const plan = buildOperationDependencyPlan(ops);
    assert.equal(plan.valid, false);
    assert.equal(plan.error.code, 'CAPTURE_DEPENDENCY_CYCLE');
}

// 8. Stable tie order by original index among ready nodes of same priority
{
    const ops = [
        { type: 'redline', target: 'First paragraph.', modified: 'A' },
        { type: 'redline', target: 'Second paragraph.', modified: 'B' },
        { type: 'redline', target: 'First paragraph.', modified: 'C' }
    ];

    const plan = buildOperationDependencyPlan(ops);
    assert.equal(plan.valid, true);
    assert.deepEqual(plan.scheduled.map(s => s.index), [0, 1, 2]);
}

// 9. orderOperationsForStableTargets throws on invalid dependency graph
{
    assert.throws(
        () => orderOperationsForStableTargets([
            { captureKey: 'dup', type: 'redline', target: 'A', modified: 'B' },
            { captureKey: 'dup', type: 'redline', target: 'C', modified: 'D' }
        ]),
        (err) => err.code === 'DUPLICATE_CAPTURE_KEY'
    );
}

// 10. Preflight marks capture consumers as 'deferred' without searching static DOM
{
    const ops = [
        { captureKey: 'new_item', type: 'redline', target: 'First paragraph.', modified: 'Created item' },
        { type: 'comment', target: { captureRef: 'new_item', select: 'item' }, textToComment: 'item', commentContent: 'Comment on capture' },
        { type: 'redline', target: 'Second paragraph.', modified: 'Updated second' }
    ];

    const preflight = preflightOperations(sampleDocXml, ops, 'Author');
    assert.equal(preflight.valid, true);
    assert.equal(preflight.status, 'ok');
    assert.equal(preflight.results.length, 3);

    // Producer
    assert.equal(preflight.results[0].status, 'ready');
    assert.equal(preflight.results[0].index, 1);

    // Consumer (deferred)
    assert.equal(preflight.results[1].status, 'deferred');
    assert.equal(preflight.results[1].resolvedBy, 'capture');
    assert.equal(preflight.results[1].captureRef, 'new_item');
    assert.equal(preflight.results[1].select, 'item');
    assert.equal(preflight.results[1].index, 2);

    // Independent op
    assert.equal(preflight.results[2].status, 'ready');
    assert.equal(preflight.results[2].index, 3);
}

// 11. Preflight fails immediately on dependency error
{
    const ops = [
        { type: 'redline', target: { captureRef: 'missing_producer' }, modified: 'X' }
    ];

    const preflight = preflightOperations(sampleDocXml, ops, 'Author');
    assert.equal(preflight.valid, false);
    assert.equal(preflight.status, 'error');
    assert.equal(preflight.error.code, 'CAPTURE_NOT_FOUND');
}

// 12. applyOperationsToDocumentXml fails cleanly on dependency plan error
{
    const ops = [
        { captureKey: 'k', type: 'redline', target: { captureRef: 'k' }, modified: 'X' }
    ];

    const result = await applyOperationsToDocumentXml(sampleDocXml, ops, 'Author');
    assert.equal(result.status, 'error');
    assert.equal(result.hasChanges, false);
    assert.equal(result.error.code, 'CAPTURE_DEPENDENCY_CYCLE');
}

// 13. Batches without capture fields preserve existing comment-first scheduling, and results are sorted by original index
{
    const ops = [
        { type: 'redline', target: 'First paragraph.', modified: 'Updated first.' },
        { type: 'comment', target: 'First paragraph.', textToComment: 'First', commentContent: 'A note' }
    ];

    const result = await applyOperationsToDocumentXml(sampleDocXml, ops, 'Author', null, { generateRedlines: true });
    assert.equal(result.hasChanges, true);
    assert.equal(result.rolledBack, undefined);
    // Comment (index 2) ran first, redline (index 1) ran second
    assert.deepEqual(result.executionOrder, [2, 1]);
    // But results array remains sorted by original index: item 1 at index 0, item 2 at index 1
    assert.equal(result.results[0].index, 1);
    assert.equal(result.results[0].type, 'redline');
    assert.equal(result.results[1].index, 2);
    assert.equal(result.results[1].type, 'comment');
}

console.log('PASS: capture dependency graph tests passed');
