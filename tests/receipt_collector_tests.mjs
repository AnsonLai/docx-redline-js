import './setup-xml-provider.mjs';
import assert from 'node:assert/strict';
import { ReceiptCollector } from '../services/receipt-collector.js';
import { DocumentOperationSession } from '../services/document-operation-session.js';
import {
    applyOperationToDocumentXml
} from '../services/standalone-operation-runner.js';
import {
    ensureCommentsArtifactsInZip
} from '../services/standalone-docx-plumbing.js';
import JSZip from 'jszip';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const sampleDocXml = `<w:document xmlns:w="${W}"><w:body><w:p><w:r><w:t>First paragraph.</w:t></w:r></w:p><w:p><w:r><w:t>Second paragraph.</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`;

// 1. ReceiptCollector unit tests
{
    const collector = new ReceiptCollector();
    collector.beginOperation(1, 'op-1', 'Author');
    collector.recordRevision('10', 'del');
    collector.recordRevision('11', 'ins');
    collector.recordComment('1');
    collector.recordNumbering('2');
    collector.recordRelationship('rId5');
    collector.recordAffectedTarget({ exactText: 'First paragraph.' });
    collector.recordWarning('Sample warning');

    const savepoint = collector.createSavepoint();
    collector.recordRevision('12', 'ins');
    assert.equal(collector.getCurrentReceipt().revisionItems.length, 3);

    // Restore savepoint
    collector.restoreSavepoint(savepoint);
    assert.equal(collector.getCurrentReceipt().revisionItems.length, 2);

    const committed = collector.commitOperation('applied');
    assert.equal(committed.operationIndex, 1);
    assert.equal(committed.operationId, 'op-1');
    assert.equal(committed.authorUsed, 'Author');
    assert.equal(committed.committed, true);
    assert.equal(committed.finalDisposition, 'applied');
    assert.equal(committed.revisionItems.length, 2);
    assert.deepEqual(committed.commentIds, ['1']);
    assert.deepEqual(committed.numberingIds, ['2']);
    assert.deepEqual(committed.relationshipIds, ['rId5']);
    assert.equal(committed.warnings.length, 1);

    assert.equal(collector.getReceipts().length, 1);
}

// 2. Text replacement records distinct deletion and insertion IDs
{
    const session = new DocumentOperationSession(sampleDocXml);
    const op = {
        type: 'redline',
        target: 'First paragraph.',
        modified: 'Modified first paragraph.',
        generateRedlines: true
    };

    const result = await applyOperationToDocumentXml(
        sampleDocXml,
        op,
        'TestAuthor',
        null,
        { _documentOperationSession: session }
    );

    assert.equal(result.hasChanges, true);
    const receipts = session.receiptCollector.getReceipts();
    assert.equal(receipts.length, 1);
    const receipt = receipts[0];
    assert.equal(receipt.attemptedDisposition, 'applied');
    assert.equal(receipt.finalDisposition, 'applied');

    const delItem = receipt.revisionItems.find(item => item.kind === 'del');
    const insItem = receipt.revisionItems.find(item => item.kind === 'ins');
    assert.ok(delItem, 'Must record deletion revision item');
    assert.ok(insItem, 'Must record insertion revision item');
    assert.notEqual(delItem.id, insItem.id, 'Deletion and insertion IDs must be distinct');
}

// 3. Formatting records property-change IDs (rPrChange / pPrChange)
{
    const session = new DocumentOperationSession(sampleDocXml);
    const op = {
        type: 'format',
        target: 'First paragraph.',
        textToFormat: 'First',
        properties: { bold: true },
        generateRedlines: true
    };

    const result = await applyOperationToDocumentXml(
        sampleDocXml,
        op,
        'TestAuthor',
        null,
        { _documentOperationSession: session }
    );

    assert.equal(result.hasChanges, true);
    const receipts = session.receiptCollector.getReceipts();
    assert.equal(receipts.length, 1);
    const receipt = receipts[0];
    const rPrItem = receipt.revisionItems.find(item => item.kind === 'rPrChange');
    assert.ok(rPrItem, 'Must record rPrChange revision item for character formatting');
    assert.ok(rPrItem.id, 'rPrChange must have an allocated revision id');
}

// 4. Comment records comment ID
{
    const session = new DocumentOperationSession(sampleDocXml);
    const op = {
        type: 'comment',
        target: 'First paragraph.',
        textToComment: 'First',
        commentContent: 'Please check this paragraph.'
    };

    const result = await applyOperationToDocumentXml(
        sampleDocXml,
        op,
        'CommentAuthor',
        null,
        { _documentOperationSession: session }
    );

    assert.equal(result.hasChanges, true);
    const receipts = session.receiptCollector.getReceipts();
    assert.equal(receipts.length, 1);
    const receipt = receipts[0];
    assert.ok(receipt.commentIds.length > 0, 'Must record placed comment ID');
}

// 5. List records numbering ID
{
    const session = new DocumentOperationSession(sampleDocXml);
    const op = {
        type: 'redline',
        target: 'Second paragraph.',
        modified: '1. Item one\n2. Item two',
        generateRedlines: true
    };

    const result = await applyOperationToDocumentXml(
        sampleDocXml,
        op,
        'ListAuthor',
        null,
        { _documentOperationSession: session }
    );

    assert.equal(result.hasChanges, true);
    const receipts = session.receiptCollector.getReceipts();
    assert.equal(receipts.length, 1);
    const receipt = receipts[0];
    assert.ok(receipt.numberingIds.length > 0, 'Must record allocated numbering ID');
}

// 6. Package plumbing records companion relationship IDs
{
    const zip = new JSZip();
    zip.file('word/_rels/document.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`);

    const collector = new ReceiptCollector();
    collector.beginOperation(1, 'plumbing-op');

    await ensureCommentsArtifactsInZip(zip, '<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>', {
        _receiptCollector: collector
    });

    assert.ok(collector.getCurrentReceipt().relationshipIds.length > 0, 'Must record created relationship ID');
    assert.ok(collector.getCurrentReceipt().relationshipIds.includes('rId2'), 'rId2 should be recorded');
}

// 7. Failed operation leaves collector state unchanged
{
    const session = new DocumentOperationSession(sampleDocXml);
    const opGood = {
        type: 'redline',
        target: 'First paragraph.',
        modified: 'First paragraph modified.',
        generateRedlines: true
    };
    await applyOperationToDocumentXml(sampleDocXml, opGood, 'Author', null, { _documentOperationSession: session });
    assert.equal(session.receiptCollector.getReceipts().length, 1);

    const opBad = {
        type: 'redline',
        target: 'Non-existent target text that cannot be found',
        modified: 'Failed replacement'
    };
    const badResult = await applyOperationToDocumentXml(session.currentDocumentXml, opBad, 'Author', null, { _documentOperationSession: session });
    assert.equal(badResult.hasChanges, false);
    assert.equal(badResult.status, 'error');

    // Collector must have only the 1 original receipt; failed op left no trace
    assert.equal(session.receiptCollector.getReceipts().length, 1);
    assert.equal(session.receiptCollector.getCurrentReceipt(), null);
}

// 8. No-op operation leaves collector state unchanged
{
    const session = new DocumentOperationSession(sampleDocXml);
    const opNoop = {
        type: 'redline',
        target: 'First paragraph.',
        modified: 'First paragraph.',
        generateRedlines: true
    };

    const noopResult = await applyOperationToDocumentXml(sampleDocXml, opNoop, 'Author', null, { _documentOperationSession: session });
    assert.equal(noopResult.hasChanges, false);

    // Collector must have 0 receipts after no-op
    assert.equal(session.receiptCollector.getReceipts().length, 0);
    assert.equal(session.receiptCollector.getCurrentReceipt(), null);
}

console.log('All ReceiptCollector tests passed successfully.');
