import assert from 'node:assert/strict';
import { openDocx as rootOpenDocx, DocxDocument as RootDocxDocument, computePackageRevisionToken as rootComputeToken } from '../index.js';
import { openDocx as nodeOpenDocx, DocxDocument as NodeDocxDocument, computePackageRevisionToken as nodeComputeToken } from '../node/docx-document.js';
import { openDocx, DocxDocument, computePackageRevisionToken } from '../document/docx-document.js';
import { zipDocx } from '../document/zip-archive.js';

const textEncoder = new TextEncoder();

function createMinimalDocxUint8Array(paragraphText = 'Hello Edge Cases') {
    const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
  <w:body>
    <w:p w14:paraId="00000001">
      <w:r><w:t>${paragraphText}</w:t></w:r>
    </w:p>
  </w:body>
</w:document>`;

    const contentTypesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`;

    const relsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

    const docRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`;

    const entries = new Map([
        ['[Content_Types].xml', textEncoder.encode(contentTypesXml)],
        ['_rels/.rels', textEncoder.encode(relsXml)],
        ['word/document.xml', textEncoder.encode(documentXml)],
        ['word/_rels/document.xml.rels', textEncoder.encode(docRelsXml)]
    ]);

    return zipDocx(entries);
}

// 1. Export identity and parity across index.js, node/docx-document.js, and document/docx-document.js
{
    assert.equal(rootOpenDocx, openDocx);
    assert.equal(nodeOpenDocx, openDocx);
    assert.equal(RootDocxDocument, DocxDocument);
    assert.equal(NodeDocxDocument, DocxDocument);
    assert.equal(rootComputeToken, computePackageRevisionToken);
    assert.equal(nodeComputeToken, computePackageRevisionToken);
}

// 2. Corrupt and invalid input handling
{
    assert.throws(() => {
        openDocx(new Uint8Array([1, 2, 3, 4, 5, 6, 7, 8]));
    }, /zip|archive/i);

    assert.throws(() => {
        openDocx(new Uint8Array(0));
    }, /zip|archive/i);
}

// 3. Optimistic concurrency with expectedRevision tokens
{
    const pkgBytes = createMinimalDocxUint8Array('Concurrency baseline');
    const doc = openDocx(pkgBytes);

    const initialToken = doc.getRevisionToken();
    assert.equal(initialToken.scope, 'package');
    assert.equal(typeof initialToken.value, 'string');

    // Mismatched scope token rejection (valid structure, but document-parts instead of package)
    const badScopeResult = await doc.applyOperations([{
        type: 'replace',
        target: { text: 'Concurrency baseline' },
        modified: 'Failed update'
    }], {
        expectedRevision: { algorithm: 'sha256', version: 1, scope: 'document-parts', value: initialToken.value }
    });
    assert.equal(badScopeResult.status, 'error');
    assert.equal(badScopeResult.error.code, 'REVISION_TOKEN_SCOPE_MISMATCH');
    assert.equal(badScopeResult.written, false);

    // Invalid token structure rejection
    const invalidTokenResult = await doc.applyOperations([{
        type: 'replace',
        target: { text: 'Concurrency baseline' },
        modified: 'Failed update'
    }], {
        expectedRevision: { invalid: true }
    });
    assert.equal(invalidTokenResult.status, 'error');
    assert.equal(invalidTokenResult.error.code, 'INVALID_REVISION_TOKEN');
    assert.equal(invalidTokenResult.written, false);

    // Matching token succeeds
    const goodResult = await doc.applyOperations([{
        type: 'replace',
        target: { text: 'Concurrency baseline' },
        modified: 'First update'
    }], {
        author: 'Alice',
        expectedRevision: initialToken
    });
    assert.equal(goodResult.status, 'ok');
    assert.equal(goodResult.written, true);

    const updatedToken = doc.getRevisionToken();
    assert.notEqual(updatedToken.value, initialToken.value);

    // Stale token rejection (optimistic locking)
    const staleResult = await doc.applyOperations([{
        type: 'replace',
        target: { text: 'First update' },
        modified: 'Second update'
    }], {
        expectedRevision: initialToken
    });
    assert.equal(staleResult.status, 'error');
    assert.equal(staleResult.error.code, 'REVISION_MISMATCH');
    assert.equal(staleResult.written, false);

    // Cross-author safety policy: Bob editing Alice's changes without policy flag is safely refused
    const crossAuthorRefused = await doc.applyOperations([{
        type: 'replace',
        target: { text: 'First update' },
        modified: 'Second update'
    }], {
        author: 'Bob',
        expectedRevision: updatedToken
    });
    assert.equal(crossAuthorRefused.status, 'error');
    assert.equal(crossAuthorRefused.results[0].error.code, 'EXISTING_REVISIONS');

    // With explicit slice-cross-author policy, Bob's edit succeeds
    const secondGoodResult = await doc.applyOperations([{
        type: 'replace',
        target: { text: 'First update' },
        modified: 'Second update'
    }], {
        author: 'Bob',
        existingRevisions: 'slice-cross-author',
        expectedRevision: updatedToken
    });
    assert.equal(secondGoodResult.status, 'ok');
    assert.equal(secondGoodResult.written, true);
    assert.equal(doc.inspect().paragraphs[0].exactText, 'Second update');
    assert.deepEqual(doc.inspect().paragraphs[0].revisionAuthors.sort(), ['Alice', 'Bob']);
}

// 4. Sequential in-memory mutations without reloading
{
    const pkgBytes = createMinimalDocxUint8Array('Sequential Step 0');
    const doc = openDocx(pkgBytes);

    for (let step = 1; step <= 3; step++) {
        const prevText = step === 1 ? 'Sequential Step 0' : `Sequential Step ${step - 1}`;
        const nextText = `Sequential Step ${step}`;

        const res = await doc.applyOperations([{
            type: 'replace',
            target: { text: prevText },
            modified: nextText
        }], { author: 'SequentialEditor' });

        assert.equal(res.status, 'ok');
        assert.equal(res.written, true);
        assert.equal(doc.inspect().paragraphs[0].exactText, nextText);
    }

    const finalBytes = doc.toUint8Array();
    const finalDoc = openDocx(finalBytes);
    const finalInspection = finalDoc.inspect();
    assert.equal(finalInspection.paragraphs[0].exactText, 'Sequential Step 3');
}

// 5. Accept and reject tracked changes in universal facade
{
    const pkgBytes = createMinimalDocxUint8Array('Tracked Changes Base');
    const doc = openDocx(pkgBytes);

    await doc.applyOperations([{
        type: 'replace',
        target: { text: 'Tracked Changes Base' },
        modified: 'Accepted Changes'
    }], { author: 'ReviewerA' });

    assert.equal(doc.inspect().paragraphs[0].revisionAuthors.length, 1);

    // Accept changes
    const acceptResult = await doc.resolveRevisions('accept', { author: 'ReviewerA' });
    assert.equal(acceptResult.hasChanges, true);
    assert.equal(acceptResult.written, true);
    assert.equal(doc.inspect().paragraphs[0].revisionAuthors.length, 0);
    assert.equal(doc.inspect().paragraphs[0].exactText, 'Accepted Changes');

    // Re-apply and reject
    await doc.applyOperations([{
        type: 'replace',
        target: { text: 'Accepted Changes' },
        modified: 'Rejected Changes'
    }], { author: 'ReviewerB' });

    assert.equal(doc.inspect().paragraphs[0].exactText, 'Rejected Changes');
    const rejectResult = await doc.resolveRevisions('reject', { author: 'ReviewerB' });
    assert.equal(rejectResult.hasChanges, true);
    assert.equal(rejectResult.written, true);
    assert.equal(doc.inspect().paragraphs[0].exactText, 'Accepted Changes');

    // Delete comments on document without comments
    const emptyCommentResult = await doc.deleteComments({ allAuthors: true });
    assert.equal(emptyCommentResult.status, 'ok');
    assert.equal(emptyCommentResult.commentsRemoved, 0);
}

// 6. Preflight checks on universal DocxDocument
{
    const pkgBytes = createMinimalDocxUint8Array('Preflight Target');
    const doc = openDocx(pkgBytes);

    const preflightValid = doc.preflight([{
        type: 'replace',
        target: { text: 'Preflight Target' },
        modified: 'Preflight Modified'
    }], 'PreflightAuthor');

    assert.equal(preflightValid.status, 'ok');

    const preflightMissing = doc.preflight([{
        type: 'replace',
        target: { text: 'Non-existent text' },
        modified: 'Should fail preflight'
    }], 'PreflightAuthor');

    assert.equal(preflightMissing.status, 'error');
}

console.log('PASS: universal facade edge cases and integration tests');
