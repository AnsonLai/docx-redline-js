import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import {
    computeDocumentPartsRevisionToken
} from '../index.js';
import {
    applyOperationsToDocumentXml,
    applyOperationToDocumentXml
} from '../services/standalone-operation-runner.js';
import { openDocx, executeCli } from '../node/index.js';
import { zipDocx } from '../node/zip-archive.js';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const initialDocXml = `<w:document xmlns:w="${W}"><w:body><w:p><w:r><w:t>Hello world</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`;
const contentTypes = `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/></Types>`;
const rels = `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`;

function createTestDocx(docXml = initialDocXml) {
    const entries = new Map([
        ['[Content_Types].xml', Buffer.from(contentTypes)],
        ['_rels/.rels', Buffer.from(rels)],
        ['word/_rels/document.xml.rels', Buffer.from(rels)],
        ['word/document.xml', Buffer.from(docXml)]
    ]);
    return zipDocx(entries);
}

// 1. Matching package token applies successfully in DocxDocument.prototype.applyOperations
{
    const zip = createTestDocx();
    const doc = openDocx(zip);
    const token = doc.getRevisionToken();

    const ops = [
        { type: 'redline', target: 'Hello world', modified: 'Hello universe' }
    ];

    const result = await doc.applyOperations(ops, {
        author: 'Alice',
        expectedRevision: token
    });

    assert.equal(result.status, 'ok');
    assert.equal(result.written, true);
    assert.equal(result.hasChanges, true);
    assert.ok(result.buffer);
}

// 2. Stale package token refuses with REVISION_MISMATCH and leaves input untouched
{
    const zip = createTestDocx();
    const doc = openDocx(zip);
    const fakeToken = {
        algorithm: 'sha256',
        version: 1,
        scope: 'package',
        value: '0000000000000000000000000000000000000000000000000000000000000000'
    };

    const ops = [
        { type: 'redline', target: 'Hello world', modified: 'Hello universe' }
    ];

    const result = await doc.applyOperations(ops, {
        author: 'Alice',
        expectedRevision: fakeToken
    });

    assert.equal(result.status, 'error');
    assert.equal(result.written, false);
    assert.equal(result.hasChanges, false);
    assert.equal(result.rolledBack, true);
    assert.equal(result.error?.code, 'REVISION_MISMATCH');
    assert.deepEqual(result.artifactsChanged, []);
    // Ensure buffer is untouched original
    assert.deepEqual(Buffer.from(result.toBuffer()), Buffer.from(zip));
}

// 3. Document-parts token cannot guard package API (REVISION_TOKEN_SCOPE_MISMATCH)
{
    const zip = createTestDocx();
    const doc = openDocx(zip);
    const partsToken = await computeDocumentPartsRevisionToken({ documentXml: initialDocXml });

    const ops = [
        { type: 'redline', target: 'Hello world', modified: 'Hello universe' }
    ];

    const result = await doc.applyOperations(ops, {
        author: 'Alice',
        expectedRevision: partsToken
    });

    assert.equal(result.status, 'error');
    assert.equal(result.written, false);
    assert.equal(result.error?.code, 'REVISION_TOKEN_SCOPE_MISMATCH');
}

// 4. Invalid token structure refuses with INVALID_REVISION_TOKEN
{
    const zip = createTestDocx();
    const doc = openDocx(zip);

    const result = await doc.applyOperations([
        { type: 'redline', target: 'Hello world', modified: 'Hello universe' }
    ], {
        author: 'Alice',
        expectedRevision: { algorithm: 'md5', version: 1, scope: 'package', value: '123' }
    });

    assert.equal(result.status, 'error');
    assert.equal(result.written, false);
    assert.equal(result.error?.code, 'INVALID_REVISION_TOKEN');
}

// 5. Lower-level applyOperationsToDocumentXml with matching document-parts token
{
    const partsToken = await computeDocumentPartsRevisionToken({ documentXml: initialDocXml });

    const ops = [
        { type: 'redline', target: 'Hello world', modified: 'Hello universe' }
    ];

    const result = await applyOperationsToDocumentXml(initialDocXml, ops, 'Alice', null, {
        expectedRevision: partsToken
    });

    assert.equal(result.hasChanges, true);
    assert.ok(result.documentXml.includes('universe'));
    assert.ok(result.documentXml.includes('w:ins'));
    assert.ok(result.documentXml.includes('w:del'));
}

// 6. Lower-level applyOperationsToDocumentXml with mismatched document-parts token
{
    const fakeToken = {
        algorithm: 'sha256',
        version: 1,
        scope: 'document-parts',
        value: '1111111111111111111111111111111111111111111111111111111111111111'
    };

    const ops = [
        { type: 'redline', target: 'Hello world', modified: 'Hello universe' }
    ];

    const result = await applyOperationsToDocumentXml(initialDocXml, ops, 'Alice', null, {
        expectedRevision: fakeToken
    });

    assert.equal(result.status, 'error');
    assert.equal(result.hasChanges, false);
    assert.equal(result.error?.code, 'REVISION_MISMATCH');
    assert.equal(result.documentXml, initialDocXml);
}

// 7. Lower-level applyOperationsToDocumentXml with package scope refuses
{
    const pkgToken = {
        algorithm: 'sha256',
        version: 1,
        scope: 'package',
        value: '1111111111111111111111111111111111111111111111111111111111111111'
    };

    const ops = [
        { type: 'redline', target: 'Hello world', modified: 'Hello universe' }
    ];

    const result = await applyOperationsToDocumentXml(initialDocXml, ops, 'Alice', null, {
        expectedRevision: pkgToken
    });

    assert.equal(result.status, 'error');
    assert.equal(result.error?.code, 'REVISION_TOKEN_SCOPE_MISMATCH');
}

// 8. Single operation applyOperationToDocumentXml with expectedRevision
{
    const partsToken = await computeDocumentPartsRevisionToken({ documentXml: initialDocXml });

    const success = await applyOperationToDocumentXml(initialDocXml, {
        type: 'redline',
        target: 'Hello world',
        modified: 'Hello universe'
    }, 'Alice', null, { expectedRevision: partsToken });

    assert.equal(success.hasChanges, true);

    const failure = await applyOperationToDocumentXml(initialDocXml, {
        type: 'redline',
        target: 'Hello world',
        modified: 'Hello universe'
    }, 'Alice', null, {
        expectedRevision: {
            algorithm: 'sha256',
            version: 1,
            scope: 'document-parts',
            value: '2222222222222222222222222222222222222222222222222222222222222222'
        }
    });

    assert.equal(failure.status, 'error');
    assert.equal(failure.hasChanges, false);
    assert.equal(failure.error?.code, 'REVISION_MISMATCH');
}

// 9. Non-atomic mode also leaves input untouched on mismatch
{
    const zip = createTestDocx();
    const doc = openDocx(zip);

    const result = await doc.applyOperations([
        { type: 'redline', target: 'Hello world', modified: 'Hello universe' }
    ], {
        author: 'Alice',
        atomic: false,
        expectedRevision: {
            algorithm: 'sha256',
            version: 1,
            scope: 'package',
            value: '3333333333333333333333333333333333333333333333333333333333333333'
        }
    });

    assert.equal(result.status, 'error');
    assert.equal(result.hasChanges, false);
    assert.equal(result.written, false);
    assert.equal(result.error?.code, 'REVISION_MISMATCH');
    assert.deepEqual(Buffer.from(result.toBuffer()), Buffer.from(zip));
}

// 10. Omitting expectedRevision preserves legacy path byte-for-byte
{
    const zip = createTestDocx();
    const doc = openDocx(zip);

    const result = await doc.applyOperations([
        { type: 'redline', target: 'Hello world', modified: 'Hello universe' }
    ], { author: 'Alice' });

    assert.equal(result.written, true);
    assert.equal(result.hasChanges, true);
}

// 11. CLI --expected-revision flag enforcement
{
    const zip = createTestDocx();
    const doc = openDocx(zip);
    const token = doc.getRevisionToken();

    const dir = await mkdtemp(path.join(tmpdir(), 'docx-token-cli-'));
    try {
        const inputPath = path.join(dir, 'input.docx');
        const opsPath = path.join(dir, 'ops.json');
        const outputPath = path.join(dir, 'output.docx');

        await writeFile(inputPath, zip);
        await writeFile(opsPath, JSON.stringify([
            { type: 'replace', target: { exactText: 'Hello world' }, modified: 'Hello CLI', author: 'Alice' }
        ]));

        // Mismatched token via CLI fails with exitCode 2
        const failRes = await executeCli([
            'apply', inputPath,
            '--operations', opsPath,
            '--expected-revision', '0000000000000000000000000000000000000000000000000000000000000000'
        ]);
        assert.equal(failRes.status, 'error');
        assert.equal(failRes.exitCode, 2);
        assert.equal(failRes.error?.code, 'REVISION_MISMATCH');

        // Matching token via CLI succeeds
        const successRes = await executeCli([
            'apply', inputPath,
            '--operations', opsPath,
            '--expected-revision', token.value,
            '--output', outputPath
        ]);
        assert.equal(successRes.status, 'ok');
        assert.equal(successRes.written, true);

        // JSON file with top-level expectedRevision
        const opsWithTokenPath = path.join(dir, 'ops-token.json');
        await writeFile(opsWithTokenPath, JSON.stringify({
            expectedRevision: token,
            operations: [
                { type: 'replace', target: { exactText: 'Hello world' }, modified: 'Hello File', author: 'Alice' }
            ]
        }));
        const fileRes = await executeCli([
            'apply', inputPath,
            '--operations', opsWithTokenPath,
            '--output', path.join(dir, 'output2.docx')
        ]);
        assert.equal(fileRes.status, 'ok');
        assert.equal(fileRes.written, true);
    } finally {
        await rm(dir, { recursive: true, force: true });
    }
}

console.log('PASS: revision token enforcement suite');

