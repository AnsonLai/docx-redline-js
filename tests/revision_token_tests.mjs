import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import {
    buildRevisionTokenFraming,
    computeRevisionToken,
    computeRevisionTokenSync,
    computeDocumentPartsRevisionToken,
    validateRevisionToken,
    normalizeOpcEntryName,
    extractDocumentPartsEntries,
    inspectDocumentParts
} from '../index.js';
import { openDocx, computePackageRevisionToken } from '../node/index.js';
import { zipDocx, unzipDocx } from '../node/zip-archive.js';

// 1. Framing binary structure tests
{
    const entries = [
        { name: 'word/document.xml', payload: '<w:document/>' },
        { name: 'word/comments.xml', payload: '<w:comments/>' }
    ];
    const { framing, scope, version, coveredParts } = buildRevisionTokenFraming({
        scope: 'document-parts',
        entries
    });

    assert.equal(scope, 'document-parts');
    assert.equal(version, 1);
    assert.deepEqual(coveredParts, ['word/comments.xml', 'word/document.xml']);

    // Check framing header
    const magicStr = new TextDecoder().decode(framing.subarray(0, 28));
    assert.equal(magicStr, 'docx-redline-revision-token\0');

    const view = new DataView(framing.buffer, framing.byteOffset, framing.byteLength);
    let offset = 28;
    const framedVersion = view.getUint32(offset, false);
    offset += 4;
    assert.equal(framedVersion, 1);

    const scopeLen = view.getUint32(offset, false);
    offset += 4;
    const framedScope = new TextDecoder().decode(framing.subarray(offset, offset + scopeLen));
    offset += scopeLen;
    assert.equal(framedScope, 'document-parts');

    const entryCount = view.getUint32(offset, false);
    offset += 4;
    assert.equal(entryCount, 2);

    // First sorted entry should be word/comments.xml
    const name1Len = view.getUint32(offset, false);
    offset += 4;
    const name1 = new TextDecoder().decode(framing.subarray(offset, offset + name1Len));
    offset += name1Len;
    assert.equal(name1, 'word/comments.xml');
}

// 2. Map / Array / Object enumeration order invariance
{
    const entryA = { name: 'b.xml', payload: 'BBB' };
    const entryB = { name: 'a.xml', payload: 'AAA' };

    const tokenOrder1 = computeRevisionTokenSync({
        scope: 'package',
        entries: [entryA, entryB],
        digestFn: b => createHash('sha256').update(b).digest('hex')
    });

    const tokenOrder2 = computeRevisionTokenSync({
        scope: 'package',
        entries: [entryB, entryA],
        digestFn: b => createHash('sha256').update(b).digest('hex')
    });

    const tokenMap = computeRevisionTokenSync({
        scope: 'package',
        entries: new Map([['b.xml', 'BBB'], ['a.xml', 'AAA']]),
        digestFn: b => createHash('sha256').update(b).digest('hex')
    });

    assert.equal(tokenOrder1.value, tokenOrder2.value);
    assert.equal(tokenOrder1.value, tokenMap.value);
    assert.deepEqual(tokenOrder1.coveredParts, ['a.xml', 'b.xml']);
}

// 3. Binary zero bytes and non-ASCII names frame unambiguously
{
    const zeroBytes = new Uint8Array([0x00, 0x01, 0x00, 0x02, 0x00]);
    const utf8Name = 'word/café-ñoño.xml';

    const token = computeRevisionTokenSync({
        scope: 'package',
        entries: [{ name: utf8Name, payload: zeroBytes }],
        digestFn: b => createHash('sha256').update(b).digest('hex')
    });

    assert.equal(token.coveredParts[0], utf8Name);
    assert.equal(typeof token.value, 'string');
    assert.equal(token.value.length, 64);
}

// 4. Duplicate normalized name rejection
{
    assert.throws(() => {
        buildRevisionTokenFraming({
            scope: 'package',
            entries: [
                { name: 'word/document.xml', payload: '1' },
                { name: 'word\\document.xml', payload: '2' }
            ]
        });
    }, /Duplicate normalized entry path/);
}

// 5. Different scopes or versions produce distinct tokens
{
    const entries = [{ name: 'part.xml', payload: 'data' }];
    const tokenDocParts = computeRevisionTokenSync({
        scope: 'document-parts',
        entries,
        digestFn: b => createHash('sha256').update(b).digest('hex')
    });
    const tokenPackage = computeRevisionTokenSync({
        scope: 'package',
        entries,
        digestFn: b => createHash('sha256').update(b).digest('hex')
    });

    assert.notEqual(tokenDocParts.value, tokenPackage.value);
    assert.equal(tokenDocParts.scope, 'document-parts');
    assert.equal(tokenPackage.scope, 'package');
}

// 6. Payload change or name change alters token
{
    const base = [{ name: 'a.xml', payload: 'data' }];
    const alteredPayload = [{ name: 'a.xml', payload: 'datb' }];
    const alteredName = [{ name: 'b.xml', payload: 'data' }];

    const tBase = computeRevisionTokenSync({ scope: 'package', entries: base, digestFn: b => createHash('sha256').update(b).digest('hex') });
    const tPayload = computeRevisionTokenSync({ scope: 'package', entries: alteredPayload, digestFn: b => createHash('sha256').update(b).digest('hex') });
    const tName = computeRevisionTokenSync({ scope: 'package', entries: alteredName, digestFn: b => createHash('sha256').update(b).digest('hex') });

    assert.notEqual(tBase.value, tPayload.value);
    assert.notEqual(tBase.value, tName.value);
}

// 7. Stable across ZIP recompression / timestamp changes
{
    const entries = new Map([
        ['[Content_Types].xml', Buffer.from('<Types/>')],
        ['word/document.xml', Buffer.from('<w:document><w:body><w:p><w:r><w:t>Hello</w:t></w:r></w:p></w:body></w:document>')]
    ]);

    const zipBuffer1 = zipDocx(entries);
    const unzipped1 = unzipDocx(zipBuffer1);
    const zipBuffer2 = zipDocx(unzipped1);

    const token1 = computePackageRevisionToken(zipBuffer1);
    const token2 = computePackageRevisionToken(zipBuffer2);
    const tokenFromEntries = computePackageRevisionToken(entries);

    assert.equal(token1.value, token2.value);
    assert.equal(token1.value, tokenFromEntries.value);
    assert.equal(token1.scope, 'package');
    assert.equal(token1.version, 1);
    assert.equal(token1.algorithm, 'sha256');
}

// 8. DocxDocument integration: inspect() and getRevisionToken()
{
    const entries = new Map([
        ['[Content_Types].xml', Buffer.from('<Types/>')],
        ['word/document.xml', Buffer.from('<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Hello world</w:t></w:r></w:p></w:body></w:document>')]
    ]);
    const zip = zipDocx(entries);
    const doc = openDocx(zip);

    const pkgToken = doc.getRevisionToken();
    assert.equal(pkgToken.scope, 'package');
    assert.equal(doc.revisionToken.value, pkgToken.value);

    const inspection = doc.inspect();
    assert.ok(inspection.revisionToken);
    assert.equal(inspection.revisionToken.scope, 'document-parts');
    assert.equal(inspection.revisionToken.algorithm, 'sha256');
    assert.equal(inspection.revisionToken.version, 1);
    assert.ok(Array.isArray(inspection.coveredParts));
    assert.ok(inspection.coveredParts.includes('word/document.xml'));
}

// 9. Web Crypto asynchronous computation test
{
    const tokenAsync = await computeRevisionToken({
        scope: 'document-parts',
        entries: [{ name: 'word/document.xml', payload: '<w:document/>' }]
    });

    const tokenSync = computeRevisionTokenSync({
        scope: 'document-parts',
        entries: [{ name: 'word/document.xml', payload: '<w:document/>' }],
        digestFn: b => createHash('sha256').update(b).digest('hex')
    });

    assert.equal(tokenAsync.value, tokenSync.value);
    assert.equal(tokenAsync.algorithm, 'sha256');
    assert.equal(tokenAsync.version, 1);
}

// 10. validateRevisionToken checks
{
    const valid = {
        algorithm: 'sha256',
        version: 1,
        scope: 'package',
        value: 'e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855'
    };
    assert.deepEqual(validateRevisionToken(valid), { valid: true });

    assert.equal(validateRevisionToken(null).valid, false);
    assert.equal(validateRevisionToken({ ...valid, algorithm: 'sha512' }).valid, false);
    assert.equal(validateRevisionToken({ ...valid, version: 2 }).valid, false);
    assert.equal(validateRevisionToken({ ...valid, scope: 'invalid' }).valid, false);
    assert.equal(validateRevisionToken({ ...valid, value: 'not-a-hash' }).valid, false);
    assert.equal(validateRevisionToken({ ...valid, value: '123' }).valid, false);
}

console.log('PASS: revision token core suite');
