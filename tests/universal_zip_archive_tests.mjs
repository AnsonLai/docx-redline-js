import assert from 'node:assert/strict';
import { unzipDocx, zipDocx, MemoryZip, crc32 } from '../document/zip-archive.js';

// 1. CRC32 calculation
{
    const textEncoder = new TextEncoder();
    const data = textEncoder.encode('123456789');
    // Standard CRC32 check value for "123456789" is 0xcbf43926
    assert.equal(crc32(data), 0xcbf43926);
    assert.equal(crc32(new Uint8Array(0)), 0);
}

// 2. Pure Uint8Array round-trip
{
    const textEncoder = new TextEncoder();
    const textDecoder = new TextDecoder();
    const entries = new Map([
        ['[Content_Types].xml', textEncoder.encode('<Types/>')],
        ['word/document.xml', textEncoder.encode('<w:document><w:body><w:p><w:r><w:t>Universal ZIP</w:t></w:r></w:p></w:body></w:document>')],
        ['nested/folder/data.bin', new Uint8Array([0, 1, 2, 3, 254, 255])],
        ['empty.txt', new Uint8Array(0)]
    ]);

    const zipBytes = zipDocx(entries);
    assert.ok(zipBytes instanceof Uint8Array);
    assert.ok(zipBytes.length > 0);

    const roundTrip = unzipDocx(zipBytes);
    assert.equal(roundTrip.size, entries.size);
    assert.deepEqual([...roundTrip.keys()], [...entries.keys()]);

    for (const [name, expectedBytes] of entries) {
        const actualBytes = roundTrip.get(name);
        assert.ok(actualBytes instanceof Uint8Array);
        assert.equal(actualBytes.length, expectedBytes.length);
        assert.deepEqual(Array.from(actualBytes), Array.from(expectedBytes));
    }

    assert.equal(textDecoder.decode(roundTrip.get('[Content_Types].xml')), '<Types/>');
}

// 3. Corrupt archive error handling
{
    assert.throws(() => unzipDocx(new Uint8Array([1, 2, 3])), /ZIP end record not found/);
    assert.throws(() => unzipDocx(new TextEncoder().encode('this is completely invalid data')), /ZIP end record not found/);

    // Build a valid zip, then corrupt central directory signature
    const valid = zipDocx(new Map([['test.txt', 'hello']]));
    const corrupt = new Uint8Array(valid);
    // Find central directory signature 0x02014b50
    for (let i = 0; i < corrupt.length - 4; i++) {
        if (corrupt[i] === 0x50 && corrupt[i+1] === 0x4b && corrupt[i+2] === 0x01 && corrupt[i+3] === 0x02) {
            corrupt[i] = 0x00; // corrupt signature
            break;
        }
    }
    assert.throws(() => unzipDocx(corrupt), /central directory is corrupt/);
}

// 4. MemoryZip tests
{
    const memZip = new MemoryZip();
    memZip.file('test.txt', 'hello world');
    const file = memZip.file('test.txt');
    assert.ok(file);
    const contentStr = await file.async('string');
    assert.equal(contentStr, 'hello world');
    const contentBytes = await file.async('uint8array');
    assert.ok(contentBytes instanceof Uint8Array);
    assert.equal(new TextDecoder().decode(contentBytes), 'hello world');
    assert.equal(memZip.file('nonexistent.txt'), null);
}

console.log('PASS: universal zip archive test suite');
