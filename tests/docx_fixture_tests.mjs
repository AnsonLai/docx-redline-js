import './setup-xml-provider.mjs';

import assert from 'assert';
import { readFileSync, existsSync } from 'fs';
import path from 'path';
import { inflateRawSync } from 'zlib';

const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const EOCD_SIGNATURE = 0x06054b50;
const CENTRAL_DIR_SIGNATURE = 0x02014b50;
const LOCAL_FILE_SIGNATURE = 0x04034b50;

function parseXmlStrict(xmlText, label) {
    const parser = new DOMParser();
    const xmlDoc = parser.parseFromString(xmlText, 'application/xml');
    const parseError = xmlDoc.getElementsByTagName('parsererror')[0];
    if (parseError) {
        throw new Error(`[XML parse error] ${label}: ${parseError.textContent || 'Unknown'}`);
    }
    return xmlDoc;
}

function findEndOfCentralDirectory(zipBuffer) {
    const minOffset = Math.max(0, zipBuffer.length - 0xffff - 22);
    for (let offset = zipBuffer.length - 22; offset >= minOffset; offset -= 1) {
        if (zipBuffer.readUInt32LE(offset) === EOCD_SIGNATURE) {
            return offset;
        }
    }
    throw new Error('Invalid zip: End of central directory not found');
}

function readEntryData(zipBuffer, localHeaderOffset, compressedSize, compressionMethod) {
    if (zipBuffer.readUInt32LE(localHeaderOffset) !== LOCAL_FILE_SIGNATURE) {
        throw new Error('Invalid zip: local file header signature mismatch');
    }

    const fileNameLength = zipBuffer.readUInt16LE(localHeaderOffset + 26);
    const extraLength = zipBuffer.readUInt16LE(localHeaderOffset + 28);
    const dataStart = localHeaderOffset + 30 + fileNameLength + extraLength;
    const compressed = zipBuffer.subarray(dataStart, dataStart + compressedSize);

    if (compressionMethod === 0) return Buffer.from(compressed);
    if (compressionMethod === 8) return inflateRawSync(compressed);
    throw new Error(`Unsupported compression method: ${compressionMethod}`);
}

function unzipEntries(zipBuffer) {
    const eocdOffset = findEndOfCentralDirectory(zipBuffer);
    const totalEntries = zipBuffer.readUInt16LE(eocdOffset + 10);
    const centralDirOffset = zipBuffer.readUInt32LE(eocdOffset + 16);

    let offset = centralDirOffset;
    const entries = new Map();

    for (let i = 0; i < totalEntries; i += 1) {
        if (zipBuffer.readUInt32LE(offset) !== CENTRAL_DIR_SIGNATURE) {
            throw new Error('Invalid zip: central directory signature mismatch');
        }

        const compressionMethod = zipBuffer.readUInt16LE(offset + 10);
        const compressedSize = zipBuffer.readUInt32LE(offset + 20);
        const fileNameLength = zipBuffer.readUInt16LE(offset + 28);
        const extraLength = zipBuffer.readUInt16LE(offset + 30);
        const commentLength = zipBuffer.readUInt16LE(offset + 32);
        const localHeaderOffset = zipBuffer.readUInt32LE(offset + 42);
        const fileName = zipBuffer.toString('utf8', offset + 46, offset + 46 + fileNameLength);

        const data = readEntryData(zipBuffer, localHeaderOffset, compressedSize, compressionMethod);
        entries.set(fileName, data);

        offset += 46 + fileNameLength + extraLength + commentLength;
    }

    return entries;
}

async function testReadRealDocxFixture() {
    const fixturePath = path.join(process.cwd(), 'tests', 'fixtures', 'sample_doc_test.docx');
    assert.ok(existsSync(fixturePath), 'docx fixture should exist at tests/fixtures/sample_doc_test.docx');

    const zipBuffer = readFileSync(fixturePath);
    const entries = unzipEntries(zipBuffer);

    assert.ok(entries.has('word/document.xml'), 'docx should contain word/document.xml');
    assert.ok(entries.has('[Content_Types].xml'), 'docx should contain [Content_Types].xml');
    assert.ok(entries.has('word/_rels/document.xml.rels'), 'docx should contain document relationships');

    const documentXml = entries.get('word/document.xml').toString('utf8');
    const contentTypesXml = entries.get('[Content_Types].xml').toString('utf8');
    const relationshipsXml = entries.get('word/_rels/document.xml.rels').toString('utf8');

    assert.ok(documentXml.includes('<w:document'), 'document.xml should include a Word document root');
    assert.ok(contentTypesXml.includes('/word/document.xml'), 'content types should include document.xml override');
    assert.ok(relationshipsXml.includes('<Relationship'), 'document relationships should include Relationship nodes');
    assert.ok(documentXml.includes('NON-DISCLOSURE AGREEMENT'), 'fixture document should include expected text content');

    const xmlDoc = parseXmlStrict(documentXml, 'word/document.xml');
    assert.ok(
        xmlDoc.getElementsByTagNameNS(NS_W, 'p').length > 0,
        'parsed document.xml should contain paragraphs'
    );
}

async function run() {
    await testReadRealDocxFixture();
    console.log('PASS: docx fixture unzip tests');
}

run().catch(err => {
    console.error('FAIL:', err.message);
    process.exit(1);
});

