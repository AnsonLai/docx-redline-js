/**
 * Small read-only ZIP helper for validation tooling.
 *
 * This intentionally stays outside the runtime package path. It supports the
 * stored and deflated entries used by DOCX packages without adding a ZIP
 * dependency to the library.
 */

import { inflateRawSync } from 'zlib';

const EOCD_SIGNATURE = 0x06054b50;
const CENTRAL_DIR_SIGNATURE = 0x02014b50;
const LOCAL_FILE_SIGNATURE = 0x04034b50;

function findEndOfCentralDirectory(zipBuffer) {
    const minOffset = Math.max(0, zipBuffer.length - 0xffff - 22);
    for (let offset = zipBuffer.length - 22; offset >= minOffset; offset -= 1) {
        if (zipBuffer.readUInt32LE(offset) === EOCD_SIGNATURE) return offset;
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

/** @returns {Map<string, Buffer>} */
export function unzipEntries(zipBuffer) {
    const eocdOffset = findEndOfCentralDirectory(zipBuffer);
    const totalEntries = zipBuffer.readUInt16LE(eocdOffset + 10);
    const centralDirOffset = zipBuffer.readUInt32LE(eocdOffset + 16);
    const entries = new Map();
    let offset = centralDirOffset;

    for (let index = 0; index < totalEntries; index += 1) {
        if (zipBuffer.readUInt32LE(offset) !== CENTRAL_DIR_SIGNATURE) {
            throw new Error('Invalid zip: central directory signature mismatch');
        }

        const compressionMethod = zipBuffer.readUInt16LE(offset + 10);
        const compressedSize = zipBuffer.readUInt32LE(offset + 20);
        const fileNameLength = zipBuffer.readUInt16LE(offset + 28);
        const extraLength = zipBuffer.readUInt16LE(offset + 30);
        const commentLength = zipBuffer.readUInt16LE(offset + 32);
        const localHeaderOffset = zipBuffer.readUInt32LE(offset + 42);
        const name = zipBuffer.toString('utf8', offset + 46, offset + 46 + fileNameLength);

        entries.set(name, readEntryData(zipBuffer, localHeaderOffset, compressedSize, compressionMethod));
        offset += 46 + fileNameLength + extraLength + commentLength;
    }

    return entries;
}
