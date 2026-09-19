/**
 * Universal ZIP archive reader and writer for DOCX containers.
 *
 * Implemented with cross-runtime standards (Uint8Array, DataView, TextEncoder/Decoder)
 * and fflate for pure-JS DEFLATE compression / INFLATE decompression.
 * Free of Node.js built-ins (node:zlib, Buffer).
 */

import { deflateSync, inflateSync } from 'fflate';

const textEncoder = new TextEncoder();
const textDecoder = new TextDecoder('utf-8');

const CRC_TABLE = (() => {
    const table = new Uint32Array(256);
    for (let n = 0; n < 256; n++) {
        let c = n;
        for (let k = 0; k < 8; k++) {
            c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
        }
        table[n] = c >>> 0;
    }
    return table;
})();

/**
 * Computes standard CRC-32 checksum.
 *
 * @param {Uint8Array} bytes
 * @returns {number}
 */
export function crc32(bytes) {
    let crc = 0xffffffff;
    for (let i = 0; i < bytes.length; i++) {
        crc = CRC_TABLE[(crc ^ bytes[i]) & 0xff] ^ (crc >>> 8);
    }
    return (crc ^ 0xffffffff) >>> 0;
}

/**
 * Normalizes binary or string input into a Uint8Array.
 *
 * @param {Uint8Array|ArrayBuffer|ArrayBufferView|string} input
 * @returns {Uint8Array}
 */
export function toUint8Array(input) {
    if (input instanceof Uint8Array) {
        return input;
    }
    if (input instanceof ArrayBuffer) {
        return new Uint8Array(input);
    }
    if (ArrayBuffer.isView(input)) {
        return new Uint8Array(input.buffer, input.byteOffset, input.byteLength);
    }
    if (typeof input === 'string') {
        return textEncoder.encode(input);
    }
    throw new TypeError('Invalid binary input: expected Uint8Array, Buffer, or ArrayBuffer.');
}

/**
 * Concatenates multiple Uint8Arrays into a single Uint8Array.
 *
 * @param {Uint8Array[]} chunks
 * @returns {Uint8Array}
 */
function concatUint8Arrays(chunks) {
    let totalLen = 0;
    for (let i = 0; i < chunks.length; i++) {
        totalLen += chunks[i].length;
    }
    const result = new Uint8Array(totalLen);
    let offset = 0;
    for (let i = 0; i < chunks.length; i++) {
        result.set(chunks[i], offset);
        offset += chunks[i].length;
    }
    return result;
}

/**
 * Unzips a DOCX ZIP archive into a map of entry names to uncompressed byte arrays.
 *
 * @param {Uint8Array|ArrayBuffer|ArrayBufferView} input - ZIP archive bytes
 * @returns {Map<string, Uint8Array>} Map of entry paths to uncompressed Uint8Array payloads
 */
export function unzipDocx(input) {
    const source = toUint8Array(input);
    if (source.length < 22) {
        throw new Error('Invalid DOCX: ZIP end record not found.');
    }

    const view = new DataView(source.buffer, source.byteOffset, source.byteLength);
    let eocd = -1;
    const minOffset = Math.max(0, source.length - 0xffff - 22);

    for (let offset = source.length - 22; offset >= minOffset; offset--) {
        if (view.getUint32(offset, true) === 0x06054b50) {
            eocd = offset;
            break;
        }
    }

    if (eocd < 0) {
        throw new Error('Invalid DOCX: ZIP end record not found.');
    }

    const count = view.getUint16(eocd + 10, true);
    const centralOffset = view.getUint32(eocd + 16, true);

    if (count === 0xffff || centralOffset === 0xffffffff) {
        throw new Error('Unsupported DOCX: ZIP64 archives are not supported.');
    }

    const entries = new Map();
    let offset = centralOffset;

    for (let index = 0; index < count; index++) {
        if (offset + 46 > source.length || view.getUint32(offset, true) !== 0x02014b50) {
            throw new Error('Invalid DOCX: central directory is corrupt.');
        }

        const flags = view.getUint16(offset + 8, true);
        const method = view.getUint16(offset + 10, true);

        if (flags & 1) {
            throw new Error('Unsupported DOCX: encrypted ZIP entries are not supported.');
        }

        const compressedSize = view.getUint32(offset + 20, true);
        const nameLength = view.getUint16(offset + 28, true);
        const extraLength = view.getUint16(offset + 30, true);
        const commentLength = view.getUint16(offset + 32, true);
        const localOffset = view.getUint32(offset + 42, true);

        if (offset + 46 + nameLength > source.length) {
            throw new Error('Invalid DOCX: central directory entry name extends beyond bounds.');
        }

        const name = textDecoder.decode(source.subarray(offset + 46, offset + 46 + nameLength));

        if (localOffset + 30 > source.length || view.getUint32(localOffset, true) !== 0x04034b50) {
            throw new Error('Invalid DOCX: local entry is corrupt.');
        }

        const localNameLength = view.getUint16(localOffset + 26, true);
        const localExtraLength = view.getUint16(localOffset + 28, true);
        const start = localOffset + 30 + localNameLength + localExtraLength;

        if (start + compressedSize > source.length) {
            throw new Error('Invalid DOCX: local entry data extends beyond archive bounds.');
        }

        const compressed = source.subarray(start, start + compressedSize);

        if (method !== 0 && method !== 8) {
            throw new Error(`Unsupported DOCX compression method: ${method}.`);
        }

        const payload = method === 8 ? inflateSync(compressed) : new Uint8Array(compressed);
        entries.set(name, payload);
        offset += 46 + nameLength + extraLength + commentLength;
    }

    return entries;
}

/**
 * Zips uncompressed entries into a standard DOCX ZIP archive.
 *
 * @param {Map<string, Uint8Array|string>|Iterable<[string, any]>|Record<string, any>} entries
 * @returns {Uint8Array} Binary ZIP archive bytes
 */
export function zipDocx(entries) {
    const entryList = entries instanceof Map
        ? Array.from(entries.entries())
        : (Symbol.iterator in Object(entries) ? Array.from(entries) : Object.entries(entries || {}));

    const localChunks = [];
    const centralChunks = [];
    let offset = 0;

    for (const [name, raw] of entryList) {
        const nameBytes = textEncoder.encode(name);
        const data = toUint8Array(raw);
        const compressed = deflateSync(data, { level: 9 });
        const method = compressed.length < data.length ? 8 : 0;
        const payload = method === 8 ? compressed : data;
        const crc = crc32(data);

        // Local header: 30 bytes
        const lh = new Uint8Array(30);
        const lhView = new DataView(lh.buffer, lh.byteOffset, 30);
        lhView.setUint32(0, 0x04034b50, true);
        lhView.setUint16(4, 20, true);
        lhView.setUint16(8, method, true);
        lhView.setUint16(12, 0x5c21, true); // deterministic timestamp
        lhView.setUint32(14, crc, true);
        lhView.setUint32(18, payload.length, true);
        lhView.setUint32(22, data.length, true);
        lhView.setUint16(26, nameBytes.length, true);
        lhView.setUint16(28, 0, true);

        localChunks.push(lh, nameBytes, payload);

        // Central directory header: 46 bytes
        const ch = new Uint8Array(46);
        const chView = new DataView(ch.buffer, ch.byteOffset, 46);
        chView.setUint32(0, 0x02014b50, true);
        chView.setUint16(4, 20, true);
        chView.setUint16(6, 20, true);
        chView.setUint16(10, method, true);
        chView.setUint16(14, 0x5c21, true);
        chView.setUint32(16, crc, true);
        chView.setUint32(20, payload.length, true);
        chView.setUint32(24, data.length, true);
        chView.setUint16(28, nameBytes.length, true);
        chView.setUint16(30, 0, true);
        chView.setUint16(32, 0, true);
        chView.setUint32(42, offset, true);

        centralChunks.push(ch, nameBytes);
        offset += lh.length + nameBytes.length + payload.length;
    }

    const directory = concatUint8Arrays(centralChunks);
    const end = new Uint8Array(22);
    const endView = new DataView(end.buffer, end.byteOffset, 22);
    endView.setUint32(0, 0x06054b50, true);
    endView.setUint16(8, entryList.length, true);
    endView.setUint16(10, entryList.length, true);
    endView.setUint32(12, directory.length, true);
    endView.setUint32(16, offset, true);

    return concatUint8Arrays([...localChunks, directory, end]);
}

/**
 * Universal in-memory ZIP helper mimicking JSZip file() interface for internal consumers.
 */
export class MemoryZip {
    constructor(entries) {
        this.entries = entries instanceof Map ? entries : new Map(Object.entries(entries || {}));
    }

    file(path, value) {
        if (value !== undefined) {
            this.entries.set(path, toUint8Array(value));
            return this;
        }
        const data = this.entries.get(path);
        if (!data) return null;
        return {
            async: async type => type === 'string' ? textDecoder.decode(data) : new Uint8Array(data)
        };
    }
}
