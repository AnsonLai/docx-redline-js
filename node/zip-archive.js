import { deflateRawSync, inflateRawSync } from 'node:zlib';

const CRC_TABLE = (() => { const table = new Uint32Array(256); for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1; table[n] = c >>> 0; } return table; })();
function crc32(buffer) { let crc = 0xffffffff; for (const byte of buffer) crc = CRC_TABLE[(crc ^ byte) & 0xff] ^ (crc >>> 8); return (crc ^ 0xffffffff) >>> 0; }

export function unzipDocx(input) {
    const source = Buffer.from(input);
    let eocd = -1;
    for (let offset = source.length - 22; offset >= Math.max(0, source.length - 0xffff - 22); offset--) if (source.readUInt32LE(offset) === 0x06054b50) { eocd = offset; break; }
    if (eocd < 0) throw new Error('Invalid DOCX: ZIP end record not found.');
    const count = source.readUInt16LE(eocd + 10); const centralOffset = source.readUInt32LE(eocd + 16);
    if (count === 0xffff || centralOffset === 0xffffffff) throw new Error('Unsupported DOCX: ZIP64 archives are not supported.');
    const entries = new Map(); let offset = centralOffset;
    for (let index = 0; index < count; index++) {
        if (source.readUInt32LE(offset) !== 0x02014b50) throw new Error('Invalid DOCX: central directory is corrupt.');
        const flags = source.readUInt16LE(offset + 8); const method = source.readUInt16LE(offset + 10);
        if (flags & 1) throw new Error('Unsupported DOCX: encrypted ZIP entries are not supported.');
        const compressedSize = source.readUInt32LE(offset + 20); const nameLength = source.readUInt16LE(offset + 28);
        const extraLength = source.readUInt16LE(offset + 30); const commentLength = source.readUInt16LE(offset + 32); const localOffset = source.readUInt32LE(offset + 42);
        const name = source.toString('utf8', offset + 46, offset + 46 + nameLength);
        if (source.readUInt32LE(localOffset) !== 0x04034b50) throw new Error('Invalid DOCX: local entry is corrupt.');
        const localNameLength = source.readUInt16LE(localOffset + 26); const localExtraLength = source.readUInt16LE(localOffset + 28);
        const start = localOffset + 30 + localNameLength + localExtraLength; const compressed = source.subarray(start, start + compressedSize);
        if (method !== 0 && method !== 8) throw new Error(`Unsupported DOCX compression method: ${method}.`);
        entries.set(name, method === 8 ? inflateRawSync(compressed) : Buffer.from(compressed));
        offset += 46 + nameLength + extraLength + commentLength;
    }
    return entries;
}

export function zipDocx(entries) {
    const local = []; const central = []; let offset = 0;
    for (const [name, raw] of entries) {
        const nameBytes = Buffer.from(name); const data = Buffer.from(raw); const compressed = deflateRawSync(data, { level: 9 });
        const method = compressed.length < data.length ? 8 : 0; const payload = method === 8 ? compressed : data; const crc = crc32(data);
        const lh = Buffer.alloc(30); lh.writeUInt32LE(0x04034b50); lh.writeUInt16LE(20,4); lh.writeUInt16LE(method,8); lh.writeUInt16LE(0x5c21,12); lh.writeUInt32LE(crc,14); lh.writeUInt32LE(payload.length,18); lh.writeUInt32LE(data.length,22); lh.writeUInt16LE(nameBytes.length,26);
        local.push(lh, nameBytes, payload);
        const ch = Buffer.alloc(46); ch.writeUInt32LE(0x02014b50); ch.writeUInt16LE(20,4); ch.writeUInt16LE(20,6); ch.writeUInt16LE(method,10); ch.writeUInt16LE(0x5c21,14); ch.writeUInt32LE(crc,16); ch.writeUInt32LE(payload.length,20); ch.writeUInt32LE(data.length,24); ch.writeUInt16LE(nameBytes.length,28); ch.writeUInt32LE(offset,42);
        central.push(ch, nameBytes); offset += lh.length + nameBytes.length + payload.length;
    }
    const directory = Buffer.concat(central); const end = Buffer.alloc(22); end.writeUInt32LE(0x06054b50); end.writeUInt16LE(entries.size,8); end.writeUInt16LE(entries.size,10); end.writeUInt32LE(directory.length,12); end.writeUInt32LE(offset,16);
    return Buffer.concat([...local, directory, end]);
}

export class MemoryZip {
    constructor(entries) { this.entries = entries; }
    file(path, value) {
        if (value !== undefined) { this.entries.set(path, Buffer.isBuffer(value) ? Buffer.from(value) : Buffer.from(String(value))); return this; }
        const data = this.entries.get(path); if (!data) return null;
        return { async: async type => type === 'string' ? data.toString('utf8') : Buffer.from(data) };
    }
}
