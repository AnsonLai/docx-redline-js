/**
 * Minimal zip writer for assembling validation .docx fixtures.
 *
 * Script-only helper (not part of the published API surface) so the package
 * keeps its no-zip-dependency guarantee while release tooling can still emit
 * real .docx files for Word/LibreOffice validation. Uses deflate via
 * node:zlib and a fixed timestamp for deterministic output.
 */

import { deflateRawSync } from 'zlib';

const CRC_TABLE = (() => {
    const table = new Uint32Array(256);
    for (let n = 0; n < 256; n++) {
        let c = n;
        for (let k = 0; k < 8; k++) {
            c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1;
        }
        table[n] = c >>> 0;
    }
    return table;
})();

function crc32(buffer) {
    let crc = 0xFFFFFFFF;
    for (let i = 0; i < buffer.length; i++) {
        crc = CRC_TABLE[(crc ^ buffer[i]) & 0xFF] ^ (crc >>> 8);
    }
    return (crc ^ 0xFFFFFFFF) >>> 0;
}

// Fixed DOS date/time (2026-01-01 00:00:00) keeps fixture bytes deterministic.
const DOS_TIME = 0;
const DOS_DATE = ((2026 - 1980) << 9) | (1 << 5) | 1;

/**
 * Builds a zip archive.
 *
 * @param {Array<{ name: string, data: Buffer|string }>} entries - Entry names
 *   must use forward slashes (OPC requirement for .docx parts).
 * @returns {Buffer}
 */
export function buildZip(entries) {
    const localParts = [];
    const centralParts = [];
    let offset = 0;

    for (const entry of entries) {
        const nameBytes = Buffer.from(entry.name, 'utf8');
        const data = Buffer.isBuffer(entry.data) ? entry.data : Buffer.from(entry.data, 'utf8');
        const crc = crc32(data);

        const deflated = deflateRawSync(data, { level: 9 });
        const useDeflate = deflated.length < data.length;
        const method = useDeflate ? 8 : 0;
        const payload = useDeflate ? deflated : data;

        const localHeader = Buffer.alloc(30);
        localHeader.writeUInt32LE(0x04034B50, 0);
        localHeader.writeUInt16LE(20, 4);            // version needed
        localHeader.writeUInt16LE(0, 6);             // flags
        localHeader.writeUInt16LE(method, 8);
        localHeader.writeUInt16LE(DOS_TIME, 10);
        localHeader.writeUInt16LE(DOS_DATE, 12);
        localHeader.writeUInt32LE(crc, 14);
        localHeader.writeUInt32LE(payload.length, 18);
        localHeader.writeUInt32LE(data.length, 22);
        localHeader.writeUInt16LE(nameBytes.length, 26);
        localHeader.writeUInt16LE(0, 28);            // extra length

        localParts.push(localHeader, nameBytes, payload);

        const centralHeader = Buffer.alloc(46);
        centralHeader.writeUInt32LE(0x02014B50, 0);
        centralHeader.writeUInt16LE(20, 4);          // version made by
        centralHeader.writeUInt16LE(20, 6);          // version needed
        centralHeader.writeUInt16LE(0, 8);           // flags
        centralHeader.writeUInt16LE(method, 10);
        centralHeader.writeUInt16LE(DOS_TIME, 12);
        centralHeader.writeUInt16LE(DOS_DATE, 14);
        centralHeader.writeUInt32LE(crc, 16);
        centralHeader.writeUInt32LE(payload.length, 20);
        centralHeader.writeUInt32LE(data.length, 24);
        centralHeader.writeUInt16LE(nameBytes.length, 28);
        centralHeader.writeUInt16LE(0, 30);          // extra length
        centralHeader.writeUInt16LE(0, 32);          // comment length
        centralHeader.writeUInt16LE(0, 34);          // disk number
        centralHeader.writeUInt16LE(0, 36);          // internal attrs
        centralHeader.writeUInt32LE(0, 38);          // external attrs
        centralHeader.writeUInt32LE(offset, 42);

        centralParts.push(centralHeader, nameBytes);
        offset += localHeader.length + nameBytes.length + payload.length;
    }

    const centralDirectory = Buffer.concat(centralParts);

    const endRecord = Buffer.alloc(22);
    endRecord.writeUInt32LE(0x06054B50, 0);
    endRecord.writeUInt16LE(0, 4);                   // disk number
    endRecord.writeUInt16LE(0, 6);                   // central dir start disk
    endRecord.writeUInt16LE(entries.length, 8);
    endRecord.writeUInt16LE(entries.length, 10);
    endRecord.writeUInt32LE(centralDirectory.length, 12);
    endRecord.writeUInt32LE(offset, 16);
    endRecord.writeUInt16LE(0, 20);                  // comment length

    return Buffer.concat([...localParts, centralDirectory, endRecord]);
}

const CONTENT_TYPES_BASE = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
%OVERRIDES%</Types>`;

const ROOT_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

/**
 * Assembles a minimal .docx package around a word/document.xml payload.
 *
 * @param {string} documentXml - Complete word/document.xml content
 * @param {{ numberingXml?: string|null }} [parts] - Optional extra parts
 * @returns {Buffer} - .docx bytes
 */
export function buildMinimalDocx(documentXml, parts = {}) {
    const overrides = [];
    const documentRels = [];
    const entries = [];

    if (parts.numberingXml) {
        overrides.push('  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>\n');
        documentRels.push('  <Relationship Id="rIdNum1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>');
    }

    entries.push({ name: '[Content_Types].xml', data: CONTENT_TYPES_BASE.replace('%OVERRIDES%', overrides.join('')) });
    entries.push({ name: '_rels/.rels', data: ROOT_RELS });
    entries.push({
        name: 'word/_rels/document.xml.rels',
        data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
${documentRels.join('\n')}
</Relationships>`
    });
    entries.push({ name: 'word/document.xml', data: documentXml });
    if (parts.numberingXml) {
        entries.push({ name: 'word/numbering.xml', data: parts.numberingXml });
    }

    return buildZip(entries);
}
