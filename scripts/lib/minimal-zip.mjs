/**
 * Minimal zip writer for assembling validation .docx fixtures.
 *
 * Script-only helper (not part of the published API surface) so the package
 * keeps its no-zip-dependency guarantee while release tooling can still emit
 * real .docx files for Word/LibreOffice validation. Uses deflate via
 * node:zlib and a fixed timestamp for deterministic output.
 */

import { deflateRawSync } from 'zlib';
import { DOMParser } from '@xmldom/xmldom';

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

const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const REL_BASE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
const CONTENT_TYPES = {
    numbering: 'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml',
    comments: 'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml',
    footnotes: 'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml',
    endnotes: 'application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml',
    header: 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml',
    footer: 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml'
};

function escapeXmlAttribute(value) {
    return String(value)
        .replace(/&/g, '&amp;')
        .replace(/"/g, '&quot;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;');
}

function parsePartXml(xml, label) {
    const errors = [];
    const document = new DOMParser({
        onError: (_level, message) => errors.push(message)
    }).parseFromString(String(xml), 'application/xml');
    if (!document?.documentElement || errors.length > 0 || document.getElementsByTagName('parsererror').length > 0) {
        throw new Error(`${label} is not well-formed XML${errors[0] ? `: ${errors[0]}` : ''}`);
    }
    return document;
}

function wordIds(document, localName) {
    return Array.from(document.getElementsByTagNameNS(NS_W, localName), node =>
        node.getAttributeNS?.(NS_W, 'id') || node.getAttribute('w:id') || node.getAttribute('id')
    );
}

function relationshipIds(document, localName) {
    return Array.from(document.getElementsByTagNameNS(NS_W, localName), node =>
        node.getAttributeNS?.(NS_R, 'id') || node.getAttribute('r:id') || node.getAttribute('id')
    ).filter(Boolean);
}

function requireRoot(document, localName, label) {
    if (document.documentElement.namespaceURI !== NS_W || document.documentElement.localName !== localName) {
        throw new Error(`${label} must have w:${localName} as its document element`);
    }
}

function validateReferencedIds(documentXml, parts) {
    const document = parsePartXml(documentXml, 'word/document.xml');
    requireRoot(document, 'document', 'word/document.xml');

    const documentCommentIds = [
        ...wordIds(document, 'commentRangeStart'),
        ...wordIds(document, 'commentRangeEnd'),
        ...wordIds(document, 'commentReference')
    ];
    if (documentCommentIds.length > 0 && !parts.commentsXml) {
        throw new Error('document comment anchors require word/comments.xml');
    }
    if (parts.commentsXml) {
        const comments = parsePartXml(parts.commentsXml, 'word/comments.xml');
        requireRoot(comments, 'comments', 'word/comments.xml');
        const defined = new Set(wordIds(comments, 'comment'));
        const starts = wordIds(document, 'commentRangeStart');
        const ends = wordIds(document, 'commentRangeEnd');
        const references = wordIds(document, 'commentReference');
        for (const id of new Set([...starts, ...ends, ...references])) {
            if (!defined.has(id)) throw new Error(`word/comments.xml does not define referenced comment ID ${id}`);
        }
        if (starts.length === 0 || JSON.stringify(starts) !== JSON.stringify(ends) || JSON.stringify(starts) !== JSON.stringify(references)) {
            throw new Error('comment start/end/reference IDs must be present and ordered identically');
        }
    }

    for (const family of ['footnote', 'endnote']) {
        const partKey = `${family}sXml`;
        const referenceIds = wordIds(document, `${family}Reference`);
        if (referenceIds.length > 0 && !parts[partKey]) {
            throw new Error(`document ${family} references require word/${family}s.xml`);
        }
        if (!parts[partKey]) continue;
        const partName = `word/${family}s.xml`;
        const notes = parsePartXml(parts[partKey], partName);
        requireRoot(notes, `${family}s`, partName);
        const defined = new Set(wordIds(notes, family));
        if (!defined.has('-1') || !defined.has('0')) {
            throw new Error(`${partName} must define separator ID -1 and continuation separator ID 0`);
        }
        for (const id of referenceIds) {
            if (!defined.has(id)) throw new Error(`${partName} does not define referenced ${family} ID ${id}`);
        }
    }

    const headerIds = new Set((parts.headers || []).map(header => header.relationshipId));
    const footerIds = new Set((parts.footers || []).map(footer => footer.relationshipId));
    for (const id of relationshipIds(document, 'headerReference')) {
        if (!headerIds.has(id)) throw new Error(`document headerReference ${id} has no configured header relationship`);
    }
    for (const id of relationshipIds(document, 'footerReference')) {
        if (!footerIds.has(id)) throw new Error(`document footerReference ${id} has no configured footer relationship`);
    }

    const hyperlinkIds = new Set((parts.externalHyperlinks || []).map(link => link.relationshipId));
    for (const id of relationshipIds(document, 'hyperlink')) {
        if (!hyperlinkIds.has(id)) throw new Error(`external hyperlink ${id} has no configured relationship`);
    }
}

function normalizeRelatedParts(parts) {
    const normalized = {
        ...parts,
        headers: (parts.headers || []).map((header, index) => ({
            partName: header.partName || `header${index + 1}.xml`,
            relationshipId: header.relationshipId || `rIdHeader${index + 1}`,
            xml: header.xml
        })),
        footers: (parts.footers || []).map((footer, index) => ({
            partName: footer.partName || `footer${index + 1}.xml`,
            relationshipId: footer.relationshipId || `rIdFooter${index + 1}`,
            xml: footer.xml
        })),
        externalHyperlinks: parts.externalHyperlinks || []
    };

    const relationshipIds = [];
    if (normalized.numberingXml) relationshipIds.push('rIdNum1');
    if (normalized.commentsXml) relationshipIds.push('rIdComments1');
    if (normalized.footnotesXml) relationshipIds.push('rIdFootnotes1');
    if (normalized.endnotesXml) relationshipIds.push('rIdEndnotes1');
    relationshipIds.push(...normalized.headers.map(item => item.relationshipId));
    relationshipIds.push(...normalized.footers.map(item => item.relationshipId));
    relationshipIds.push(...normalized.externalHyperlinks.map(item => item.relationshipId));
    if (relationshipIds.some(id => !/^rId[A-Za-z0-9._-]+$/.test(String(id)))) {
        throw new Error('every relationship ID must use a non-empty rId-prefixed token');
    }
    if (new Set(relationshipIds).size !== relationshipIds.length) throw new Error('document relationship IDs must be unique');

    for (const [family, items, pattern] of [
        ['header', normalized.headers, /^header[1-9][0-9]*\.xml$/],
        ['footer', normalized.footers, /^footer[1-9][0-9]*\.xml$/]
    ]) {
        for (const item of items) {
            if (!pattern.test(item.partName)) throw new Error(`${family} part name must match ${family}<number>.xml`);
            if (typeof item.xml !== 'string' || item.xml.length === 0) throw new Error(`${family} ${item.partName} requires XML content`);
            const xml = parsePartXml(item.xml, `word/${item.partName}`);
            requireRoot(xml, family === 'header' ? 'hdr' : 'ftr', `word/${item.partName}`);
        }
    }
    const partNames = [...normalized.headers, ...normalized.footers].map(item => item.partName);
    if (new Set(partNames).size !== partNames.length) throw new Error('header/footer part names must be unique');

    for (const link of normalized.externalHyperlinks) {
        if (!link.relationshipId || !link.target) throw new Error('external hyperlinks require relationshipId and target');
        let url;
        try { url = new URL(link.target); } catch { throw new Error(`invalid external hyperlink target: ${link.target}`); }
        if (!url.protocol || url.protocol === 'file:') throw new Error(`unsupported external hyperlink target: ${link.target}`);
    }
    return normalized;
}

/**
 * Creates and validates the deterministic entry set used by buildMinimalDocx.
 * Exported for package-integrity tests; this remains script-only tooling.
 */
export function buildMinimalDocxEntries(documentXml, parts = {}) {
    const normalized = normalizeRelatedParts(parts);
    validateReferencedIds(documentXml, normalized);
    const overrides = [];
    const documentRels = [];
    const entries = [];

    const addXmlPart = (partName, contentType, relationshipId, relationshipType, xml) => {
        overrides.push(`  <Override PartName="/word/${partName}" ContentType="${contentType}"/>\n`);
        documentRels.push(`  <Relationship Id="${escapeXmlAttribute(relationshipId)}" Type="${REL_BASE}${relationshipType}" Target="${escapeXmlAttribute(partName)}"/>`);
        entries.push({ name: `word/${partName}`, data: xml });
    };

    if (normalized.numberingXml) addXmlPart('numbering.xml', CONTENT_TYPES.numbering, 'rIdNum1', 'numbering', normalized.numberingXml);
    if (normalized.commentsXml) addXmlPart('comments.xml', CONTENT_TYPES.comments, 'rIdComments1', 'comments', normalized.commentsXml);
    if (normalized.footnotesXml) addXmlPart('footnotes.xml', CONTENT_TYPES.footnotes, 'rIdFootnotes1', 'footnotes', normalized.footnotesXml);
    if (normalized.endnotesXml) addXmlPart('endnotes.xml', CONTENT_TYPES.endnotes, 'rIdEndnotes1', 'endnotes', normalized.endnotesXml);
    normalized.headers.forEach(item => addXmlPart(item.partName, CONTENT_TYPES.header, item.relationshipId, 'header', item.xml));
    normalized.footers.forEach(item => addXmlPart(item.partName, CONTENT_TYPES.footer, item.relationshipId, 'footer', item.xml));
    normalized.externalHyperlinks.forEach(link => {
        documentRels.push(`  <Relationship Id="${escapeXmlAttribute(link.relationshipId)}" Type="${REL_BASE}hyperlink" Target="${escapeXmlAttribute(link.target)}" TargetMode="External"/>`);
    });

    entries.unshift({ name: 'word/document.xml', data: documentXml });
    entries.unshift({
        name: 'word/_rels/document.xml.rels',
        data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
${documentRels.join('\n')}
</Relationships>`
    });
    entries.unshift({ name: '_rels/.rels', data: ROOT_RELS });
    entries.unshift({ name: '[Content_Types].xml', data: CONTENT_TYPES_BASE.replace('%OVERRIDES%', overrides.join('')) });
    return entries;
}

/**
 * Assembles a minimal .docx package around a word/document.xml payload.
 *
 * @param {string} documentXml - Complete word/document.xml content
 * @param {Object} [parts] - Optional numbering/comments/notes/header/footer/
 *   external-hyperlink package parts. This helper is development-only.
 * @returns {Buffer} - .docx bytes
 */
export function buildMinimalDocx(documentXml, parts = {}) {
    return buildZip(buildMinimalDocxEntries(documentXml, parts));
}
