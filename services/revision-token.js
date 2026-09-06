/**
 * Revision token framing and calculation for docx-redline-js.
 *
 * Implements unambiguous big-endian binary framing to generate cryptographic
 * revision tokens over package entries or document parts.
 */

const textEncoder = new TextEncoder();

/**
 * Normalizes an OPC entry or part name to canonical forward-slash form.
 *
 * @param {string} name - Raw entry name
 * @returns {string} Normalized entry name
 */
export function normalizeOpcEntryName(name) {
    if (typeof name !== 'string') {
        throw new TypeError(`Entry name must be a string, got ${typeof name}`);
    }
    let normalized = name.replace(/\\/g, '/');
    normalized = normalized.replace(/^\/+/, '');
    while (normalized.startsWith('./')) {
        normalized = normalized.slice(2);
    }
    normalized = normalized.replace(/\/+/g, '/');
    if (!normalized) {
        throw new Error(`Invalid empty entry name: "${name}"`);
    }
    return normalized;
}

/**
 * Builds the binary framing buffer for a revision token according to the specification:
 *
 * magic = "docx-redline-revision-token\0"
 * version = uint32be(1)
 * scopeLength + scope
 * entryCount
 * for each entry sorted by normalized name:
 *   nameLength + name
 *   payloadLength + payload
 *
 * @param {{ scope: string, entries: Array<{ name: string, payload: Uint8Array|string|Buffer }|Array> }} options
 * @returns {{ framing: Uint8Array, scope: string, version: number, coveredParts: string[] }}
 */
export function buildRevisionTokenFraming({ scope, entries = [] }) {
    if (typeof scope !== 'string' || !scope) {
        throw new TypeError('Revision token scope must be a non-empty string.');
    }
    const magicBytes = textEncoder.encode('docx-redline-revision-token\0');
    const version = 1;
    const scopeBytes = textEncoder.encode(scope);

    const normalizedEntries = [];
    const seenNames = new Set();

    const rawList = Array.isArray(entries)
        ? entries
        : (entries instanceof Map ? Array.from(entries.entries()) : Object.entries(entries || {}));

    for (const item of rawList) {
        if (!item) continue;
        const rawName = Array.isArray(item) ? item[0] : item.name;
        const rawPayload = Array.isArray(item) ? item[1] : (item.payload ?? item.bytes);

        if (rawName == null) continue;
        const normName = normalizeOpcEntryName(String(rawName));
        if (seenNames.has(normName)) {
            throw new Error(`Duplicate normalized entry path detected: "${normName}"`);
        }
        seenNames.add(normName);

        let payloadBytes;
        if (typeof rawPayload === 'string') {
            payloadBytes = textEncoder.encode(rawPayload);
        } else if (rawPayload instanceof Uint8Array) {
            payloadBytes = rawPayload;
        } else if (rawPayload && typeof rawPayload.length === 'number') {
            payloadBytes = new Uint8Array(rawPayload);
        } else {
            payloadBytes = new Uint8Array(0);
        }

        normalizedEntries.push({
            name: normName,
            nameBytes: textEncoder.encode(normName),
            payloadBytes
        });
    }

    // Sort entries by normalized name in binary codepoint order
    normalizedEntries.sort((a, b) => (a.name < b.name ? -1 : a.name > b.name ? 1 : 0));

    let totalSize = magicBytes.length + 4 + 4 + scopeBytes.length + 4;
    for (const e of normalizedEntries) {
        totalSize += 4 + e.nameBytes.length + 4 + e.payloadBytes.length;
    }

    const buffer = new Uint8Array(totalSize);
    const view = new DataView(buffer.buffer, buffer.byteOffset, buffer.byteLength);
    let offset = 0;

    // magic
    buffer.set(magicBytes, offset);
    offset += magicBytes.length;

    // version uint32be
    view.setUint32(offset, version, false);
    offset += 4;

    // scopeLength uint32be
    view.setUint32(offset, scopeBytes.length, false);
    offset += 4;

    // scope
    buffer.set(scopeBytes, offset);
    offset += scopeBytes.length;

    // entryCount uint32be
    view.setUint32(offset, normalizedEntries.length, false);
    offset += 4;

    for (const e of normalizedEntries) {
        // nameLength uint32be
        view.setUint32(offset, e.nameBytes.length, false);
        offset += 4;
        buffer.set(e.nameBytes, offset);
        offset += e.nameBytes.length;

        // payloadLength uint32be
        view.setUint32(offset, e.payloadBytes.length, false);
        offset += 4;
        buffer.set(e.payloadBytes, offset);
        offset += e.payloadBytes.length;
    }

    return {
        framing: buffer,
        scope,
        version,
        coveredParts: normalizedEntries.map(e => e.name)
    };
}

/**
 * Calculates a SHA-256 revision token asynchronously using Web Crypto or a custom digest function.
 *
 * @param {{ scope: string, entries: Array<any>, digestFn?: (bytes: Uint8Array) => Promise<string> }} options
 * @returns {Promise<{ algorithm: 'sha256', version: number, scope: string, value: string, coveredParts: string[] }>}
 */
export async function computeRevisionToken({ scope, entries = [], digestFn = null }) {
    const { framing, coveredParts, version } = buildRevisionTokenFraming({ scope, entries });

    let hashHex = '';
    if (typeof digestFn === 'function') {
        hashHex = await digestFn(framing);
    } else if (typeof globalThis.crypto?.subtle?.digest === 'function') {
        const hashBuf = await globalThis.crypto.subtle.digest('SHA-256', framing);
        const hashBytes = new Uint8Array(hashBuf);
        hashHex = Array.from(hashBytes).map(b => b.toString(16).padStart(2, '0')).join('');
    } else {
        throw new Error('No crypto provider available for SHA-256 revision token computation.');
    }

    return {
        algorithm: 'sha256',
        version,
        scope,
        value: hashHex,
        coveredParts
    };
}

/**
 * Calculates a SHA-256 revision token synchronously using a provided digest function.
 *
 * @param {{ scope: string, entries: Array<any>, digestFn: (bytes: Uint8Array) => string }} options
 * @returns {{ algorithm: 'sha256', version: number, scope: string, value: string, coveredParts: string[] }}
 */
export function computeRevisionTokenSync({ scope, entries = [], digestFn }) {
    if (typeof digestFn !== 'function') {
        throw new TypeError('computeRevisionTokenSync requires a synchronous digestFn.');
    }
    const { framing, coveredParts, version } = buildRevisionTokenFraming({ scope, entries });
    const hashHex = digestFn(framing);
    return {
        algorithm: 'sha256',
        version,
        scope,
        value: hashHex,
        coveredParts
    };
}

/**
 * Extracts entry records from document parts descriptor.
 *
 * @param {object} parts - Document parts object
 * @returns {Array<{ name: string, payload: any }>}
 */
export function extractDocumentPartsEntries(parts) {
    const entries = [];
    if (parts?.documentXml) {
        entries.push({ name: 'word/document.xml', payload: parts.documentXml });
    }
    if (parts?.commentsXml) {
        entries.push({ name: 'word/comments.xml', payload: parts.commentsXml });
    }
    if (parts?.numberingXml) {
        entries.push({ name: 'word/numbering.xml', payload: parts.numberingXml });
    }
    if (parts?.stylesXml) {
        entries.push({ name: 'word/styles.xml', payload: parts.stylesXml });
    }
    if (parts?.parts instanceof Map) {
        for (const [name, payload] of parts.parts.entries()) {
            entries.push({ name, payload });
        }
    } else if (parts?.additionalParts && typeof parts.additionalParts === 'object') {
        for (const [name, payload] of Object.entries(parts.additionalParts)) {
            entries.push({ name, payload });
        }
    }
    return entries;
}

/**
 * Computes a revision token for document parts asynchronously.
 *
 * @param {object} parts - Document parts
 * @param {object} [options={}] - Options
 * @returns {Promise<{ algorithm: 'sha256', version: number, scope: 'document-parts', value: string, coveredParts: string[] }>}
 */
export async function computeDocumentPartsRevisionToken(parts, options = {}) {
    const entries = extractDocumentPartsEntries(parts);
    return computeRevisionToken({
        scope: 'document-parts',
        entries,
        digestFn: options.digestFn
    });
}

/**
 * Validates the structure and syntax of an incoming revision token object.
 *
 * @param {any} token
 * @returns {{ valid: boolean, error?: { code: string, message: string } }}
 */
export function validateRevisionToken(token) {
    if (!token || typeof token !== 'object') {
        return { valid: false, error: { code: 'INVALID_REVISION_TOKEN', message: 'Revision token must be an object.' } };
    }
    if (token.algorithm !== 'sha256') {
        return { valid: false, error: { code: 'INVALID_REVISION_TOKEN', message: `Unsupported revision token algorithm: "${token.algorithm}". Expected "sha256".` } };
    }
    if (token.version !== 1) {
        return { valid: false, error: { code: 'INVALID_REVISION_TOKEN', message: `Unsupported revision token version: "${token.version}". Expected 1.` } };
    }
    if (token.scope !== 'document-parts' && token.scope !== 'package') {
        return { valid: false, error: { code: 'INVALID_REVISION_TOKEN', message: `Unsupported revision token scope: "${token.scope}". Expected "document-parts" or "package".` } };
    }
    if (typeof token.value !== 'string' || !/^[0-9a-f]{64}$/i.test(token.value.trim())) {
        return { valid: false, error: { code: 'INVALID_REVISION_TOKEN', message: 'Revision token value must be a 64-character hex string.' } };
    }
    return { valid: true };
}

/**
 * Compares two revision token values using timing-safe byte comparison.
 *
 * @param {string} a
 * @param {string} b
 * @returns {boolean}
 */
export function areRevisionTokensEqual(a, b) {
    if (typeof a !== 'string' || typeof b !== 'string') return false;
    const aNorm = a.trim().toLowerCase();
    const bNorm = b.trim().toLowerCase();
    if (aNorm.length !== bNorm.length) return false;
    const aBuf = textEncoder.encode(aNorm);
    const bBuf = textEncoder.encode(bNorm);
    let diff = 0;
    for (let i = 0; i < aBuf.length; i++) {
        diff |= aBuf[i] ^ bBuf[i];
    }
    return diff === 0;
}
