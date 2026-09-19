/**
 * Node.js compatibility layer for ZIP archives in DOCX packages.
 *
 * Delegates compression and decompression to the universal pure-JS implementation
 * in document/zip-archive.js while returning Node.js Buffers for backward compatibility.
 */

import {
    unzipDocx as universalUnzipDocx,
    zipDocx as universalZipDocx
} from '../document/zip-archive.js';

export function unzipDocx(input) {
    const rawEntries = universalUnzipDocx(input);
    const entries = new Map();
    for (const [name, payload] of rawEntries) {
        entries.set(name, Buffer.from(payload.buffer, payload.byteOffset, payload.byteLength));
    }
    return entries;
}

export function zipDocx(entries) {
    return Buffer.from(universalZipDocx(entries));
}

export class MemoryZip {
    constructor(entries) {
        this.entries = entries instanceof Map ? entries : new Map(Object.entries(entries || {}));
    }

    file(path, value) {
        if (value !== undefined) {
            this.entries.set(path, Buffer.isBuffer(value) ? Buffer.from(value) : Buffer.from(String(value)));
            return this;
        }
        const data = this.entries.get(path);
        if (!data) return null;
        return {
            async: async type => type === 'string' ? data.toString('utf8') : Buffer.from(data)
        };
    }
}
