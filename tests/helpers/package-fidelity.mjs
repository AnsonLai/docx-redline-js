import { createHash } from 'node:crypto';
import { unzipDocx } from '../../node/zip-archive.js';

/**
 * Normalizes a ZIP entry name to canonical OPC forward-slash form.
 *
 * @param {string} name
 * @returns {string}
 */
export function normalizeOpcEntryName(name) {
    if (typeof name !== 'string') {
        throw new TypeError(`ZIP entry name must be a string, got ${typeof name}`);
    }
    let normalized = name.replace(/\\/g, '/');
    normalized = normalized.replace(/^\/+/, '');
    while (normalized.startsWith('./')) {
        normalized = normalized.slice(2);
    }
    normalized = normalized.replace(/\/+/g, '/');
    if (!normalized) {
        throw new Error(`Invalid empty ZIP entry name: "${name}"`);
    }
    return normalized;
}

/**
 * Inventories every uncompressed entry in a ZIP package.
 *
 * @param {Buffer|Uint8Array|Map<string, Buffer>|object} input
 * @returns {{ entries: Map<string, { name: string, originalName: string, size: number, sha256: string, bytes: Buffer }>, hasEntry: (name: string) => boolean, getEntry: (name: string) => object|null }}
 */
export function inventoryPackage(input) {
    let rawEntries;
    if (Buffer.isBuffer(input) || input instanceof Uint8Array) {
        rawEntries = unzipDocx(input);
    } else if (input instanceof Map) {
        rawEntries = input;
    } else if (input?.entries instanceof Map) {
        rawEntries = input.entries;
    } else if (typeof input?.toBuffer === 'function') {
        rawEntries = unzipDocx(input.toBuffer());
    } else {
        throw new TypeError('inventoryPackage requires a Buffer, Uint8Array, Map of entries, or DocxDocument/MemoryZip.');
    }

    const normalizedMap = new Map();
    for (const [rawName, rawPayload] of rawEntries.entries()) {
        const normName = normalizeOpcEntryName(rawName);
        if (normalizedMap.has(normName)) {
            throw new Error(`Duplicate normalized ZIP entry path detected: "${normName}" (from "${rawName}")`);
        }
        const bytes = Buffer.isBuffer(rawPayload) ? rawPayload : Buffer.from(rawPayload);
        const sha256 = createHash('sha256').update(bytes).digest('hex');
        normalizedMap.set(normName, {
            name: normName,
            originalName: rawName,
            size: bytes.length,
            sha256,
            bytes
        });
    }

    const sortedEntries = new Map(
        Array.from(normalizedMap.entries()).sort(([a], [b]) => a.localeCompare(b))
    );

    return {
        entries: sortedEntries,
        hasEntry(name) {
            return sortedEntries.has(normalizeOpcEntryName(name));
        },
        getEntry(name) {
            return sortedEntries.get(normalizeOpcEntryName(name)) || null;
        }
    };
}

/**
 * Compares two package inventories or buffers and asserts exact payload preservation
 * for all entries except those explicitly allowed to change.
 *
 * @param {Buffer|Map|object} beforeInput
 * @param {Buffer|Map|object} afterInput
 * @param {string[]|Set<string>} [allowedChangedEntries=[]]
 * @returns {{ identical: boolean, report: string, unexpectedChanged: Array<object>, unexpectedAdded: Array<object>, unexpectedRemoved: Array<object>, changed: Array<object>, added: Array<object>, removed: Array<object>, unchanged: Array<object>, allowedModifications: Array<object> }}
 */
export function comparePackageEntries(beforeInput, afterInput, allowedChangedEntries = []) {
    const beforeInventory = beforeInput?.entries instanceof Map ? beforeInput : inventoryPackage(beforeInput);
    const afterInventory = afterInput?.entries instanceof Map ? afterInput : inventoryPackage(afterInput);

    const allowedSet = new Set(
        (Array.isArray(allowedChangedEntries) ? allowedChangedEntries : Array.from(allowedChangedEntries || []))
            .map(normalizeOpcEntryName)
    );

    const changed = [];
    const unexpectedChanged = [];
    const added = [];
    const unexpectedAdded = [];
    const removed = [];
    const unexpectedRemoved = [];
    const unchanged = [];
    const allowedModifications = [];

    const beforeNames = new Set(beforeInventory.entries.keys());
    const afterNames = new Set(afterInventory.entries.keys());

    for (const [name, beforeEntry] of beforeInventory.entries) {
        if (!afterNames.has(name)) {
            const isAllowed = allowedSet.has(name);
            const record = { name, before: beforeEntry };
            removed.push(record);
            if (isAllowed) {
                allowedModifications.push({ kind: 'removed', name, before: beforeEntry });
            } else {
                unexpectedRemoved.push(record);
            }
        } else {
            const afterEntry = afterInventory.entries.get(name);
            if (beforeEntry.sha256 === afterEntry.sha256) {
                unchanged.push({ name, size: beforeEntry.size, sha256: beforeEntry.sha256 });
            } else {
                const isAllowed = allowedSet.has(name);
                const record = {
                    name,
                    before: { size: beforeEntry.size, sha256: beforeEntry.sha256 },
                    after: { size: afterEntry.size, sha256: afterEntry.sha256 }
                };
                changed.push(record);
                if (isAllowed) {
                    allowedModifications.push({ kind: 'changed', ...record });
                } else {
                    unexpectedChanged.push(record);
                }
            }
        }
    }

    for (const [name, afterEntry] of afterInventory.entries) {
        if (!beforeNames.has(name)) {
            const isAllowed = allowedSet.has(name);
            const record = { name, after: afterEntry };
            added.push(record);
            if (isAllowed) {
                allowedModifications.push({ kind: 'added', name, after: afterEntry });
            } else {
                unexpectedAdded.push(record);
            }
        }
    }

    const identical = unexpectedChanged.length === 0 && unexpectedAdded.length === 0 && unexpectedRemoved.length === 0;

    let report = '';
    if (!identical) {
        const lines = ['Package fidelity comparison failed: unexpected changes detected in package entries:'];
        if (unexpectedChanged.length > 0) {
            lines.push(`  Unexpected modified entries (${unexpectedChanged.length}):`);
            for (const item of unexpectedChanged) {
                lines.push(`    - ${item.name}: size ${item.before.size} -> ${item.after.size}, sha256 ${item.before.sha256.slice(0, 8)}... -> ${item.after.sha256.slice(0, 8)}...`);
            }
        }
        if (unexpectedAdded.length > 0) {
            lines.push(`  Unexpected added entries (${unexpectedAdded.length}):`);
            for (const item of unexpectedAdded) {
                lines.push(`    - ${item.name}: size ${item.after.size}, sha256 ${item.after.sha256.slice(0, 8)}...`);
            }
        }
        if (unexpectedRemoved.length > 0) {
            lines.push(`  Unexpected removed entries (${unexpectedRemoved.length}):`);
            for (const item of unexpectedRemoved) {
                lines.push(`    - ${item.name}: size ${item.before.size}, sha256 ${item.before.sha256.slice(0, 8)}...`);
            }
        }
        report = lines.join('\n');
    }

    return {
        identical,
        report,
        unexpectedChanged,
        unexpectedAdded,
        unexpectedRemoved,
        changed,
        added,
        removed,
        unchanged,
        allowedModifications
    };
}

/**
 * Asserts that two packages have identical entries except for explicitly allowed changes.
 *
 * @param {Buffer|Map|object} beforeInput
 * @param {Buffer|Map|object} afterInput
 * @param {string[]|Set<string>} [allowedChangedEntries=[]]
 * @returns {object}
 */
export function assertPackageFidelity(beforeInput, afterInput, allowedChangedEntries = []) {
    const result = comparePackageEntries(beforeInput, afterInput, allowedChangedEntries);
    if (!result.identical) {
        const error = new Error(`[PACKAGE_FIDELITY_VIOLATION] ${result.report}`);
        error.code = 'PACKAGE_FIDELITY_VIOLATION';
        error.comparison = result;
        throw error;
    }
    return result;
}
