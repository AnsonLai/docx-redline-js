/**
 * Universal DOCX document facade and package-level operation runner.
 *
 * Implements complete-document manipulation using cross-runtime standards
 * (Uint8Array, TextEncoder/Decoder) and the universal ZIP archive layer.
 * Runs in Node.js, browsers, Cloudflare Workers, Deno, and sandboxed runtimes.
 */

import { getDefaultAuthor } from '../adapters/config.js';
import { inspectDocumentParts } from '../services/document-inspection.js';
import { applyOperationsToDocumentXml, preflightOperations } from '../services/standalone-operation-runner.js';
import { createDynamicNumberingIdState, mergeNumberingXmlBySchemaOrder } from '../services/numbering-helpers.js';
import { ensureCommentsArtifactsInZip, ensureCommentsExtendedArtifactsInZip, ensureNumberingArtifactsInZip, validateDocxPackage } from '../services/standalone-docx-plumbing.js';
import { validateRedlineOoxml } from '../core/redline-validation.js';
import { subtractValidationIssueMultiset, validationErrors } from '../core/validation-delta.js';
import { acceptTrackedChangesInOoxml, rejectTrackedChangesInOoxml, deleteCommentsByAuthorInOoxml } from '../services/revision-comment-management.js';
import { createSerializer, parseOoxmlSafe } from '../adapters/xml-adapter.js';
import { sha256 } from '../core/sha256.js';
import { MemoryZip, unzipDocx, zipDocx, toUint8Array } from './zip-archive.js';
import { computeRevisionTokenSync, validateRevisionToken, areRevisionTokensEqual } from '../services/revision-token.js';
import { createRetryPlan, normalizeErrorWithRecovery } from '../services/error-recovery.js';

const textEncoder = new TextEncoder();
const textDecoder = new TextDecoder('utf-8');

const text = (entries, path) => {
    const entry = entries.get(path);
    if (!entry) return null;
    if (typeof entry === 'string') return entry;
    return textDecoder.decode(entry);
};

const cloneEntries = entries => new Map([...entries].map(([name, data]) => [name, new Uint8Array(data)]));

function toBufferCompatible(uint8Arr) {
    if (typeof Buffer !== 'undefined' && typeof Buffer.from === 'function') {
        return Buffer.from(uint8Arr.buffer, uint8Arr.byteOffset, uint8Arr.byteLength);
    }
    return uint8Arr;
}

function areByteArraysEqual(a, b) {
    if (a === b) return true;
    if (!a || !b) return false;
    if (a.length !== b.length) return false;
    for (let i = 0; i < a.length; i++) {
        if (a[i] !== b[i]) return false;
    }
    return true;
}

function rolledBackOperationPayload(operationResult, operationCount) {
    const rollbackReceipt = receipt => {
        if (!receipt || typeof receipt !== 'object') return receipt;
        if (receipt.attemptedDisposition !== 'applied' && receipt.committed !== true) return { ...receipt };
        return { ...receipt, committed: false, finalDisposition: 'rolled_back' };
    };
    const receipts = (operationResult?.receipts || []).map(rollbackReceipt);
    const receiptByIndex = new Map(receipts.map(receipt => [receipt.operationIndex, receipt]));
    const results = (operationResult?.results || []).map(result => ({
        ...result,
        ...(result.receipt ? {
            receipt: receiptByIndex.get(result.index) || rollbackReceipt(result.receipt)
        } : {})
    }));
    return {
        results,
        receipts,
        executionOrder: operationResult?.executionOrder || [],
        authorsUsed: [],
        rolledBack: true,
        retryPlan: createRetryPlan({
            atomic: true,
            rolledBack: true,
            results,
            receipts,
            operationCount
        })
    };
}

/**
 * Computes a package-scoped revision token over all uncompressed entries in a DOCX archive.
 *
 * @param {Uint8Array|ArrayBuffer|Map<string, Uint8Array>|DocxDocument|object} input
 * @returns {{ algorithm: 'sha256', version: number, scope: 'package', value: string, coveredParts: string[] }}
 */
export function computePackageRevisionToken(input) {
    let entries;
    if (input instanceof Uint8Array || (typeof Buffer !== 'undefined' && Buffer.isBuffer?.(input))) {
        entries = unzipDocx(input);
    } else if (input instanceof Map) {
        entries = input;
    } else if (input?.entries instanceof Map) {
        entries = input.entries;
    } else if (typeof input?.toUint8Array === 'function') {
        entries = unzipDocx(input.toUint8Array());
    } else if (typeof input?.toBuffer === 'function') {
        entries = unzipDocx(input.toBuffer());
    } else if (input instanceof ArrayBuffer || ArrayBuffer.isView(input)) {
        entries = unzipDocx(toUint8Array(input));
    } else {
        throw new TypeError('computePackageRevisionToken requires a Uint8Array, Buffer, Map of entries, or DocxDocument.');
    }
    return computeRevisionTokenSync({
        scope: 'package',
        entries
    });
}

function nextCommentId(entries) {
    const ids = `${text(entries, 'word/document.xml') || ''} ${text(entries, 'word/comments.xml') || ''}`.match(/(?:w:)?id=["'](\d+)["']/g) || [];
    let next = ids.reduce((max, token) => Math.max(max, Number(token.match(/\d+/)?.[0] || 0)), 0) + 1;
    return () => next++;
}

function existingCommentDetails(entries) {
    const commentsXml = text(entries, 'word/comments.xml');
    if (!commentsXml) return {};
    const parsed = parseOoxmlSafe(commentsXml, 'application/xml');
    if (!parsed.doc || parsed.error) return {};
    const details = {};
    for (const comment of Array.from(parsed.doc.getElementsByTagNameNS('*', 'comment'))) {
        const id = comment.getAttribute('w:id') || comment.getAttribute('id');
        if (id === '') continue;
        details[id] = {
            id,
            author: comment.getAttribute('w:author') || comment.getAttribute('author') || '',
            text: String(comment.textContent || '').trim(),
            paraId: Array.from(comment.getElementsByTagNameNS('*', 'p'))[0]?.getAttribute('w14:paraId') || null
        };
    }
    return details;
}

function packageFailure(sourceBytes, code, message) {
    const buf = toBufferCompatible(sourceBytes);
    return {
        status: 'error',
        hasChanges: false,
        written: false,
        rolledBack: true,
        error: normalizeErrorWithRecovery({ code, message }),
        retryPlan: createRetryPlan({ atomic: true, rolledBack: true }),
        artifactsChanged: [],
        uint8Array: sourceBytes,
        toUint8Array: () => new Uint8Array(sourceBytes),
        buffer: buf,
        toBuffer: () => toBufferCompatible(sourceBytes)
    };
}

/**
 * Universal DOCX document representation.
 */
export class DocxDocument {
    constructor(input) {
        this.originalBytes = toUint8Array(input);
        this.entries = unzipDocx(this.originalBytes);
    }

    get originalBuffer() {
        return toBufferCompatible(this.originalBytes);
    }

    inspect(options = {}) {
        const digestFn = options.digestFn || (bytes => sha256(bytes));
        return inspectDocumentParts({
            documentXml: text(this.entries, 'word/document.xml'),
            commentsXml: text(this.entries, 'word/comments.xml'),
            commentsExtendedXml: text(this.entries, 'word/commentsExtended.xml'),
            numberingXml: text(this.entries, 'word/numbering.xml'),
            stylesXml: text(this.entries, 'word/styles.xml')
        }, { ...options, digestFn });
    }

    getRevisionToken() {
        return computePackageRevisionToken(this.entries);
    }

    get revisionToken() {
        return this.getRevisionToken();
    }

    preflight(operations, author = getDefaultAuthor(), options = {}) {
        return preflightOperations(
            text(this.entries, 'word/document.xml'),
            operations,
            author || getDefaultAuthor(),
            { ...options, _existingCommentDetails: existingCommentDetails(this.entries) }
        );
    }

    toUint8Array() {
        return zipDocx(this.entries);
    }

    toBuffer() {
        return toBufferCompatible(this.toUint8Array());
    }

    async applyOperations(operations, options = {}) {
        const failedApply = error => {
            const results = [];
            const receipts = [];
            const buf = toBufferCompatible(this.originalBytes);
            return {
                status: 'error',
                hasChanges: false,
                written: false,
                rolledBack: true,
                results,
                receipts,
                artifactsChanged: [],
                error: normalizeErrorWithRecovery(error),
                retryPlan: createRetryPlan({
                    atomic: true,
                    rolledBack: true,
                    results,
                    receipts,
                    operationCount: Array.isArray(operations) ? operations.length : 0
                }),
                validation: { originalIssues: [], generatedIssues: [] },
                uint8Array: this.originalBytes,
                toUint8Array: () => new Uint8Array(this.originalBytes),
                buffer: buf,
                toBuffer: () => toBufferCompatible(this.originalBytes)
            };
        };

        if (options?.expectedRevision) {
            const tokenValidation = validateRevisionToken(options.expectedRevision);
            if (!tokenValidation.valid) {
                return failedApply({
                    code: tokenValidation.error?.code || 'INVALID_REVISION_TOKEN',
                    message: tokenValidation.error?.message || 'Invalid revision token.'
                });
            }
            if (options.expectedRevision.scope !== 'package') {
                return failedApply({
                    code: 'REVISION_TOKEN_SCOPE_MISMATCH',
                    message: `Revision token scope mismatch: expected 'package', got '${options.expectedRevision.scope}'.`,
                    expectedScope: 'package',
                    actualScope: options.expectedRevision.scope
                });
            }
            const currentToken = computePackageRevisionToken(this.entries);
            if (!areRevisionTokensEqual(currentToken.value, options.expectedRevision.value)) {
                return failedApply({
                    code: 'REVISION_MISMATCH',
                    message: `Document revision mismatch: expected '${options.expectedRevision.value}', current is '${currentToken.value}'.`,
                    expectedRevision: options.expectedRevision,
                    currentRevision: currentToken
                });
            }
        }

        const originalEntries = this.entries;
        const working = cloneEntries(originalEntries);
        const zip = new MemoryZip(working);
        const documentXml = text(working, 'word/document.xml');
        let originalIssues = [];
        let operationResult = null;

        try {
            if (!documentXml) throw new Error('Missing word/document.xml.');
            const baseline = validateRedlineOoxml(documentXml);
            originalIssues = baseline.issues.map(issue => ({ source: 'word/document.xml', ...issue }));
            try {
                await validateDocxPackage(new MemoryZip(cloneEntries(originalEntries)));
            } catch (error) {
                originalIssues.push({ source: 'package', code: 'PACKAGE_VALIDATION', severity: 'error', message: error.message });
            }

            const context = {
                numberingIdState: createDynamicNumberingIdState(text(working, 'word/numbering.xml') || undefined),
                commentsXml: text(working, 'word/comments.xml'),
                commentsExtendedXml: text(working, 'word/commentsExtended.xml')
            };
            const { expectedRevision: _pkgExpectedRevision, ...runnerOptions } = options;
            const author = options.author || getDefaultAuthor();
            const result = operationResult = await applyOperationsToDocumentXml(documentXml, operations, author, context, {
                ...runnerOptions,
                atomic: options.atomic === true,
                strictTargets: options.strictTargets !== false,
                _existingCommentDetails: existingCommentDetails(working),
                commentIdAllocator: nextCommentId(working)
            });

            if (result.rolledBack || result.status === 'error') {
                const buf = toBufferCompatible(this.originalBytes);
                return {
                    ...result,
                    ...rolledBackOperationPayload(result, Array.isArray(operations) ? operations.length : 0),
                    written: false,
                    artifactsChanged: [],
                    validation: { originalIssues, generatedIssues: [] },
                    uint8Array: this.originalBytes,
                    toUint8Array: () => new Uint8Array(this.originalBytes),
                    buffer: buf,
                    toBuffer: () => toBufferCompatible(this.originalBytes)
                };
            }

            if (!result.hasChanges) {
                const buf = toBufferCompatible(this.originalBytes);
                return {
                    ...result,
                    status: result.status || 'ok',
                    written: false,
                    artifactsChanged: [],
                    validation: { originalIssues, generatedIssues: [] },
                    uint8Array: this.originalBytes,
                    toUint8Array: () => new Uint8Array(this.originalBytes),
                    buffer: buf,
                    toBuffer: () => toBufferCompatible(this.originalBytes)
                };
            }

            working.set('word/document.xml', textEncoder.encode(result.documentXml));
            await ensureNumberingArtifactsInZip(zip, result.numberingXmlParts, { mergeNumberingXml: mergeNumberingXmlBySchemaOrder });

            const existingCommentsXml = text(working, 'word/comments.xml');
            const commentsXmlForPackaging = result.commentsXml || existingCommentsXml;
            await ensureCommentsArtifactsInZip(zip, commentsXmlForPackaging, {
                replaceExisting: result.commentsXmlMode === 'replace' || (!result.commentsXml && !!existingCommentsXml)
            });

            const existingCommentsExtendedXml = text(working, 'word/commentsExtended.xml');
            const commentsExtendedXmlForPackaging = result.commentsExtendedXml || existingCommentsExtendedXml;
            await ensureCommentsExtendedArtifactsInZip(zip, commentsExtendedXmlForPackaging, {
                replaceExisting: result.commentsExtendedXmlMode === 'replace' || (!result.commentsExtendedXml && !!existingCommentsExtendedXml)
            });

            if (options.validate !== false) {
                const generated = validateRedlineOoxml(result.documentXml);
                const outputIssues = generated.issues.map(issue => ({ source: 'word/document.xml', ...issue }));
                try {
                    await validateDocxPackage(zip);
                } catch (error) {
                    outputIssues.push({ source: 'package', code: 'PACKAGE_VALIDATION', severity: 'error', message: error.message });
                }
                const introduced = subtractValidationIssueMultiset(outputIssues, originalIssues);
                const introducedErrors = validationErrors(introduced);
                if (introducedErrors.length) {
                    const codes = [...new Set(introducedErrors.map(issue => issue.code))].join(', ');
                    throw Object.assign(
                        new Error(`Applied operations introduced invalid revision markup (${codes}); these are generated-output issues, not pre-existing input issues.`),
                        { issues: introducedErrors }
                    );
                }
            }

            this.entries = working;
            const outputBytes = this.toUint8Array();
            this.originalBytes = outputBytes;
            const outputBuf = toBufferCompatible(outputBytes);

            const artifactsChanged = [...working]
                .filter(([name, data]) => !originalEntries.has(name) || !areByteArraysEqual(data, originalEntries.get(name)))
                .map(([name]) => name);

            return {
                ...result,
                status: result.status || 'ok',
                written: true,
                artifactsChanged,
                validation: { originalIssues, generatedIssues: [] },
                uint8Array: outputBytes,
                toUint8Array: () => new Uint8Array(outputBytes),
                buffer: outputBuf,
                inspection: this.inspect(),
                toBuffer: () => toBufferCompatible(outputBytes)
            };
        } catch (error) {
            this.entries = originalEntries;
            const generatedIssues = error.issues || [{ source: 'package', code: 'PACKAGE_OPERATION_FAILED', severity: 'error', message: error.message }];
            const rollbackPayload = rolledBackOperationPayload(
                operationResult,
                Array.isArray(operations) ? operations.length : 0
            );
            const buf = toBufferCompatible(this.originalBytes);
            return {
                ...rollbackPayload,
                status: 'error',
                hasChanges: false,
                written: false,
                artifactsChanged: [],
                error: normalizeErrorWithRecovery({
                    code: 'PACKAGE_OPERATION_FAILED',
                    message: error.message,
                    stage: 'package',
                    issues: generatedIssues
                }),
                validation: { originalIssues, generatedIssues },
                issues: generatedIssues,
                uint8Array: this.originalBytes,
                toUint8Array: () => new Uint8Array(this.originalBytes),
                buffer: buf,
                toBuffer: () => toBufferCompatible(this.originalBytes)
            };
        }
    }

    async resolveRevisions(action, options = {}) {
        const transform = action === 'accept' ? acceptTrackedChangesInOoxml : action === 'reject' ? rejectTrackedChangesInOoxml : null;
        if (!transform) return packageFailure(this.originalBytes, 'INVALID_ACTION', `Unknown revision action: ${action}`);

        const sourceBytes = this.originalBytes;
        const working = cloneEntries(this.entries);
        const zip = new MemoryZip(working);
        const result = transform(text(working, 'word/document.xml'), { author: options.author, allAuthors: options.allAuthors === true });

        if (result.status === 'error' || result.error) {
            return packageFailure(sourceBytes, result.error?.code || 'REVISION_OPERATION_FAILED', result.error?.message || 'Revision operation failed.');
        }

        if (!result.hasChanges) {
            const buf = toBufferCompatible(sourceBytes);
            return {
                ...result,
                status: 'ok',
                written: false,
                artifactsChanged: [],
                uint8Array: sourceBytes,
                toUint8Array: () => new Uint8Array(sourceBytes),
                buffer: buf,
                toBuffer: () => toBufferCompatible(sourceBytes)
            };
        }

        working.set('word/document.xml', textEncoder.encode(result.oxml));
        try {
            if (options.validate !== false) await validateDocxPackage(zip);
        } catch (error) {
            return packageFailure(sourceBytes, 'PACKAGE_VALIDATION', error.message);
        }

        this.entries = working;
        const outputBytes = this.toUint8Array();
        this.originalBytes = outputBytes;
        const outputBuf = toBufferCompatible(outputBytes);

        return {
            ...result,
            status: 'ok',
            written: true,
            artifactsChanged: ['word/document.xml'],
            uint8Array: outputBytes,
            toUint8Array: () => new Uint8Array(outputBytes),
            buffer: outputBuf,
            toBuffer: () => toBufferCompatible(outputBytes)
        };
    }

    async deleteComments(options = {}) {
        const sourceBytes = this.originalBytes;
        const working = cloneEntries(this.entries);
        const commentsXml = text(working, 'word/comments.xml');
        const buf = toBufferCompatible(sourceBytes);

        if (!commentsXml) {
            return {
                status: 'ok',
                hasChanges: false,
                written: false,
                commentsRemoved: 0,
                referencesRemoved: 0,
                artifactsChanged: [],
                uint8Array: sourceBytes,
                toUint8Array: () => new Uint8Array(sourceBytes),
                buffer: buf,
                toBuffer: () => toBufferCompatible(sourceBytes)
            };
        }

        const parsed = parseOoxmlSafe(commentsXml, 'application/xml');
        if (!parsed.doc || parsed.error) {
            return packageFailure(sourceBytes, 'PARSE_ERROR', parsed.error?.message || 'Could not parse comments.xml.');
        }

        const matches = comment => options.allAuthors === true || (comment.getAttribute('w:author') || comment.getAttribute('author')) === options.author;
        const ids = new Set(
            Array.from(parsed.doc.getElementsByTagNameNS('*', 'comment'))
                .filter(matches)
                .map(node => node.getAttribute('w:id') || node.getAttribute('id'))
                .filter(Boolean)
        );

        if (!ids.size) {
            return {
                status: 'ok',
                hasChanges: false,
                written: false,
                commentsRemoved: 0,
                referencesRemoved: 0,
                artifactsChanged: [],
                uint8Array: sourceBytes,
                toUint8Array: () => new Uint8Array(sourceBytes),
                buffer: buf,
                toBuffer: () => toBufferCompatible(sourceBytes)
            };
        }

        const commentsExtendedXml = text(working, 'word/commentsExtended.xml');
        let extendedParsed = null;
        const removedParaIds = new Set();

        if (commentsExtendedXml) {
            extendedParsed = parseOoxmlSafe(commentsExtendedXml, 'application/xml');
            if (!extendedParsed.doc || extendedParsed.error) {
                return packageFailure(sourceBytes, 'PARSE_ERROR', extendedParsed.error?.message || 'Could not parse commentsExtended.xml.');
            }
            const idByParaId = new Map();
            for (const comment of Array.from(parsed.doc.getElementsByTagNameNS('*', 'comment'))) {
                const id = comment.getAttribute('w:id') || comment.getAttribute('id');
                const paragraph = Array.from(comment.getElementsByTagNameNS('*', 'p'))[0];
                const paraId = paragraph?.getAttribute('w14:paraId') || paragraph?.getAttribute('paraId');
                if (paraId) idByParaId.set(paraId.toUpperCase(), id);
            }
            for (const [paraId, id] of idByParaId) {
                if (ids.has(id)) removedParaIds.add(paraId);
            }
            let expanded = true;
            while (expanded) {
                expanded = false;
                for (const entry of Array.from(extendedParsed.doc.getElementsByTagNameNS('*', 'commentEx'))) {
                    const paraId = (entry.getAttribute('w15:paraId') || entry.getAttribute('paraId') || '').toUpperCase();
                    const parentParaId = (entry.getAttribute('w15:paraIdParent') || entry.getAttribute('paraIdParent') || '').toUpperCase();
                    if (parentParaId && removedParaIds.has(parentParaId) && !removedParaIds.has(paraId)) {
                        removedParaIds.add(paraId);
                        if (idByParaId.has(paraId)) ids.add(idByParaId.get(paraId));
                        expanded = true;
                    }
                }
            }
        }

        const commentsResult = deleteCommentsByAuthorInOoxml(commentsXml, { author: options.author, allAuthors: options.allAuthors === true });
        const remainingParsed = parseOoxmlSafe(commentsResult.oxml, 'application/xml');
        if (!remainingParsed.doc || remainingParsed.error) {
            return packageFailure(sourceBytes, 'PARSE_ERROR', remainingParsed.error?.message || 'Could not parse updated comments.xml.');
        }

        for (const comment of Array.from(remainingParsed.doc.getElementsByTagNameNS('*', 'comment'))) {
            const id = comment.getAttribute('w:id') || comment.getAttribute('id');
            if (ids.has(id)) comment.parentNode?.removeChild(comment);
        }

        const documentParsed = parseOoxmlSafe(text(working, 'word/document.xml'), 'application/xml');
        if (!documentParsed.doc || documentParsed.error) {
            return packageFailure(sourceBytes, 'PARSE_ERROR', documentParsed.error?.message || 'Could not parse document.xml.');
        }

        let referencesRemoved = 0;
        for (const name of ['commentRangeStart', 'commentRangeEnd', 'commentReference']) {
            for (const node of Array.from(documentParsed.doc.getElementsByTagNameNS('*', name))) {
                const id = node.getAttribute('w:id') || node.getAttribute('id');
                if (!ids.has(id) || !node.parentNode) continue;
                const parent = node.parentNode;
                parent.removeChild(node);
                referencesRemoved += 1;
                if (name === 'commentReference' && parent.localName === 'r' && !Array.from(parent.childNodes || []).some(child => child.nodeType === 1 && child.localName !== 'rPr')) {
                    parent.parentNode?.removeChild(parent);
                }
            }
        }

        const serializer = createSerializer();
        working.set('word/comments.xml', textEncoder.encode(serializer.serializeToString(remainingParsed.doc)));

        if (extendedParsed?.doc) {
            for (const entry of Array.from(extendedParsed.doc.getElementsByTagNameNS('*', 'commentEx'))) {
                const paraId = (entry.getAttribute('w15:paraId') || entry.getAttribute('paraId') || '').toUpperCase();
                if (removedParaIds.has(paraId)) entry.parentNode?.removeChild(entry);
            }
            working.set('word/commentsExtended.xml', textEncoder.encode(serializer.serializeToString(extendedParsed.doc)));
        }

        working.set('word/document.xml', textEncoder.encode(serializer.serializeToString(documentParsed.doc)));

        try {
            if (options.validate !== false) await validateDocxPackage(new MemoryZip(working));
        } catch (error) {
            return packageFailure(sourceBytes, 'PACKAGE_VALIDATION', error.message);
        }

        this.entries = working;
        const outputBytes = this.toUint8Array();
        this.originalBytes = outputBytes;
        const outputBuf = toBufferCompatible(outputBytes);

        return {
            status: 'ok',
            hasChanges: true,
            written: true,
            commentsRemoved: ids.size,
            referencesRemoved,
            artifactsChanged: ['word/document.xml', 'word/comments.xml', ...(extendedParsed?.doc ? ['word/commentsExtended.xml'] : [])],
            uint8Array: outputBytes,
            toUint8Array: () => new Uint8Array(outputBytes),
            buffer: outputBuf,
            toBuffer: () => toBufferCompatible(outputBytes)
        };
    }
}

/**
 * Opens a DOCX document package from a byte array or Buffer.
 *
 * @param {Uint8Array|ArrayBuffer|ArrayBufferView|string} input - Binary DOCX archive
 * @returns {DocxDocument}
 */
export function openDocx(input) {
    return new DocxDocument(input);
}
