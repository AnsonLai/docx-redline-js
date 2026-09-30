/**
 * Universal DOCX document facade and package-level operation runner.
 *
 * Implements complete-document manipulation using cross-runtime standards
 * (Uint8Array, TextEncoder/Decoder) and the universal ZIP archive layer.
 * Runs in Node.js, browsers, Cloudflare Workers, Deno, and sandboxed runtimes.
 */

import { getDefaultAuthor } from '../adapters/config.js';
import { applyThreadResolutionToParts, commentParaId, reconcileCommentSiblingParts } from '../services/comment-thread-parts.js';
import { collectRevisionIds, countRevisionsInsideFields, discoverHeaderFooterParts, isHeaderFooterPath, renumberCollidingRevisionIds, resolvePartSelector } from '../services/headers-footers.js';
import { inspectDocumentParts } from '../services/document-inspection.js';
import { applyOperationsToDocumentXml, preflightOperations } from '../services/standalone-operation-runner.js';
import { createDynamicNumberingIdState, mergeNumberingXmlBySchemaOrder } from '../services/numbering-helpers.js';
import { ensureCommentsArtifactsInZip, ensureCommentsExtendedArtifactsInZip, ensureNumberingArtifactsInZip, repairKnownContentTypes, validateDocxPackage } from '../services/standalone-docx-plumbing.js';
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

/** Keeps commentsIds/commentsExtensible in step with comments.xml after comments were added or removed. */
function reconcileCommentSiblings(working) {
    const updated = reconcileCommentSiblingParts({
        commentsXml: text(working, 'word/comments.xml'),
        commentsIdsXml: text(working, 'word/commentsIds.xml'),
        commentsExtensibleXml: text(working, 'word/commentsExtensible.xml')
    });
    if (updated.commentsIdsXml) working.set('word/commentsIds.xml', textEncoder.encode(updated.commentsIdsXml));
    if (updated.commentsExtensibleXml) working.set('word/commentsExtensible.xml', textEncoder.encode(updated.commentsExtensibleXml));
}

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

// Word does not support comments in headers or footers, so these are refused when `part` is set.
const COMMENT_OPERATION_TYPES = new Set(['comment', 'comment_reply', 'comment_resolve']);

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
            paraId: commentParaId(comment) || null
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
        const inspection = inspectDocumentParts({
            documentXml: text(this.entries, 'word/document.xml'),
            commentsXml: text(this.entries, 'word/comments.xml'),
            commentsExtendedXml: text(this.entries, 'word/commentsExtended.xml'),
            numberingXml: text(this.entries, 'word/numbering.xml'),
            stylesXml: text(this.entries, 'word/styles.xml')
        }, { ...options, digestFn });
        if (inspection.status === 'error') return inspection;
        return { ...inspection, headersFooters: this.inspectHeadersFooters() };
    }

    /**
     * Header and footer parts the document uses, with their paragraph text, so operations can target them with
     * `part`. `hasFields` marks parts containing PAGE/NUMPAGES-style fields (edit around them, not inside).
     */
    inspectHeadersFooters() {
        return discoverHeaderFooterParts({
            documentXml: text(this.entries, 'word/document.xml'),
            relsXml: text(this.entries, 'word/_rels/document.xml.rels'),
            settingsXml: text(this.entries, 'word/settings.xml')
        }).map(part => {
            const xml = text(this.entries, part.path) || '';
            const partInspection = xml ? inspectDocumentParts({ documentXml: xml }) : { paragraphs: [] };
            return {
                ...part,
                hasFields: /<w:fldChar\b|<w:fldSimple\b/.test(xml),
                paragraphs: (partInspection.paragraphs || []).map(paragraph => ({ index: paragraph.index, text: paragraph.text }))
            };
        });
    }

    getRevisionToken() {
        return computePackageRevisionToken(this.entries);
    }

    get revisionToken() {
        return this.getRevisionToken();
    }

    preflight(operations, author = getDefaultAuthor(), options = {}) {
        const list = Array.isArray(operations) ? operations : [];
        const runBody = ops => preflightOperations(
            text(this.entries, 'word/document.xml'),
            ops,
            author || getDefaultAuthor(),
            { ...options, _existingCommentDetails: existingCommentDetails(this.entries) }
        );
        if (!list.some(operation => operation && typeof operation === 'object' && operation.part != null)) return runBody(operations);

        const parts = this.inspectHeadersFooters();
        const bodyOps = [];
        const results = [];
        const conflicts = [];
        const authorsUsed = new Set();
        let valid = true;
        list.forEach((operation, i) => {
            if (!operation || typeof operation !== 'object' || operation.part == null) {
                bodyOps.push({ operation, origIndex: i });
                return;
            }
            const resolved = resolvePartSelector(parts, operation.part);
            const refusal = COMMENT_OPERATION_TYPES.has(operation.type)
                ? { code: 'COMMENT_IN_HEADER_FOOTER', message: 'Word does not support comments in headers or footers.' }
                : resolved.error;
            if (refusal) {
                valid = false;
                results.push({ index: i + 1, type: operation.type || 'redline', status: 'error', error: normalizeErrorWithRecovery({ ...refusal, operationIndex: i + 1 }) });
                return;
            }
            const { part: _selector, ...rest } = operation;
            const partResult = preflightOperations(text(this.entries, resolved.part.path), [rest], author || getDefaultAuthor(), { ...options, _existingCommentDetails: {} });
            valid = valid && partResult.valid;
            for (const result of partResult.results || []) results.push({ ...result, index: i + 1, part: resolved.part.path });
            conflicts.push(...(partResult.conflicts || []));
            for (const name of partResult.authorsUsed || []) authorsUsed.add(name);
        });
        let requiredArtifacts = { comments: false, numbering: false };
        if (bodyOps.length) {
            const body = runBody(bodyOps.map(item => item.operation));
            valid = valid && body.valid;
            for (const result of body.results || []) results.push({ ...result, index: bodyOps[result.index - 1].origIndex + 1 });
            conflicts.push(...(body.conflicts || []));
            for (const name of body.authorsUsed || []) authorsUsed.add(name);
            requiredArtifacts = body.requiredArtifacts || requiredArtifacts;
        }
        results.sort((a, b) => a.index - b.index);
        const ok = valid && conflicts.length === 0;
        return { valid: ok, status: ok ? 'ok' : 'error', results, conflicts, authorsUsed: [...authorsUsed], requiredArtifacts };
    }

    toUint8Array() {
        return zipDocx(this.entries);
    }

    toBuffer() {
        return toBufferCompatible(this.toUint8Array());
    }

    /**
     * Applies operations to the document. Operations without `part` target the body (word/document.xml).
     * An operation with `part` targets one header or footer part:
     * `{ part: { kind: 'footer', type: 'default', section: 0 } | 'word/footer2.xml', target: { exactText }, ... }`.
     * Discover parts with `inspect().headersFooters`.
     */
    async applyOperations(operations, options = {}) {
        const list = Array.isArray(operations) ? operations : [];
        if (!list.some(operation => operation && typeof operation === 'object' && operation.part != null)) {
            return this.applyBodyOperations(operations, options);
        }
        return this.applyOperationsWithParts(list, options);
    }

    async applyOperationsWithParts(operations, options = {}) {
        const author = options.author || getDefaultAuthor();
        const atomic = options.atomic === true;
        const snapshotEntries = this.entries;
        const snapshotBytes = this.originalBytes;
        const total = operations.length;
        const restore = () => { this.entries = snapshotEntries; this.originalBytes = snapshotBytes; };
        const original = () => ({
            uint8Array: snapshotBytes,
            toUint8Array: () => new Uint8Array(snapshotBytes),
            buffer: toBufferCompatible(snapshotBytes),
            toBuffer: () => toBufferCompatible(snapshotBytes)
        });
        const refuse = (error, results = [], receipts = [], extra = {}) => ({
            status: 'error', hasChanges: false, written: false, rolledBack: true, results, receipts, executionOrder: [], authorsUsed: [],
            artifactsChanged: [],
            error: normalizeErrorWithRecovery(error),
            retryPlan: createRetryPlan({ atomic: true, rolledBack: true, results, receipts, operationCount: total }),
            validation: { originalIssues: [], generatedIssues: [] },
            ...extra,
            ...original()
        });

        if (options.expectedRevision) {
            const precondition = await this.applyBodyOperations([], { expectedRevision: options.expectedRevision });
            if (precondition.status === 'error') return precondition;
        }

        // Resolve every selector before touching anything, so a bad selector fails the whole request.
        const parts = discoverHeaderFooterParts({
            documentXml: text(this.entries, 'word/document.xml'),
            relsXml: text(this.entries, 'word/_rels/document.xml.rels'),
            settingsXml: text(this.entries, 'word/settings.xml')
        });
        const bodyOps = [];
        const partOps = [];
        for (let i = 0; i < operations.length; i++) {
            const operation = operations[i];
            if (!operation || typeof operation !== 'object' || operation.part == null) {
                bodyOps.push({ operation, origIndex: i });
                continue;
            }
            if (COMMENT_OPERATION_TYPES.has(operation.type)) {
                return refuse({ code: 'COMMENT_IN_HEADER_FOOTER', message: `Operation ${i + 1}: Word does not support comments in headers or footers.`, operationIndex: i + 1 });
            }
            const resolved = resolvePartSelector(parts, operation.part);
            if (resolved.error) return refuse({ ...resolved.error, message: `Operation ${i + 1}: ${resolved.error.message}`, operationIndex: i + 1 });
            const { part: _selector, ...rest } = operation;
            partOps.push({ operation: rest, origIndex: i, part: resolved.part });
        }

        const results = [];
        const receipts = [];
        const executionOrder = [];
        const authorsUsed = new Set();
        const originalIssues = [];
        let anyChange = false;
        let anyFailure = false;
        let bodyResult = null;

        if (bodyOps.length) {
            bodyResult = await this.applyBodyOperations(bodyOps.map(item => item.operation), { ...options, expectedRevision: undefined });
            const toOriginal = index => bodyOps[index - 1]?.origIndex + 1;
            for (const result of bodyResult.results || []) {
                results.push({
                    ...result,
                    index: toOriginal(result.index),
                    ...(result.receipt ? { receipt: { ...result.receipt, operationIndex: toOriginal(result.receipt.operationIndex) } } : {})
                });
            }
            for (const receipt of bodyResult.receipts || []) receipts.push({ ...receipt, operationIndex: toOriginal(receipt.operationIndex) });
            for (const index of bodyResult.executionOrder || []) executionOrder.push(toOriginal(index));
            for (const name of bodyResult.authorsUsed || []) authorsUsed.add(name);
            originalIssues.push(...(bodyResult.validation?.originalIssues || []));
            if (bodyResult.status === 'error' || bodyResult.rolledBack) {
                anyFailure = true;
                if (atomic) {
                    restore();
                    return { ...bodyResult, results, receipts, executionOrder, written: false, artifactsChanged: [], rolledBack: true, ...original() };
                }
            }
            anyChange = bodyResult.hasChanges === true;
        }

        const working = cloneEntries(this.entries);
        const zip = new MemoryZip(working);
        const changedParts = new Set();
        const { expectedRevision: _ignored, ...runnerOptions } = options;
        const failPart = (item, error) => {
            anyFailure = true;
            const failure = normalizeErrorWithRecovery({ ...error, operationIndex: item.origIndex + 1 });
            results.push({ index: item.origIndex + 1, type: item.operation.type || 'redline', status: 'error', authorUsed: author, part: item.part.path, error: failure });
        };

        for (const item of partOps) {
            const path = item.part.path;
            const xml = text(working, path);
            if (!xml) {
                failPart(item, { code: 'PART_NOT_FOUND', message: `Part '${path}' is referenced but missing from the package.` });
                if (atomic) break;
                continue;
            }
            const numberingText = text(working, 'word/numbering.xml');
            const outcome = await applyOperationsToDocumentXml(xml, [item.operation], author, { numberingIdState: createDynamicNumberingIdState(numberingText || undefined) }, {
                ...runnerOptions, atomic: true, strictTargets: options.strictTargets !== false, _existingCommentDetails: {}
            });
            const single = outcome.results?.[0];
            if (outcome.rolledBack || outcome.status === 'error') {
                failPart(item, single?.error || outcome.error || { code: 'OPERATION_FAILED', message: 'Operation failed.' });
                if (atomic) break;
                continue;
            }
            if (!outcome.hasChanges) {
                results.push({ ...single, index: item.origIndex + 1, part: path });
                continue;
            }
            if (countRevisionsInsideFields(outcome.documentXml) > countRevisionsInsideFields(xml)) {
                failPart(item, { code: 'FIELD_EDIT_REFUSED', message: `The edit would change a field (for example PAGE or NUMPAGES) in ${path}. Word recomputes fields; edit the text around them instead.` });
                if (atomic) break;
                continue;
            }
            const taken = new Set();
            for (const [name, data] of working) {
                if (name === path || !(name === 'word/document.xml' || isHeaderFooterPath(name))) continue;
                for (const id of collectRevisionIds(textDecoder.decode(data))) taken.add(id);
            }
            const renumbered = renumberCollidingRevisionIds(outcome.documentXml, taken);
            working.set(path, textEncoder.encode(renumbered.xml));
            // Receipts must describe what is in the file: the part it lives in and its final revision ids.
            const describeReceipt = receipt => (receipt?.revisionItems
                ? { ...receipt, revisionItems: receipt.revisionItems.map(revision => ({ ...revision, id: renumbered.idMap.get(String(revision.id)) ?? revision.id, partName: path })) }
                : receipt);
            if (outcome.numberingXmlParts?.length) {
                await ensureNumberingArtifactsInZip(zip, outcome.numberingXmlParts, { mergeNumberingXml: mergeNumberingXmlBySchemaOrder });
            }
            changedParts.add(path);
            anyChange = true;
            results.push({
                ...single,
                index: item.origIndex + 1,
                part: path,
                partSections: item.part.appliesToSections,
                ...(single?.receipt ? { receipt: { ...describeReceipt(single.receipt), operationIndex: item.origIndex + 1 } } : {})
            });
            if (outcome.receipts?.[0]) receipts.push({ ...describeReceipt(outcome.receipts[0]), operationIndex: item.origIndex + 1 });
            executionOrder.push(item.origIndex + 1);
            for (const name of outcome.authorsUsed || []) authorsUsed.add(name);
        }

        // Each changed part must not introduce revision-markup problems, and the package must still validate.
        const generatedIssues = [];
        if (!(atomic && anyFailure)) {
            for (const path of changedParts) {
                const before = validateRedlineOoxml(text(this.entries, path)).issues.map(issue => ({ source: path, ...issue }));
                const after = validateRedlineOoxml(text(working, path)).issues.map(issue => ({ source: path, ...issue }));
                generatedIssues.push(...validationErrors(subtractValidationIssueMultiset(after, before)));
            }
            if (options.validate !== false && changedParts.size) {
                try {
                    await validateDocxPackage(zip);
                } catch (error) {
                    generatedIssues.push({ source: 'package', code: 'PACKAGE_VALIDATION', severity: 'error', message: error.message });
                }
            }
        }
        if ((atomic && anyFailure) || generatedIssues.length) {
            restore();
            const error = generatedIssues.length
                ? { code: 'PACKAGE_OPERATION_FAILED', message: `Applied operations introduced invalid markup (${[...new Set(generatedIssues.map(issue => issue.code))].join(', ')}).`, issues: generatedIssues }
                : { code: 'BATCH_OPERATION_FAILED', message: 'Atomic batch rolled back because one or more operations failed.' };
            const rolledBackResults = results
                .map(result => (result.receipt ? { ...result, receipt: { ...result.receipt, committed: false, finalDisposition: 'rolled_back' } } : result))
                .sort((a, b) => a.index - b.index);
            const rolledBackReceipts = receipts.map(receipt => ({ ...receipt, committed: false, finalDisposition: 'rolled_back' }));
            return refuse(error, rolledBackResults, rolledBackReceipts, { validation: { originalIssues, generatedIssues }, issues: generatedIssues });
        }

        results.sort((a, b) => a.index - b.index);
        receipts.sort((a, b) => a.operationIndex - b.operationIndex);
        const status = !anyFailure ? 'ok' : (anyChange ? 'partial' : 'error');
        const failureError = anyFailure
            ? { error: normalizeErrorWithRecovery({ code: 'BATCH_OPERATION_FAILED', message: 'One or more operations failed; the others were applied.' }) }
            : {};
        if (!changedParts.size) {
            // No part changed; the body call (if any) already committed its own result.
            const buf = toBufferCompatible(this.originalBytes);
            return {
                ...(bodyResult || {}), results, receipts, executionOrder, authorsUsed: [...authorsUsed], status, hasChanges: anyChange,
                written: bodyResult?.written === true, artifactsChanged: bodyResult?.artifactsChanged || [],
                validation: { originalIssues, generatedIssues: [] }, ...failureError,
                uint8Array: this.originalBytes, toUint8Array: () => new Uint8Array(this.originalBytes), buffer: buf, toBuffer: () => toBufferCompatible(this.originalBytes)
            };
        }

        this.entries = working;
        const outputBytes = this.toUint8Array();
        this.originalBytes = outputBytes;
        const artifactsChanged = [...working]
            .filter(([name, data]) => !snapshotEntries.has(name) || !areByteArraysEqual(data, snapshotEntries.get(name)))
            .map(([name]) => name);
        return {
            ...(bodyResult || {}),
            status, hasChanges: true, written: true, results, receipts, executionOrder, authorsUsed: [...authorsUsed], artifactsChanged,
            validation: { originalIssues, generatedIssues: [] }, ...failureError,
            uint8Array: outputBytes, toUint8Array: () => new Uint8Array(outputBytes), buffer: toBufferCompatible(outputBytes),
            inspection: this.inspect(), toBuffer: () => toBufferCompatible(outputBytes)
        };
    }

    async applyBodyOperations(operations, options = {}) {
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

            // Resolving a thread only changes the comment parts; keep document.xml byte-identical.
            const bodyChanged = (result.results || []).some(item => item?.status !== 'error' && item?.operationType !== 'comment_resolve');
            if (bodyChanged) working.set('word/document.xml', textEncoder.encode(result.documentXml));
            await ensureNumberingArtifactsInZip(zip, result.numberingXmlParts, { mergeNumberingXml: mergeNumberingXmlBySchemaOrder });

            const existingCommentsXml = text(working, 'word/comments.xml');
            const commentsXmlForPackaging = result.commentsXml || existingCommentsXml;
            await ensureCommentsArtifactsInZip(zip, commentsXmlForPackaging, {
                replaceExisting: result.commentsXmlMode === 'replace' || (!result.commentsXml && !!existingCommentsXml)
            });

            // Only touch the extended part (and its content type/rel) when an operation changed it.
            // Rewriting it on every save is what stamped a bad content type onto every
            // already-commented document.
            if (result.commentsExtendedXml) {
                await ensureCommentsExtendedArtifactsInZip(zip, result.commentsExtendedXml, {
                    replaceExisting: result.commentsExtendedXmlMode === 'replace'
                });
            }
            await repairKnownContentTypes(zip);
            reconcileCommentSiblings(working);

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
        const originalEntries = this.entries;
        const working = cloneEntries(this.entries);
        const zip = new MemoryZip(working);
        const result = transform(text(working, 'word/document.xml'), { author: options.author, allAuthors: options.allAuthors === true });

        if (result.status === 'error' || result.error) {
            return packageFailure(sourceBytes, result.error?.code || 'REVISION_OPERATION_FAILED', result.error?.message || 'Revision operation failed.');
        }

        // Headers and footers carry their own tracked changes; resolve them with the same filter.
        const partUpdates = new Map();
        for (const part of this.inspectHeadersFooters()) {
            const partXml = text(working, part.path);
            if (!partXml) continue;
            const partResult = transform(partXml, { author: options.author, allAuthors: options.allAuthors === true });
            if (partResult.status === 'error' || partResult.error) {
                return packageFailure(sourceBytes, partResult.error?.code || 'REVISION_OPERATION_FAILED', `${part.path}: ${partResult.error?.message || 'Revision operation failed.'}`);
            }
            if (partResult.hasChanges) partUpdates.set(part.path, partResult.oxml);
        }
        const hasChanges = result.hasChanges || partUpdates.size > 0;

        if (!hasChanges) {
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

        if (result.hasChanges) working.set('word/document.xml', textEncoder.encode(result.oxml));
        for (const [path, xml] of partUpdates) working.set(path, textEncoder.encode(xml));
        try {
            await repairKnownContentTypes(zip);
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
            hasChanges: true,
            artifactsChanged: [...working]
                .filter(([name, data]) => !originalEntries.has(name) || !areByteArraysEqual(data, originalEntries.get(name)))
                .map(([name]) => name),
            uint8Array: outputBytes,
            toUint8Array: () => new Uint8Array(outputBytes),
            buffer: outputBuf,
            toBuffer: () => toBufferCompatible(outputBytes)
        };
    }

    /**
     * Resolves (or reopens) the comment thread containing `commentId`.
     * Resolved is a thread-level state in Word, so the root and every reply are updated together.
     */
    async resolveComment(commentId, options = {}) {
        const sourceBytes = this.originalBytes;
        const working = cloneEntries(this.entries);
        const resolved = options.resolved !== false;
        const noChange = extra => {
            const buf = toBufferCompatible(sourceBytes);
            return {
                status: 'ok', hasChanges: false, written: false, artifactsChanged: [], ...extra,
                uint8Array: sourceBytes, toUint8Array: () => new Uint8Array(sourceBytes), buffer: buf, toBuffer: () => toBufferCompatible(sourceBytes)
            };
        };

        const result = applyThreadResolutionToParts({
            commentsXml: text(working, 'word/comments.xml'),
            commentsExtendedXml: text(working, 'word/commentsExtended.xml'),
            commentId,
            resolved
        });
        if (result.status === 'error') return packageFailure(sourceBytes, result.error.code, result.error.message);
        if (!result.hasChanges) return noChange({ resolved, threadRootId: result.threadRootId, commentIds: result.commentIds });

        const zip = new MemoryZip(working);
        try {
            if (result.commentsXml) working.set('word/comments.xml', textEncoder.encode(result.commentsXml));
            await ensureCommentsExtendedArtifactsInZip(zip, result.commentsExtendedXml, { replaceExisting: true });
            await repairKnownContentTypes(zip);
            if (options.validate !== false) await validateDocxPackage(zip);
        } catch (error) {
            return packageFailure(sourceBytes, 'PACKAGE_VALIDATION', error.message);
        }

        const artifactsChanged = [...working]
            .filter(([name, data]) => !this.entries.has(name) || !areByteArraysEqual(data, this.entries.get(name)))
            .map(([name]) => name);
        this.entries = working;
        const outputBytes = this.toUint8Array();
        this.originalBytes = outputBytes;
        return {
            status: 'ok', hasChanges: true, written: true, resolved,
            threadRootId: result.threadRootId, commentIds: result.commentIds, artifactsChanged,
            uint8Array: outputBytes, toUint8Array: () => new Uint8Array(outputBytes),
            buffer: toBufferCompatible(outputBytes), toBuffer: () => toBufferCompatible(outputBytes)
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

        // `ids` deletes specific comments (a thread root also takes its replies); otherwise filter by author.
        const requestedIds = Array.isArray(options.ids) ? new Set(options.ids.map(String)) : null;
        if (requestedIds) {
            const known = new Set(Array.from(parsed.doc.getElementsByTagNameNS('*', 'comment')).map(node => node.getAttribute('w:id') || node.getAttribute('id')));
            const missing = [...requestedIds].filter(id => !known.has(id));
            if (missing.length) return packageFailure(sourceBytes, 'COMMENT_NOT_FOUND', `Comment id(s) not found: ${missing.join(', ')}.`);
        }
        const matches = comment => (requestedIds
            ? requestedIds.has(comment.getAttribute('w:id') || comment.getAttribute('id'))
            : options.allAuthors === true || (comment.getAttribute('w:author') || comment.getAttribute('author')) === options.author);
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
                const paraId = commentParaId(comment);
                if (paraId) idByParaId.set(paraId, id);
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

        const remainingXml = requestedIds
            ? commentsXml
            : deleteCommentsByAuthorInOoxml(commentsXml, { author: options.author, allAuthors: options.allAuthors === true }).oxml;
        const remainingParsed = parseOoxmlSafe(remainingXml, 'application/xml');
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
            reconcileCommentSiblings(working);
            const finalZip = new MemoryZip(working);
            await repairKnownContentTypes(finalZip);
            if (options.validate !== false) await validateDocxPackage(finalZip);
        } catch (error) {
            return packageFailure(sourceBytes, 'PACKAGE_VALIDATION', error.message);
        }

        const artifactsChanged = [...working]
            .filter(([name, data]) => !this.entries.has(name) || !areByteArraysEqual(data, this.entries.get(name)))
            .map(([name]) => name);
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
            artifactsChanged,
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
