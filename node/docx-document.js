import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { configureXmlProvider } from '../adapters/xml-adapter.js';
import { inspectDocumentParts } from '../services/document-inspection.js';
import { applyOperationsToDocumentXml, preflightOperations } from '../services/standalone-operation-runner.js';
import { createDynamicNumberingIdState, mergeNumberingXmlBySchemaOrder } from '../services/numbering-helpers.js';
import { ensureCommentsArtifactsInZip, ensureNumberingArtifactsInZip, validateDocxPackage } from '../services/standalone-docx-plumbing.js';
import { validateRedlineOoxml } from '../core/redline-validation.js';
import { acceptTrackedChangesInOoxml, rejectTrackedChangesInOoxml, deleteCommentsByAuthorInOoxml } from '../services/revision-comment-management.js';
import { createSerializer, parseOoxmlSafe } from '../adapters/xml-adapter.js';
import { createHash } from 'node:crypto';
import { MemoryZip, unzipDocx, zipDocx } from './zip-archive.js';
import { computeRevisionTokenSync } from '../services/revision-token.js';

configureXmlProvider({ DOMParser, XMLSerializer });
const text = (entries, path) => entries.get(path)?.toString('utf8') || null;
const cloneEntries = entries => new Map([...entries].map(([name, data]) => [name, Buffer.from(data)]));

/**
 * Computes a package-scoped revision token over all uncompressed entries in a DOCX archive.
 *
 * @param {Buffer|Uint8Array|Map<string, Buffer>|DocxDocument|object} input
 * @returns {{ algorithm: 'sha256', version: number, scope: 'package', value: string, coveredParts: string[] }}
 */
export function computePackageRevisionToken(input) {
    let entries;
    if (Buffer.isBuffer(input) || input instanceof Uint8Array) {
        entries = unzipDocx(input);
    } else if (input instanceof Map) {
        entries = input;
    } else if (input?.entries instanceof Map) {
        entries = input.entries;
    } else if (typeof input?.toBuffer === 'function') {
        entries = unzipDocx(input.toBuffer());
    } else {
        throw new TypeError('computePackageRevisionToken requires a Buffer, Uint8Array, Map of entries, or DocxDocument.');
    }
    return computeRevisionTokenSync({
        scope: 'package',
        entries,
        digestFn: bytes => createHash('sha256').update(bytes).digest('hex')
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
            text: String(comment.textContent || '').trim()
        };
    }
    return details;
}

export class DocxDocument {
    constructor(buffer) { this.originalBuffer = Buffer.from(buffer); this.entries = unzipDocx(this.originalBuffer); }
    inspect(options = {}) {
        const digestFn = options.digestFn || (bytes => createHash('sha256').update(bytes).digest('hex'));
        return inspectDocumentParts({
            documentXml: text(this.entries, 'word/document.xml'),
            commentsXml: text(this.entries, 'word/comments.xml'),
            numberingXml: text(this.entries, 'word/numbering.xml'),
            stylesXml: text(this.entries, 'word/styles.xml')
        }, { ...options, digestFn });
    }
    getRevisionToken() { return computePackageRevisionToken(this.entries); }
    get revisionToken() { return this.getRevisionToken(); }
    preflight(operations, author, options = {}) { return preflightOperations(text(this.entries, 'word/document.xml'), operations, author, { ...options, _existingCommentDetails: existingCommentDetails(this.entries) }); }
    toBuffer() { return zipDocx(this.entries); }
    async applyOperations(operations, options = {}) {
        const originalEntries = this.entries; const working = cloneEntries(originalEntries); const zip = new MemoryZip(working);
        const documentXml = text(working, 'word/document.xml');
        let originalIssues = [];
        try {
            if (!documentXml) throw new Error('Missing word/document.xml.');
            const baseline = validateRedlineOoxml(documentXml);
            originalIssues = baseline.issues.map(issue => ({ source: 'word/document.xml', ...issue }));
            try { await validateDocxPackage(new MemoryZip(cloneEntries(originalEntries))); }
            catch (error) { originalIssues.push({ source: 'package', code: 'PACKAGE_VALIDATION', severity: 'error', message: error.message }); }
            const context = { numberingIdState: createDynamicNumberingIdState(text(working, 'word/numbering.xml') || undefined) };
            const result = await applyOperationsToDocumentXml(documentXml, operations, options.author, context, {
                ...options, atomic: options.atomic !== false, strictTargets: options.strictTargets !== false,
                _existingCommentDetails: existingCommentDetails(working),
                commentIdAllocator: nextCommentId(working)
            });
            if (result.rolledBack || result.status === 'error') return { ...result, written: false, artifactsChanged: [], validation: { originalIssues, generatedIssues: [] }, buffer: this.originalBuffer, toBuffer: () => Buffer.from(this.originalBuffer) };
            if (!result.hasChanges) return { ...result, status: result.status || 'ok', written: false, artifactsChanged: [], validation: { originalIssues, generatedIssues: [] }, buffer: Buffer.from(this.originalBuffer), toBuffer: () => Buffer.from(this.originalBuffer) };
            working.set('word/document.xml', Buffer.from(result.documentXml));
            await ensureNumberingArtifactsInZip(zip, result.numberingXmlParts, { mergeNumberingXml: mergeNumberingXmlBySchemaOrder });
            await ensureCommentsArtifactsInZip(zip, result.commentsXml);
            if (options.validate !== false) {
                const generated = validateRedlineOoxml(result.documentXml);
                const baselineErrors = new Set(baseline.issues.filter(i => i.severity === 'error').map(i => `${i.code}:${i.message}`));
                const introduced = generated.issues.filter(i => i.severity === 'error' && !baselineErrors.has(`${i.code}:${i.message}`));
                if (introduced.length) throw Object.assign(new Error('Generated document introduced invalid revision markup.'), { issues: introduced });
                await validateDocxPackage(zip);
            }
            this.entries = working;
            const output = this.toBuffer();
            this.originalBuffer = Buffer.from(output);
            const artifactsChanged = [...working].filter(([name, data]) => !originalEntries.has(name) || !data.equals(originalEntries.get(name))).map(([name]) => name);
            return { ...result, written: true, artifactsChanged, validation: { originalIssues, generatedIssues: [] }, buffer: output, inspection: this.inspect(), toBuffer: () => Buffer.from(output) };
        } catch (error) {
            this.entries = originalEntries;
            const generatedIssues = error.issues || [{ source: 'package', code: 'PACKAGE_OPERATION_FAILED', severity: 'error', message: error.message }];
            return { status: 'error', hasChanges: false, written: false, rolledBack: true, results: [], artifactsChanged: [], error: { code: 'PACKAGE_OPERATION_FAILED', message: error.message }, validation: { originalIssues, generatedIssues }, issues: generatedIssues, buffer: Buffer.from(this.originalBuffer), toBuffer: () => Buffer.from(this.originalBuffer) };
        }
    }

    async resolveRevisions(action, options = {}) {
        const transform = action === 'accept' ? acceptTrackedChangesInOoxml : action === 'reject' ? rejectTrackedChangesInOoxml : null;
        if (!transform) return packageFailure(this.originalBuffer, 'INVALID_ACTION', `Unknown revision action: ${action}`);
        const source = Buffer.from(this.originalBuffer); const working = cloneEntries(this.entries); const zip = new MemoryZip(working);
        const result = transform(text(working, 'word/document.xml'), { author: options.author, allAuthors: options.allAuthors === true });
        if (result.status === 'error' || result.error) return packageFailure(source, result.error?.code || 'REVISION_OPERATION_FAILED', result.error?.message || 'Revision operation failed.');
        if (!result.hasChanges) return { ...result, status: 'ok', written: false, artifactsChanged: [], buffer: source, toBuffer: () => Buffer.from(source) };
        working.set('word/document.xml', Buffer.from(result.oxml));
        try { if (options.validate !== false) await validateDocxPackage(zip); }
        catch (error) { return packageFailure(source, 'PACKAGE_VALIDATION', error.message); }
        this.entries = working; const output = this.toBuffer(); this.originalBuffer = Buffer.from(output);
        return { ...result, status: 'ok', written: true, artifactsChanged: ['word/document.xml'], buffer: output, toBuffer: () => Buffer.from(output) };
    }

    async deleteComments(options = {}) {
        const source = Buffer.from(this.originalBuffer); const working = cloneEntries(this.entries); const commentsXml = text(working, 'word/comments.xml');
        if (!commentsXml) return { status: 'ok', hasChanges: false, written: false, commentsRemoved: 0, referencesRemoved: 0, artifactsChanged: [], buffer: source, toBuffer: () => Buffer.from(source) };
        const parsed = parseOoxmlSafe(commentsXml, 'application/xml');
        if (!parsed.doc || parsed.error) return packageFailure(source, 'PARSE_ERROR', parsed.error?.message || 'Could not parse comments.xml.');
        const matches = comment => options.allAuthors === true || (comment.getAttribute('w:author') || comment.getAttribute('author')) === options.author;
        const ids = new Set(Array.from(parsed.doc.getElementsByTagNameNS('*', 'comment')).filter(matches).map(node => node.getAttribute('w:id') || node.getAttribute('id')).filter(Boolean));
        if (!ids.size) return { status: 'ok', hasChanges: false, written: false, commentsRemoved: 0, referencesRemoved: 0, artifactsChanged: [], buffer: source, toBuffer: () => Buffer.from(source) };
        const commentsResult = deleteCommentsByAuthorInOoxml(commentsXml, { author: options.author, allAuthors: options.allAuthors === true });
        const documentParsed = parseOoxmlSafe(text(working, 'word/document.xml'), 'application/xml');
        if (!documentParsed.doc || documentParsed.error) return packageFailure(source, 'PARSE_ERROR', documentParsed.error?.message || 'Could not parse document.xml.');
        let referencesRemoved = 0;
        for (const name of ['commentRangeStart', 'commentRangeEnd', 'commentReference']) for (const node of Array.from(documentParsed.doc.getElementsByTagNameNS('*', name))) {
            const id = node.getAttribute('w:id') || node.getAttribute('id'); if (!ids.has(id) || !node.parentNode) continue;
            const parent = node.parentNode; parent.removeChild(node); referencesRemoved += 1;
            if (name === 'commentReference' && parent.localName === 'r' && !Array.from(parent.childNodes || []).some(child => child.nodeType === 1 && child.localName !== 'rPr')) parent.parentNode?.removeChild(parent);
        }
        working.set('word/comments.xml', Buffer.from(commentsResult.oxml));
        working.set('word/document.xml', Buffer.from(createSerializer().serializeToString(documentParsed.doc)));
        try { if (options.validate !== false) await validateDocxPackage(new MemoryZip(working)); }
        catch (error) { return packageFailure(source, 'PACKAGE_VALIDATION', error.message); }
        this.entries = working; const output = this.toBuffer(); this.originalBuffer = Buffer.from(output);
        return { status: 'ok', hasChanges: true, written: true, commentsRemoved: commentsResult.commentsRemoved, referencesRemoved, artifactsChanged: ['word/document.xml', 'word/comments.xml'], buffer: output, toBuffer: () => Buffer.from(output) };
    }
}

function packageFailure(source, code, message) { return { status: 'error', hasChanges: false, written: false, rolledBack: true, error: { code, message }, artifactsChanged: [], buffer: Buffer.from(source), toBuffer: () => Buffer.from(source) }; }
export function openDocx(input) { return new DocxDocument(input); }
