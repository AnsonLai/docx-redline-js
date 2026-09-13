import { randomUUID } from 'node:crypto';
import { openDocx } from '../node/index.js';

/**
 * Demonstration-only defaults for an AI-agent document session.
 *
 * This module is intentionally outside the package's published `files` and
 * `exports`. It is a testable example of how a host can keep its policy separate
 * from the library while delegating all DOCX work to the Node facade.
 */
export const DEFAULT_AGENT_PROFILE = Object.freeze({
    name: 'legal-agent-example',
    author: 'AI Redliner',
    atomic: true,
    continueOnError: true,
    strictTargets: true,
    validate: true,
    generateRedlines: true,
    existingRevisions: 'merge-same-author'
});

const APPLY_OPTION_KEYS = Object.freeze([
    'author',
    'atomic',
    'continueOnError',
    'strictTargets',
    'validate',
    'generateRedlines',
    'existingRevisions'
]);

function copyProfile(profile = {}) {
    return Object.freeze({ ...DEFAULT_AGENT_PROFILE, ...profile });
}

function effectiveApplyOptions(profile, overrides = {}) {
    return Object.fromEntries(APPLY_OPTION_KEYS.map(key => [
        key,
        overrides[key] === undefined ? profile[key] : overrides[key]
    ]));
}

function failedOperationResults(result) {
    return Array.isArray(result?.results)
        ? result.results.filter(item => item?.status === 'error')
        : [];
}

function successfulResult(result) {
    return result?.status !== 'error'
        && result?.status !== 'partial'
        && failedOperationResults(result).length === 0;
}

function boundedInteger(value, fallback, maximum) {
    if (value === undefined || value === null) return fallback;
    const parsed = Number(value);
    return Number.isInteger(parsed) && parsed >= 0 && parsed <= maximum ? parsed : null;
}

function targetDescriptor(paragraph) {
    return {
        exactText: paragraph.exactText,
        ...(paragraph.paragraphId ? { paragraphId: paragraph.paragraphId } : {}),
        ...(paragraph.fingerprint ? { fingerprint: paragraph.fingerprint } : {}),
        inTable: paragraph.inTable === true,
        revisionView: paragraph.revisionView === 'rejected' ? 'rejected' : 'accepted'
    };
}

function invalidRequest(message, field = null) {
    return {
        ok: false,
        complete: false,
        changed: false,
        status: 'error',
        outputBytes: null,
        partialOutputBytes: null,
        error: {
            code: 'INVALID_AGENT_REQUEST',
            message,
            ...(field ? { field } : {}),
            recovery: {
                action: 'change_request',
                sameArgumentsSafe: false,
                requiresUserAuthorization: false
            }
        }
    };
}

/**
 * Creates a stateful, demonstration-only agent session over one DOCX buffer.
 */
export class ExampleAgentDocumentSession {
    constructor(inputBytes, options = {}) {
        this.id = options.sessionId || randomUUID();
        this.profile = copyProfile(options.profile);
        this.document = openDocx(inputBytes);
        this.packageRevision = this.document.getRevisionToken();
        this.nextHandle = 1;
        this.targets = new Map();
        this.handlesByIdentity = new Map();
        this.retiredHandles = new Set();
    }

    getRevisionToken() {
        return { ...this.packageRevision, coveredParts: [...(this.packageRevision.coveredParts || [])] };
    }

    toBuffer() {
        return this.document.toBuffer();
    }

    inspect(options = {}) {
        const around = boundedInteger(options.around, 0, 20);
        if (around == null) {
            return {
                ...invalidRequest('around must be an integer from 0 through 20.', 'around'),
                sessionId: this.id
            };
        }

        const revisionView = options.revisionView === 'rejected' ? 'rejected' : 'accepted';
        const inspection = this.document.inspect({ revisionView });
        if (inspection.status === 'error') {
            return {
                ok: false,
                status: 'error',
                sessionId: this.id,
                packageRevision: this.getRevisionToken(),
                error: inspection.error,
                targets: [],
                warnings: inspection.warnings || []
            };
        }

        const paragraphs = inspection.paragraphs || [];
        const search = options.search == null ? null : String(options.search).toLowerCase();
        const indexes = Array.isArray(options.indexes) ? new Set(options.indexes.map(Number)) : null;
        const rangeStart = options.range ? Number(options.range.start ?? options.range[0]) : null;
        const rangeEnd = options.range ? Number(options.range.end ?? options.range[1]) : null;
        const directMatches = paragraphs.filter(paragraph => {
            if (search != null && !paragraph.exactText.toLowerCase().includes(search)) return false;
            if (indexes && !indexes.has(paragraph.index)) return false;
            if (options.range && !(paragraph.index >= rangeStart && paragraph.index <= rangeEnd)) return false;
            if (options.revisedOnly === true && paragraph.hasRevisions !== true) return false;
            if (typeof options.inTable === 'boolean' && paragraph.inTable !== options.inTable) return false;
            if (options.skipEmpty !== false && paragraph.exactText.trim().length === 0) return false;
            return true;
        });

        const directIndexes = new Set(directMatches.map(paragraph => paragraph.index));
        const selectedIndexes = new Set(directIndexes);
        if (around > 0) {
            for (const paragraph of directMatches) {
                const sourcePosition = paragraphs.indexOf(paragraph);
                const start = Math.max(0, sourcePosition - around);
                const end = Math.min(paragraphs.length - 1, sourcePosition + around);
                for (let position = start; position <= end; position += 1) {
                    const candidate = paragraphs[position];
                    if (options.skipEmpty !== false && candidate.exactText.trim().length === 0) continue;
                    selectedIndexes.add(candidate.index);
                }
            }
        }

        const selected = paragraphs.filter(paragraph => selectedIndexes.has(paragraph.index));
        const targets = selected.map(paragraph => {
            const descriptor = targetDescriptor(paragraph);
            const identity = [
                this.packageRevision.value,
                descriptor.revisionView,
                descriptor.paragraphId || paragraph.index,
                descriptor.fingerprint || ''
            ].join(':');
            let handle = this.handlesByIdentity.get(identity);
            if (!handle) {
                handle = `T${this.nextHandle++}`;
                this.handlesByIdentity.set(identity, handle);
                this.targets.set(handle, {
                    handle,
                    descriptor,
                    packageRevision: this.packageRevision.value,
                    paragraph
                });
            }
            return {
                handle,
                exactText: paragraph.exactText,
                humanReference: paragraph.humanReference,
                nearestHeading: paragraph.nearestHeading,
                inTable: paragraph.inTable,
                list: paragraph.list,
                role: directIndexes.has(paragraph.index) ? 'match' : 'context'
            };
        });

        return {
            ok: true,
            status: 'ok',
            sessionId: this.id,
            packageRevision: this.getRevisionToken(),
            effectiveProfile: { ...this.profile },
            targets,
            counts: {
                matches: directMatches.length,
                context: targets.filter(target => target.role === 'context').length,
                returned: targets.length
            },
            warnings: inspection.warnings || []
        };
    }

    resolveHandle(handle) {
        const target = this.targets.get(handle);
        if (target && target.packageRevision === this.packageRevision.value) return { target };
        if (target || this.retiredHandles.has(handle)) {
            return {
                error: {
                    code: 'STALE_TARGET_HANDLE',
                    message: `Target handle ${String(handle)} belongs to an earlier document revision.`,
                    recovery: {
                        action: 'reinspect',
                        sameArgumentsSafe: false,
                        requiresReinspection: true,
                        requiresUserAuthorization: false
                    }
                }
            };
        }
        return {
            error: {
                code: 'TARGET_HANDLE_NOT_FOUND',
                message: `Target handle ${String(handle)} is not registered in session ${this.id}.`,
                recovery: {
                    action: 'reinspect',
                    sameArgumentsSafe: false,
                    requiresReinspection: true,
                    requiresUserAuthorization: false
                }
            }
        };
    }

    translateEdit(edit, index) {
        if (!edit || typeof edit !== 'object' || Array.isArray(edit)) {
            return { error: invalidRequest(`Edit ${index + 1} must be an object.`, `edits[${index}]`).error };
        }
        const resolution = this.resolveHandle(edit.target);
        if (resolution.error) return { error: resolution.error };
        const target = { ...resolution.target.descriptor };
        const common = {
            operationId: typeof edit.operationId === 'string' ? edit.operationId : `agent-edit-${index + 1}`,
            target,
            ...(typeof edit.author === 'string' ? { author: edit.author } : {}),
            ...(typeof edit.existingRevisions === 'string' ? { existingRevisions: edit.existingRevisions } : {})
        };

        if (typeof edit.desiredText === 'string') {
            return { operation: { ...common, type: 'redline', modified: edit.desiredText } };
        }
        if (edit.action === 'delete') {
            return { operation: { ...common, type: 'delete' } };
        }
        if (typeof edit.commentContent === 'string' && edit.commentContent.length > 0) {
            return {
                operation: {
                    ...common,
                    type: 'comment',
                    commentContent: edit.commentContent,
                    ...(typeof edit.textToComment === 'string' ? { textToComment: edit.textToComment } : {})
                }
            };
        }
        return {
            error: invalidRequest(
                `Edit ${index + 1} must provide desiredText, action: "delete", or commentContent.`,
                `edits[${index}]`
            ).error
        };
    }

    async applyEdits(edits, options = {}) {
        if (!Array.isArray(edits) || edits.length === 0) {
            return { ...invalidRequest('edits must be a non-empty array.', 'edits'), sessionId: this.id };
        }
        const translated = edits.map((edit, index) => this.translateEdit(edit, index));
        const invalidIndex = translated.findIndex(item => item.error);
        if (invalidIndex >= 0) {
            return {
                ok: false,
                complete: false,
                changed: false,
                status: 'error',
                sessionId: this.id,
                outputBytes: null,
                partialOutputBytes: null,
                error: translated[invalidIndex].error,
                retryPlan: {
                    base: 'original',
                    committedIndexes: [],
                    failedIndexes: [invalidIndex + 1],
                    unattemptedIndexes: translated.map((_, index) => index + 1).filter(index => index !== invalidIndex + 1),
                    replayWholeBatch: true,
                    sameArgumentsSafe: false
                }
            };
        }

        const effectiveOptions = effectiveApplyOptions(this.profile, options);
        const result = await this.document.applyOperations(
            translated.map(item => item.operation),
            { ...effectiveOptions, expectedRevision: this.getRevisionToken() }
        );
        const failed = failedOperationResults(result);
        const ok = successfulResult(result);
        const packageChanged = result.written === true;
        const output = packageChanged ? result.toBuffer() : this.document.toBuffer();

        if (packageChanged) {
            for (const handle of this.targets.keys()) this.retiredHandles.add(handle);
            this.targets.clear();
            this.handlesByIdentity.clear();
            this.packageRevision = this.document.getRevisionToken();
        }

        const executedIndexes = new Set((result.results || []).map(item => item.index));
        const committedIndexes = (result.receipts || [])
            .filter(receipt => receipt?.committed === true)
            .map(receipt => receipt.operationIndex);
        const failedIndexes = failed.map(item => item.index);
        const unattemptedIndexes = translated
            .map((_, index) => index + 1)
            .filter(index => !executedIndexes.has(index));
        const retryBase = result.rolledBack === true || effectiveOptions.atomic === true ? 'original' : 'output';

        return {
            ok,
            complete: ok,
            changed: result.hasChanges === true,
            status: result.status || 'ok',
            sessionId: this.id,
            effectiveProfile: { name: this.profile.name, ...effectiveOptions },
            packageRevision: this.getRevisionToken(),
            outputBytes: ok ? Buffer.from(output) : null,
            partialOutputBytes: !ok && packageChanged ? Buffer.from(output) : null,
            error: result.error || null,
            results: result.results || [],
            receipts: result.receipts || [],
            warnings: result.warnings || [],
            validation: result.validation || null,
            retryPlan: ok ? null : {
                base: retryBase,
                committedIndexes,
                failedIndexes,
                unattemptedIndexes,
                replayWholeBatch: retryBase === 'original',
                sameArgumentsSafe: false
            }
        };
    }

    async resolveReview(action, filter = {}) {
        if (!['accept', 'reject', 'delete-comments'].includes(action)) {
            return { ...invalidRequest('action must be accept, reject, or delete-comments.', 'action'), sessionId: this.id };
        }
        if (!filter.allAuthors && typeof filter.author !== 'string') {
            return { ...invalidRequest('resolveReview requires author or allAuthors: true.', 'filter'), sessionId: this.id };
        }
        const options = {
            ...(filter.allAuthors ? { allAuthors: true } : { author: filter.author }),
            validate: filter.validate !== false
        };
        const result = action === 'delete-comments'
            ? await this.document.deleteComments(options)
            : await this.document.resolveRevisions(action, options);
        const ok = successfulResult(result);
        const packageChanged = result.written === true;
        if (packageChanged) {
            for (const handle of this.targets.keys()) this.retiredHandles.add(handle);
            this.targets.clear();
            this.handlesByIdentity.clear();
            this.packageRevision = this.document.getRevisionToken();
        }
        return {
            ok,
            complete: ok,
            changed: result.hasChanges === true,
            status: result.status || 'ok',
            sessionId: this.id,
            packageRevision: this.getRevisionToken(),
            outputBytes: ok ? Buffer.from(this.document.toBuffer()) : null,
            error: result.error || null,
            warnings: result.warnings || [],
            validation: result.validation || null
        };
    }
}

export function createExampleAgentSession(inputBytes, options = {}) {
    return new ExampleAgentDocumentSession(inputBytes, options);
}
