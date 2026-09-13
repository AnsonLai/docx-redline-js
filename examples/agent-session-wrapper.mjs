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

function patchError(code, message, field, recoveryAction, details = {}) {
    return {
        code,
        message,
        field,
        ...details,
        recovery: {
            action: recoveryAction,
            sameArgumentsSafe: false,
            requiresReinspection: false,
            requiresUserAuthorization: false
        }
    };
}

function occurrenceOffsets(sourceText, find) {
    const offsets = [];
    let from = 0;
    while (from <= sourceText.length - find.length) {
        const start = sourceText.indexOf(find, from);
        if (start < 0) break;
        offsets.push(start);
        from = start + 1;
    }
    return offsets;
}

function offsetCandidate(sourceText, start, length) {
    const excerptStart = Math.max(0, start - 40);
    const excerptEnd = Math.min(sourceText.length, start + length + 40);
    return {
        start,
        end: start + length,
        excerpt: sourceText.slice(excerptStart, excerptEnd)
    };
}

/**
 * Compiles exact, source-relative replacement intents into complete desired text.
 * All ranges are resolved before mutation so replacements are simultaneous.
 */
export function compileExactReplacements(sourceText, replacements, field = 'replacements') {
    if (typeof sourceText !== 'string') {
        return {
            ok: false,
            error: invalidRequest('Localized replacements require string source text.', field).error
        };
    }
    if (!Array.isArray(replacements) || replacements.length === 0) {
        return {
            ok: false,
            error: invalidRequest('replacements must be a non-empty array.', field).error
        };
    }

    const resolved = [];
    for (let index = 0; index < replacements.length; index += 1) {
        const replacement = replacements[index];
        const itemField = `${field}[${index}]`;
        if (!replacement || typeof replacement !== 'object' || Array.isArray(replacement)) {
            return { ok: false, error: invalidRequest(`${itemField} must be an object.`, itemField).error };
        }
        if (typeof replacement.find !== 'string' || replacement.find.length === 0) {
            return { ok: false, error: invalidRequest(`${itemField}.find must be a non-empty string.`, `${itemField}.find`).error };
        }
        if (typeof replacement.replace !== 'string') {
            return { ok: false, error: invalidRequest(`${itemField}.replace must be a string.`, `${itemField}.replace`).error };
        }
        const occurrence = replacement.occurrence == null ? null : Number(replacement.occurrence);
        if (occurrence != null && (!Number.isInteger(occurrence) || occurrence < 1)) {
            return { ok: false, error: invalidRequest(`${itemField}.occurrence must be a positive integer.`, `${itemField}.occurrence`).error };
        }

        const offsets = occurrenceOffsets(sourceText, replacement.find);
        if (offsets.length === 0 || (occurrence != null && occurrence > offsets.length)) {
            return {
                ok: false,
                error: patchError(
                    'PATCH_SOURCE_NOT_FOUND',
                    occurrence == null
                        ? `Exact patch source was not found: "${replacement.find}".`
                        : `Occurrence ${occurrence} of exact patch source was not found: "${replacement.find}".`,
                    itemField,
                    'change_patch',
                    { matchCount: offsets.length }
                )
            };
        }
        if (occurrence == null && offsets.length > 1) {
            return {
                ok: false,
                error: patchError(
                    'AMBIGUOUS_PATCH_SOURCE',
                    `Exact patch source matched ${offsets.length} locations; provide occurrence.`,
                    itemField,
                    'choose_occurrence',
                    { candidates: offsets.map(start => offsetCandidate(sourceText, start, replacement.find.length)) }
                )
            };
        }

        const start = offsets[(occurrence || 1) - 1];
        resolved.push({
            requestIndex: index,
            start,
            end: start + replacement.find.length,
            find: replacement.find,
            replace: replacement.replace,
            occurrence: occurrence || 1
        });
    }

    resolved.sort((left, right) => left.start - right.start || left.end - right.end || left.requestIndex - right.requestIndex);
    const unique = [];
    for (const replacement of resolved) {
        const previous = unique.at(-1);
        if (previous && replacement.start === previous.start && replacement.end === previous.end) {
            if (replacement.replace !== previous.replace) {
                return {
                    ok: false,
                    error: patchError(
                        'CONFLICTING_PATCHES',
                        `Localized replacements ${previous.requestIndex + 1} and ${replacement.requestIndex + 1} assign different text to the same source range.`,
                        field,
                        'combine_patches',
                        { replacementIndexes: [previous.requestIndex + 1, replacement.requestIndex + 1] }
                    )
                };
            }
            continue;
        }
        if (previous && replacement.start < previous.end) {
            return {
                ok: false,
                error: patchError(
                    'OVERLAPPING_PATCHES',
                    `Localized replacements ${previous.requestIndex + 1} and ${replacement.requestIndex + 1} overlap.`,
                    field,
                    'combine_patches',
                    { replacementIndexes: [previous.requestIndex + 1, replacement.requestIndex + 1] }
                )
            };
        }
        unique.push(replacement);
    }

    let cursor = 0;
    let desiredText = '';
    for (const replacement of unique) {
        desiredText += sourceText.slice(cursor, replacement.start);
        desiredText += replacement.replace;
        cursor = replacement.end;
    }
    desiredText += sourceText.slice(cursor);

    return {
        ok: true,
        desiredText,
        replacements: unique.map(replacement => ({
            requestIndex: replacement.requestIndex + 1,
            start: replacement.start,
            end: replacement.end,
            occurrence: replacement.occurrence,
            removedLength: replacement.end - replacement.start,
            insertedLength: replacement.replace.length
        }))
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

    registerParagraph(paragraph, role = 'match') {
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
            role
        };
    }

    retireCurrentHandles() {
        for (const handle of this.targets.keys()) this.retiredHandles.add(handle);
        this.targets.clear();
        this.handlesByIdentity.clear();
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
        const targets = selected.map(paragraph => this.registerParagraph(
            paragraph,
            directIndexes.has(paragraph.index) ? 'match' : 'context'
        ));

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

        const intentCount = [
            typeof edit.desiredText === 'string',
            edit.replacements !== undefined,
            edit.action === 'delete',
            typeof edit.commentContent === 'string' && edit.commentContent.length > 0
        ].filter(Boolean).length;
        if (intentCount > 1) {
            return {
                error: invalidRequest(
                    `Edit ${index + 1} must provide exactly one edit intent.`,
                    `edits[${index}]`
                ).error
            };
        }

        if (typeof edit.desiredText === 'string') {
            return {
                operation: { ...common, type: 'redline', modified: edit.desiredText },
                sourceHandle: edit.target
            };
        }
        if (edit.replacements !== undefined) {
            const compiled = compileExactReplacements(
                resolution.target.descriptor.exactText,
                edit.replacements,
                `edits[${index}].replacements`
            );
            if (!compiled.ok) return { error: compiled.error };
            return {
                operation: { ...common, type: 'redline', modified: compiled.desiredText },
                sourceHandle: edit.target,
                compilation: {
                    operationId: common.operationId,
                    sourceLength: resolution.target.descriptor.exactText.length,
                    desiredLength: compiled.desiredText.length,
                    replacements: compiled.replacements
                }
            };
        }
        if (edit.action === 'delete') {
            return { operation: { ...common, type: 'delete' }, sourceHandle: edit.target };
        }
        if (typeof edit.commentContent === 'string' && edit.commentContent.length > 0) {
            return {
                operation: {
                    ...common,
                    type: 'comment',
                    commentContent: edit.commentContent,
                    ...(typeof edit.textToComment === 'string' ? { textToComment: edit.textToComment } : {})
                },
                sourceHandle: edit.target
            };
        }
        return {
            error: invalidRequest(
                `Edit ${index + 1} must provide desiredText, replacements, action: "delete", or commentContent.`,
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
        let refreshedTargets = [];

        if (packageChanged) {
            this.retireCurrentHandles();
            this.packageRevision = this.document.getRevisionToken();
            const currentParagraphs = this.document.inspect().paragraphs || [];
            refreshedTargets = translated.flatMap((item, index) => {
                const operationResult = (result.results || []).find(entry => entry.index === index + 1);
                if (!operationResult || operationResult.status === 'error' || item.operation.type === 'delete') return [];
                const paragraphId = item.operation.target.paragraphId;
                let paragraph = paragraphId
                    ? currentParagraphs.find(candidate => candidate.paragraphId === paragraphId)
                    : null;
                if (!paragraph && typeof item.operation.modified === 'string') {
                    const matches = currentParagraphs.filter(candidate => candidate.exactText === item.operation.modified);
                    paragraph = matches.length === 1 ? matches[0] : null;
                }
                if (!paragraph || paragraph.exactText.trim().length === 0) return [];
                return [{
                    operationId: item.operation.operationId,
                    previousHandle: item.sourceHandle,
                    target: this.registerParagraph(paragraph)
                }];
            });
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
            compiledEdits: translated.flatMap(item => item.compilation ? [item.compilation] : []),
            refreshedTargets,
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
            this.retireCurrentHandles();
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
