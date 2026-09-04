/**
 * Public document-operation normalization and validation.
 *
 * Keep compatibility aliases at this boundary so the runner and preflight
 * logic can operate on one internal shape.
 */

const SUPPORTED_OPERATION_TYPES = new Set([
    'redline',
    'replace',
    'format',
    'list-change',
    'table-reconciliation',
    'insert',
    'delete',
    'comment',
    'highlight'
]);

function isRecord(value) {
    return !!value && typeof value === 'object' && !Array.isArray(value);
}

function nonEmptyString(value) {
    return typeof value === 'string' && value.trim().length > 0;
}

export function getCanonicalOperationType(operation) {
    const type = operation?.type;
    if (type === 'comment' || type === 'highlight') return type;
    return 'redline';
}

export function normalizeTargetDescriptor(target, legacyTargetRef = null) {
    if (!isRecord(target)) {
        return {
            text: typeof target === 'string' ? target : '',
            index: legacyTargetRef ?? null,
            paragraphId: null,
            occurrence: null,
            inTable: null,
            fingerprint: null
        };
    }

    return {
        text: typeof target.exactText === 'string'
            ? target.exactText
            : (typeof target.text === 'string' ? target.text : ''),
        index: target.index ?? target.paragraphIndex ?? legacyTargetRef ?? null,
        paragraphId: nonEmptyString(target.paragraphId) ? target.paragraphId.trim() : null,
        occurrence: Number.isInteger(target.occurrence) && target.occurrence > 0
            ? target.occurrence
            : null,
        inTable: typeof target.inTable === 'boolean' ? target.inTable : null,
        fingerprint: nonEmptyString(target.fingerprint)
            ? target.fingerprint.trim()
            : (nonEmptyString(target.sourceFingerprint) ? target.sourceFingerprint.trim() : null)
    };
}

export function normalizeDocumentOperation(operation) {
    const source = isRecord(operation) ? operation : {};
    const targetDescriptor = normalizeTargetDescriptor(source.target, source.targetRef);
    const targetEndDescriptor = isRecord(source.targetEnd)
        ? normalizeTargetDescriptor(source.targetEnd, source.targetEndRef)
        : null;
    const kind = getCanonicalOperationType(source);

    return {
        ...source,
        operationKind: kind,
        targetDescriptor,
        target: targetDescriptor.text,
        targetRef: targetDescriptor.index,
        targetEndRef: targetEndDescriptor?.index ?? source.targetEndRef ?? null,
        ...(source.type === 'delete' && source.modified == null ? { modified: '' } : {}),
        ...(kind === 'comment' && !nonEmptyString(source.textToComment) && nonEmptyString(targetDescriptor.text)
            ? { textToComment: targetDescriptor.text }
            : {})
    };
}

export function validateDocumentOperation(operation) {
    if (!isRecord(operation)) {
        return {
            valid: false,
            error: { code: 'INVALID_OPERATION', message: 'Operation must be an object.' }
        };
    }

    const rawType = operation.type;
    if (rawType != null && rawType !== '' && !SUPPORTED_OPERATION_TYPES.has(rawType)) {
        return {
            valid: false,
            error: {
                code: 'INVALID_OPERATION',
                message: `Unsupported operation type: "${String(rawType)}".`
            }
        };
    }

    const normalized = normalizeDocumentOperation(operation);
    const target = normalized.targetDescriptor;
    if (!nonEmptyString(target.text) && target.index == null && !target.paragraphId) {
        return {
            valid: false,
            error: {
                code: 'INVALID_OPERATION',
                message: 'Operation target must provide text, a paragraph index, or a paragraphId.'
            }
        };
    }

    if (target.occurrence != null && !nonEmptyString(target.text)) {
        return {
            valid: false,
            error: {
                code: 'INVALID_OPERATION',
                message: 'Target occurrence requires target text.'
            }
        };
    }

    if (normalized.operationKind === 'redline' && typeof normalized.modified !== 'string') {
        return {
            valid: false,
            error: { code: 'INVALID_OPERATION', message: 'Redline operations require a string "modified" field.' }
        };
    }

    if (normalized.operationKind === 'comment') {
        if (!nonEmptyString(normalized.commentContent)) {
            return {
                valid: false,
                error: { code: 'INVALID_OPERATION', message: 'Comment operations require a non-empty "commentContent" field.' }
            };
        }
    }

    if (normalized.operationKind === 'highlight') {
        if (!nonEmptyString(normalized.textToHighlight)) {
            return {
                valid: false,
                error: { code: 'INVALID_OPERATION', message: 'Highlight operations require a non-empty "textToHighlight" field.' }
            };
        }
        if (normalized.color != null && !nonEmptyString(normalized.color)) {
            return {
                valid: false,
                error: { code: 'INVALID_OPERATION', message: 'Highlight color must be a non-empty string when provided.' }
            };
        }
    }

    if (normalized.author != null && !nonEmptyString(normalized.author)) {
        return {
            valid: false,
            error: { code: 'INVALID_OPERATION', message: 'Operation author must be a non-empty string when provided.' }
        };
    }

    return { valid: true, operation: normalized };
}

export function resolveDocumentOperationAuthor(operation, batchAuthor, fallbackAuthor) {
    if (nonEmptyString(operation?.author)) return operation.author.trim();
    if (nonEmptyString(batchAuthor)) return batchAuthor.trim();
    return nonEmptyString(fallbackAuthor) ? fallbackAuthor.trim() : 'Author';
}
