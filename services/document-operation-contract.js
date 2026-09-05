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
            fingerprint: null,
            revisionView: 'accepted'
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
            : (nonEmptyString(target.sourceFingerprint) ? target.sourceFingerprint.trim() : null),
        revisionView: target.revisionView === 'rejected' ? 'rejected' : 'accepted'
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
        ...(source.type === 'delete' && source.modified == null ? { modified: '' } : {})
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
    if (isRecord(operation.target) && operation.target.revisionView != null) {
        if (operation.target.revisionView !== 'accepted' && operation.target.revisionView !== 'rejected') {
            return {
                valid: false,
                error: {
                    code: 'INVALID_OPERATION',
                    message: 'Target revisionView must be "accepted" or "rejected" when provided.'
                }
            };
        }
    }

    if (isRecord(operation.targetEnd) && operation.targetEnd.revisionView != null) {
        if (operation.targetEnd.revisionView !== 'accepted' && operation.targetEnd.revisionView !== 'rejected') {
            return {
                valid: false,
                error: {
                    code: 'INVALID_OPERATION',
                    message: 'Target revisionView must be "accepted" or "rejected" when provided.'
                }
            };
        }
    }

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

    if (normalized.structuredContent != null && typeof normalized.structuredContent !== 'boolean') {
        return {
            valid: false,
            error: { code: 'INVALID_OPERATION', message: 'structuredContent must be a boolean when provided.' }
        };
    }

    if (normalized.pairReplacements != null && typeof normalized.pairReplacements !== 'boolean') {
        return {
            valid: false,
            error: { code: 'INVALID_OPERATION', message: 'pairReplacements must be a boolean when provided.' }
        };
    }

    if (normalized.insertionAffinity != null) {
        if (!isRecord(normalized.insertionAffinity)) {
            return {
                valid: false,
                error: { code: 'INVALID_OPERATION', message: 'insertionAffinity must be an object when provided.' }
            };
        }
        const { formatting, hyperlink, revision, bookmark, comment } = normalized.insertionAffinity;
        if (formatting != null && !['left', 'right', 'none'].includes(formatting)) {
            return {
                valid: false,
                error: { code: 'INVALID_OPERATION', message: 'insertionAffinity.formatting must be "left", "right", or "none".' }
            };
        }
        if (hyperlink != null && !['inside', 'outside', 'preserve'].includes(hyperlink)) {
            return {
                valid: false,
                error: { code: 'INVALID_OPERATION', message: 'insertionAffinity.hyperlink must be "inside", "outside", or "preserve".' }
            };
        }
        if (revision != null && !['coalesce_same_author', 'separate'].includes(revision)) {
            return {
                valid: false,
                error: { code: 'INVALID_OPERATION', message: 'insertionAffinity.revision must be "coalesce_same_author" or "separate".' }
            };
        }
        if (bookmark != null && !['inside', 'outside'].includes(bookmark)) {
            return {
                valid: false,
                error: { code: 'INVALID_OPERATION', message: 'insertionAffinity.bookmark must be "inside" or "outside".' }
            };
        }
        if (comment != null && !['inside', 'outside'].includes(comment)) {
            return {
                valid: false,
                error: { code: 'INVALID_OPERATION', message: 'insertionAffinity.comment must be "inside" or "outside".' }
            };
        }
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
