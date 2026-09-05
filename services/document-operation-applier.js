import { getDefaultAuthor } from '../adapters/config.js';
import {
    normalizeDocumentOperation,
    resolveDocumentOperationAuthor,
    validateDocumentOperation
} from './document-operation-contract.js';
import { DocumentOperationSession } from './document-operation-session.js';
import {
    applyCommentToParagraphByExactText,
    applyHighlightToParagraphByExactText,
    applyToParagraphByExactText
} from './document-operation-mutations.js';

export function normalizeOperationError(error) {
    return {
        code: typeof error?.code === 'string' && error.code ? error.code : 'OPERATION_ERROR',
        message: error?.message || String(error),
        ...(Array.isArray(error?.candidates) ? { candidates: error.candidates } : {})
    };
}

/**
 * Validates and dispatches one structured operation against full document XML.
 * Result metadata is assembled here so every mutation path exposes the same
 * operation type, author, and target-resolution audit fields.
 */
export async function applyOperationToDocumentXml(documentXml, op, author, runtimeContext = null, options = {}) {
    const validation = validateDocumentOperation(op);
    if (!validation.valid) {
        return {
            documentXml,
            hasChanges: false,
            status: 'error',
            error: validation.error,
            operationType: normalizeDocumentOperation(op).operationKind,
            authorUsed: resolveDocumentOperationAuthor(op, author, getDefaultAuthor())
        };
    }

    const operation = validation.operation || normalizeDocumentOperation(op);
    const authorUsed = resolveDocumentOperationAuthor(operation, author, getDefaultAuthor());
    const session = options?._documentOperationSession instanceof DocumentOperationSession
        ? options._documentOperationSession
        : new DocumentOperationSession(documentXml, options);
    if (!session.valid) {
        return {
            documentXml,
            hasChanges: false,
            status: 'error',
            error: session.parseResult.error,
            warnings: session.parseResult.warnings,
            operationType: operation.operationKind,
            authorUsed
        };
    }

    const resolutionCapture = {};
    const savepoint = session.createSavepoint();
    const operationOptions = {
        ...options,
        ...(typeof operation.generateRedlines === 'boolean' ? { generateRedlines: operation.generateRedlines } : {}),
        ...(operation.existingRevisions ? { existingRevisions: operation.existingRevisions } : {}),
        ...(operation.structuredContent === true ? { structuredContent: true } : {}),
        targetDescriptor: operation.targetDescriptor,
        _resolutionCapture: resolutionCapture,
        _revisionIdAllocator: session.revisionIdAllocator,
        _documentOperationSession: session
    };

    try {
        let result;
        if (operation.operationKind === 'highlight') {
            result = await applyHighlightToParagraphByExactText(
                documentXml,
                operation.target,
                operation.textToHighlight,
                operation.color,
                authorUsed,
                operation.targetRef,
                runtimeContext,
                operationOptions
            );
        } else if (operation.operationKind === 'comment') {
            result = await applyCommentToParagraphByExactText(
                documentXml,
                operation.target,
                operation.textToComment,
                operation.commentContent,
                authorUsed,
                operation.targetRef,
                runtimeContext,
                operationOptions
            );
        } else {
            result = await applyToParagraphByExactText(
                documentXml,
                operation.target,
                operation.modified,
                authorUsed,
                operation.targetRef,
                operation.targetEndRef,
                runtimeContext,
                operationOptions
            );
        }
        if (
            operation.operationKind === 'comment'
            && result?.hasChanges !== true
            && result?.status !== 'error'
            && !result?.error
        ) {
            result = {
                ...result,
                status: 'error',
                error: {
                    code: 'COMMENT_NOT_APPLIED',
                    message: 'The comment operation completed without placing a comment.'
                }
            };
        }
        const isError = result?.status === 'error' || !!result?.error;
        if (isError || result?.hasChanges !== true) {
            session.restoreSavepoint(savepoint);
        } else {
            session.markMutationCommitted();
            result.documentXml = session.deferSerialization
                ? session.currentDocumentXml
                : session.serializeCurrent();
        }
        return {
            ...result,
            operationType: operation.operationKind,
            authorUsed,
            ...resolutionCapture
        };
    } catch (error) {
        session.restoreSavepoint(savepoint);
        const normalizedError = normalizeOperationError(error);
        return {
            documentXml,
            hasChanges: false,
            status: 'error',
            error: normalizedError,
            warnings: [normalizedError.message],
            operationType: operation.operationKind,
            authorUsed,
            ...resolutionCapture
        };
    }
}
