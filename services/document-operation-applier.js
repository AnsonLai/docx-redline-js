import { getDefaultAuthor } from '../adapters/config.js';
import {
    normalizeDocumentOperation,
    resolveDocumentOperationAuthor,
    validateDocumentOperation
} from './document-operation-contract.js';
import { DocumentOperationSession } from './document-operation-session.js';
import {
    validateRevisionToken,
    computeDocumentPartsRevisionToken,
    areRevisionTokensEqual
} from './revision-token.js';
import {
    applyCommentToParagraphByExactText,
    applyFormattingToParagraphByExactText,
    applyHighlightToParagraphByExactText,
    applyParagraphFormatToParagraphByExactText,
    applyToParagraphByExactText
} from './document-operation-mutations.js';
import { applyCommentReplyToParts } from './comment-replies.js';
import {
    deriveCapturedEntity,
    invalidateAffectedCaptures
} from './capture-engine.js';
import {
    createEmptyReceipt,
    reconcileReceiptsAgainstOutput
} from './receipt-collector.js';

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
 * */
export async function applyOperationToDocumentXml(documentXml, op, author, runtimeContext = null, options = {}) {
    const operationIndex = typeof options._operationIndex === 'number' ? options._operationIndex : 1;
    const validation = validateDocumentOperation(op);
    if (!validation.valid) {
        const authorUsed = resolveDocumentOperationAuthor(op, author, getDefaultAuthor());
        return {
            documentXml,
            hasChanges: false,
            status: 'error',
            error: validation.error,
            operationType: normalizeDocumentOperation(op).operationKind,
            authorUsed,
            receipt: createEmptyReceipt(operationIndex, op?.operationId, authorUsed, 'refused')
        };
    }

    const operation = validation.operation || normalizeDocumentOperation(op);
    const authorUsed = resolveDocumentOperationAuthor(operation, author, getDefaultAuthor());

    if (operation.operationKind !== 'comment_reply' && operation.targetDescriptor?.revisionView === 'rejected') {
        return {
            documentXml,
            hasChanges: false,
            status: 'error',
            error: {
                code: 'UNSUPPORTED_REVISION_VIEW_MUTATION',
                message: 'Targeting rejected revision view for mutation is not supported yet.'
            },
            operationType: operation.operationKind,
            authorUsed,
            receipt: createEmptyReceipt(operationIndex, operation.operationId, authorUsed, 'refused')
        };
    }

    if (!options?._documentOperationSession && options?.expectedRevision) {
        const tokenValidation = validateRevisionToken(options.expectedRevision);
        if (!tokenValidation.valid) {
            return {
                documentXml,
                hasChanges: false,
                status: 'error',
                error: {
                    code: tokenValidation.error?.code || 'INVALID_REVISION_TOKEN',
                    message: tokenValidation.error?.message || 'Invalid revision token.'
                },
                operationType: operation.operationKind,
                authorUsed,
                receipt: createEmptyReceipt(operationIndex, operation.operationId, authorUsed, 'refused')
            };
        }
        if (options.expectedRevision.scope !== 'document-parts') {
            return {
                documentXml,
                hasChanges: false,
                status: 'error',
                error: {
                    code: 'REVISION_TOKEN_SCOPE_MISMATCH',
                    message: `Revision token scope mismatch: expected 'document-parts', got '${options.expectedRevision.scope}'.`
                },
                operationType: operation.operationKind,
                authorUsed,
                receipt: createEmptyReceipt(operationIndex, operation.operationId, authorUsed, 'refused')
            };
        }
        const currentToken = await computeDocumentPartsRevisionToken({
            documentXml,
            commentsXml: runtimeContext?.commentsXml || options.commentsXml,
            commentsExtendedXml: runtimeContext?.commentsExtendedXml || options.commentsExtendedXml,
            numberingXml: runtimeContext?.numberingXml || options.numberingXml,
            stylesXml: runtimeContext?.stylesXml || options.stylesXml
        }, options);
        if (!areRevisionTokensEqual(currentToken.value, options.expectedRevision.value)) {
            return {
                documentXml,
                hasChanges: false,
                status: 'error',
                error: {
                    code: 'REVISION_MISMATCH',
                    message: `Document revision mismatch: expected '${options.expectedRevision.value}', current is '${currentToken.value}'.`
                },
                operationType: operation.operationKind,
                authorUsed,
                receipt: createEmptyReceipt(operationIndex, operation.operationId, authorUsed, 'refused')
            };
        }
    }

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
            authorUsed,
            receipt: createEmptyReceipt(operationIndex, operation.operationId, authorUsed, 'refused')
        };
    }

    const resolutionCapture = {};
    const savepoint = session.createSavepoint();
    session.receiptCollector?.beginOperation(
        operationIndex,
        operation.operationId,
        authorUsed
    );
    const operationWarnings = [];
    const operationOptions = {
        ...options,
        ...(typeof operation.generateRedlines === 'boolean' ? { generateRedlines: operation.generateRedlines } : {}),
        ...(operation.existingRevisions ? { existingRevisions: operation.existingRevisions } : {}),
        structuredContent: typeof operation.structuredContent === 'boolean' ? operation.structuredContent : (options.structuredContent !== false),
        explicitStructuredContent: operation.structuredContent === true,
        pairReplacements: typeof operation.pairReplacements === 'boolean' ? operation.pairReplacements : (options.pairReplacements !== false),
        ...(operation.insertionAffinity ? { insertionAffinity: operation.insertionAffinity } : {}),
        ...(operation.formattingRevisionPolicy ? { formattingRevisionPolicy: operation.formattingRevisionPolicy } : {}),
        targetDescriptor: operation.targetDescriptor,
        _resolutionCapture: resolutionCapture,
        _revisionIdAllocator: session.revisionIdAllocator,
        _documentOperationSession: session,
        _mutationLiveNodes: [],
        _mutationRemovedNodes: [],
        onInfo: (msg) => {
            if (typeof options?.onInfo === 'function') options.onInfo(msg);
        },
        onWarn: (msg) => {
            operationWarnings.push(String(msg));
            if (typeof options?.onWarn === 'function') options.onWarn(msg);
        }
    };

    try {
        let result;
        if (operation.operationKind === 'comment_reply') {
            const existingCommentsXml = session.commentsXml || runtimeContext?.commentsXml || options.commentsXml;
            const existingExtendedXml = session.commentsExtendedXml || runtimeContext?.commentsExtendedXml || options.commentsExtendedXml;
            let commentId = typeof options.commentIdAllocator === 'function' ? options.commentIdAllocator() : null;
            if (commentId == null && existingCommentsXml) {
                const ids = [...existingCommentsXml.matchAll(/<(?:w:)?comment\b[^>]*\b(?:w:)?id=["'](\d+)["']/g)].map(match => Number(match[1]));
                commentId = (ids.length ? Math.max(...ids) : -1) + 1;
            }
            if (commentId == null) {
                result = { documentXml, hasChanges: false, status: 'error', error: { code: 'COMMENTS_PART_MISSING', message: 'A comment reply requires an existing comments part.' } };
            } else {
                result = applyCommentReplyToParts({
                    commentsXml: existingCommentsXml,
                    commentsExtendedXml: existingExtendedXml,
                    parentCommentId: operation.parentCommentId,
                    commentId,
                    commentContent: operation.commentContent,
                    author: authorUsed,
                    date: operation.date || new Date().toISOString()
                });
                result.documentXml = documentXml;
                if (result.hasChanges && session.receiptCollector) session.receiptCollector.recordComment(commentId);
            }
        } else if (operation.operationKind === 'highlight') {
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
        } else if (operation.operationKind === 'format') {
            result = await applyFormattingToParagraphByExactText(
                documentXml,
                operation.target,
                operation.textToFormat,
                operation.properties,
                authorUsed,
                operation.targetRef,
                runtimeContext,
                operationOptions
            );
        } else if (operation.operationKind === 'paragraph-format') {
            result = await applyParagraphFormatToParagraphByExactText(
                documentXml,
                operation.target,
                operation.properties,
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
            (operation.operationKind === 'comment' || operation.operationKind === 'comment_reply')
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
        if (operationWarnings.length > 0 && result) {
            const merged = Array.from(new Set([
                ...(Array.isArray(result.warnings) ? result.warnings : []),
                ...operationWarnings
            ]));
            result.warnings = merged;
        }
        const isError = result?.status === 'error' || !!result?.error;
        let operationReceipt = null;
        if (isError || result?.hasChanges !== true) {
            session.restoreSavepoint(savepoint);
            const disposition = isError ? 'refused' : 'no_change';
            operationReceipt = createEmptyReceipt(
                operationIndex,
                operation.operationId,
                authorUsed,
                disposition
            );
            if (Array.isArray(result?.warnings)) {
                for (const w of result.warnings) {
                    operationReceipt.warnings.push(String(w));
                }
            }
        } else {
            session.markMutationCommitted(operation.operationKind !== 'comment_reply');
            if (operation.captureKey && session.captureTable) {
                session.captureTable.set(
                    operation.captureKey,
                    deriveCapturedEntity(session, operation, operationOptions._mutationLiveNodes)
                );
            }
            if (operationOptions._mutationRemovedNodes?.length > 0 && session.captureTable) {
                invalidateAffectedCaptures(session.captureTable, operationOptions._mutationRemovedNodes);
            }
            if (resolutionCapture.resolvedTarget) {
                session.receiptCollector?.recordAffectedTarget(resolutionCapture.resolvedTarget);
            }
            if (Array.isArray(result.warnings)) {
                for (const w of result.warnings) {
                    session.receiptCollector?.recordWarning(w);
                }
            }
            operationReceipt = session.receiptCollector?.commitOperation('applied');
            result.documentXml = session.deferSerialization
                ? session.currentDocumentXml
                : session.serializeCurrent();

            if (!session.deferSerialization && operationReceipt) {
                const reconciliation = reconcileReceiptsAgainstOutput({
                    documentXml: result.documentXml,
                    commentsXml: result.commentsXml || null,
                    numberingXml: result.numberingXml || null
                }, [operationReceipt]);
                if (!reconciliation.valid) {
                    session.restoreSavepoint(savepoint);
                    operationReceipt.finalDisposition = 'rolled_back';
                    operationReceipt.committed = false;
                    return {
                        documentXml,
                        hasChanges: false,
                        status: 'error',
                        error: reconciliation.error,
                        warnings: [reconciliation.error.message],
                        operationType: operation.operationKind,
                        authorUsed,
                        receipt: operationReceipt,
                        ...resolutionCapture
                    };
                }
            }
        }
        return {
            ...result,
            operationType: operation.operationKind,
            authorUsed,
            receipt: operationReceipt,
            ...resolutionCapture
        };
    } catch (error) {
        session.restoreSavepoint(savepoint);
        const normalizedError = normalizeOperationError(error);
        const operationReceipt = createEmptyReceipt(
            operationIndex,
            operation.operationId,
            authorUsed,
            'refused'
        );
        operationReceipt.warnings.push(normalizedError.message);
        return {
            documentXml,
            hasChanges: false,
            status: 'error',
            error: normalizedError,
            warnings: [normalizedError.message],
            operationType: operation.operationKind,
            authorUsed,
            receipt: operationReceipt,
            ...resolutionCapture
        };
    }
}
