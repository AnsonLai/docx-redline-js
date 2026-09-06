import { createSerializer, parseOoxmlSafe } from '../adapters/xml-adapter.js';
import { getDefaultAuthor } from '../adapters/config.js';
import {
    normalizeDocumentOperation,
    resolveDocumentOperationAuthor
} from './document-operation-contract.js';
import {
    applyOperationToDocumentXml,
    normalizeOperationError
} from './document-operation-applier.js';
import {
    cloneBatchRuntimeContext,
    commitBatchRuntimeContext,
    DocumentOperationSession
} from './document-operation-session.js';
import {
    validateRevisionToken,
    computeDocumentPartsRevisionToken,
    areRevisionTokensEqual
} from './revision-token.js';

const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

function operationTargetPriority(op) {
    return op?.type === 'comment' ? 0 : 1;
}

export function orderOperationsForStableTargets(operations = []) {
    return (Array.isArray(operations) ? operations : [])
        .map((operation, index) => ({ operation, index }))
        .sort((a, b) => operationTargetPriority(a.operation) - operationTargetPriority(b.operation) || a.index - b.index)
        .map(item => item.operation);
}

function mergeCommentsXml(existingXml, incomingXml) {
    if (!incomingXml) return existingXml;
    if (!existingXml) return incomingXml;
    const existingDoc = parseOoxmlSafe(existingXml, 'application/xml')?.doc;
    const incomingDoc = parseOoxmlSafe(incomingXml, 'application/xml')?.doc;
    if (!existingDoc || !incomingDoc) return existingXml;
    const serializer = createSerializer();
    const existingRoot = existingDoc.documentElement;
    const existingIds = new Set(
        Array.from(existingDoc.getElementsByTagNameNS(NS_W, 'comment'))
            .map(node => node.getAttribute('w:id') || node.getAttribute('id'))
            .filter(Boolean)
    );
    for (const comment of Array.from(incomingDoc.getElementsByTagNameNS(NS_W, 'comment'))) {
        const id = comment.getAttribute('w:id') || comment.getAttribute('id');
        if (id && existingIds.has(id)) continue;
        existingRoot.appendChild(existingDoc.importNode(comment, true));
        if (id) existingIds.add(id);
    }
    return serializer.serializeToString(existingDoc);
}

export async function applyOperationsToDocumentXml(documentXml, operations, author, runtimeContext = null, options = {}) {
    if (options?.expectedRevision) {
        const tokenValidation = validateRevisionToken(options.expectedRevision);
        if (!tokenValidation.valid) {
            return {
                documentXml,
                hasChanges: false,
                commentsXml: null,
                numberingXmlParts: [],
                results: [],
                executionOrder: [],
                authorsUsed: [],
                status: 'error',
                error: {
                    code: tokenValidation.error?.code || 'INVALID_REVISION_TOKEN',
                    message: tokenValidation.error?.message || 'Invalid revision token.'
                }
            };
        }
        if (options.expectedRevision.scope !== 'document-parts') {
            return {
                documentXml,
                hasChanges: false,
                commentsXml: null,
                numberingXmlParts: [],
                results: [],
                executionOrder: [],
                authorsUsed: [],
                status: 'error',
                error: {
                    code: 'REVISION_TOKEN_SCOPE_MISMATCH',
                    message: `Revision token scope mismatch: expected 'document-parts', got '${options.expectedRevision.scope}'.`
                }
            };
        }
        const currentToken = await computeDocumentPartsRevisionToken({
            documentXml,
            commentsXml: runtimeContext?.commentsXml || options.commentsXml,
            numberingXml: runtimeContext?.numberingXml || options.numberingXml,
            stylesXml: runtimeContext?.stylesXml || options.stylesXml
        }, options);
        if (!areRevisionTokensEqual(currentToken.value, options.expectedRevision.value)) {
            return {
                documentXml,
                hasChanges: false,
                commentsXml: null,
                numberingXmlParts: [],
                results: [],
                executionOrder: [],
                authorsUsed: [],
                status: 'error',
                error: {
                    code: 'REVISION_MISMATCH',
                    message: `Document revision mismatch: expected '${options.expectedRevision.value}', current is '${currentToken.value}'.`
                }
            };
        }
    }

    const session = new DocumentOperationSession(documentXml, {
        ...options,
        _deferDocumentSerialization: true
    });
    if (!session.valid) {
        return {
            documentXml,
            hasChanges: false,
            commentsXml: null,
            numberingXmlParts: [],
            results: [],
            executionOrder: [],
            authorsUsed: [],
            status: 'error',
            error: session.parseResult.error,
            warnings: session.parseResult.warnings
        };
    }
    const sourceOperations = Array.isArray(operations) ? operations : [];
    const scheduled = sourceOperations
        .map((operation, index) => ({ operation, index }))
        .sort((a, b) => operationTargetPriority(a.operation) - operationTargetPriority(b.operation) || a.index - b.index);

    const atomic = options.atomic !== false;
    const continueOnError = options.continueOnError !== false;
    const context = cloneBatchRuntimeContext(runtimeContext);
    if (!(context.targetRefSnapshot instanceof Map)) {
        context.targetRefSnapshot = session.initialTargetReferenceSnapshot;
    }
    session.runtimeContext = context;

    let hasChanges = false;
    const results = session.results;
    const executionOrder = session.executionOrder;
    const authorsUsed = session.authorsUsed;
    let operationFailed = false;

    for (const { operation, index } of scheduled) {
        executionOrder.push(index + 1);
        try {
            const result = await applyOperationToDocumentXml(
                session.currentDocumentXml,
                operation,
                author,
                context,
                {
                    ...options,
                    _documentOperationSession: session,
                    _deferDocumentSerialization: true
                }
            );
            hasChanges = hasChanges || result.hasChanges === true;
            session.commentsXml = mergeCommentsXml(session.commentsXml, result.commentsXml || null);
            if (result.numberingXml) session.numberingXmlParts.push(result.numberingXml);
            const isError = result.status === 'error' || !!result.error;
            operationFailed = operationFailed || isError;
            if (!isError && result.hasChanges && result.authorUsed) authorsUsed.add(result.authorUsed);
            results.push({
                index: index + 1,
                type: operation?.type || 'redline',
                status: isError ? 'error' : (result.hasChanges ? 'applied' : 'no_change'),
                operationType: result.operationType || 'redline',
                authorUsed: result.authorUsed,
                ...(result.resolvedBy ? { resolvedBy: result.resolvedBy } : {}),
                ...(result.resolvedTarget ? { resolvedTarget: result.resolvedTarget } : {}),
                ...(result.resolvedAnchor ? { resolvedAnchor: result.resolvedAnchor } : {}),
                ...(Array.isArray(result.warnings) && result.warnings.length > 0 ? { warnings: result.warnings } : {}),
                ...(result.error ? { error: result.error } : {})
            });
            if (isError && !continueOnError) break;
        } catch (error) {
            const normalizedError = normalizeOperationError(error);
            operationFailed = true;
            const authorUsed = resolveDocumentOperationAuthor(operation, author, getDefaultAuthor());
            results.push({
                index: index + 1,
                type: operation?.type || 'redline',
                status: 'error',
                operationType: normalizeDocumentOperation(operation).operationKind,
                authorUsed,
                warnings: [normalizedError.message],
                error: normalizedError
            });
            if (!continueOnError) break;
        }
    }

    results.sort((a, b) => a.index - b.index);
    let rolledBack = atomic && operationFailed;
    let outputDocumentXml = documentXml;
    let serializationError = null;
    if (!rolledBack && hasChanges) {
        try {
            outputDocumentXml = session.serializeCurrent();
        } catch (error) {
            serializationError = normalizeOperationError(error);
            rolledBack = true;
        }
    }
    if (!rolledBack) commitBatchRuntimeContext(runtimeContext, context);

    return {
        documentXml: rolledBack ? session.rollback() : outputDocumentXml,
        hasChanges: rolledBack ? false : hasChanges,
        commentsXml: rolledBack ? null : session.commentsXml,
        numberingXmlParts: rolledBack ? [] : session.numberingXmlParts,
        results,
        executionOrder,
        authorsUsed: rolledBack ? [] : Array.from(authorsUsed),
        ...(rolledBack ? {
            rolledBack: true,
            status: 'error',
            error: {
                code: serializationError ? 'DOCUMENT_SERIALIZATION_FAILED' : 'BATCH_OPERATION_FAILED',
                message: serializationError?.message
                    || 'Atomic batch rolled back because one or more operations failed.'
            }
        } : {})
    };
}
