import { createSerializer, parseOoxmlSafe } from '../adapters/xml-adapter.js';
import { getDefaultAuthor } from '../adapters/config.js';
import { buildTargetReferenceSnapshot } from '../core/paragraph-targeting.js';
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

const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

function operationTargetPriority(op) {
    return op?.type === 'comment' ? 0 : 1;
}

export function orderOperationsForStableTargets(operations = []) {
    return (Array.isArray(operations) ? operations : [])
        .map((operation, index) => ({ operation, index }))
        .sort((a, b) => operationTargetPriority(a.operation) - operationTargetPriority(b.operation) || a.index - b.index)
        .map(entry => entry.operation);
}

function mergeCommentsXml(existingXml, incomingXml) {
    if (!incomingXml) return existingXml || null;
    if (!existingXml) return incomingXml;

    const serializer = createSerializer();
    const existingDoc = parseOoxmlSafe(existingXml, 'application/xml').doc;
    const incomingDoc = parseOoxmlSafe(incomingXml, 'application/xml').doc;
    if (!existingDoc || !incomingDoc) return existingXml;
    const existingRoot = existingDoc.documentElement;
    const existingIds = new Set(
        Array.from(existingRoot.getElementsByTagNameNS(NS_W, 'comment'))
            .map(comment => comment.getAttribute('w:id') || comment.getAttribute('id'))
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
    const session = new DocumentOperationSession(documentXml, options);
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
        context.targetRefSnapshot = buildTargetReferenceSnapshot(session.document);
    }

    let commentsXml = null;
    let hasChanges = false;
    const numberingXmlParts = [];
    const results = [];
    const executionOrder = [];
    const authorsUsed = new Set();
    let operationFailed = false;

    for (const { operation, index } of scheduled) {
        executionOrder.push(index + 1);
        try {
            const result = await applyOperationToDocumentXml(
                session.currentDocumentXml,
                operation,
                author,
                context,
                options
            );
            session.setDocumentXml(result.documentXml);
            hasChanges = hasChanges || result.hasChanges === true;
            commentsXml = mergeCommentsXml(commentsXml, result.commentsXml || null);
            if (result.numberingXml) numberingXmlParts.push(result.numberingXml);
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
    const rolledBack = atomic && operationFailed;
    if (!rolledBack) commitBatchRuntimeContext(runtimeContext, context);

    return {
        documentXml: rolledBack ? session.rollback() : session.currentDocumentXml,
        hasChanges: rolledBack ? false : hasChanges,
        commentsXml: rolledBack ? null : commentsXml,
        numberingXmlParts: rolledBack ? [] : numberingXmlParts,
        results,
        executionOrder,
        authorsUsed: rolledBack ? [] : Array.from(authorsUsed),
        ...(rolledBack ? {
            rolledBack: true,
            status: 'error',
            error: {
                code: 'BATCH_OPERATION_FAILED',
                message: 'Atomic batch rolled back because one or more operations failed.'
            }
        } : {})
    };
}
