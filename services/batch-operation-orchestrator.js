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
import {
    createEmptyReceipt,
    reconcileReceiptsAgainstOutput
} from './receipt-collector.js';

const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

function operationTargetPriority(op) {
    return op?.type === 'comment' ? 0 : 1;
}

/**
 * Builds a dependency graph and topological execution plan for a batch of operations.
 *
 * @param {Array<object>} operations
 * @returns {{ valid: boolean, scheduled?: Array<{ operation: object, index: number }>, captureProducers?: Map<string, number>, error?: { code: string, message: string } }}
 */
export function buildOperationDependencyPlan(operations = []) {
    const list = Array.isArray(operations) ? operations : [];
    const captureProducers = new Map();

    // 1. Collect and validate capture keys
    for (let i = 0; i < list.length; i++) {
        const op = list[i];
        const captureKey = op?.captureKey ?? op?.targetDescriptor?.captureKey;
        if (captureKey != null) {
            if (typeof captureKey !== 'string' || !captureKey.trim() || captureKey.trim().length > 256) {
                return {
                    valid: false,
                    error: {
                        code: 'INVALID_OPERATION',
                        message: `Operation ${i + 1} declared an invalid captureKey: must be a non-empty string under 256 characters.`
                    }
                };
            }
            const key = captureKey.trim();
            if (captureProducers.has(key)) {
                const priorIndex = captureProducers.get(key);
                return {
                    valid: false,
                    error: {
                        code: 'DUPLICATE_CAPTURE_KEY',
                        message: `Duplicate capture key "${key}" declared by operations at indices ${priorIndex + 1} and ${i + 1}.`
                    }
                };
            }
            captureProducers.set(key, i);
        }
    }

    // 2. Build dependency edges
    const inDegrees = new Array(list.length).fill(0);
    const dependents = Array.from({ length: list.length }, () => new Set());
    const dependencies = Array.from({ length: list.length }, () => new Set());

    for (let i = 0; i < list.length; i++) {
        const op = list[i];
        const captureRef = op?.target?.captureRef ?? op?.targetDescriptor?.captureRef;
        if (captureRef != null) {
            if (typeof captureRef !== 'string' || !captureRef.trim() || captureRef.trim().length > 256) {
                return {
                    valid: false,
                    error: {
                        code: 'INVALID_OPERATION',
                        message: `Operation ${i + 1} referenced an invalid captureRef: must be a non-empty string under 256 characters.`
                    }
                };
            }
            const ref = captureRef.trim();
            if (!captureProducers.has(ref)) {
                return {
                    valid: false,
                    error: {
                        code: 'CAPTURE_NOT_FOUND',
                        message: `Capture "${ref}" referenced by operation ${i + 1} was not found in the batch.`
                    }
                };
            }
            const producerIndex = captureProducers.get(ref);
            if (producerIndex === i) {
                return {
                    valid: false,
                    error: {
                        code: 'CAPTURE_DEPENDENCY_CYCLE',
                        message: `Capture dependency cycle detected: operation ${i + 1} references its own capture "${ref}".`
                    }
                };
            }
            if (!dependencies[i].has(producerIndex)) {
                dependencies[i].add(producerIndex);
                dependents[producerIndex].add(i);
                inDegrees[i]++;
            }
        }
    }

    // 3. Stable topological sort with comment priority among ready nodes
    const ready = [];
    for (let i = 0; i < list.length; i++) {
        if (inDegrees[i] === 0) {
            ready.push({ operation: list[i], index: i });
        }
    }

    const scheduled = [];
    while (ready.length > 0) {
        ready.sort((a, b) => operationTargetPriority(a.operation) - operationTargetPriority(b.operation) || a.index - b.index);
        const next = ready.shift();
        scheduled.push(next);

        for (const depIndex of dependents[next.index]) {
            inDegrees[depIndex]--;
            if (inDegrees[depIndex] === 0) {
                ready.push({ operation: list[depIndex], index: depIndex });
            }
        }
    }

    if (scheduled.length < list.length) {
        return {
            valid: false,
            error: {
                code: 'CAPTURE_DEPENDENCY_CYCLE',
                message: 'Capture dependency cycle detected among batch operations.'
            }
        };
    }

    return {
        valid: true,
        scheduled,
        captureProducers
    };
}

export function orderOperationsForStableTargets(operations = []) {
    const plan = buildOperationDependencyPlan(operations);
    if (!plan.valid) {
        throw Object.assign(new Error(plan.error.message), { code: plan.error.code });
    }
    return plan.scheduled.map(item => item.operation);
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
    const sourceOperations = Array.isArray(operations) ? operations : [];
    const defaultAuthor = getDefaultAuthor();
    const emptyReceipts = () => sourceOperations.map((op, i) => createEmptyReceipt(
        i + 1,
        op?.operationId,
        resolveDocumentOperationAuthor(op, author, defaultAuthor),
        'not_attempted'
    ));

    if (options?.expectedRevision) {
        const tokenValidation = validateRevisionToken(options.expectedRevision);
        if (!tokenValidation.valid) {
            return {
                documentXml,
                hasChanges: false,
                commentsXml: null,
                numberingXmlParts: [],
                results: [],
                receipts: emptyReceipts(),
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
                receipts: emptyReceipts(),
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
                receipts: emptyReceipts(),
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

    const session = options?._documentOperationSession instanceof DocumentOperationSession
        ? options._documentOperationSession
        : new DocumentOperationSession(documentXml, {
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
            receipts: emptyReceipts(),
            executionOrder: [],
            authorsUsed: [],
            status: 'error',
            error: session.parseResult.error,
            warnings: session.parseResult.warnings
        };
    }
    const dependencyPlan = buildOperationDependencyPlan(sourceOperations);
    if (!dependencyPlan.valid) {
        return {
            documentXml,
            hasChanges: false,
            commentsXml: null,
            numberingXmlParts: [],
            results: [],
            receipts: emptyReceipts(),
            executionOrder: [],
            authorsUsed: [],
            status: 'error',
            error: dependencyPlan.error
        };
    }
    const scheduled = dependencyPlan.scheduled;

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
                    _deferDocumentSerialization: true,
                    _operationIndex: index + 1
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
                ...(result.error ? { error: result.error } : {}),
                ...(result.receipt ? { receipt: result.receipt } : {})
            });
            if (isError && !continueOnError) break;
        } catch (error) {
            const normalizedError = normalizeOperationError(error);
            operationFailed = true;
            const authorUsed = resolveDocumentOperationAuthor(operation, author, getDefaultAuthor());
            const errorReceipt = createEmptyReceipt(
                index + 1,
                operation?.operationId,
                authorUsed,
                'refused'
            );
            errorReceipt.warnings.push(normalizedError.message);
            results.push({
                index: index + 1,
                type: operation?.type || 'redline',
                status: 'error',
                operationType: normalizeDocumentOperation(operation).operationKind,
                authorUsed,
                warnings: [normalizedError.message],
                error: normalizedError,
                receipt: errorReceipt
            });
            if (!continueOnError) break;
        }
    }

    results.sort((a, b) => a.index - b.index);

    const allReceipts = [];
    for (let i = 0; i < sourceOperations.length; i++) {
        const opIndex = i + 1;
        const executed = results.find(r => r.index === opIndex);
        if (executed?.receipt) {
            allReceipts.push(executed.receipt);
        } else {
            const unattemptedReceipt = createEmptyReceipt(
                opIndex,
                sourceOperations[i]?.operationId,
                resolveDocumentOperationAuthor(sourceOperations[i], author, defaultAuthor),
                'not_attempted'
            );
            allReceipts.push(unattemptedReceipt);
        }
    }

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

    let reconciliationError = null;
    if (!rolledBack && hasChanges) {
        const reconciliation = reconcileReceiptsAgainstOutput({
            documentXml: outputDocumentXml,
            commentsXml: session.commentsXml,
            numberingXmlParts: session.numberingXmlParts
        }, allReceipts);
        if (!reconciliation.valid) {
            reconciliationError = reconciliation.error;
            if (atomic) {
                rolledBack = true;
            }
        }
    }

    if (rolledBack) {
        for (const r of results) {
            if (r.receipt && r.receipt.attemptedDisposition === 'applied') {
                r.receipt.finalDisposition = 'rolled_back';
                r.receipt.committed = false;
            }
        }
        for (const receipt of allReceipts) {
            if (receipt.attemptedDisposition === 'applied') {
                receipt.finalDisposition = 'rolled_back';
                receipt.committed = false;
            }
        }
    } else {
        for (const r of results) {
            if (r.receipt && r.receipt.attemptedDisposition === 'applied') {
                r.receipt.committed = true;
                r.receipt.finalDisposition = 'applied';
            }
        }
        for (const receipt of allReceipts) {
            if (receipt.attemptedDisposition === 'applied') {
                receipt.committed = true;
                receipt.finalDisposition = 'applied';
            }
        }
    }

    if (!rolledBack) commitBatchRuntimeContext(runtimeContext, context);

    return {
        documentXml: rolledBack ? session.rollback() : outputDocumentXml,
        hasChanges: rolledBack ? false : hasChanges,
        commentsXml: rolledBack ? null : session.commentsXml,
        numberingXmlParts: rolledBack ? [] : session.numberingXmlParts,
        results: [...results].sort((a, b) => (a.index || 0) - (b.index || 0)),
        receipts: allReceipts.sort((a, b) => a.operationIndex - b.operationIndex),
        executionOrder,
        authorsUsed: rolledBack ? [] : Array.from(authorsUsed),
        ...(rolledBack ? {
            rolledBack: true,
            status: 'error',
            error: {
                code: reconciliationError ? reconciliationError.code : (serializationError ? 'DOCUMENT_SERIALIZATION_FAILED' : 'BATCH_OPERATION_FAILED'),
                message: reconciliationError?.message
                    || serializationError?.message
                    || 'Atomic batch rolled back because one or more operations failed.'
            }
        } : (reconciliationError ? {
            status: 'error',
            error: reconciliationError
        } : {}))
    };
}
