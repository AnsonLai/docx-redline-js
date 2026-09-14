import {
    buildParagraphMetadataIndex,
    getDocumentParagraphNodes,
    getParagraphId,
    resolveTargetParagraph
} from '../core/paragraph-targeting.js';
import {
    normalizeDocumentOperation,
    normalizeTargetDescriptor,
    validateDocumentOperation
} from './document-operation-contract.js';
import { compileExactReplacements } from './localized-replacement-compiler.js';

const TEXT_WRITE_KINDS = new Set(['redline', 'restore', 'rejected-insert']);
const FORMAT_WRITE_KINDS = new Set(['highlight', 'format', 'paragraph-format']);

function normalizedError(error) {
    const details = {};
    if (error && typeof error === 'object') {
        for (const [key, value] of Object.entries(error)) {
            if (key === 'name' || key === 'stack' || key === 'message' || key === 'code') continue;
            details[key] = value;
        }
    }
    return {
        code: typeof error?.code === 'string' && error.code ? error.code : 'OPERATION_ERROR',
        message: error?.message || String(error),
        ...details
    };
}

function isParagraphAttached(document, paragraph) {
    return !!paragraph && getDocumentParagraphNodes(document).includes(paragraph);
}

function paragraphNodesFromMutation(nodes = []) {
    const paragraphs = [];
    for (const node of nodes) {
        if (!node || node.nodeType !== 1) continue;
        if (node.localName === 'p') paragraphs.push(node);
        for (const paragraph of Array.from(node.getElementsByTagNameNS?.('*', 'p') || [])) {
            if (!paragraphs.includes(paragraph)) paragraphs.push(paragraph);
        }
    }
    return paragraphs;
}

/** Session-local source identities over live paragraph nodes. */
export class SourceTargetRegistry {
    constructor(document) {
        this.document = document;
        this.nextId = 1;
        this.entries = new Map();
        this.idsByNode = new WeakMap();
    }

    register(paragraph, metadata = {}) {
        const existing = this.idsByNode.get(paragraph);
        if (existing) return existing;
        const sourceId = `S${this.nextId++}`;
        this.entries.set(sourceId, {
            sourceId,
            paragraph,
            sourceIndex: metadata.index || null,
            paragraphId: metadata.paragraphId || null,
            fingerprint: metadata.fingerprint || null,
            text: metadata.text || '',
            revisionView: metadata.revisionView === 'rejected' ? 'rejected' : 'accepted',
            consumed: false,
            consumedByOperation: null
        });
        this.idsByNode.set(paragraph, sourceId);
        return sourceId;
    }

    describe(sourceId) {
        const entry = this.entries.get(sourceId);
        if (!entry) return null;
        return {
            sourceId,
            index: entry.sourceIndex,
            paragraphId: entry.paragraphId,
            fingerprint: entry.fingerprint,
            text: entry.text,
            revisionView: entry.revisionView
        };
    }

    resolve(sourceId, document = this.document) {
        const entry = this.entries.get(sourceId);
        if (!entry) {
            return {
                error: {
                    code: 'SOURCE_TARGET_NOT_FOUND',
                    message: `Compiled source target ${String(sourceId)} was not found.`
                }
            };
        }
        if (entry.consumed) {
            return {
                error: {
                    code: 'TARGET_CONSUMED_BY_OPERATION',
                    message: `Compiled source target ${sourceId} was consumed by operation ${entry.consumedByOperation}.`,
                    sourceTarget: this.describe(sourceId),
                    consumedByOperation: entry.consumedByOperation
                }
            };
        }
        if (isParagraphAttached(document, entry.paragraph)) {
            return { paragraph: entry.paragraph, resolvedBy: 'batch_source' };
        }
        return {
            error: {
                code: 'TARGET_CONSUMED_BY_OPERATION',
                message: `Compiled source target ${sourceId} is no longer present in the live document.`,
                sourceTarget: this.describe(sourceId),
                consumedByOperation: entry.consumedByOperation
            }
        };
    }

    createSavepoint(document = this.document) {
        const paragraphs = getDocumentParagraphNodes(document);
        return new Map(Array.from(this.entries, ([sourceId, entry]) => [sourceId, {
            paragraphIndex: paragraphs.indexOf(entry.paragraph),
            consumed: entry.consumed,
            consumedByOperation: entry.consumedByOperation
        }]));
    }

    restoreSavepoint(document, snapshot) {
        this.document = document;
        this.idsByNode = new WeakMap();
        const paragraphs = getDocumentParagraphNodes(document);
        for (const [sourceId, entry] of this.entries) {
            const saved = snapshot instanceof Map ? snapshot.get(sourceId) : null;
            entry.consumed = saved?.consumed === true;
            entry.consumedByOperation = saved?.consumedByOperation || null;
            entry.paragraph = Number.isInteger(saved?.paragraphIndex) && saved.paragraphIndex >= 0
                ? (paragraphs[saved.paragraphIndex] || null)
                : null;
            if (entry.paragraph) this.idsByNode.set(entry.paragraph, sourceId);
        }
    }

    commitMutation(sourceId, removedNodes, liveNodes, operationIndex) {
        const entry = this.entries.get(sourceId);
        if (!entry || !Array.isArray(removedNodes) || !removedNodes.includes(entry.paragraph)) return;
        const candidates = paragraphNodesFromMutation(liveNodes);
        const sameId = entry.paragraphId
            ? candidates.filter(paragraph => getParagraphId(paragraph) === entry.paragraphId)
            : [];
        const successor = sameId.length === 1
            ? sameId[0]
            : (candidates.length === 1 ? candidates[0] : null);
        if (successor) {
            entry.paragraph = successor;
            this.idsByNode.set(successor, sourceId);
            return;
        }
        entry.paragraph = null;
        entry.consumed = true;
        entry.consumedByOperation = operationIndex;
    }
}

function resolvedTarget(metadata, revisionView) {
    return {
        index: metadata.index,
        paragraphId: metadata.paragraphId,
        fingerprint: metadata.fingerprint,
        text: metadata.text,
        inTable: metadata.inTable,
        revisionView
    };
}

function resolveDescriptor(xmlDoc, descriptor, kind, strictTargets, metadataIndices, callbacks = {}) {
    const revisionView = descriptor?.revisionView === 'rejected' ? 'rejected' : 'accepted';
    if (!metadataIndices[revisionView]) {
        metadataIndices[revisionView] = buildParagraphMetadataIndex(xmlDoc, { revisionView });
    }
    const index = metadataIndices[revisionView];
    const resolution = resolveTargetParagraph(xmlDoc, {
        targetText: descriptor.text,
        targetRef: descriptor.index,
        targetDescriptor: descriptor,
        opType: kind,
        strictAmbiguity: strictTargets,
        paragraphMetadataIndex: index,
        metadataIndices,
        onInfo: callbacks.onInfo,
        onWarn: callbacks.onWarn
    });
    const metadata = index.byParagraph.get(resolution.paragraph);
    return {
        paragraph: resolution.paragraph,
        resolvedBy: resolution.resolvedBy,
        metadata: resolvedTarget(metadata, revisionView)
    };
}

function buildConflict(code, message, bindings, target, action) {
    return {
        code,
        message,
        operationIndexes: bindings.map(binding => binding.index).sort((a, b) => a - b),
        target,
        recovery: {
            action,
            sameArgumentsSafe: false,
            requiresReinspection: false,
            requiresUserAuthorization: false
        }
    };
}

/** Resolve source targets once and derive source-level conflicts before mutation. */
export function compileOperationBatch(xmlDoc, operations = [], options = {}) {
    const sourceOperations = Array.isArray(operations) ? operations : [];
    const strictTargets = options.strictTargets !== false;
    const registry = new SourceTargetRegistry(xmlDoc);
    const metadataIndices = {
        accepted: buildParagraphMetadataIndex(xmlDoc, { revisionView: 'accepted' }),
        rejected: null
    };
    const bindings = [];
    const compiledOperations = [];

    for (let index = 0; index < sourceOperations.length; index += 1) {
        const sourceOperation = sourceOperations[index];
        const validation = validateDocumentOperation(sourceOperation);
        const operation = validation.operation || normalizeDocumentOperation(sourceOperation);
        const binding = {
            index: index + 1,
            operationKind: operation.operationKind,
            operationId: operation.operationId,
            sourceIds: [],
            warnings: []
        };
        const compiled = { ...sourceOperation };

        const dynamicOccurrence = operation.targetDescriptor?.occurrence != null
            && !operation.targetDescriptor?.paragraphId
            && !operation.targetDescriptor?.fingerprint;
        const compositeTextTarget = typeof operation.targetDescriptor?.text === 'string'
            && /\r|\n/.test(operation.targetDescriptor.text)
            && !operation.targetEndDescriptor;
        if (dynamicOccurrence || compositeTextTarget) binding.dynamic = true;

        if (
            validation.valid
            && operation.operationKind !== 'comment_reply'
            && !operation.targetDescriptor?.captureRef
            && !binding.dynamic
        ) {
            try {
                const start = resolveDescriptor(
                    xmlDoc,
                    operation.targetDescriptor,
                    operation.operationKind,
                    strictTargets,
                    metadataIndices,
                    {
                        onInfo: options.onInfo,
                        onWarn: warning => binding.warnings.push(String(warning))
                    }
                );
                const startId = registry.register(start.paragraph, start.metadata);
                binding.sourceIds.push(startId);
                binding.resolvedBy = start.resolvedBy;
                binding.resolvedTarget = start.metadata;
                compiled._compiledSourceId = startId;
                compiled._compiledResolvedBy = start.resolvedBy;
                if (binding.warnings.length > 0) compiled._compiledWarnings = binding.warnings;

                if (Array.isArray(operation.replacements)) {
                    const patchCompilation = compileExactReplacements(
                        start.metadata.text,
                        operation.replacements
                    );
                    if (!patchCompilation.ok) {
                        throw Object.assign(new Error(patchCompilation.error.message), patchCompilation.error);
                    }
                    compiled.modified = patchCompilation.desiredText;
                    delete compiled.replacements;
                    compiled._localizedReplacementCompilation = {
                        sourceText: start.metadata.text,
                        desiredText: patchCompilation.desiredText,
                        replacements: patchCompilation.replacements
                    };
                    const canonicalValidation = validateDocumentOperation(compiled);
                    if (!canonicalValidation.valid) {
                        throw Object.assign(
                            new Error(canonicalValidation.error.message),
                            canonicalValidation.error
                        );
                    }
                } else if (operation.operationKind === 'restore' && sourceOperation.modified === undefined) {
                    compiled.modified = start.metadata.text;
                    const canonicalValidation = validateDocumentOperation(compiled);
                    if (!canonicalValidation.valid) {
                        throw Object.assign(
                            new Error(canonicalValidation.error.message),
                            canonicalValidation.error
                        );
                    }
                }

                const targetEndDescriptor = operation.targetEndDescriptor
                    || (operation.targetEndRef != null
                        ? normalizeTargetDescriptor(null, operation.targetEndRef, operation.targetDescriptor.revisionView)
                        : null);
                if (targetEndDescriptor) {
                    const end = resolveDescriptor(
                        xmlDoc,
                        targetEndDescriptor,
                        operation.operationKind,
                        strictTargets,
                        metadataIndices,
                        {
                            onInfo: options.onInfo,
                            onWarn: warning => binding.warnings.push(String(warning))
                        }
                    );
                    const endId = registry.register(end.paragraph, end.metadata);
                    compiled._compiledSourceEndId = endId;
                    const paragraphs = getDocumentParagraphNodes(xmlDoc);
                    const startIndex = paragraphs.indexOf(start.paragraph);
                    const endIndex = paragraphs.indexOf(end.paragraph);
                    if (startIndex < 0 || endIndex < startIndex) {
                        throw Object.assign(new Error('Target range is not a forward contiguous paragraph range.'), {
                            code: 'TARGET_RANGE_INVALID'
                        });
                    }
                    for (const paragraph of paragraphs.slice(startIndex, endIndex + 1)) {
                        const viewIndex = metadataIndices[operation.targetDescriptor.revisionView] || metadataIndices.accepted;
                        const metadata = viewIndex.byParagraph.get(paragraph);
                        binding.sourceIds.push(registry.register(paragraph, metadata || {}));
                    }
                    binding.sourceIds = [...new Set(binding.sourceIds)];
                    binding.resolvedTargetEnd = end.metadata;
                }
            } catch (error) {
                binding.error = normalizedError(error);
            }
        }
        bindings.push(binding);
        compiledOperations.push(compiled);
    }

    const existingCaptureKeys = new Set(compiledOperations.map(operation => operation?.captureKey).filter(Boolean));
    let nextImplicitCapture = 1;
    for (const binding of bindings) {
        if (binding.error?.code !== 'TARGET_NOT_FOUND') continue;
        const consumerIndex = binding.index - 1;
        const consumer = normalizeDocumentOperation(sourceOperations[consumerIndex]);
        const targetText = consumer.targetDescriptor?.text;
        if (!targetText) continue;
        const producers = sourceOperations.flatMap((candidate, producerIndex) => {
            if (producerIndex === consumerIndex || typeof candidate?.modified !== 'string') return [];
            const createdParagraphs = candidate.modified.split(/\r?\n/).map(text => text.trim()).filter(Boolean);
            return createdParagraphs.includes(targetText.trim()) ? [{ candidate, producerIndex }] : [];
        });
        if (producers.length !== 1) continue;
        const [{ candidate: producer, producerIndex }] = producers;
        let captureKey = producer.captureKey || compiledOperations[producerIndex].captureKey;
        if (!captureKey) {
            do {
                captureKey = `__batch_created_${nextImplicitCapture++}`;
            } while (existingCaptureKeys.has(captureKey));
            existingCaptureKeys.add(captureKey);
            compiledOperations[producerIndex].captureKey = captureKey;
        }
        compiledOperations[consumerIndex].target = { captureRef: captureKey, select: targetText };
        delete compiledOperations[consumerIndex]._compiledSourceId;
        delete compiledOperations[consumerIndex]._compiledResolvedBy;
        delete binding.error;
        binding.sourceIds = [];
        binding.resolvedBy = 'created_content_dependency';
        binding.createdByOperation = producerIndex + 1;
        binding.captureRef = captureKey;
    }

    const bySource = new Map();
    for (const binding of bindings) {
        if (binding.error || binding.dynamic) continue;
        for (const sourceId of binding.sourceIds) {
            if (!bySource.has(sourceId)) bySource.set(sourceId, []);
            bySource.get(sourceId).push(binding);
        }
    }

    const conflicts = [];
    const conflictKeys = new Set();
    for (const [sourceId, sourceBindings] of bySource) {
        const textWrites = sourceBindings.filter(binding => TEXT_WRITE_KINDS.has(binding.operationKind));
        const formatWrites = sourceBindings.filter(binding => FORMAT_WRITE_KINDS.has(binding.operationKind));
        if (textWrites.length > 1) {
            const key = `text:${textWrites.map(binding => binding.index).sort().join(',')}`;
            if (!conflictKeys.has(key)) {
                conflictKeys.add(key);
                conflicts.push(buildConflict(
                    'OVERLAPPING_SOURCE_TARGETS',
                    `Operations ${textWrites.map(binding => binding.index).join(', ')} contain incompatible text writes to the same batch-start source target.`,
                    textWrites,
                    registry.describe(sourceId),
                    'consolidate_operations'
                ));
            }
        }
        if (textWrites.length > 0 && formatWrites.length > 0) {
            const affected = [...textWrites, ...formatWrites];
            const key = `format:${affected.map(binding => binding.index).sort().join(',')}`;
            if (!conflictKeys.has(key)) {
                conflictKeys.add(key);
                conflicts.push(buildConflict(
                    'REVISION_ORDER_CONFLICT',
                    `Operations ${affected.map(binding => binding.index).join(', ')} combine text and formatting writes on the same batch-start source target.`,
                    affected,
                    registry.describe(sourceId),
                    'split_or_consolidate_operations'
                ));
            }
        }
    }

    return {
        valid: conflicts.length === 0 && bindings.every(binding => !binding.error),
        compiledOperations,
        bindings,
        conflicts,
        registry
    };
}
