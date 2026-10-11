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
import { getParagraphListInfo } from '../core/list-targeting.js';
import { parseTypedListMarker } from './list-format.js';
import { isTrackedDeletedParagraph } from './numbering-helpers.js';

// list-format strips typed markers (a text write) and binds numPr (a pPr write). On its own target it conflicts with
// every other text or format write except one compatible single-paragraph redline (`checkListFormatComposition`).
const TEXT_WRITE_KINDS = new Set(['redline', 'restore', 'rejected-insert', 'list-format']);
const FORMAT_WRITE_KINDS = new Set(['highlight', 'format', 'paragraph-format']);
const WORD_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
export const LIST_FORMAT_COMPOSITION_CONFLICT = 'LIST_FORMAT_COMPOSITION_CONFLICT';
const COMPOSABLE_REDLINE_TYPES = new Set(['redline', 'replace']);

/**
 * The typed-marker prefix of a paragraph text as list-format reads it: the exact text a convert-plain strip removes
 * ("1. ", "A. ", "- "), `unsupported:<marker>` for a marker shape list-format refuses, or null for no marker.
 */
function typedMarkerKey(text) {
    const value = String(text ?? '');
    const parsed = parseTypedListMarker(value);
    if (!parsed) return null;
    if (parsed.unsupported) return `unsupported:${parsed.marker}`;
    return value.slice(0, parsed.prefixLength);
}

/**
 * Rule for one `redline` and one `list-format` writing the same batch-start paragraph (the list-format targets it, or
 * rebinds it through `scope: 'list'`). They compose when the redline changes only the paragraph's text after its typed
 * marker and the list-format owns the marker and numbering: e.g. "1. Old" -> "1. New" as a redline plus
 * `convert-plain`. Execution order is fixed: the redline runs first (its target binds at batch start), then the
 * list-format strips the unchanged marker from the redlined paragraph. Returns null when compatible, else a reason.
 */
function checkListFormatComposition(redline, sourceText) {
    const operation = redline.operation;
    if (!COMPOSABLE_REDLINE_TYPES.has(operation?.type ?? 'redline')
        || operation.targetEnd != null || operation.targetEndRef != null
        || operation.anchor != null || operation.structuredContent === true || operation.part != null
        || operation.target?.revisionView === 'rejected'
        || typeof operation.modified !== 'string' || /[\r\n]/.test(operation.modified)) {
        return {
            reason: 'redline-shape',
            hint: 'Pair a list-format only with a single-paragraph redline or replace of that paragraph: modified without line breaks, no targetEnd, anchor, part or structuredContent.'
        };
    }
    if (operation.modified === sourceText) {
        return {
            reason: 'redline-unchanged',
            hint: 'The redline does not change the text; drop it and send only the list-format.'
        };
    }
    const before = typedMarkerKey(sourceText);
    const after = typedMarkerKey(operation.modified);
    if (before !== after) {
        return {
            reason: 'marker-changed',
            sourceMarker: before,
            modifiedMarker: after,
            hint: 'The redline adds, removes or changes the typed marker. Keep the marker text exactly as it is (the list-format strips it), and express a new label through list-format (strip "none" with format/start, or change-format).'
        };
    }
    return null;
}

/** Live body paragraphs a `scope: 'list'` change-format/join rebinds outside its named target. */
function listScopeRebindParagraphs(xmlDoc, registry, binding, operation) {
    const mode = operation.list?.mode;
    if (mode !== 'change-format' && mode !== 'join') return [];
    if ((operation.list?.scope || 'list') !== 'list') return [];
    const targets = binding.sourceIds.map(sourceId => registry.entries.get(sourceId)?.paragraph).filter(Boolean);
    const info = targets.map(paragraph => getParagraphListInfo(paragraph)).find(Boolean);
    if (!info) return [];
    return Array.from(xmlDoc.getElementsByTagNameNS(WORD_NS, 'p')).filter(paragraph => {
        if (targets.includes(paragraph) || isTrackedDeletedParagraph(paragraph)) return false;
        return getParagraphListInfo(paragraph)?.numId === info.numId;
    });
}

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

    /** The paragraph that will stand for sourceId once this mutation commits (no state change). */
    previewSuccessor(sourceId, removedNodes, liveNodes) {
        const entry = this.entries.get(sourceId);
        if (!entry || !entry.paragraph) return null;
        if (!Array.isArray(removedNodes) || !removedNodes.includes(entry.paragraph)) return entry.paragraph;
        const candidates = paragraphNodesFromMutation(liveNodes);
        const sameId = entry.paragraphId
            ? candidates.filter(paragraph => getParagraphId(paragraph) === entry.paragraphId)
            : [];
        return sameId.length === 1
            ? sameId[0]
            : (candidates.length === 1 ? candidates[0] : null);
    }

    commitMutation(sourceId, removedNodes, liveNodes, operationIndex) {
        const entry = this.entries.get(sourceId);
        if (!entry || !Array.isArray(removedNodes) || !removedNodes.includes(entry.paragraph)) return;
        const successor = this.previewSuccessor(sourceId, removedNodes, liveNodes);
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

function buildConflict(code, message, bindings, target, action, details = {}, hint = null) {
    return {
        code,
        message,
        operationIndexes: bindings.map(binding => binding.index).sort((a, b) => a - b),
        target,
        ...details,
        recovery: {
            action,
            sameArgumentsSafe: false,
            requiresReinspection: false,
            requiresUserAuthorization: false,
            ...(hint ? { hint } : {})
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
            && !operation.targetEndDescriptor
            // Localized replacements always address one paragraph (the contract refuses ranges and
            // captures), so "\n" in the target text is a soft line break (w:br), not a paragraph
            // boundary. Strong or weak (index + exactText) it binds at batch start; leaving it
            // dynamic would skip replacement compilation and apply an empty "modified".
            && !Array.isArray(operation.replacements);
        if (dynamicOccurrence || compositeTextTarget) binding.dynamic = true;

        if (
            validation.valid
            && operation.operationKind !== 'comment_reply'
            && operation.operationKind !== 'comment_resolve'
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

                // list-format join reads (never writes) the paragraph it continues; bind it at batch start too.
                if (operation.operationKind === 'list-format' && operation.list?.continueFrom && typeof operation.list.continueFrom === 'object') {
                    let continueFrom;
                    try {
                        continueFrom = resolveDescriptor(
                            xmlDoc,
                            normalizeTargetDescriptor(operation.list.continueFrom, null, 'accepted'),
                            operation.operationKind,
                            strictTargets,
                            metadataIndices,
                            {
                                onInfo: options.onInfo,
                                onWarn: warning => binding.warnings.push(String(warning))
                            }
                        );
                    } catch (error) {
                        throw Object.assign(error instanceof Error ? error : new Error(String(error)), { field: 'list.continueFrom' });
                    }
                    compiled._compiledContinueFromId = registry.register(continueFrom.paragraph, continueFrom.metadata);
                    binding.resolvedContinueFrom = continueFrom.metadata;
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
    const rebindsBySource = new Map();
    for (const binding of bindings) {
        if (binding.error || binding.dynamic) continue;
        for (const sourceId of binding.sourceIds) {
            if (!bySource.has(sourceId)) bySource.set(sourceId, []);
            bySource.get(sourceId).push(binding);
        }
        // A scope "list" change-format/join writes pPr on every live item of the target's list, so those items are
        // registered as format writes of the operation: another write to them must compose or is refused.
        if (binding.operationKind === 'list-format') {
            const operation = compiledOperations[binding.index - 1];
            for (const paragraph of listScopeRebindParagraphs(xmlDoc, registry, binding, operation)) {
                const sourceId = registry.register(paragraph, metadataIndices.accepted.byParagraph.get(paragraph) || {});
                if (binding.sourceIds.includes(sourceId)) continue;
                (binding.rebindSourceIds ||= []).push(sourceId);
                if (!bySource.has(sourceId)) bySource.set(sourceId, []);
                if (!rebindsBySource.has(sourceId)) rebindsBySource.set(sourceId, []);
                rebindsBySource.get(sourceId).push(binding);
            }
        }
    }

    const conflicts = [];
    const conflictKeys = new Set();
    const pushConflict = (key, conflict) => {
        if (conflictKeys.has(key)) return;
        conflictKeys.add(key);
        conflicts.push(conflict);
    };
    for (const [sourceId, targetBindings] of bySource) {
        const rebinds = rebindsBySource.get(sourceId) || [];
        const writers = [...targetBindings, ...rebinds]
            .filter(binding => TEXT_WRITE_KINDS.has(binding.operationKind) || FORMAT_WRITE_KINDS.has(binding.operationKind));
        const listFormats = writers.filter(binding => binding.operationKind === 'list-format');
        const others = writers.filter(binding => binding.operationKind !== 'list-format');
        if (listFormats.length === 1 && others.length === 1 && others[0].operationKind === 'redline') {
            const [listFormat] = listFormats;
            const [redline] = others;
            const describedSource = registry.describe(sourceId);
            const incompatible = checkListFormatComposition(
                { operation: compiledOperations[redline.index - 1] },
                redline.resolvedTarget?.text ?? describedSource?.text ?? ''
            );
            if (incompatible) {
                const affected = [redline, listFormat];
                pushConflict(`list:${affected.map(binding => binding.index).sort().join(',')}`, buildConflict(
                    LIST_FORMAT_COMPOSITION_CONFLICT,
                    `Operations ${affected.map(binding => binding.index).sort((a, b) => a - b).join(', ')} combine a redline and a list-format on the same paragraph in a way that does not compose (${incompatible.reason}).`,
                    affected,
                    describedSource,
                    'split_text_and_list_changes',
                    { reason: incompatible.reason, ...(incompatible.sourceMarker !== undefined ? { sourceMarker: incompatible.sourceMarker, modifiedMarker: incompatible.modifiedMarker } : {}) },
                    incompatible.hint
                ));
                continue;
            }
            // Deterministic order: the list-format runs after the redline, whatever the batch order.
            const compiledListFormat = compiledOperations[listFormat.index - 1];
            const runAfter = new Set(compiledListFormat._compiledRunAfter || []);
            runAfter.add(redline.index - 1);
            compiledListFormat._compiledRunAfter = [...runAfter].sort((a, b) => a - b);
            listFormat.composedWith = [...new Set([...(listFormat.composedWith || []), redline.index])];
            redline.composedWith = [...new Set([...(redline.composedWith || []), listFormat.index])];
            continue;
        }
        if (rebinds.length > 0 && writers.length > 1) {
            pushConflict(`rebind:${writers.map(binding => binding.index).sort().join(',')}`, buildConflict(
                LIST_FORMAT_COMPOSITION_CONFLICT,
                `Operation ${rebinds.map(binding => binding.index).join(', ')} (list-format, scope "list") rebinds this list item, and operations ${writers.filter(binding => !rebinds.includes(binding)).map(binding => binding.index).join(', ')} also write it.`,
                writers,
                registry.describe(sourceId),
                'split_text_and_list_changes',
                { reason: 'list-rebind', rebindingOperationIndexes: rebinds.map(binding => binding.index) },
                'A list-format with scope "list" rewrites every item of the list; in the same batch only one single-paragraph redline per item can compose with it. Send the other change in a separate apply, or use scope "range".'
            ));
            continue;
        }
        const sourceBindings = targetBindings;
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
