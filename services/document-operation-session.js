import { createSerializer, parseOoxmlSafe } from '../adapters/xml-adapter.js';
import {
    RevisionIdAllocator,
    seedRevisionIdsFromDocument,
    setRevisionIdAllocatorForDocument
} from '../core/types.js';

/**
 * Owns the document state used by one standalone operation or batch.
 *
 * Phase 2 establishes this boundary without changing the legacy per-operation
 * serialization behavior. Phase 1 can move mutation onto this live document
 * later without changing the applier or orchestrator contracts.
 */
export class DocumentOperationSession {
    constructor(documentXml, options = {}) {
        this.originalDocumentXml = typeof documentXml === 'string' ? documentXml : '';
        this.currentDocumentXml = this.originalDocumentXml;
        this.serializer = createSerializer();
        this.parseResult = parseOoxmlSafe(this.originalDocumentXml, 'application/xml');
        this.document = this.parseResult.doc || null;
        this.revisionIdAllocator = null;
        this.paragraphIndex = null;
        this.invalidated = false;

        if (this.document) {
            this.revisionIdAllocator = options?._revisionIdAllocator instanceof RevisionIdAllocator
                ? options._revisionIdAllocator
                : new RevisionIdAllocator();
            seedRevisionIdsFromDocument(this.document, this.revisionIdAllocator);
            setRevisionIdAllocatorForDocument(this.document, this.revisionIdAllocator);
        }
    }

    get valid() {
        return !!this.document && !this.parseResult.error;
    }

    setDocumentXml(documentXml) {
        this.currentDocumentXml = documentXml;
        this.invalidated = true;
        this.paragraphIndex = null;
    }

    invalidateParagraphIndex() {
        this.invalidated = true;
        this.paragraphIndex = null;
    }

    serialize() {
        return this.document ? this.serializer.serializeToString(this.document) : this.currentDocumentXml;
    }

    rollback() {
        this.currentDocumentXml = this.originalDocumentXml;
        return this.originalDocumentXml;
    }
}

export function prepareRevisionAllocator(xmlDoc, options = {}) {
    const allocator = options?._revisionIdAllocator instanceof RevisionIdAllocator
        ? options._revisionIdAllocator
        : new RevisionIdAllocator();
    seedRevisionIdsFromDocument(xmlDoc, allocator);
    setRevisionIdAllocatorForDocument(xmlDoc, allocator);
    return allocator;
}

export function cloneBatchRuntimeContext(runtimeContext) {
    if (!runtimeContext || typeof runtimeContext !== 'object') return {};

    const context = { ...runtimeContext };
    if (runtimeContext.listFallbackSharedNumIdByKey instanceof Map) {
        context.listFallbackSharedNumIdByKey = new Map(runtimeContext.listFallbackSharedNumIdByKey);
    }
    if (runtimeContext.tableStructuralRedlineKeys instanceof Set) {
        context.tableStructuralRedlineKeys = new Set(runtimeContext.tableStructuralRedlineKeys);
    }
    if (runtimeContext.numberingIdState && typeof runtimeContext.numberingIdState === 'object') {
        context.numberingIdState = {
            ...runtimeContext.numberingIdState,
            usedNumIds: runtimeContext.numberingIdState.usedNumIds instanceof Set
                ? new Set(runtimeContext.numberingIdState.usedNumIds)
                : runtimeContext.numberingIdState.usedNumIds,
            usedAbstractNumIds: runtimeContext.numberingIdState.usedAbstractNumIds instanceof Set
                ? new Set(runtimeContext.numberingIdState.usedAbstractNumIds)
                : runtimeContext.numberingIdState.usedAbstractNumIds
        };
    }
    if (runtimeContext.listFallbackSequenceState && typeof runtimeContext.listFallbackSequenceState === 'object') {
        context.listFallbackSequenceState = {
            ...runtimeContext.listFallbackSequenceState,
            explicitByNumberingKey: runtimeContext.listFallbackSequenceState.explicitByNumberingKey instanceof Map
                ? new Map(runtimeContext.listFallbackSequenceState.explicitByNumberingKey)
                : runtimeContext.listFallbackSequenceState.explicitByNumberingKey
        };
    }
    return context;
}

export function commitBatchRuntimeContext(runtimeContext, context) {
    if (!runtimeContext || typeof runtimeContext !== 'object') return;

    if (runtimeContext.listFallbackSharedNumIdByKey instanceof Map && context.listFallbackSharedNumIdByKey instanceof Map) {
        runtimeContext.listFallbackSharedNumIdByKey.clear();
        for (const entry of context.listFallbackSharedNumIdByKey) runtimeContext.listFallbackSharedNumIdByKey.set(...entry);
        context.listFallbackSharedNumIdByKey = runtimeContext.listFallbackSharedNumIdByKey;
    }
    if (runtimeContext.tableStructuralRedlineKeys instanceof Set && context.tableStructuralRedlineKeys instanceof Set) {
        runtimeContext.tableStructuralRedlineKeys.clear();
        for (const value of context.tableStructuralRedlineKeys) runtimeContext.tableStructuralRedlineKeys.add(value);
        context.tableStructuralRedlineKeys = runtimeContext.tableStructuralRedlineKeys;
    }
    if (runtimeContext.numberingIdState && context.numberingIdState) {
        for (const key of ['usedNumIds', 'usedAbstractNumIds']) {
            if (runtimeContext.numberingIdState[key] instanceof Set && context.numberingIdState[key] instanceof Set) {
                runtimeContext.numberingIdState[key].clear();
                for (const value of context.numberingIdState[key]) runtimeContext.numberingIdState[key].add(value);
                context.numberingIdState[key] = runtimeContext.numberingIdState[key];
            }
        }
        Object.assign(runtimeContext.numberingIdState, context.numberingIdState);
        context.numberingIdState = runtimeContext.numberingIdState;
    }
    if (runtimeContext.listFallbackSequenceState && context.listFallbackSequenceState) {
        const originalMap = runtimeContext.listFallbackSequenceState.explicitByNumberingKey;
        const updatedMap = context.listFallbackSequenceState.explicitByNumberingKey;
        if (originalMap instanceof Map && updatedMap instanceof Map) {
            originalMap.clear();
            for (const entry of updatedMap) originalMap.set(...entry);
            context.listFallbackSequenceState.explicitByNumberingKey = originalMap;
        }
        Object.assign(runtimeContext.listFallbackSequenceState, context.listFallbackSequenceState);
        context.listFallbackSequenceState = runtimeContext.listFallbackSequenceState;
    }
    Object.assign(runtimeContext, context);
}
