import { createSerializer, parseOoxmlSafe } from '../adapters/xml-adapter.js';
import {
    RevisionIdAllocator,
    seedRevisionIdsFromDocument,
    setRevisionIdAllocatorForDocument
} from '../core/types.js';
import {
    buildParagraphMetadataIndex,
    buildTargetReferenceSnapshot
} from '../core/paragraph-targeting.js';

/**
 * Owns the document state used by one standalone operation or batch.
 *
 * The session keeps accuracy-sensitive mutation decisions unchanged while a
 * batch shares one live DOM. Savepoints protect against partial failed/no-op
 * mutations, and the original input string remains the rollback authority.
 */
export class DocumentOperationSession {
    constructor(documentXml, options = {}) {
        this.originalDocumentXml = typeof documentXml === 'string' ? documentXml : '';
        this.currentDocumentXml = this.originalDocumentXml;
        this.serializer = createSerializer();
        this.parseResult = parseOoxmlSafe(this.originalDocumentXml, 'application/xml');
        this.document = this.parseResult.doc || null;
        this.revisionIdAllocator = null;
        this.initialTargetReferenceSnapshot = null;
        this.paragraphIndex = null;
        this.invalidated = false;
        this.hasChanges = false;
        this.deferSerialization = options?._deferDocumentSerialization === true;
        this.instrumentation = options?._sessionInstrumentation || null;
        this.runtimeContext = null;
        this.commentsXml = null;
        this.numberingXmlParts = [];
        this.results = [];
        this.executionOrder = [];
        this.authorsUsed = new Set();
        this.captureTable = new Map();
        this.nextCaptureParaId = 1;

        if (this.document) {
            this.instrumentation?.onDocumentParse?.(this.originalDocumentXml);
            this.revisionIdAllocator = options?._revisionIdAllocator instanceof RevisionIdAllocator
                ? options._revisionIdAllocator
                : new RevisionIdAllocator();
            seedRevisionIdsFromDocument(this.document, this.revisionIdAllocator);
            setRevisionIdAllocatorForDocument(this.document, this.revisionIdAllocator);
            const paragraphIndex = this.getParagraphMetadataIndex();
            this.initialTargetReferenceSnapshot = buildTargetReferenceSnapshot(this.document, paragraphIndex);
        }
    }

    get valid() {
        return !!this.document && !this.parseResult.error;
    }

    setDocumentXml(documentXml) {
        this.currentDocumentXml = documentXml;
        this.invalidateParagraphIndex();
    }

    invalidateParagraphIndex() {
        this.invalidated = true;
        this.paragraphIndex = null;
    }

    serialize() {
        if (!this.document) return this.currentDocumentXml;
        this.instrumentation?.onDocumentSerialize?.(this.document);
        return this.serializer.serializeToString(this.document);
    }

    serializeCurrent() {
        if (!this.hasChanges) return this.originalDocumentXml;
        this.currentDocumentXml = this.serialize();
        return this.currentDocumentXml;
    }

    markMutationCommitted() {
        this.hasChanges = true;
        this.invalidateParagraphIndex();
    }

    generateParagraphId() {
        return (0x40000000 + (this.nextCaptureParaId++)).toString(16).toUpperCase();
    }

    createSavepoint() {
        if (!this.document) return null;
        return {
            document: this.document.cloneNode(true),
            allocatorNextId: this.revisionIdAllocator?.nextId,
            allocatorOccupiedIds: this.revisionIdAllocator?.occupiedIds instanceof Set
                ? new Set(this.revisionIdAllocator.occupiedIds)
                : null,
            hasChanges: this.hasChanges,
            currentDocumentXml: this.currentDocumentXml,
            captureTable: cloneCaptureTable(this.captureTable),
            nextCaptureParaId: this.nextCaptureParaId
        };
    }

    restoreSavepoint(savepoint) {
        if (!savepoint?.document) return;
        this.document = savepoint.document;
        this.hasChanges = savepoint.hasChanges;
        this.currentDocumentXml = savepoint.currentDocumentXml;
        this.captureTable = savepoint.captureTable ? cloneCaptureTable(savepoint.captureTable) : new Map();
        if (typeof savepoint.nextCaptureParaId === 'number') {
            this.nextCaptureParaId = savepoint.nextCaptureParaId;
        }
        if (this.revisionIdAllocator) {
            this.revisionIdAllocator.nextId = savepoint.allocatorNextId;
            if (savepoint.allocatorOccupiedIds instanceof Set) {
                this.revisionIdAllocator.occupiedIds = new Set(savepoint.allocatorOccupiedIds);
            }
            setRevisionIdAllocatorForDocument(this.document, this.revisionIdAllocator);
        }
        this.invalidateParagraphIndex();
    }

    getParagraphIndex() {
        return this.getParagraphMetadataIndex().entries;
    }

    getParagraphMetadataIndex() {
        if (this.paragraphIndex) return this.paragraphIndex;
        this.paragraphIndex = buildParagraphMetadataIndex(this.document);
        this.invalidated = false;
        return this.paragraphIndex;
    }

    rollback() {
        this.currentDocumentXml = this.originalDocumentXml;
        this.hasChanges = false;
        this.captureTable.clear();
        return this.originalDocumentXml;
    }
}

function cloneCaptureTable(table) {
    const cloned = new Map();
    if (table instanceof Map) {
        for (const [k, v] of table.entries()) {
            cloned.set(k, JSON.parse(JSON.stringify(v)));
        }
    }
    return cloned;
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
