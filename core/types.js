/**
 * OOXML Reconciliation Pipeline - Core Types
 * 
 * Data types and enums for the reconciliation system.
 */

import { getDefaultAuthor } from '../adapters/config.js';

// WordprocessingML namespace
export const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
export const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

/**
 * Diff operation types from word-level diffing
 */
export const DiffOp = Object.freeze({
    EQUAL: 'equal',
    DELETE: 'delete',
    INSERT: 'insert'
});

/**
 * Run types in the run model
 */
export const RunKind = Object.freeze({
    TEXT: 'run',
    DELETION: 'deletion',
    INSERTION: 'insertion',
    HYPERLINK: 'hyperlink',
    BOOKMARK: 'bookmark',
    FIELD: 'field',
    // Container delimiters for preserving hierarchy
    CONTAINER_START: 'container_start',
    CONTAINER_END: 'container_end',
    // Multi-paragraph support
    PARAGRAPH_START: 'paragraph_start'
});

/**
 * Container types that wrap runs
 */
export const ContainerKind = Object.freeze({
    SDT: 'sdt',                 // Content Control
    SMART_TAG: 'smartTag',      // Smart Tag
    CUSTOM_XML: 'customXml',    // Custom XML
    FIELD_COMPLEX: 'fldComplex' // Complex field (fldChar-based)
});

/**
 * Content types for block-level detection
 */
export const ContentType = Object.freeze({
    PARAGRAPH: 'paragraph',
    BULLET_LIST: 'bullet_list',
    NUMBERED_LIST: 'numbered_list',
    TABLE: 'table'
});

/**
 * Supported numbering formats
 */
export const NumberFormat = Object.freeze({
    DECIMAL: 'decimal',           // 1, 2, 3
    LOWER_ALPHA: 'lowerLetter',   // a, b, c
    UPPER_ALPHA: 'upperLetter',   // A, B, C
    LOWER_ROMAN: 'lowerRoman',    // i, ii, iii
    UPPER_ROMAN: 'upperRoman',    // I, II, III
    BULLET: 'bullet',             // •
    OUTLINE: 'outline'            // 1.1.2.3
});

/**
 * Numbering suffixes/formats
 */
export const NumberSuffix = Object.freeze({
    PERIOD: 'period',             // 1.
    PAREN_RIGHT: 'parenRight',    // 1)
    PAREN_BOTH: 'parenBoth',      // (1)
    NONE: 'none'
});

/**
 * @typedef {Object} RunEntry
 * @property {string} kind - RunKind value
 * @property {string} text - Text content of the run
 * @property {string} [rPrXml] - Serialized run properties (formatting)
 * @property {Element|null} [pPrElement] - Lazy paragraph properties element for PARAGRAPH_START entries
 * @property {string} [pPrXml] - Serialized paragraph properties for PARAGRAPH_START entries
 * @property {number} startOffset - Start offset in accepted text
 * @property {number} endOffset - End offset in accepted text
 * @property {string} [author] - Author for track changes
 * @property {string} [nodeXml] - Original XML for special elements
 */

/**
 * @typedef {Object} DiffOperation
 * @property {string} type - DiffOp value
 * @property {number} startOffset - Start offset in original text
 * @property {number} endOffset - End offset in original text
 * @property {string} text - Text content of the operation
 */

/**
 * @typedef {Object} FormatHint
 * @property {number} start - Start offset in clean text
 * @property {number} end - End offset in clean text
 * @property {Object} format - Format flags (bold, italic, etc.)
 */

/**
 * @typedef {Object} IngestionResult
 * @property {RunEntry[]} runModel - Array of run entries
 * @property {string} acceptedText - Reconstructed text from runs
 * @property {Element|null} pPr - Paragraph properties element
 */

/**
 * @typedef {Object} PreprocessResult
 * @property {string} cleanText - Text with markdown stripped
 * @property {FormatHint[]} formatHints - Position-based format information
 */

/**
 * @typedef {Object} ReconciliationResult
 * @property {string} ooxml - The reconciled OOXML output
 * @property {boolean} isValid - Whether validation passed
 * @property {string[]} warnings - Any warnings during processing
 * @property {'package'|'document'|'fragment'} [sourceType] - Shape of the OOXML payload when known
 */

/**
 * @typedef {Object} SerializationOptions
 * @property {string} [author] - Author for generated track changes (defaults to configured default author)
 * @property {boolean} [generateRedlines=true] - Toggle track-change wrappers
 * @property {string|null} [font=null] - Optional font override for generated runs
 */

/**
 * @typedef {Object} DocumentFragmentOptions
 * @property {boolean} [includeNumbering=false] - Include numbering relationship/part
 * @property {string|null} [numberingXml=null] - Custom numbering part payload
 * @property {boolean} [appendTrailingParagraph=true] - Append trailing blank paragraph
 */

/**
 * Escapes XML special characters
 * @param {string} str - String to escape
 * @returns {string} Escaped string
 */
export function escapeXml(str) {
    if (!str) return '';
    return str
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&apos;');
}

const DEFAULT_REVISION_ID_START = 1000;
const MAX_PRACTICAL_REVISION_ID = 2147483647;
const REVISION_ID_SAFETY_MARGIN = 10000;
const REVISION_ELEMENT_NAMES = new Set([
    'ins',
    'del',
    'moveFrom',
    'moveTo',
    'rPrChange',
    'pPrChange',
    'cellIns',
    'cellDel',
    'comment'
]);
const revisionAllocatorByDocument = new WeakMap();

function isRevisionIdElement(element) {
    if (!element || element.nodeType !== 1) return false;
    const localName = String(element.localName || element.nodeName || '').replace(/^.*:/, '');
    if (!REVISION_ELEMENT_NAMES.has(localName)) return false;
    return !element.namespaceURI || element.namespaceURI === NS_W || String(element.nodeName || '').startsWith('w:');
}

function readWordId(element) {
    const raw = element?.getAttributeNS?.(NS_W, 'id')
        || element?.getAttribute?.('w:id')
        || element?.getAttribute?.('id');
    const parsed = Number.parseInt(String(raw ?? ''), 10);
    return Number.isInteger(parsed) && parsed >= 0 ? parsed : null;
}

/**
 * Document-scoped allocator for Word revision IDs.
 */
export class RevisionIdAllocator {
    constructor(startValue = DEFAULT_REVISION_ID_START) {
        this.startValue = Number.isInteger(startValue) && startValue >= 0
            ? startValue
            : DEFAULT_REVISION_ID_START;
        this.nextId = this.startValue;
        this.occupiedIds = new Set();
    }

    seed(xmlDoc) {
        let maxFound = -1;
        const traversalRoot = xmlDoc?.nodeType === 9 ? xmlDoc.documentElement : xmlDoc;
        let node = traversalRoot || null;

        while (node) {
            if (isRevisionIdElement(node)) {
                const id = readWordId(node);
                if (id != null) {
                    this.occupiedIds.add(id);
                    maxFound = Math.max(maxFound, id);
                }
            }

            if (node.firstChild) {
                node = node.firstChild;
                continue;
            }
            while (node && node !== traversalRoot && !node.nextSibling) node = node.parentNode;
            node = node && node !== traversalRoot ? node.nextSibling : null;
        }

        const highRiskBoundary = MAX_PRACTICAL_REVISION_ID - REVISION_ID_SAFETY_MARGIN;
        this.nextId = maxFound >= highRiskBoundary
            ? this.startValue
            : Math.max(this.nextId, maxFound + 1);
        this.advanceToAvailableId();
        return this.nextId;
    }

    advanceToAvailableId() {
        const highRiskBoundary = MAX_PRACTICAL_REVISION_ID - REVISION_ID_SAFETY_MARGIN;
        if (this.nextId >= highRiskBoundary) this.nextId = this.startValue;
        while (this.occupiedIds.has(this.nextId)) {
            this.nextId += 1;
            if (this.nextId >= highRiskBoundary) this.nextId = this.startValue;
        }
    }

    next() {
        this.advanceToAvailableId();
        const id = this.nextId;
        this.occupiedIds.add(id);
        this.nextId += 1;
        return id;
    }
}

let defaultRevisionIdAllocator = new RevisionIdAllocator();

export function setRevisionIdAllocatorForDocument(xmlNode, allocator) {
    const xmlDoc = xmlNode?.nodeType === 9 ? xmlNode : xmlNode?.ownerDocument;
    if (xmlDoc && allocator instanceof RevisionIdAllocator) {
        revisionAllocatorByDocument.set(xmlDoc, allocator);
    }
    return allocator;
}

export function getRevisionIdAllocatorForDocument(xmlNode) {
    const xmlDoc = xmlNode?.nodeType === 9 ? xmlNode : xmlNode?.ownerDocument;
    return xmlDoc ? (revisionAllocatorByDocument.get(xmlDoc) || null) : null;
}

export function createRevisionIdAllocator(xmlDoc, startValue = DEFAULT_REVISION_ID_START) {
    const allocator = new RevisionIdAllocator(startValue);
    allocator.seed(xmlDoc);
    setRevisionIdAllocatorForDocument(xmlDoc, allocator);
    return allocator;
}

/**
 * Gets the next unique revision ID for track changes
 * @returns {number} Unique revision ID
 */
export function getNextRevisionId() {
    return defaultRevisionIdAllocator.next();
}

/**
 * Returns the canonical ISO timestamp used for track-change metadata.
 *
 * @param {Date} [date] - Optional date source (for tests)
 * @returns {string}
 */
export function getRevisionTimestamp(date = new Date()) {
    return date.toISOString();
}

/**
 * Creates shared revision metadata for OOXML track-change tags.
 *
 * @param {string} [author] - Track-change author (defaults to configured default author)
 * @param {RevisionIdAllocator|Document|Element|null} [allocatorOrNode] - Scoped allocator or registered OOXML node
 * @returns {{ id: number, author: string, date: string }}
 */
export function createRevisionMetadata(author, allocatorOrNode = null) {
    const resolvedAuthor = typeof author === 'string' && author.trim()
        ? author.trim()
        : getDefaultAuthor();
    const allocator = allocatorOrNode instanceof RevisionIdAllocator
        ? allocatorOrNode
        : (getRevisionIdAllocatorForDocument(allocatorOrNode) || defaultRevisionIdAllocator);

    return {
        id: allocator.next(),
        author: resolvedAuthor,
        date: getRevisionTimestamp()
    };
}

/**
 * @typedef {Object} ReplacementRevisionEvent
 * @property {number} deletionId - Unique revision ID for deletion
 * @property {number} insertionId - Unique revision ID for insertion
 * @property {string} author - Change author
 * @property {string} date - Shared ISO timestamp
 */

/**
 * Creates paired revision metadata for a replacement event.
 * Allocates two unique IDs but shares author and timestamp.
 *
 * @param {string} [author]
 * @param {RevisionIdAllocator|Document|Element|null} [allocatorOrNode=null]
 * @returns {ReplacementRevisionEvent}
 */
export function createReplacementRevisionEvent(author, allocatorOrNode = null) {
    const resolvedAuthor = typeof author === 'string' && author.trim()
        ? author.trim()
        : getDefaultAuthor();
    const allocator = allocatorOrNode instanceof RevisionIdAllocator
        ? allocatorOrNode
        : (getRevisionIdAllocatorForDocument(allocatorOrNode) || defaultRevisionIdAllocator);
    const date = getRevisionTimestamp();
    return {
        deletionId: allocator.next(),
        insertionId: allocator.next(),
        author: resolvedAuthor,
        date
    };
}

/**
 * Seeds the revision ID counter above any existing Word revision/comment id values.
 *
 * @param {Document|Element} xmlDoc - Parsed OOXML document or element
 * @param {RevisionIdAllocator} [allocator=defaultRevisionIdAllocator] - Allocator to seed
 * @returns {number} Next revision id after seeding
 */
export function seedRevisionIdsFromDocument(xmlDoc, allocator = defaultRevisionIdAllocator) {
    const resolvedAllocator = allocator instanceof RevisionIdAllocator
        ? allocator
        : defaultRevisionIdAllocator;
    const nextId = resolvedAllocator.seed(xmlDoc);
    setRevisionIdAllocatorForDocument(xmlDoc, resolvedAllocator);
    return nextId;
}

/**
 * Resets the revision ID counter (for testing)
 * @param {number} [startValue=1000] - Value to reset to
 */
export function resetRevisionIdCounter(startValue = 1000) {
    defaultRevisionIdAllocator = new RevisionIdAllocator(startValue);
}
