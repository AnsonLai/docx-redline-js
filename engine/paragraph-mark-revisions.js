/**
 * Paragraph-mark revisions: the single place that constructs a tracked `w:ins` / `w:del` under
 * `w:pPr/w:rPr` (the paragraph mark). Every write path (reconstruction, list generation, table
 * generation, document-level shortcuts) goes through this module so that no path can forget a rule
 * the others have (see docs/plans revision-fidelity-invariants, WP-6).
 *
 * Word semantics this module encodes:
 *   - an inserted mark (`w:ins`) is removed by Reject All, merging the paragraph into the next one;
 *   - a deleted mark (`w:del`) is restored by Reject All, so a deleted source paragraph keeps its own
 *     deleted paragraph (never collapse N deleted paragraphs into one);
 *   - Reject All of a correct output has exactly as many paragraphs as the source.
 *
 * Depends only on `core/`. Both DOM callers and string-OOXML callers share one implementation:
 * the DOM entry points are `markParagraphInserted` / `markParagraphDeleted` /
 * `setParagraphPropertiesMarkRevision`; string builders are `paragraphMarkRevisionXml` and
 * `paragraphMarkPPrXml` (both format the same metadata the DOM path writes).
 */

import { createRevisionMetadata, escapeXml } from '../core/types.js';
import { createWordElement } from '../core/word-xml.js';

/** @typedef {'ins'|'del'} ParagraphMarkKind */

/**
 * Revision context shared by every entry point.
 *
 * @typedef {Object} ParagraphMarkContext
 * @property {string} [author] - Revision author (falls back to the configured default author)
 * @property {string} [date] - Optional ISO timestamp override (default: current revision timestamp)
 * @property {Object} [revisionIdAllocator] - RevisionIdAllocator, or a document/node whose allocator
 *   is used (default: the paragraph's owner document, then the process-wide allocator). Ids are
 *   allocated through this so receipts record them.
 * @property {{id: (string|number), author: string, date: string}} [metadata] - Pre-allocated metadata
 *   (skips allocation; used when the caller already reserved the id)
 */

function localNameOf(node) {
    if (!node || node.nodeType !== 1) return '';
    return node.localName || String(node.nodeName || '').replace(/^.*:/, '');
}

function directWordChild(node, localName) {
    for (const child of Array.from(node?.childNodes || [])) {
        if (child.nodeType === 1 && localNameOf(child) === localName) return child;
    }
    return null;
}

function normalizeKind(kind) {
    if (kind === 'ins' || kind === 'del') return kind;
    throw new Error(`Unsupported paragraph-mark revision kind: ${kind}`);
}

function allocateMetadata(kind, ctx, fallbackNode) {
    const context = ctx || {};
    if (context.metadata) return context.metadata;
    const allocatorOrNode = context.revisionIdAllocator
        || (fallbackNode?.nodeType === 9 ? fallbackNode : fallbackNode?.ownerDocument)
        || null;
    const metadata = createRevisionMetadata(context.author, allocatorOrNode, kind);
    if (context.date) metadata.date = context.date;
    return metadata;
}

function buildMarker(xmlDoc, kind, metadata) {
    const marker = createWordElement(xmlDoc, kind === 'ins' ? 'w:ins' : 'w:del');
    marker.setAttribute('w:id', String(metadata.id));
    marker.setAttribute('w:author', metadata.author);
    marker.setAttribute('w:date', metadata.date);
    return marker;
}

function removeMarkers(rPr) {
    for (const child of Array.from(rPr.childNodes || [])) {
        const name = localNameOf(child);
        if (name === 'ins' || name === 'del') rPr.removeChild(child);
    }
}

function ensureParagraphProperties(xmlDoc, paragraph) {
    let pPr = directWordChild(paragraph, 'pPr');
    if (!pPr) {
        pPr = createWordElement(xmlDoc, 'w:pPr');
        paragraph.insertBefore(pPr, paragraph.firstChild || null);
    } else if (paragraph.firstChild !== pPr) {
        paragraph.insertBefore(pPr, paragraph.firstChild || null);
    }
    return pPr;
}

function ensureMarkRunProperties(xmlDoc, pPr) {
    let rPr = directWordChild(pPr, 'rPr');
    if (!rPr) {
        rPr = createWordElement(xmlDoc, 'w:rPr');
        pPr.appendChild(rPr);
    } else if (pPr.lastChild !== rPr) {
        pPr.appendChild(rPr);
    }
    return rPr;
}

/**
 * Sets (replacing any previous ins/del) the paragraph-mark revision on a `w:pPr` element. Used by
 * callers that hold detached paragraph properties (list generation parses a `pPr` string).
 *
 * @param {Element} pPr - `w:pPr` element
 * @param {ParagraphMarkKind} kind
 * @param {ParagraphMarkContext} [ctx]
 * @returns {Element} The marker element
 */
export function setParagraphPropertiesMarkRevision(pPr, kind, ctx = {}) {
    const resolved = normalizeKind(kind);
    const xmlDoc = pPr.ownerDocument;
    const metadata = allocateMetadata(resolved, ctx, pPr);
    const rPr = ensureMarkRunProperties(xmlDoc, pPr);
    removeMarkers(rPr);
    const marker = buildMarker(xmlDoc, resolved, metadata);
    rPr.appendChild(marker);
    return marker;
}

function markParagraph(paragraph, kind, ctx) {
    const xmlDoc = paragraph.ownerDocument;
    const pPr = ensureParagraphProperties(xmlDoc, paragraph);
    return setParagraphPropertiesMarkRevision(pPr, kind, {
        ...ctx,
        revisionIdAllocator: ctx?.revisionIdAllocator || xmlDoc
    });
}

/**
 * Tracks the paragraph's mark as inserted (`pPr/rPr/ins`). Reject All removes the mark and merges
 * the paragraph into the next one.
 *
 * @param {Element} paragraph - `w:p`
 * @param {ParagraphMarkContext} [ctx]
 * @returns {Element} The `w:ins` marker
 */
export function markParagraphInserted(paragraph, ctx = {}) {
    return markParagraph(paragraph, 'ins', ctx);
}

/**
 * Tracks the paragraph's mark as deleted (`pPr/rPr/del`). The paragraph content is expected to be
 * moved into `w:del` wrappers by the caller; Reject All restores the paragraph.
 *
 * @param {Element} paragraph - `w:p`
 * @param {ParagraphMarkContext} [ctx]
 * @returns {Element} The `w:del` marker
 */
export function markParagraphDeleted(paragraph, ctx = {}) {
    return markParagraph(paragraph, 'del', ctx);
}

/**
 * Removes any tracked ins/del from the paragraph mark (and the `w:rPr` when that leaves it empty).
 *
 * @param {Element} paragraph - `w:p`
 * @returns {boolean} True when a marker was removed
 */
export function clearParagraphMarkRevision(paragraph) {
    const pPr = directWordChild(paragraph, 'pPr');
    const rPr = directWordChild(pPr, 'rPr');
    if (!rPr) return false;
    const before = rPr.childNodes.length;
    removeMarkers(rPr);
    const removed = rPr.childNodes.length !== before;
    if (removed && !Array.from(rPr.childNodes).some(child => child.nodeType === 1)) {
        pPr.removeChild(rPr);
    }
    return removed;
}

/**
 * Re-applies a previously captured inserted-mark identity (id/author/date kept verbatim) when the
 * paragraph does not carry an `ins` mark any more. Does not allocate a new id.
 *
 * @param {Element} paragraph - `w:p`
 * @param {{id?: (string|number), author: string, date?: string}} metadata
 * @returns {Element|null} The added marker, or null when an inserted mark already exists
 */
export function restoreParagraphMarkInserted(paragraph, metadata) {
    const xmlDoc = paragraph.ownerDocument;
    const pPr = ensureParagraphProperties(xmlDoc, paragraph);
    const rPr = directWordChild(pPr, 'rPr') || (() => {
        const created = createWordElement(xmlDoc, 'w:rPr');
        pPr.appendChild(created);
        return created;
    })();
    if (directWordChild(rPr, 'ins')) return null;
    const marker = createWordElement(xmlDoc, 'w:ins');
    if (metadata.id) marker.setAttribute('w:id', String(metadata.id));
    marker.setAttribute('w:author', metadata.author);
    if (metadata.date) marker.setAttribute('w:date', metadata.date);
    rPr.appendChild(marker);
    return marker;
}

/**
 * String form of the marker, for callers that build OOXML strings. Allocation and formatting are the
 * same as the DOM path (allocator from `ctx.revisionIdAllocator`, receipts record the id).
 *
 * @param {ParagraphMarkKind} kind
 * @param {ParagraphMarkContext} [ctx]
 * @returns {string} e.g. `<w:ins w:id="7" w:author="A" w:date="..."/>`
 */
export function paragraphMarkRevisionXml(kind, ctx = {}) {
    const resolved = normalizeKind(kind);
    const metadata = allocateMetadata(resolved, ctx, null);
    return `<w:${resolved} w:id="${metadata.id}" w:author="${escapeXml(metadata.author)}" w:date="${metadata.date}"/>`;
}

/**
 * Paragraph properties carrying only the paragraph-mark revision, for `serializeToOoxml(..., pPr)`.
 *
 * @param {ParagraphMarkKind} kind
 * @param {ParagraphMarkContext} [ctx]
 * @returns {string} `<w:pPr><w:rPr>...</w:rPr></w:pPr>`
 */
export function paragraphMarkPPrXml(kind, ctx = {}) {
    return `<w:pPr><w:rPr>${paragraphMarkRevisionXml(kind, ctx)}</w:rPr></w:pPr>`;
}

/**
 * @typedef {Object} ParagraphMarkAlignmentEntry
 * @property {number|null} [sourceIndex] - Index into sourceParagraphs, or null/undefined for a new paragraph
 * @property {number|null} [targetIndex] - Index into targetParagraphs, or null/undefined for a deleted paragraph
 * @property {boolean} [propertyChange] - Aligned paragraph whose properties change (tracked as pPrChange)
 */

/**
 * @typedef {Object} ParagraphMarkPlanEntry
 * @property {'keep'|'insertMark'|'deleteMark'|'propertyChange'} action
 * @property {number|null} sourceIndex
 * @property {number|null} targetIndex
 */

/**
 * Plans what each paragraph's mark must be so that Reject All restores exactly the source
 * paragraphs. Pure: it inspects no DOM, it only checks that the alignment accounts for every source
 * and target paragraph exactly once and returns the action per aligned entry in alignment order.
 *
 * Guarantees (throws a descriptive Error when the alignment violates them):
 *   - the untracked source-side marks (keep + propertyChange + deleteMark restored by Reject All)
 *     equal the number of source paragraphs;
 *   - every target paragraph without a source paragraph carries an inserted mark;
 *   - one deleted paragraph per deleted source paragraph (never N collapsed into one).
 *
 * @param {{sourceParagraphs: Array, targetParagraphs: Array, alignment: ParagraphMarkAlignmentEntry[]}} input
 * @param {ParagraphMarkContext} [_ctx] Reserved: the plan does not allocate ids (the mark helpers do)
 * @returns {ParagraphMarkPlanEntry[]}
 */
export function planParagraphMarks({ sourceParagraphs = [], targetParagraphs = [], alignment = [] } = {}, _ctx = {}) {
    const seenSource = new Set();
    const seenTarget = new Set();
    const plan = alignment.map((entry, position) => {
        const sourceIndex = Number.isInteger(entry?.sourceIndex) ? entry.sourceIndex : null;
        const targetIndex = Number.isInteger(entry?.targetIndex) ? entry.targetIndex : null;
        if (sourceIndex === null && targetIndex === null) {
            throw new Error(`Paragraph-mark plan entry ${position} has neither a source nor a target paragraph.`);
        }
        if (sourceIndex !== null) {
            if (sourceIndex < 0 || sourceIndex >= sourceParagraphs.length || seenSource.has(sourceIndex)) {
                throw new Error(`Paragraph-mark plan references source paragraph ${sourceIndex} out of range or more than once.`);
            }
            seenSource.add(sourceIndex);
        }
        if (targetIndex !== null) {
            if (targetIndex < 0 || targetIndex >= targetParagraphs.length || seenTarget.has(targetIndex)) {
                throw new Error(`Paragraph-mark plan references target paragraph ${targetIndex} out of range or more than once.`);
            }
            seenTarget.add(targetIndex);
        }
        let action;
        if (sourceIndex === null) action = 'insertMark';
        else if (targetIndex === null) action = 'deleteMark';
        else action = entry.propertyChange ? 'propertyChange' : 'keep';
        return { action, sourceIndex, targetIndex };
    });

    if (seenSource.size !== sourceParagraphs.length) {
        throw new Error(`Paragraph-mark plan drops ${sourceParagraphs.length - seenSource.size} source paragraph(s): every source paragraph needs a keep or a deleted mark.`);
    }
    if (seenTarget.size !== targetParagraphs.length) {
        throw new Error(`Paragraph-mark plan leaves ${targetParagraphs.length - seenTarget.size} target paragraph(s) without a source or an inserted mark.`);
    }
    const sourceSideMarks = plan.filter(entry => entry.action !== 'insertMark').length;
    if (sourceSideMarks !== sourceParagraphs.length) {
        throw new Error('Paragraph-mark plan would change the paragraph count seen by Reject All.');
    }
    return plan;
}
