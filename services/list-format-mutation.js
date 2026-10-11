/**
 * DOM core of `list-format` `convert-plain` (canonical list operations plan, WP-4).
 *
 * Converts plain paragraphs IN PLACE: the paragraph element, its `w14:paraId`, bookmarks, comments and direct
 * formatting stay; the typed marker (when stripped) becomes a tracked `w:del` of exactly the marker text, and the
 * new `w:numPr` binding is recorded as a tracked `w:pPrChange` through the same path as
 * `paragraph-format.numbering` (`snapshotAndAttachPPrChange` + `applyParagraphPropertiesToPPr`). Word's Reject All
 * restores the literal marker and the unnumbered paragraph; Accept All shows a real list label.
 *
 * Every refusal is decided in a planning pass before anything is mutated. Current properties are read only
 * through `core/current-properties.js`; no paragraph-mark revision is created (a binding is a property change).
 */

import { parseOoxmlSafe } from '../adapters/xml-adapter.js';
import {
    getCurrentChild,
    getCurrentNumIdElements,
    getCurrentParagraphProperties,
    getDirectWordChild
} from '../core/current-properties.js';
import { extractCanonicalParagraphText } from '../core/paragraph-text.js';
import { refreshRunPropertyChangeIds } from '../core/revision-cloning.js';
import { createWordElement, isWordElement } from '../core/word-xml.js';
import { getParagraphListInfo } from '../core/list-targeting.js';
import { createTrackChange, snapshotAndAttachPPrChange } from '../engine/run-builders.js';
import { applyParagraphPropertiesToPPr } from '../engine/rpr-helpers.js';
import {
    createRunFromPieces,
    getRunContentPieces,
    getRunTextLength,
    sliceRunPieces
} from '../engine/surgical-run-splitting.js';
import { allocateNumberingDefinition, mergeNumberingParts } from './numbering-allocation.js';
import { createDynamicNumberingIdState } from './numbering-helpers.js';
import {
    LIST_FORMAT_STYLE_NUMBERED,
    LIST_FORMAT_TARGET_IS_LIST,
    LIST_MARKER_AMBIGUOUS,
    LIST_MARKER_NOT_FOUND,
    LIST_MARKER_UNSUPPORTED,
    listFormatSequenceKey,
    listLevelSpec,
    markerValue,
    parseTypedListMarker,
    resolveMarkerFormat
} from './list-format.js';

const PARAGRAPH_FORMAT_REVISION_CONFLICT = 'PARAGRAPH_FORMAT_REVISION_CONFLICT';
/** Paragraph children that carry no visible text and may precede the marker. */
const ZERO_WIDTH_CHILDREN = new Set([
    'pPr', 'bookmarkStart', 'bookmarkEnd', 'proofErr', 'commentRangeStart', 'commentRangeEnd', 'permStart', 'permEnd'
]);
/** Run children a marker-bearing run may contain (anything else, such as a field or drawing, is refused). */
const SPLITTABLE_RUN_CHILDREN = new Set(['rPr', 't', 'tab', 'br', 'cr', 'noBreakHyphen', 'softHyphen', 'lastRenderedPageBreak']);
const MARKER_RUN_CHILDREN = new Set(['rPr', 't', 'tab', 'lastRenderedPageBreak']);

const excerpt = (text, limit = 80) => {
    const value = String(text ?? '');
    return value.length > limit ? `${value.slice(0, limit)}...` : value;
};
const localName = node => String(node?.localName || node?.nodeName || '').replace(/^.*:/, '');
const readVal = element => element?.getAttribute?.('w:val') ?? element?.getAttribute?.('val') ?? null;

export function refusal(code, message, extra = {}, recovery = null) {
    return { ok: false, error: { code, message, ...extra, ...(recovery ? { recovery } : {}) } };
}

// ---------------------------------------------------------------------------------------------------------
// Style-linked numbering
// ---------------------------------------------------------------------------------------------------------

export function parseStyles(stylesXml) {
    const text = String(stylesXml || '');
    if (!text.trim()) return null;
    const parsed = parseOoxmlSafe(text, 'application/xml');
    return parsed?.doc || null;
}

/**
 * The numId a paragraph style (or its `w:basedOn` ancestors) binds, or null. A paragraph that is numbered by its
 * style is already a list item; converting it would stack a second binding over the style's list.
 */
export function styleNumbering(stylesDoc, styleId) {
    if (!stylesDoc) return null;
    const styles = Array.from(stylesDoc.getElementsByTagNameNS('*', 'style'))
        .filter(style => (style.getAttribute('w:type') ?? style.getAttribute('type')) === 'paragraph');
    const byId = new Map(styles.map(style => [style.getAttribute('w:styleId') ?? style.getAttribute('styleId'), style]));
    let style = styleId
        ? byId.get(styleId)
        : styles.find(candidate => ['1', 'true', 'on'].includes(candidate.getAttribute('w:default') ?? candidate.getAttribute('default')));
    for (let depth = 0; style && depth < 16; depth += 1) {
        const numPr = getCurrentChild(getDirectWordChild(style, 'pPr'), 'numPr');
        const numId = readVal(getCurrentChild(numPr, 'numId'));
        if (numId != null) return numId === '0' ? null : { numId, styleId: style.getAttribute('w:styleId') ?? style.getAttribute('styleId') };
        style = byId.get(readVal(getDirectWordChild(style, 'basedOn')));
    }
    return null;
}

// ---------------------------------------------------------------------------------------------------------
// Marker run location
// ---------------------------------------------------------------------------------------------------------

/**
 * Finds the plain runs that carry the first `prefixLength` characters of the paragraph. Returns the run segments
 * or a reason when the marker is not plain run text (inside a hyperlink, field, tracked change, ...).
 */
function locateMarkerRuns(paragraph, prefixLength, expectedPrefix) {
    const segments = [];
    let offset = 0;
    let collected = '';
    for (const child of Array.from(paragraph.childNodes || [])) {
        if (offset >= prefixLength) break;
        if (child.nodeType !== 1) continue;
        const name = localName(child);
        if (isWordElement(child, 'r')) {
            const pieces = getRunContentPieces(child);
            const length = getRunTextLength(pieces);
            if (length === 0) {
                const hasContent = Array.from(child.childNodes || []).some(node => node.nodeType === 1 && !MARKER_RUN_CHILDREN.has(localName(node)));
                if (hasContent) return { reason: 'run-content' };
                continue;
            }
            const elementNames = Array.from(child.childNodes || []).filter(node => node.nodeType === 1).map(localName);
            if (elementNames.some(item => !SPLITTABLE_RUN_CHILDREN.has(item))) return { reason: 'run-content' };
            const overlap = Math.min(prefixLength, offset + length) - offset;
            // The marker part of the run may hold only text and tabs.
            let covered = 0;
            for (const piece of pieces) {
                if (covered >= overlap) break;
                if (!isWordElement(piece.node, 't') && !isWordElement(piece.node, 'tab')) return { reason: 'run-content' };
                covered = piece.end;
            }
            segments.push({ run: child, start: offset, end: offset + length, pieces });
            collected += pieces.map(piece => piece.text).join('');
            offset += length;
            continue;
        }
        if (ZERO_WIDTH_CHILDREN.has(name)) continue;
        return { reason: name === 'ins' || name === 'del' || name === 'moveTo' || name === 'moveFrom' ? 'tracked-change' : 'container' };
    }
    if (offset < prefixLength || collected.slice(0, prefixLength) !== expectedPrefix) return { reason: 'text-mismatch' };
    return { segments };
}

/** Removes the first `prefixLength` characters: a tracked `w:del` when tracking, a direct removal otherwise. */
function stripMarkerPrefix(xmlDoc, segments, prefixLength, { author, generateRedlines, revisionIdAllocator }) {
    for (const segment of segments) {
        const { run, start, end, pieces } = segment;
        const localEnd = Math.min(prefixLength, end) - start;
        if (localEnd <= 0) continue;
        const parent = run.parentNode;
        const rPr = getDirectWordChild(run, 'rPr');
        if (generateRedlines) {
            const deletedRun = createRunFromPieces(xmlDoc, sliceRunPieces(xmlDoc, pieces, 0, localEnd, true), rPr);
            parent.insertBefore(createTrackChange(xmlDoc, 'del', deletedRun, author), run);
        }
        if (localEnd < end - start) {
            const keptRun = createRunFromPieces(xmlDoc, sliceRunPieces(xmlDoc, pieces, localEnd, end - start, false), rPr);
            // The deleted copy keeps any rPrChange id; the kept copy needs its own.
            if (generateRedlines) refreshRunPropertyChangeIds(keptRun, revisionIdAllocator);
            parent.insertBefore(keptRun, run);
        }
        parent.removeChild(run);
    }
}

// ---------------------------------------------------------------------------------------------------------
// Numbering
// ---------------------------------------------------------------------------------------------------------

function numIdIsLive(xmlDoc, numId) {
    if (numId == null) return false;
    return getCurrentNumIdElements(xmlDoc.documentElement).some(element => readVal(element) === String(numId));
}

function sequenceMapOf(runtimeContext) {
    if (!runtimeContext || typeof runtimeContext !== 'object') return null;
    if (!runtimeContext.listFallbackSequenceState || typeof runtimeContext.listFallbackSequenceState !== 'object') {
        runtimeContext.listFallbackSequenceState = { explicitByNumberingKey: new Map() };
    }
    const state = runtimeContext.listFallbackSequenceState;
    if (!(state.explicitByNumberingKey instanceof Map)) state.explicitByNumberingKey = new Map();
    return state.explicitByNumberingKey;
}

export function numberingIdStateOf(runtimeContext, options) {
    if (runtimeContext?.numberingIdState) return runtimeContext.numberingIdState;
    const state = createDynamicNumberingIdState(runtimeContext?.numberingXml || options?.numberingXml || undefined);
    if (runtimeContext && typeof runtimeContext === 'object') runtimeContext.numberingIdState = state;
    return state;
}

// ---------------------------------------------------------------------------------------------------------
// Plan + apply
// ---------------------------------------------------------------------------------------------------------

/**
 * Converts plain paragraphs (one target or a same-container range) to list items.
 *
 * @param {Document} xmlDoc
 * @param {Element[]} paragraphs - the resolved target range, in document order
 * @param {{ format: string|null, kind: string|null, level: number, strip: 'marker'|'none', start: number|null }} request
 * @param {{
 *   author: string, generateRedlines?: boolean, revisionIdAllocator?: object|null,
 *   runtimeContext?: object|null, stylesXml?: string|null, numberingXml?: string|null,
 *   paragraphIndexOf?: (paragraph: Element) => number
 * }} context
 * @returns {{ ok: true, format: string, items: object[], skipped: object[], numberingXml: string|null, allocatedNumIds: number[] }
 *   | { ok: false, error: object }}
 */
export function convertPlainParagraphsToList(xmlDoc, paragraphs, request, context = {}) {
    const generateRedlines = context.generateRedlines !== false;
    const author = context.author;
    const indexOf = typeof context.paragraphIndexOf === 'function' ? context.paragraphIndexOf : () => null;
    const stylesDoc = parseStyles(context.stylesXml);
    const strip = request.strip === 'none' ? 'none' : 'marker';

    // 1. Classify every paragraph without mutating anything.
    const entries = [];
    const skipped = [];
    for (const paragraph of paragraphs) {
        const paragraphIndex = indexOf(paragraph);
        const text = extractCanonicalParagraphText(paragraph, { revisionView: 'accepted' });
        const listInfo = getParagraphListInfo(paragraph);
        if (listInfo) {
            return refusal(LIST_FORMAT_TARGET_IS_LIST,
                `Paragraph ${paragraphIndex ?? ''} is already a list item (numId ${listInfo.numId}, level ${listInfo.ilvl}); convert-plain converts plain paragraphs only.`.replace('  ', ' '),
                { paragraphIndex, numId: listInfo.numId, level: listInfo.ilvl, targetExcerpt: excerpt(text) },
                { action: 'use_paragraph_numbering', hint: 'Narrow the range to plain paragraphs. To move a list item to another existing list or level, use paragraph-format with properties.numbering.' });
        }
        const current = getCurrentParagraphProperties(paragraph);
        const suppressed = current.numPr?.numId === '0';
        const styled = suppressed ? null : styleNumbering(stylesDoc, current.pStyle);
        if (styled) {
            return refusal(LIST_FORMAT_STYLE_NUMBERED,
                `Paragraph ${paragraphIndex ?? ''} is numbered by its style "${styled.styleId}" (numId ${styled.numId}); it is already a list item through the style.`.replace('  ', ' '),
                { paragraphIndex, styleId: styled.styleId, numId: styled.numId, targetExcerpt: excerpt(text) },
                { action: 'convert_in_host', hint: 'Style-linked numbering is not rebound by list-format-v1. Change the paragraph style or the style\'s numbering in the host.' });
        }
        if (!text.trim()) {
            skipped.push({ paragraph, paragraphIndex, reason: 'empty' });
            continue;
        }
        if (strip === 'none') {
            entries.push({ paragraph, paragraphIndex, text, parsed: null });
            continue;
        }
        const parsed = parseTypedListMarker(text);
        if (parsed?.unsupported) {
            const detail = {
                multilevel: 'multi-level markers such as "2.1." need a multi-level list definition',
                parenthesized: 'parenthesized markers such as "(a)" are not supported',
                zero: 'a marker value of 0 cannot start a Word list'
            }[parsed.unsupported];
            return refusal(LIST_MARKER_UNSUPPORTED,
                `The typed marker "${parsed.marker}" is not supported by list-format-v1: ${detail}.`,
                { reason: parsed.unsupported, marker: parsed.marker, paragraphIndex, targetExcerpt: excerpt(text) },
                { action: 'convert_in_host', hint: 'Convert this paragraph in the host, or leave the marker as literal text.' });
        }
        if (!parsed) {
            // A non-list paragraph between converted headers stays body text; the list continues across it.
            skipped.push({ paragraph, paragraphIndex, reason: 'no-marker' });
            continue;
        }
        entries.push({ paragraph, paragraphIndex, text, parsed });
    }

    if (entries.length === 0) {
        return refusal(LIST_MARKER_NOT_FOUND,
            strip === 'marker'
                ? 'No target paragraph starts with a typed list marker (for example "A. ", "1. ", "iv. " or "- "), so there is nothing to strip.'
                : 'The target range has no non-empty paragraph to convert.',
            { targetExcerpt: excerpt(skipped[0] ? extractCanonicalParagraphText(skipped[0].paragraph) : '') },
            { action: 'change_request', hint: 'Target the paragraphs that carry the typed markers, or pass strip: "none" with list.format to bind paragraphs whose text has no marker.', field: 'list.strip', allowedValue: 'none' });
    }

    // 2. One format for the whole run.
    let format = request.format;
    if (strip === 'marker') {
        const resolved = resolveMarkerFormat(entries.map(entry => entry.parsed), request.format);
        if (resolved.error) {
            const { reason } = resolved.error;
            if (reason === 'ambiguous') {
                return refusal(LIST_MARKER_AMBIGUOUS,
                    `The typed marker "${resolved.error.marker}" reads as ${resolved.error.formats.join(' or ')}; list-format will not guess.`,
                    { marker: resolved.error.marker, formats: resolved.error.formats },
                    { action: 'set_list_format', hint: `Pass list.format (${resolved.error.formats.join(' or ')}).`, field: 'list.format', allowedValues: resolved.error.formats });
            }
            const message = reason === 'format-mismatch'
                ? `The typed marker "${resolved.error.marker}" is not a ${resolved.error.requestedFormat} marker.`
                : `The target range mixes list markers (${[...new Set(resolved.error.markers)].join(', ')}); one convert-plain operation converts one list format.`;
            return refusal(LIST_MARKER_UNSUPPORTED, message,
                { reason, ...(resolved.error.marker ? { marker: resolved.error.marker } : {}), ...(resolved.error.markers ? { markers: resolved.error.markers } : {}) },
                { action: 'change_request', hint: reason === 'format-mismatch' ? 'Omit list.format to use the typed marker, or correct it.' : 'Split the range into one operation per marker format.' });
        }
        format = resolved.format;
    }
    if (request.kind && (request.kind === 'bullet') !== (format === 'bullet')) {
        return refusal(LIST_MARKER_UNSUPPORTED, `The typed markers are ${format} markers, not a ${request.kind} list.`,
            { reason: 'format-mismatch', requestedKind: request.kind },
            { action: 'change_request', hint: 'Omit list.kind, or target paragraphs whose markers match it.' });
    }
    const terminator = strip === 'marker' ? entries[0].parsed.terminator : '.';
    const bulletChar = strip === 'marker' ? (entries[0].parsed.bulletChar || null) : null;
    const level = request.level || 0;

    // 3. Paragraph-property snapshots and marker runs.
    for (const entry of entries) {
        const pendingChange = getCurrentChild(getDirectWordChild(entry.paragraph, 'pPr'), 'pPrChange');
        const pendingAuthor = pendingChange?.getAttribute('w:author') ?? pendingChange?.getAttribute('author') ?? '';
        if (generateRedlines && pendingChange && pendingAuthor !== author) {
            return refusal(PARAGRAPH_FORMAT_REVISION_CONFLICT,
                `Paragraph ${entry.paragraphIndex ?? ''} already has a pending paragraph-property change by "${pendingAuthor || 'unattributed'}". Overwriting its snapshot would change what Reject All restores.`.replace('  ', ' '),
                { pendingAuthor, paragraphIndex: entry.paragraphIndex },
                { action: 'resolve_pending_revision', hint: 'Ask the user to accept or reject the existing paragraph formatting change first; this tool does not resolve revisions by other reviewers.' });
        }
        entry.keepSnapshot = !!pendingChange;
        if (strip === 'marker') {
            const located = locateMarkerRuns(entry.paragraph, entry.parsed.prefixLength, entry.text.slice(0, entry.parsed.prefixLength));
            if (located.reason) {
                return refusal(LIST_MARKER_UNSUPPORTED,
                    `The typed marker "${entry.parsed.marker}" is not plain run text (${located.reason}), so it cannot be stripped as one tracked deletion.`,
                    { reason: 'marker-shape', shape: located.reason, marker: entry.parsed.marker, paragraphIndex: entry.paragraphIndex, targetExcerpt: excerpt(entry.text) },
                    { action: 'convert_in_host', hint: 'The marker sits inside a hyperlink, field, content control or pending revision. Resolve that markup first or convert the paragraph in the host.' });
            }
            entry.segments = located.segments;
        }
    }

    // 4. Values and list groups. A value that does not continue the previous one starts a new list, so every
    //    label Word renders after Accept All is the label the paragraph showed as text.
    const sequence = sequenceMapOf(context.runtimeContext);
    const sequenceKey = listFormatSequenceKey({ format, level, terminator, bulletChar });
    const previous = sequence?.get(sequenceKey) || null;
    const previousLive = previous && numIdIsLive(xmlDoc, previous.numId);
    let nextValue = strip === 'none'
        ? (request.start ?? (previousLive && Number.isInteger(previous.nextStartAt) ? previous.nextStartAt : 1))
        : null;
    for (const entry of entries) {
        if (format === 'bullet') entry.value = null;
        else if (strip === 'marker') entry.value = markerValue(entry.parsed, format);
        else entry.value = nextValue++;
    }

    const numberingIdState = numberingIdStateOf(context.runtimeContext, context);
    const spec = listLevelSpec({ format, level, terminator, bulletChar });
    const parts = [];
    const allocatedNumIds = [];
    let groupNumId = null;
    let lastValue = null;
    for (const [index, entry] of entries.entries()) {
        const continues = index > 0 && (format === 'bullet' || entry.value === lastValue + 1);
        if (!continues) {
            const reusePrevious = index === 0 && previousLive
                && (format === 'bullet' || previous.nextStartAt === entry.value);
            if (reusePrevious) {
                groupNumId = String(previous.numId);
            } else {
                const allocated = allocateNumberingDefinition({
                    numberingIdState,
                    levels: [spec],
                    startOverrides: entry.value != null && entry.value !== 1 ? { [level]: entry.value } : undefined
                });
                if (!allocated) {
                    return refusal('NUMBERING_ID_EXHAUSTED', 'No numbering ids are available for a new list definition.', {},
                        { action: 'report_library_failure' });
                }
                groupNumId = String(allocated.numId);
                allocatedNumIds.push(allocated.numId);
                parts.push(allocated.numberingXml);
            }
        }
        entry.numId = groupNumId;
        lastValue = entry.value;
    }

    // 5. Mutate.
    for (const entry of entries) {
        let pPr = getDirectWordChild(entry.paragraph, 'pPr');
        if (!pPr) {
            pPr = createWordElement(xmlDoc, 'w:pPr');
            entry.paragraph.insertBefore(pPr, entry.paragraph.firstChild);
        }
        // A same-author pending snapshot already records the earliest "before" state; keep it.
        if (generateRedlines && !entry.keepSnapshot) snapshotAndAttachPPrChange(xmlDoc, pPr, author);
        applyParagraphPropertiesToPPr(xmlDoc, pPr, { numbering: { numId: entry.numId, level } });
        if (strip === 'marker') {
            stripMarkerPrefix(xmlDoc, entry.segments, entry.parsed.prefixLength, {
                author,
                generateRedlines,
                revisionIdAllocator: context.revisionIdAllocator || null
            });
        }
    }

    const last = entries[entries.length - 1];
    sequence?.set(sequenceKey, { numId: last.numId, nextStartAt: last.value != null ? last.value + 1 : null });

    return {
        ok: true,
        format,
        level,
        strip,
        items: entries.map(entry => ({
            paragraph: entry.paragraph,
            paragraphIndex: entry.paragraphIndex,
            marker: entry.parsed?.marker ?? null,
            value: entry.value,
            numId: entry.numId
        })),
        skipped,
        numberingXml: parts.length > 0 ? mergeNumberingParts('', parts) : null,
        allocatedNumIds
    };
}
