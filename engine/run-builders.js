/**
 * OOXML run/track-change builders.
 *
 * This module centralizes creation of `w:r`, `w:ins`, `w:del`, `w:rPrChange`,
 * and `w:pPrChange` elements used by surgical and reconstruction modes.
 */

import { extractFormatFromRPr, RPR_SCHEMA_ORDER } from './rpr-helpers.js';
import { createRevisionMetadata } from '../core/types.js';
import { getFirstElementByTag } from '../core/xml-query.js';
import { createWordElement } from '../core/word-xml.js';
import { clonePropertiesWithoutRevisionHistory, refreshRunPropertyChangeIds } from '../core/revision-cloning.js';

/**
 * Creates an insertion/deletion wrapper.
 *
 * @param {Document} xmlDoc - XML document
 * @param {'ins'|'del'} type - Wrapper type
 * @param {Element|null} run - Optional run to append
 * @param {string} author - Change author
 * @param {Object|null} [revisionMetadata=null] - Optional explicit revision metadata
 * @returns {Element}
 */
export function createTrackChange(xmlDoc, type, run, author, revisionMetadata = null) {
    const wrapper = createWordElement(xmlDoc, type === 'ins' ? 'w:ins' : 'w:del');
    const metadata = revisionMetadata || createRevisionMetadata(author, xmlDoc, type === 'ins' ? 'ins' : 'del');
    wrapper.setAttribute('w:id', String(metadata.id));
    wrapper.setAttribute('w:author', metadata.author);
    wrapper.setAttribute('w:date', metadata.date);
    if (run) {
        wrapper.appendChild(run);
    }
    return wrapper;
}

/**
 * Creates a text run with optional formatting.
 *
 * @param {Document} xmlDoc - XML document
 * @param {string} text - Text content
 * @param {Element|null} rPr - Run properties
 * @param {boolean} isDelete - Use `w:delText` instead of `w:t`
 * @returns {Element}
 */
export function createTextRun(xmlDoc, text, rPr, isDelete) {
    const run = createWordElement(xmlDoc, 'w:r');
    if (rPr) run.appendChild(rPr.cloneNode(true));

    if (!isDelete) {
        appendVisibleTextPieces(xmlDoc, run, text);
        return run;
    }

    const textEl = createWordElement(xmlDoc, isDelete ? 'w:delText' : 'w:t');
    textEl.setAttribute('xml:space', 'preserve');
    textEl.textContent = text;
    run.appendChild(textEl);

    return run;
}

/**
 * Creates an array of runs with formatting applied from hints.
 *
 * @param {Document} xmlDoc - XML document
 * @param {string} text - Text to split and format
 * @param {Element|null} baseRPr - Base run properties
 * @param {Array} formatHints - Formatting hints
 * @param {number} baseOffset - Absolute base offset
 * @param {string} [author] - Change author
 * @param {boolean} [generateRedlines] - Whether to create rPrChange
 * @returns {Element[]}
 */
export function createFormattedRuns(xmlDoc, text, baseRPr, formatHints, baseOffset, author, generateRedlines) {
    if (!text) return [];

    const breaks = new Set([0, text.length]);
    for (const hint of formatHints) {
        const localStart = Math.max(0, hint.start - baseOffset);
        const localEnd = Math.min(text.length, hint.end - baseOffset);
        if (localStart >= 0 && localStart < text.length) breaks.add(localStart);
        if (localEnd > 0 && localEnd <= text.length) breaks.add(localEnd);
    }

    const sortedBreaks = Array.from(breaks).sort((a, b) => a - b);
    const runs = [];

    for (let i = 0; i < sortedBreaks.length - 1; i++) {
        const start = sortedBreaks[i];
        const end = sortedBreaks[i + 1];
        const segment = text.slice(start, end);
        if (!segment) continue;

        const segmentBaseOffset = baseOffset + start;
        const segmentEndOffset = baseOffset + end;

        const applicableHints = formatHints.filter(h =>
            h.start <= segmentBaseOffset && h.end >= segmentEndOffset
        );

        const combinedFormat = { ...extractFormatFromRPr(baseRPr) };
        applicableHints.forEach(h => {
            if (h.format) Object.assign(combinedFormat, h.format);
        });

        // During a text edit, missing Markdown is not an instruction to clear
        // the source run's formatting. Only synchronize explicitly hinted spans.
        const formattedRPr = applicableHints.length > 0
            ? injectFormattingToRPr(xmlDoc, baseRPr, combinedFormat, author, generateRedlines)
            : baseRPr?.cloneNode(true) || null;
        runs.push(createTextRunWithRPrElement(xmlDoc, segment, formattedRPr, false));
    }

    // Pieces of one source run each copy its pending w:rPrChange; only the first keeps the id.
    const seenChangeIds = new Set();
    for (const run of runs) {
        const change = getFirstElementByTag(getFirstElementByTag(run, 'w:rPr'), 'w:rPrChange');
        if (!change) continue;
        const id = change.getAttribute('w:id');
        if (seenChangeIds.has(id)) {
            refreshRunPropertyChangeIds(change);
            seenChangeIds.add(change.getAttribute('w:id'));
        } else {
            seenChangeIds.add(id);
        }
    }

    return runs;
}

/**
 * Creates a text run with an existing rPr element (no clone).
 *
 * @param {Document} xmlDoc - XML document
 * @param {string} text - Text content
 * @param {Element|null} rPrElement - Run properties element
 * @param {boolean} isDelete - Use `w:delText` instead of `w:t`
 * @returns {Element}
 */
export function createTextRunWithRPrElement(xmlDoc, text, rPrElement, isDelete) {
    const run = createWordElement(xmlDoc, 'w:r');
    if (rPrElement) run.appendChild(rPrElement);

    if (!isDelete) {
        appendVisibleTextPieces(xmlDoc, run, text);
        return run;
    }

    const textEl = createWordElement(xmlDoc, isDelete ? 'w:delText' : 'w:t');
    textEl.setAttribute('xml:space', 'preserve');
    textEl.textContent = text;
    run.appendChild(textEl);

    return run;
}

function appendVisibleTextPieces(xmlDoc, run, text) {
    const source = String(text || '');
    const parts = source.split(/(\t|\n|\u2011)/);

    for (const part of parts) {
        if (!part) continue;
        if (part === '\t') {
            run.appendChild(createWordElement(xmlDoc, 'w:tab'));
            continue;
        }
        if (part === '\n') {
            run.appendChild(createWordElement(xmlDoc, 'w:br'));
            continue;
        }
        if (part === '\u2011') {
            run.appendChild(createWordElement(xmlDoc, 'w:noBreakHyphen'));
            continue;
        }

        const textEl = createWordElement(xmlDoc, 'w:t');
        if (/^\s|\s$/.test(part)) {
            textEl.setAttribute('xml:space', 'preserve');
        }
        textEl.textContent = part;
        run.appendChild(textEl);
    }
}

/**
 * Creates a new rPr synchronized to the requested core formatting flags.
 *
 * @param {Document} xmlDoc - XML document
 * @param {Element|null} baseRPr - Base run properties
 * @param {Object|null} format - Format flags
 * @param {string} [author] - Change author
 * @param {boolean} [generateRedlines] - Whether to create rPrChange
 * @returns {Element}
 */
export function injectFormattingToRPr(xmlDoc, baseRPr, format, author, generateRedlines) {
    const rPr = createWordElement(xmlDoc, 'w:rPr');

    if (baseRPr) {
        Array.from(baseRPr.childNodes).forEach(child => {
            if (!['w:b', 'w:bCs', 'w:i', 'w:iCs', 'w:u', 'w:strike', 'w:rPrChange'].includes(child.nodeName)) {
                rPr.appendChild(child.cloneNode(true));
            }
        });
    }

    const activeFormat = format || { bold: false, italic: false, underline: false, strikethrough: false };

    if (author && generateRedlines) {
        createRPrChange(xmlDoc, rPr, author, baseRPr);
    }

    const syncElement = (tagName, isOn, valOn = null, valOff = '0') => {
        const el = createWordElement(xmlDoc, tagName);
        if (isOn) {
            if (valOn) el.setAttribute('w:val', valOn);
        } else if (valOff) {
            el.setAttribute('w:val', valOff);
        }

        const myIndex = RPR_SCHEMA_ORDER.indexOf(tagName);
        const myPriority = myIndex === -1 ? 999 : myIndex;

        let inserted = false;
        for (const child of Array.from(rPr.childNodes)) {
            if (child.nodeType !== 1) continue;
            const childIndex = RPR_SCHEMA_ORDER.indexOf(child.nodeName);
            const childPriority = childIndex === -1 ? 999 : childIndex;
            if (childPriority > myPriority) {
                rPr.insertBefore(el, child);
                inserted = true;
                break;
            }
        }
        if (!inserted) rPr.appendChild(el);
    };

    syncElement('w:b', !!activeFormat.bold, '1', '0');
    syncElement('w:bCs', !!activeFormat.bold, '1', '0');
    syncElement('w:i', !!activeFormat.italic, '1', '0');
    syncElement('w:iCs', !!activeFormat.italic, '1', '0');
    syncElement('w:u', !!activeFormat.underline, 'single', 'none');
    syncElement('w:strike', !!activeFormat.strikethrough, '1', '0');

    return rPr;
}

/**
 * Creates and attaches a `w:rPrChange` snapshot.
 *
 * @param {Document} xmlDoc - XML document
 * @param {Element} rPr - Target run properties
 * @param {string} author - Change author
 * @param {string} dateStr - ISO date string
 * @param {Element} [sourceNode] - Optional source for previous state snapshot
 * @returns {Element}
 */
export function snapshotAndAttachRPrChange(xmlDoc, rPr, author, dateStr, sourceNode) {
    const source = sourceNode || rPr;
    const existing = getFirstElementByTag(rPr, 'w:rPrChange');
    const pending = pendingSameAuthorRPrChange(source, author);
    let rPrChange;
    if (pending) {
        // The run's formatting is already a pending change by this author. Its snapshot is the
        // state Reject All must restore, so it is kept; a snapshot of the current (already
        // changed) properties would make Reject All keep the earlier change.
        rPrChange = pending.cloneNode(true);
        if (reusedRPrChanges.has(pending)) {
            // The source run was split: only the first piece keeps the revision id.
            const metadata = createRevisionMetadata(author, xmlDoc, 'rPrChange');
            rPrChange.setAttribute('w:id', String(metadata.id));
        } else {
            reusedRPrChanges.add(pending);
        }
    } else {
        rPrChange = createWordElement(xmlDoc, 'w:rPrChange');
        const metadata = createRevisionMetadata(author, xmlDoc, 'rPrChange');
        rPrChange.setAttribute('w:id', String(metadata.id));
        rPrChange.setAttribute('w:author', metadata.author);
        rPrChange.setAttribute('w:date', dateStr || metadata.date);
        rPrChange.appendChild(clonePropertiesWithoutRevisionHistory(source)
            || createWordElement(xmlDoc, 'w:rPr'));
    }

    if (existing) {
        rPr.removeChild(existing);
    }

    rPr.appendChild(rPrChange);
    return rPrChange;
}

const reusedRPrChanges = new WeakSet();

function pendingSameAuthorRPrChange(source, author) {
    const change = Array.from(source?.childNodes || []).find(
        child => child.nodeType === 1 && (child.nodeName === 'w:rPrChange' || child.localName === 'rPrChange')
    );
    if (!change) return null;
    const changeAuthor = change.getAttribute('w:author') || '';
    return changeAuthor.trim().toLowerCase() === String(author || '').trim().toLowerCase() ? change : null;
}

/**
 * Creates `w:rPrChange` for track formatting changes.
 *
 * @param {Document} xmlDoc - XML document
 * @param {Element} rPr - Run properties target
 * @param {string} author - Change author
 * @param {Element} [previousRPrArg] - Optional explicit previous-state source
 * @returns {void}
 */
function createRPrChange(xmlDoc, rPr, author, previousRPrArg) {
    snapshotAndAttachRPrChange(xmlDoc, rPr, author, null, previousRPrArg || rPr);
}

/**
 * Creates and attaches a `w:pPrChange` snapshot to paragraph properties.
 *
 * @param {Document} xmlDoc - XML document
 * @param {Element} pPr - Target paragraph properties
 * @param {string} author - Change author
 * @param {string} [dateStr] - ISO date string
 * @param {Element} [sourceNode] - Optional source for previous state snapshot
 * @returns {Element}
 */
export function snapshotAndAttachPPrChange(xmlDoc, pPr, author, dateStr, sourceNode) {
    const pPrChange = createWordElement(xmlDoc, 'w:pPrChange');
    const metadata = createRevisionMetadata(author, xmlDoc, 'pPrChange');
    pPrChange.setAttribute('w:id', String(metadata.id));
    pPrChange.setAttribute('w:author', metadata.author);
    pPrChange.setAttribute('w:date', dateStr || metadata.date);

    const source = sourceNode || pPr;
    const previousPPr = clonePropertiesWithoutRevisionHistory(source)
        || createWordElement(xmlDoc, 'w:pPr');

    pPrChange.appendChild(previousPPr);

    const existing = Array.from(pPr.childNodes).find(
        c => c.nodeType === 1 && (c.nodeName === 'w:pPrChange' || c.localName === 'pPrChange')
    );
    if (existing) {
        pPr.removeChild(existing);
    }

    pPr.appendChild(pPrChange);
    return pPrChange;
}
