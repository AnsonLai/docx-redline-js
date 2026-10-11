/**
 * `list-format` operation: contract shape, typed-marker grammar and numbering level specs (canonical list
 * operations plan, WP-4). Pure helpers only; the DOM mutation lives in `document-operation-mutations.js`.
 *
 * `convert-plain` binds plain paragraphs to a NEW list definition. With `strip: 'marker'` (default) every
 * paragraph's typed marker (`A.`, `1)`, `iv.`, `-`) states the list format and the item's value; the marker is
 * removed with a tracked deletion and the binding is recorded as a tracked `w:pPrChange`. With `strip: 'none'`
 * the text is left alone and `list.format` / `list.start` state the format and first value.
 *
 * The grammar is deliberately narrow: anything it cannot classify exactly is refused with a recovery-enveloped
 * code rather than guessed (single roman-looking letters, mixed families, multi-level `2.1.`, `(a)`).
 */

export const LIST_FORMAT_MODES = Object.freeze(['convert-plain']);
export const LIST_FORMAT_FORMATS = Object.freeze(['decimal', 'upperLetter', 'lowerLetter', 'upperRoman', 'lowerRoman', 'bullet']);
export const LIST_FORMAT_STRIP_VALUES = Object.freeze(['marker', 'none']);
const LIST_KEYS = Object.freeze(['mode', 'kind', 'format', 'start', 'level', 'strip']);
/** Keys of the planned contract that later work packages implement (change-format, join, detach). */
const LATER_KEYS = Object.freeze(['lvlText', 'continueFrom', 'scope', 'materializeMarker']);

export const LIST_MARKER_NOT_FOUND = 'LIST_MARKER_NOT_FOUND';
export const LIST_MARKER_AMBIGUOUS = 'LIST_MARKER_AMBIGUOUS';
export const LIST_MARKER_UNSUPPORTED = 'LIST_MARKER_UNSUPPORTED';
export const LIST_FORMAT_TARGET_IS_LIST = 'LIST_FORMAT_TARGET_IS_LIST';
export const LIST_FORMAT_STYLE_NUMBERED = 'LIST_FORMAT_STYLE_NUMBERED';

const BULLET_CHARS = '•\\-*–·▪◦';
const SPACE = '[ \\t\\u00a0]';
const NUMBERED_MARKER = new RegExp(`^(${SPACE}*)(?:(\\d{1,3})|([A-Za-z]{1,7}))([.)])(${SPACE}+)(?=\\S)`);
const BULLET_MARKER = new RegExp(`^(${SPACE}*)([${BULLET_CHARS}])(${SPACE}+)(?=\\S)`);
const MULTILEVEL_MARKER = new RegExp(`^${SPACE}*\\d+(?:\\.\\d+)+\\.?${SPACE}+\\S`);
const PARENTHESIZED_MARKER = new RegExp(`^${SPACE}*\\((?:\\d{1,3}|[A-Za-z]{1,7})\\)${SPACE}+\\S`);
const ROMAN = /^M{0,3}(CM|CD|D?C{0,3})(XC|XL|L?X{0,3})(IX|IV|V?I{0,3})$/;
const ROMAN_VALUES = { I: 1, V: 5, X: 10, L: 50, C: 100, D: 500, M: 1000 };

function isRecord(value) {
    return !!value && typeof value === 'object' && !Array.isArray(value);
}

/** Validates `operation.list` for `list-format`. Returns an error message (with `field`) or null. */
export function validateListFormatShape(list) {
    if (!isRecord(list)) return { message: 'list-format operations require a "list" object.', field: 'list' };
    const later = Object.keys(list).filter(key => LATER_KEYS.includes(key));
    if (later.length > 0) {
        return { message: `list.${later[0]} is not supported by list-format-v1 (convert-plain only).`, field: `list.${later[0]}` };
    }
    const unknown = Object.keys(list).filter(key => !LIST_KEYS.includes(key));
    if (unknown.length > 0) return { message: `list has unsupported ${unknown.length === 1 ? 'key' : 'keys'}: ${unknown.join(', ')}.`, field: `list.${unknown[0]}` };
    if (list.mode !== 'convert-plain') {
        return { message: 'list.mode must be "convert-plain" (change-format, join and detach are not available in list-format-v1).', field: 'list.mode' };
    }
    if (list.format != null && !LIST_FORMAT_FORMATS.includes(list.format)) {
        return { message: `list.format must be one of: ${LIST_FORMAT_FORMATS.join(', ')}.`, field: 'list.format' };
    }
    if (list.kind != null && list.kind !== 'bullet' && list.kind !== 'numbered') {
        return { message: 'list.kind must be "bullet" or "numbered".', field: 'list.kind' };
    }
    if (list.kind != null && list.format != null && (list.kind === 'bullet') !== (list.format === 'bullet')) {
        return { message: `list.kind "${list.kind}" contradicts list.format "${list.format}".`, field: 'list.kind' };
    }
    if (list.level != null && !(Number.isInteger(list.level) && list.level >= 0 && list.level <= 8)) {
        return { message: 'list.level must be an integer from 0 to 8.', field: 'list.level' };
    }
    if (list.strip != null && !LIST_FORMAT_STRIP_VALUES.includes(list.strip)) {
        return { message: 'list.strip must be "marker" or "none".', field: 'list.strip' };
    }
    const strip = list.strip || 'marker';
    if (list.start != null) {
        if (strip === 'marker') {
            return { message: 'list.start is derived from the typed marker when strip is "marker"; omit it, or use strip: "none".', field: 'list.start' };
        }
        if (!(Number.isInteger(list.start) && list.start >= 1 && list.start <= 32767)) {
            return { message: 'list.start must be an integer from 1 to 32767.', field: 'list.start' };
        }
    }
    if (strip === 'none' && list.format == null && list.kind !== 'bullet') {
        return { message: 'strip: "none" leaves the text alone, so list.format (or kind: "bullet") must state the list format.', field: 'list.format' };
    }
    return null;
}

/** Normalized `list` (defaults applied); call after `validateListFormatShape`. */
export function normalizeListFormatRequest(list) {
    const format = list.format || (list.kind === 'bullet' ? 'bullet' : null);
    return {
        mode: list.mode,
        format,
        kind: format ? (format === 'bullet' ? 'bullet' : 'numbered') : (list.kind || null),
        level: Number.isInteger(list.level) ? list.level : 0,
        strip: list.strip || 'marker',
        start: Number.isInteger(list.start) ? list.start : null
    };
}

function romanValue(letters) {
    const upper = letters.toUpperCase();
    if (!upper || !ROMAN.test(upper)) return null;
    let total = 0;
    for (let i = 0; i < upper.length; i += 1) {
        const value = ROMAN_VALUES[upper[i]];
        const next = ROMAN_VALUES[upper[i + 1]] || 0;
        total += value < next ? -value : value;
    }
    return total;
}

/**
 * Parses the typed marker at the start of a paragraph's text.
 *
 * Returns `null` when the text has no marker, `{ unsupported: reason, marker }` for a marker shape list-format-v1
 * refuses, or `{ marker, terminator, prefixLength, candidates: [{ format, value }], bulletChar? }`. `prefixLength`
 * covers leading spaces, the marker and the whitespace after it: exactly the text a strip removes. A single
 * letter that is also a roman numeral (`I.`, `C.`) has two candidates; the caller disambiguates.
 *
 * @param {string} text
 */
export function parseTypedListMarker(text) {
    const value = String(text ?? '');
    if (MULTILEVEL_MARKER.test(value)) {
        return { unsupported: 'multilevel', marker: value.trim().split(/\s/)[0] };
    }
    if (PARENTHESIZED_MARKER.test(value)) {
        return { unsupported: 'parenthesized', marker: value.trim().split(/\s/)[0] };
    }
    const bullet = BULLET_MARKER.exec(value);
    if (bullet) {
        return {
            marker: bullet[2],
            terminator: '',
            bulletChar: bullet[2],
            prefixLength: bullet[0].length,
            candidates: [{ format: 'bullet', value: null }]
        };
    }
    const numbered = NUMBERED_MARKER.exec(value);
    if (!numbered) return null;
    const [whole, , digits, letters, terminator] = numbered;
    const marker = `${digits ?? letters}${terminator}`;
    const base = { marker, terminator, prefixLength: whole.length };
    if (digits != null) {
        const number = Number.parseInt(digits, 10);
        if (number < 1) return { unsupported: 'zero', marker };
        return { ...base, candidates: [{ format: 'decimal', value: number }] };
    }
    const upper = letters === letters.toUpperCase();
    const lower = letters === letters.toLowerCase();
    if (!upper && !lower) return null;
    const candidates = [];
    if (letters.length === 1) {
        candidates.push({ format: upper ? 'upperLetter' : 'lowerLetter', value: letters.toLowerCase().charCodeAt(0) - 96 });
    }
    const roman = romanValue(letters);
    if (roman != null) candidates.push({ format: upper ? 'upperRoman' : 'lowerRoman', value: roman });
    if (candidates.length === 0) return null;
    return { ...base, candidates };
}

/**
 * Chooses one list format for a run of parsed markers.
 *
 * With `requestedFormat`, every marker must offer it. Otherwise the formats every marker offers are
 * intersected; several remaining formats (a lone `I.`, or `C.` / `D.`) are resolved only when exactly one of
 * them makes the values a consecutive sequence (`C., D.` are letters 3, 4; roman 100, 500 are not). Returns
 * `{ format }` or `{ error: { reason, ... } }`.
 */
export function resolveMarkerFormat(markers, requestedFormat = null) {
    if (markers.length === 0) return { error: { reason: 'none' } };
    const terminators = new Set(markers.map(item => item.terminator));
    const bulletChars = new Set(markers.filter(item => item.bulletChar).map(item => item.bulletChar));
    if (terminators.size > 1 || bulletChars.size > 1) {
        return { error: { reason: 'mixed', markers: markers.map(item => item.marker) } };
    }
    if (requestedFormat) {
        const mismatch = markers.find(item => !item.candidates.some(candidate => candidate.format === requestedFormat));
        if (mismatch) return { error: { reason: 'format-mismatch', marker: mismatch.marker, requestedFormat } };
        return { format: requestedFormat };
    }
    let shared = markers[0].candidates.map(candidate => candidate.format);
    for (const item of markers.slice(1)) {
        const formats = item.candidates.map(candidate => candidate.format);
        shared = shared.filter(format => formats.includes(format));
    }
    if (shared.length === 0) return { error: { reason: 'mixed', markers: markers.map(item => item.marker) } };
    if (shared.length === 1) return { format: shared[0] };
    const consecutive = shared.filter(format => {
        if (markers.length < 2) return false;
        const values = markers.map(item => item.candidates.find(candidate => candidate.format === format).value);
        return values.every((value, index) => index === 0 || value === values[index - 1] + 1);
    });
    if (consecutive.length === 1) return { format: consecutive[0] };
    return { error: { reason: 'ambiguous', marker: markers[0].marker, formats: shared } };
}

/** The value a marker has under a resolved format (null for bullets). */
export function markerValue(parsed, format) {
    return parsed.candidates.find(candidate => candidate.format === format)?.value ?? null;
}

/**
 * The `w:lvl` redefinition for `allocateNumberingDefinition`. The level text keeps the typed terminator, so
 * the label Word renders after Accept All is the text the marker showed (`A.` stays `A.`, `1)` stays `1)`).
 */
export function listLevelSpec({ format, level, terminator = '.', bulletChar = null }) {
    if (format === 'bullet') {
        return { ilvl: level, numFmt: 'bullet', lvlText: bulletChar === '-' || bulletChar === '*' || !bulletChar ? '•' : bulletChar };
    }
    return { ilvl: level, numFmt: format, lvlText: `%${level + 1}${terminator || '.'}` };
}

/** Sequence-state key for continuing a converted list across operations of one batch. */
export function listFormatSequenceKey({ format, level, terminator = '', bulletChar = null }) {
    return `list-format:${format}:${level}:${format === 'bullet' ? (bulletChar || '') : (terminator || '.')}`;
}
