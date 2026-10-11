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
 *
 * `change-format` and `join` (WP-5, capability `list-format-change-v1`) rebind existing list items; their DOM core is
 * `services/list-format-change.js`. This module validates their shape (`scope`, `lvlText`, `continueFrom`).
 */

export const LIST_FORMAT_MODES = Object.freeze(['convert-plain', 'change-format', 'join']);
export const LIST_FORMAT_FORMATS = Object.freeze(['decimal', 'upperLetter', 'lowerLetter', 'upperRoman', 'lowerRoman', 'bullet']);
export const LIST_FORMAT_STRIP_VALUES = Object.freeze(['marker', 'none']);
export const LIST_FORMAT_SCOPES = Object.freeze(['list', 'range']);
/** Keys each mode accepts (`list-format-v1`: convert-plain; `list-format-change-v1`: change-format and join). */
const MODE_KEYS = Object.freeze({
    'convert-plain': Object.freeze(['mode', 'kind', 'format', 'start', 'level', 'strip']),
    'change-format': Object.freeze(['mode', 'kind', 'format', 'lvlText', 'start', 'level', 'scope']),
    join: Object.freeze(['mode', 'continueFrom', 'scope'])
});
/** Keys of the planned contract that a later work package implements (detach). */
const LATER_KEYS = Object.freeze(['materializeMarker']);
const ALL_KEYS = new Set([...Object.values(MODE_KEYS).flat(), ...LATER_KEYS]);

export const LIST_MARKER_NOT_FOUND = 'LIST_MARKER_NOT_FOUND';
export const LIST_MARKER_AMBIGUOUS = 'LIST_MARKER_AMBIGUOUS';
export const LIST_MARKER_UNSUPPORTED = 'LIST_MARKER_UNSUPPORTED';
export const LIST_FORMAT_TARGET_IS_LIST = 'LIST_FORMAT_TARGET_IS_LIST';
export const LIST_FORMAT_STYLE_NUMBERED = 'LIST_FORMAT_STYLE_NUMBERED';
// change-format and join (list-format-change-v1)
export const LIST_FORMAT_TARGET_NOT_LIST = 'LIST_FORMAT_TARGET_NOT_LIST';
export const LIST_FORMAT_MIXED_RANGE = 'LIST_FORMAT_MIXED_RANGE';
export const LIST_FORMAT_STYLE_LINKED = 'LIST_FORMAT_STYLE_LINKED';
export const LIST_FORMAT_SPLIT_UNSUPPORTED = 'LIST_FORMAT_SPLIT_UNSUPPORTED';
export const LIST_FORMAT_DEFINITION_UNSUPPORTED = 'LIST_FORMAT_DEFINITION_UNSUPPORTED';
export const LIST_FORMAT_JOIN_INVALID = 'LIST_FORMAT_JOIN_INVALID';
export const LIST_FORMAT_UNCHANGED = 'LIST_FORMAT_UNCHANGED';

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

const fieldError = (message, field) => ({ message, field });

function validateFormatAndKind(list) {
    if (list.format != null && !LIST_FORMAT_FORMATS.includes(list.format)) {
        return fieldError(`list.format must be one of: ${LIST_FORMAT_FORMATS.join(', ')}.`, 'list.format');
    }
    if (list.kind != null && list.kind !== 'bullet' && list.kind !== 'numbered') {
        return fieldError('list.kind must be "bullet" or "numbered".', 'list.kind');
    }
    if (list.kind != null && list.format != null && (list.kind === 'bullet') !== (list.format === 'bullet')) {
        return fieldError(`list.kind "${list.kind}" contradicts list.format "${list.format}".`, 'list.kind');
    }
    if (list.level != null && !(Number.isInteger(list.level) && list.level >= 0 && list.level <= 8)) {
        return fieldError('list.level must be an integer from 0 to 8.', 'list.level');
    }
    return null;
}

/**
 * Checks a `w:lvlText` value. Numbered levels use `%1`..`%9` placeholders and, once the level is known, must contain
 * their own placeholder and no deeper one; bullet levels are one to four literal characters. Returns a message or null.
 */
export function validateLevelText(lvlText, { bullet = false, level = null } = {}) {
    if (typeof lvlText !== 'string' || lvlText.length < 1 || lvlText.length > 32) return 'list.lvlText must be a string of 1 to 32 characters.';
    const placeholders = lvlText.match(/%./g) || [];
    if (bullet) {
        if (lvlText.includes('%') || lvlText.length > 4) return 'A bullet list.lvlText is the bullet itself: 1 to 4 characters without % placeholders.';
        return null;
    }
    if (placeholders.length === 0) return 'A numbered list.lvlText needs a %1..%9 placeholder (for example "%1." or "(%1)").';
    if (placeholders.some(item => !/^%[1-9]$/.test(item))) return 'list.lvlText placeholders must be %1 to %9.';
    if (level != null) {
        if (!placeholders.includes(`%${level + 1}`)) return `list.lvlText for level ${level} must contain its own placeholder %${level + 1}.`;
        if (placeholders.some(item => Number(item.slice(1)) > level + 1)) return `list.lvlText for level ${level} cannot reference a deeper level than %${level + 1}.`;
    }
    return null;
}

function validateContinueFrom(target) {
    if (!isRecord(target)) {
        return fieldError('join needs list.continueFrom: a paragraph target { paragraphId, fingerprint?, exactText? } of an item of the list to continue.', 'list.continueFrom');
    }
    const strong = ['paragraphId', 'fingerprint', 'exactText', 'text'].some(key => typeof target[key] === 'string' && target[key].trim());
    if (!strong) return fieldError('list.continueFrom needs paragraphId, fingerprint or exactText.', 'list.continueFrom');
    if (target.captureRef != null || target.occurrence != null || (target.revisionView != null && target.revisionView !== 'accepted')) {
        return fieldError('list.continueFrom targets a batch-start accepted-view paragraph (no captureRef, occurrence or rejected view).', 'list.continueFrom');
    }
    return null;
}

/** Validates `operation.list` for `list-format`. Returns an error message (with `field`) or null. */
export function validateListFormatShape(list) {
    if (!isRecord(list)) return fieldError('list-format operations require a "list" object.', 'list');
    const unknown = Object.keys(list).filter(key => !ALL_KEYS.has(key));
    if (unknown.length > 0) return fieldError(`list has unsupported ${unknown.length === 1 ? 'key' : 'keys'}: ${unknown.join(', ')}.`, `list.${unknown[0]}`);
    const later = Object.keys(list).filter(key => LATER_KEYS.includes(key));
    if (later.length > 0) return fieldError(`list.${later[0]} is not supported yet (detach is a later release).`, `list.${later[0]}`);
    if (!LIST_FORMAT_MODES.includes(list.mode)) {
        return fieldError(`list.mode must be one of: ${LIST_FORMAT_MODES.join(', ')} (detach is not available yet).`, 'list.mode');
    }
    const foreign = Object.keys(list).filter(key => !MODE_KEYS[list.mode].includes(key));
    if (foreign.length > 0) {
        return fieldError(`list.${foreign[0]} does not apply to mode "${list.mode}" (it accepts: ${MODE_KEYS[list.mode].filter(key => key !== 'mode').join(', ')}).`, `list.${foreign[0]}`);
    }
    if (list.scope != null && !LIST_FORMAT_SCOPES.includes(list.scope)) {
        return fieldError('list.scope must be "list" (every paragraph of the target\'s list) or "range" (only the target range).', 'list.scope');
    }
    const shapeError = validateFormatAndKind(list);
    if (shapeError) return shapeError;

    if (list.mode === 'join') return validateContinueFrom(list.continueFrom);

    if (list.mode === 'change-format') {
        if (list.format == null && list.kind == null && list.lvlText == null && list.start == null) {
            return fieldError('change-format needs at least one of list.format, list.kind, list.lvlText or list.start.', 'list.format');
        }
        if (list.kind === 'numbered' && list.format == null) {
            return fieldError('kind "numbered" does not say which numbering format; pass list.format (decimal, upperLetter, lowerLetter, upperRoman or lowerRoman).', 'list.format');
        }
        if (list.lvlText != null) {
            const bullet = (list.format ?? (list.kind === 'bullet' ? 'bullet' : null)) === 'bullet';
            const message = validateLevelText(list.lvlText, { bullet, level: Number.isInteger(list.level) ? list.level : null });
            if (message) return fieldError(message, 'list.lvlText');
        }
        if (list.start != null) {
            if (!(Number.isInteger(list.start) && list.start >= 1 && list.start <= 32767)) return fieldError('list.start must be an integer from 1 to 32767.', 'list.start');
            if (list.scope === 'range') {
                return fieldError('scope "range" continues the source count across the split; list.start applies to scope "list" only.', 'list.start');
            }
        }
        return null;
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
        // change-format: null means "the level the target paragraphs use".
        explicitLevel: Number.isInteger(list.level) ? list.level : null,
        strip: list.strip || 'marker',
        start: Number.isInteger(list.start) ? list.start : null,
        lvlText: typeof list.lvlText === 'string' ? list.lvlText : null,
        scope: list.scope || 'list',
        continueFrom: list.continueFrom ?? null
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
