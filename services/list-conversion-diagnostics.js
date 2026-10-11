/**
 * Interim list-conversion safety diagnostics (canonical list operations plan, WP-1).
 *
 * These detect requests that the text-reconciliation path cannot honour as list structure and
 * report them instead of returning a silent `ok`:
 * - `LIST_MARKER_NOT_BOUND` / `UNSUPPORTED_LIST_CONVERSION`: a plain paragraph was edited so its
 *   text starts with a list marker, but the marker was written as literal text (no `w:numPr`).
 * - `LIST_FORMAT_NOT_APPLIED` / `UNSUPPORTED_LIST_FORMAT_CHANGE`: a list item was asked to change
 *   kind or numbering format, but the paragraph keeps its source numbering definition.
 * - `UNKNOWN_PARAGRAPH_FORMAT_PROPERTY`: a `paragraph-format` key that is not implemented.
 *
 * They are independent of the Reject All verifier (`REVISION_FIDELITY_VIOLATION`). The pure
 * detectors below take plain data; policy resolution decides warning versus error.
 */

import { parseOoxmlSafe } from '../adapters/xml-adapter.js';
import { parseListItem } from '../pipeline/list-markers.js';
import { parseTypedListMarker } from './list-format.js';

export const LIST_MARKER_NOT_BOUND = 'LIST_MARKER_NOT_BOUND';
export const UNSUPPORTED_LIST_CONVERSION = 'UNSUPPORTED_LIST_CONVERSION';
export const LIST_FORMAT_NOT_APPLIED = 'LIST_FORMAT_NOT_APPLIED';
export const UNSUPPORTED_LIST_FORMAT_CHANGE = 'UNSUPPORTED_LIST_FORMAT_CHANGE';
export const UNKNOWN_PARAGRAPH_FORMAT_PROPERTY = 'UNKNOWN_PARAGRAPH_FORMAT_PROPERTY';
export const INVALID_PARAGRAPH_NUMBERING = 'INVALID_PARAGRAPH_NUMBERING';
export const PARAGRAPH_NUMBERING_NOT_FOUND = 'PARAGRAPH_NUMBERING_NOT_FOUND';
export const PARAGRAPH_FORMAT_REVISION_CONFLICT = 'PARAGRAPH_FORMAT_REVISION_CONFLICT';

/** Keys `paragraph-format.properties` implements today (closed set). */
export const PARAGRAPH_FORMAT_PROPERTY_KEYS = Object.freeze([
    'alignment', 'keepNext', 'keepLines', 'pageBreakBefore', 'style', 'numbering'
]);

const NUMBERING_STYLE_TO_NUM_FMT = Object.freeze({
    bullet: 'bullet',
    decimal: 'decimal',
    lowerAlpha: 'lowerLetter',
    upperAlpha: 'upperLetter',
    lowerRoman: 'lowerRoman',
    upperRoman: 'upperRoman'
});
const COMPARABLE_NUM_FMTS = new Set(['bullet', 'decimal', 'lowerLetter', 'upperLetter', 'lowerRoman', 'upperRoman']);

/**
 * `listMarkers: 'literal'` is the only accepted value; it suppresses these diagnostics for
 * callers that intentionally want literal marker text.
 */
export function isLiteralListMarkerOptOut(value) {
    return value === 'literal';
}

/**
 * Resolves diagnostic policy: `off` (opt-out), `error` (strict profile) or `warn` (default).
 * Strict is requested explicitly with `strictListConversion` (the CLI agent profile sets it) or
 * implied by an atomic batch.
 */
export function resolveListDiagnosticPolicy(options = {}) {
    if (isLiteralListMarkerOptOut(options.listMarkers)) return 'off';
    const strict = options.strictListConversion !== undefined
        ? options.strictListConversion === true
        : options.atomic === true;
    return strict ? 'error' : 'warn';
}

function excerpt(text, limit = 80) {
    const value = String(text ?? '');
    return value.length > limit ? `${value.slice(0, limit)}...` : value;
}

/** Parses single-line text as one list marker plus content, or `null`. */
export function parseLiteralListMarker(text) {
    const raw = String(text ?? '');
    if (!raw.trim() || /[\r\n]/.test(raw)) return null;
    const parsed = parseListItem(raw, { allowZeroSpaceAfterMarker: false, indentSpaces: 2 });
    if (!parsed || !parsed.marker) return null;
    // A bare number or number-dot-number ("7 days", "2.1 Scope") is far more often prose than a list
    // marker; require an explicit terminator for numbered markers so the diagnostics stay low-noise.
    if (parsed.markerType !== 'bullet' && !/[.)]$/.test(parsed.marker)) return null;
    return {
        marker: parsed.marker,
        kind: parsed.markerType === 'bullet' ? 'bullet' : 'numbered',
        numberingStyle: parsed.numberingStyle,
        numFmt: NUMBERING_STYLE_TO_NUM_FMT[parsed.numberingStyle] || 'decimal',
        text: String(parsed.text || '').trim()
    };
}

/**
 * Detects a plain (non-list) paragraph whose edited text starts a list marker the original
 * did not carry. Returns the requested marker, or `null` when there is nothing to report.
 */
export function detectUnboundListMarker({ originalText, modifiedText, targetIsListItem }) {
    if (targetIsListItem) return null;
    const requested = parseLiteralListMarker(modifiedText);
    if (!requested) return null;
    const current = parseLiteralListMarker(originalText);
    if (current && current.kind === requested.kind && current.numFmt === requested.numFmt) return null;
    return requested;
}

/**
 * The `list-format` operation that performs the conversion a text edit could not: bind the paragraph as a list
 * item in the requested format starting at the typed value, leaving the (unmarked) text alone. Only offered when
 * the original paragraph has no marker of its own, because `strip: "none"` would keep that marker as text.
 */
function suggestedListFormatOperation(requested, originalHasMarker) {
    if (originalHasMarker) return null;
    const format = requested.numFmt === 'bullet' ? 'bullet' : requested.numFmt;
    const list = { mode: 'convert-plain', format, strip: 'none' };
    if (format !== 'bullet') {
        const parsed = parseTypedListMarker(`${requested.marker} x`);
        const value = parsed?.candidates?.find(candidate => candidate.format === format)?.value;
        if (Number.isInteger(value) && value > 1) list.start = value;
    }
    return { type: 'list-format', list };
}

/** Builds the structured diagnostic (warning or error payload) for an unbound marker. */
export function buildUnboundListMarkerDiagnostic(requested, { policy, targetText, operationIndex } = {}) {
    const isError = policy === 'error';
    const code = isError ? UNSUPPORTED_LIST_CONVERSION : LIST_MARKER_NOT_BOUND;
    const message = isError
        ? `Converting a plain paragraph to a list item is not supported by text edits: the new text begins with the list marker "${requested.marker}", which would be written as literal text with no list binding. Use a list-format convert-plain operation on the paragraph (capability list-format-v1), or pass listMarkers: "literal" to keep the marker as text.`
        : `Modified text begins with the list marker "${requested.marker}", but the marker was written as literal text; no list binding was created. A list-format convert-plain operation creates the list binding.`;
    const suggestedOperation = suggestedListFormatOperation(requested, !!parseLiteralListMarker(targetText));
    return {
        code,
        severity: isError ? 'error' : 'warning',
        message,
        listMarker: requested.marker,
        requestedList: { kind: requested.kind, format: requested.numFmt },
        ...(targetText ? { targetExcerpt: excerpt(targetText) } : {}),
        ...(Number.isInteger(operationIndex) ? { operationIndex } : {}),
        recovery: {
            action: 'use_list_format_convert_plain',
            hint: suggestedOperation
                ? 'Keep the paragraph text as it is and send a list-format operation on the same target (recovery.suggestedOperation plus your target), or pass listMarkers: "literal" to write the marker as text intentionally.'
                : 'Edit the text with a redline that starts with the typed marker, then convert it with a list-format convert-plain operation (strip: "marker"), or pass listMarkers: "literal" to write the marker as text intentionally.',
            ...(suggestedOperation ? { suggestedOperation } : {}),
            field: 'listMarkers',
            allowedValue: 'literal'
        }
    };
}

function levelDefinition(numberingDoc, numId, ilvl) {
    const byTag = tag => Array.from(numberingDoc.getElementsByTagNameNS('*', tag));
    const attr = (node, name) => node?.getAttribute?.(`w:${name}`) ?? node?.getAttribute?.(name) ?? null;
    const child = (node, tag) => Array.from(node?.childNodes || []).find(c => c.nodeType === 1 && c.localName === tag) || null;
    const num = byTag('num').find(node => String(attr(node, 'numId')) === String(numId));
    if (!num) return null;
    const override = Array.from(num.childNodes).filter(c => c.nodeType === 1 && c.localName === 'lvlOverride')
        .find(c => Number(attr(c, 'ilvl') || 0) === ilvl);
    const overrideFmt = attr(child(child(override, 'lvl'), 'numFmt'), 'val');
    if (overrideFmt) return { numFmt: overrideFmt };
    const abstractId = attr(child(num, 'abstractNumId'), 'val');
    const abstract = byTag('abstractNum').find(node => String(attr(node, 'abstractNumId')) === String(abstractId));
    if (!abstract) return null;
    const lvl = Array.from(abstract.childNodes).filter(c => c.nodeType === 1 && c.localName === 'lvl')
        .find(c => Number(attr(c, 'ilvl') || 0) === ilvl);
    if (!lvl) return null;
    return { numFmt: attr(child(lvl, 'numFmt'), 'val') || 'decimal' };
}

/** Reads `{ numFmt, kind }` of a numbering level from the source numbering part, or `null`. */
export function readSourceListLevel(sourceNumberingXml, numId, ilvl = 0) {
    if (!String(sourceNumberingXml || '').trim() || numId == null) return null;
    const parsed = parseOoxmlSafe(String(sourceNumberingXml));
    if (!parsed?.doc) return null;
    const level = levelDefinition(parsed.doc, numId, Number(ilvl) || 0);
    if (!level) return null;
    return { numFmt: level.numFmt, kind: level.numFmt === 'bullet' ? 'bullet' : 'numbered' };
}

/**
 * Compares the format a list edit asked for with the source level definition. Returns the
 * mismatch description or `null` when it matches, is unknown, or cannot be compared safely.
 */
export function detectListFormatMismatch({ requested, sourceLevel }) {
    if (!requested || !sourceLevel) return null;
    const requestedKind = requested.kind === 'bullet' ? 'bullet' : 'numbered';
    if (requestedKind !== sourceLevel.kind) return { requested, source: sourceLevel };
    if (
        requestedKind === 'numbered'
        && COMPARABLE_NUM_FMTS.has(requested.numFmt)
        && COMPARABLE_NUM_FMTS.has(sourceLevel.numFmt)
        && requested.numFmt !== sourceLevel.numFmt
    ) {
        return { requested, source: sourceLevel };
    }
    return null;
}

export function buildListFormatDiagnostic(mismatch, { policy, numId, level, marker, targetText, operationIndex } = {}) {
    const isError = policy === 'error';
    const code = isError ? UNSUPPORTED_LIST_FORMAT_CHANGE : LIST_FORMAT_NOT_APPLIED;
    const describe = (kind, format) => (kind === 'bullet' ? 'a bullet list' : `a numbered list${format ? ` (${format})` : ''}`);
    const requestedFormat = mismatch.requested.kind === 'bullet' ? 'bullet' : (mismatch.requested.numFmt || mismatch.requested.format);
    const requestedText = describe(mismatch.requested.kind, requestedFormat);
    const sourceText = describe(mismatch.source.kind, mismatch.source.numFmt);
    const suggestedOperation = COMPARABLE_NUM_FMTS.has(requestedFormat)
        ? { type: 'list-format', list: { mode: 'change-format', format: requestedFormat, level: Number(level) || 0, scope: 'list' } }
        : null;
    return {
        code,
        severity: isError ? 'error' : 'warning',
        message: `${isError ? 'Changing the list format is not supported by text edits' : 'The requested list format was not applied'}: the list item keeps its source numbering definition (numId ${numId}, ${sourceText}); the edit asked for ${requestedText}${marker ? ` ("${marker}")` : ''}. A list-format change-format operation changes the list's format (capability list-format-change-v1).`,
        requestedList: { kind: mismatch.requested.kind, format: requestedFormat },
        sourceList: { numId: String(numId), level: Number(level) || 0, kind: mismatch.source.kind, format: mismatch.source.numFmt },
        ...(targetText ? { targetExcerpt: excerpt(targetText) } : {}),
        ...(Number.isInteger(operationIndex) ? { operationIndex } : {}),
        recovery: {
            action: 'use_list_format_change_format',
            hint: suggestedOperation
                ? 'Text edits cannot change a list\'s kind or numbering format. Send recovery.suggestedOperation with a target on one item of the list as its own operation (scope "list" rebinds every item; scope "range" only a range that runs to the end of the list). Make any text edit in a separate redline without the new markers, targeting a different paragraph or a separate call. Or pass listMarkers: "literal" to accept the current result.'
                : 'Text edits cannot change a list\'s kind or numbering format. Use a list-format change-format operation, or pass listMarkers: "literal" to accept the current result.',
            ...(suggestedOperation ? { suggestedOperation } : {}),
            field: 'listMarkers',
            allowedValue: 'literal'
        }
    };
}

/**
 * Validates the shape of `paragraph-format.properties.numbering`: `null` (remove the binding) or
 * `{ numId: positive integer, level?: integer 0-8 }`. Returns an error message or null.
 */
export function validateParagraphNumberingShape(numbering) {
    if (numbering === null) return null;
    if (typeof numbering !== 'object' || Array.isArray(numbering)) {
        return 'numbering must be null (remove the list binding) or an object { numId, level }.';
    }
    const extra = Object.keys(numbering).filter(key => key !== 'numId' && key !== 'level');
    if (extra.length > 0) return `numbering has unsupported ${extra.length === 1 ? 'key' : 'keys'}: ${extra.join(', ')}. Use only numId and level.`;
    const numId = numbering.numId;
    const numIdOk = (Number.isInteger(numId) && numId >= 1) || (typeof numId === 'string' && /^[1-9]\d*$/.test(numId));
    if (!numIdOk) return 'numbering.numId must be a positive integer (numId 0 is not a list; pass numbering: null to remove the binding).';
    const level = numbering.level;
    if (level !== undefined && !(Number.isInteger(level) && level >= 0 && level <= 8)) {
        return 'numbering.level must be an integer from 0 to 8.';
    }
    return null;
}

/** Finds `paragraph-format.properties` keys that are not implemented. */
export function findUnknownParagraphFormatKeys(properties) {
    if (!properties || typeof properties !== 'object') return [];
    return Object.keys(properties).filter(key => !PARAGRAPH_FORMAT_PROPERTY_KEYS.includes(key));
}

export function buildUnknownParagraphFormatDiagnostic(keys) {
    const list = keys.map(key => `"${key}"`).join(', ');
    return {
        code: UNKNOWN_PARAGRAPH_FORMAT_PROPERTY,
        severity: 'warning',
        message: `paragraph-format ignored unsupported ${keys.length === 1 ? 'property' : 'properties'} ${list}. Supported: ${PARAGRAPH_FORMAT_PROPERTY_KEYS.join(', ')}. A future minor release will reject unknown keys.`,
        properties: keys,
        recovery: { action: 'change_request', hint: `Use only: ${PARAGRAPH_FORMAT_PROPERTY_KEYS.join(', ')}.` }
    };
}

/** `CODE: message` string form used in `warnings` arrays. */
export function diagnosticWarningText(diagnostic) {
    return `${diagnostic.code}: ${diagnostic.message}`;
}

function normalizeLine(text) {
    return String(text ?? '').replace(/[\s\u00a0]+/g, ' ').trim();
}

/**
 * Compares the list markers in a multi-line edit with the list items it edits.
 *
 * `sourceItems` are `{ text, numId, ilvl }` for the existing list paragraphs. A modified line
 * that keeps a source item's text and carries a marker states the format requested for that
 * item. Inserted marker lines are not compared: the insertion heuristics deliberately add
 * nested bullets under numbered items and the reverse.
 */
export function detectRequestedListFormatMismatch({ modifiedText, sourceItems, sourceNumberingXml }) {
    if (!Array.isArray(sourceItems) || sourceItems.length === 0) return null;
    const lines = String(modifiedText ?? '').split(/\r?\n/).filter(line => line.trim());
    const parsed = lines.map(line => ({ line, item: parseLiteralListMarker(line) }));
    const retained = new Set();
    for (const source of sourceItems) {
        const sourceText = normalizeLine(source.text);
        const match = parsed.find((entry, index) => !retained.has(index)
            && normalizeLine(entry.item ? entry.item.text : entry.line) === sourceText);
        if (!match) continue;
        retained.add(parsed.indexOf(match));
        if (!match.item) continue;
        const sourceLevel = readSourceListLevel(sourceNumberingXml, source.numId, source.ilvl);
        const mismatch = detectListFormatMismatch({ requested: match.item, sourceLevel });
        if (mismatch) return { mismatch, marker: match.item.marker, numId: source.numId, level: source.ilvl, targetText: source.text };
    }
    return null;
}
