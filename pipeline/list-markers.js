/**
 * Shared list-marker detection/parsing helpers.
 *
 * Keeps marker parsing consistent across router, pipeline, and patching flows.
 */

const LIST_MARKER_CORE = String.raw`(?:\d+(?:\.\d+)*\.?|\((?:\d+|[a-zA-Z]|[ivxlcIVXLC]+)\)|[a-zA-Z]\.|\d+\.|[ivxlcIVXLC]+\.|[-*+\u2022])`;

const LINE_REGEX_STRICT = new RegExp(`^(\\s*)((?:${LIST_MARKER_CORE})\\s+)`);
const LINE_REGEX_LOOSE = new RegExp(`^(\\s*)((?:${LIST_MARKER_CORE})\\s*)`);
const MULTILINE_REGEX_STRICT = new RegExp(`^(\\s*)((?:${LIST_MARKER_CORE})\\s+)`, 'm');
const MULTILINE_REGEX_LOOSE = new RegExp(`^(\\s*)((?:${LIST_MARKER_CORE})\\s*)`, 'm');

/**
 * Determines whether text should be treated as list-target content.
 * Strict mode requires at least one whitespace after the marker.
 *
 * @param {string} text - Candidate text
 * @returns {boolean}
 */
export function isListTargetStrict(text) {
    if (typeof text !== 'string') return false;
    return text.includes('\n') && MULTILINE_REGEX_STRICT.test(text);
}

/**
 * Determines whether text should be treated as list-target content.
 * Loose mode allows markers with optional trailing whitespace.
 *
 * @param {string} text - Candidate text
 * @returns {boolean}
 */
export function isListTargetLoose(text) {
    if (typeof text !== 'string') return false;
    return text.includes('\n') && MULTILINE_REGEX_LOOSE.test(text.trim());
}

/**
 * Matches a list marker at the start of a line.
 *
 * @param {string} line - Input line
 * @param {Object} [options={}] - Match options
 * @param {boolean} [options.allowZeroSpaceAfterMarker=false] - Allow zero spaces after marker
 * @returns {RegExpMatchArray|null}
 */
export function matchListMarker(line, options = {}) {
    const { allowZeroSpaceAfterMarker = false } = options;
    const regex = allowZeroSpaceAfterMarker ? LINE_REGEX_LOOSE : LINE_REGEX_STRICT;
    return line.match(regex);
}

/**
 * Extracts the marker text from a line.
 *
 * @param {string} line - Input line
 * @param {Object} [options={}] - Match options
 * @param {boolean} [options.allowZeroSpaceAfterMarker=false] - Allow zero spaces after marker
 * @returns {string}
 */
export function extractListMarker(line, options = {}) {
    const match = matchListMarker(line, options);
    return match ? match[2].trim() : '';
}

/**
 * Strips the marker (and its immediate trailing spacing) from a line.
 *
 * @param {string} line - Input line
 * @param {Object} [options={}] - Strip options
 * @param {boolean} [options.allowZeroSpaceAfterMarker=false] - Allow zero spaces after marker
 * @returns {string}
 */
export function stripListMarker(line, options = {}) {
    const { allowZeroSpaceAfterMarker = false } = options;
    const regex = allowZeroSpaceAfterMarker ? LINE_REGEX_LOOSE : LINE_REGEX_STRICT;
    return line.replace(regex, '');
}

/**
 * Classifies a parsed marker using the shared list vocabulary.
 *
 * @param {string} marker - Marker without trailing whitespace
 * @returns {'bullet'|'numbered'}
 */
export function classifyListMarker(marker) {
    return /^[-*+\u2022]$/.test(String(marker || '').trim()) ? 'bullet' : 'numbered';
}

/**
 * Infers the Word numbering style represented by a numbered marker.
 *
 * @param {string} marker - Marker text
 * @returns {'bullet'|'decimal'|'lowerAlpha'|'upperAlpha'|'lowerRoman'|'upperRoman'}
 */
export function inferNumberingStyleFromMarker(marker) {
    const value = String(marker || '').trim();
    if (classifyListMarker(value) === 'bullet') return 'bullet';
    if (/^\d+(?:\.\d+)*\.?$/.test(value) || /^\(\d+\)$/.test(value)) return 'decimal';
    if (/^[ivxlcdm]+\.$/.test(value)) return 'lowerRoman';
    if (/^[IVXLCDM]{2,}\.$/.test(value)) return 'upperRoman';
    if (/^[a-z]\.$/.test(value)) return 'lowerAlpha';
    if (/^[A-Z]\.$/.test(value)) return 'upperAlpha';
    return 'decimal';
}

/**
 * Returns an outline level encoded directly in a composite decimal marker.
 * Indentation-derived levels remain the caller's responsibility.
 *
 * @param {string} marker - Marker text
 * @returns {number|null}
 */
export function parseOutlineLevelFromMarker(marker) {
    const value = String(marker || '').trim();
    if (!/^\d+(?:\.\d+)+\.?$/.test(value)) return null;
    return Math.max(0, value.replace(/\.$/, '').split('.').length - 1);
}

/**
 * Parses one line into the common list-item representation.
 *
 * @param {string} line - Input line
 * @param {{allowZeroSpaceAfterMarker?: boolean, indentSpaces?: number}} [options]
 * @returns {{line:string,text:string,marker:string,indent:number,level:number,markerType:'bullet'|'numbered',listType:'bullet'|'numbered',numberingStyle:string,outlineLevel:number|null}|null}
 */
export function parseListItem(line, options = {}) {
    const match = matchListMarker(String(line || ''), options);
    if (!match) return null;
    const marker = match[2].trim();
    const indent = (match[1] || '').length;
    const indentSpaces = Math.max(1, Number(options.indentSpaces) || 2);
    const markerType = classifyListMarker(marker);
    return {
        line: String(line || ''),
        text: stripListMarker(String(line || ''), options),
        marker,
        indent,
        level: Math.min(8, Math.floor(indent / indentSpaces)),
        markerType,
        listType: markerType,
        numberingStyle: inferNumberingStyleFromMarker(marker),
        outlineLevel: markerType === 'numbered' ? parseOutlineLevelFromMarker(marker) : null
    };
}
