/**
 * Configurable runtime defaults for the reconciliation core.
 * Callers can set these during bootstrap.
 */

const fallbackAuthor = () => (typeof process !== 'undefined' && process.env?.DOCX_REDLINE_AUTHOR) || 'AI Redliner';

let _defaultAuthor = fallbackAuthor();
let _platform = 'Unknown';

/**
 * Set the default track-change author for revision metadata.
 *
 * @param {string} author
 */
export function setDefaultAuthor(author) {
    _defaultAuthor = typeof author === 'string' && author.trim() ? author.trim() : fallbackAuthor();
}

/**
 * Get the current default track-change author.
 *
 * @returns {string}
 */
export function getDefaultAuthor() {
    return _defaultAuthor;
}

/**
 * Set the platform identifier (e.g. 'Win32', 'Mac', 'OfficeOnline').
 *
 * @param {string} platform
 */
export function setPlatform(platform) {
    _platform = typeof platform === 'string' && platform.trim() ? platform.trim() : 'Unknown';
}

/**
 * Get the current platform identifier.
 *
 * @returns {string}
 */
export function getPlatform() {
    return _platform;
}
