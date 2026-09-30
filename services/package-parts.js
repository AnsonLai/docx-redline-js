/**
 * Registry of OPC package parts this library reads or writes.
 *
 * Single source of truth for part paths, content types and relationship types so the
 * package plumbing never hard-codes them. Every `verified` entry is checked against
 * documents saved by real Word in tests/package_parts_registry_tests.mjs
 * (fixtures under tests/fixtures/word-authored/).
 */

const WORD_ML = 'application/vnd.openxmlformats-officedocument.wordprocessingml';
const REL_OFFICE_DOC = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

/**
 * @typedef {object} PackagePartSpec
 * @property {string} id
 * @property {string|null} path       Fixed part path, or null when the name is not fixed (headers/footers).
 * @property {RegExp|null} pathPattern Matches every part name of this kind, when the name is not fixed.
 * @property {string} contentType     The value Word writes in [Content_Types].xml.
 * @property {string[]} legacyContentTypes Known-wrong values written by earlier versions; safe to rewrite.
 * @property {string} relType         Relationship type from word/_rels/document.xml.rels.
 * @property {string|null} relTarget  Relative target for a fixed-path part.
 * @property {boolean} verified       Confirmed against a Word-saved document.
 */

/** @type {ReadonlyArray<Readonly<PackagePartSpec>>} */
const PART_SPECS = Object.freeze([
    { id: 'document', path: 'word/document.xml', pathPattern: null, contentType: `${WORD_ML}.document.main+xml`, legacyContentTypes: [], relType: `${REL_OFFICE_DOC}/officeDocument`, relTarget: null, verified: true },
    { id: 'numbering', path: 'word/numbering.xml', pathPattern: null, contentType: `${WORD_ML}.numbering+xml`, legacyContentTypes: [], relType: `${REL_OFFICE_DOC}/numbering`, relTarget: 'numbering.xml', verified: false },
    { id: 'comments', path: 'word/comments.xml', pathPattern: null, contentType: `${WORD_ML}.comments+xml`, legacyContentTypes: [], relType: `${REL_OFFICE_DOC}/comments`, relTarget: 'comments.xml', verified: true },
    {
        id: 'commentsExtended', path: 'word/commentsExtended.xml', pathPattern: null,
        contentType: `${WORD_ML}.commentsExtended+xml`,
        // Written by this package through 0.8.0; Word does not recognize it and offers to repair the file.
        legacyContentTypes: ['application/vnd.ms-word.commentsExtended+xml'],
        relType: 'http://schemas.microsoft.com/office/2011/relationships/commentsExtended', relTarget: 'commentsExtended.xml', verified: true
    },
    { id: 'commentsIds', path: 'word/commentsIds.xml', pathPattern: null, contentType: `${WORD_ML}.commentsIds+xml`, legacyContentTypes: [], relType: 'http://schemas.microsoft.com/office/2016/09/relationships/commentsIds', relTarget: 'commentsIds.xml', verified: true },
    { id: 'commentsExtensible', path: 'word/commentsExtensible.xml', pathPattern: null, contentType: `${WORD_ML}.commentsExtensible+xml`, legacyContentTypes: [], relType: 'http://schemas.microsoft.com/office/2018/08/relationships/commentsExtensible', relTarget: 'commentsExtensible.xml', verified: true },
    // Not yet produced by any Word fixture (Word only writes people.xml for signed-in authors): unverified.
    { id: 'people', path: 'word/people.xml', pathPattern: null, contentType: `${WORD_ML}.people+xml`, legacyContentTypes: [], relType: 'http://schemas.microsoft.com/office/2011/relationships/people', relTarget: 'people.xml', verified: false },
    { id: 'header', path: null, pathPattern: /^word\/header\d+\.xml$/i, contentType: `${WORD_ML}.header+xml`, legacyContentTypes: [], relType: `${REL_OFFICE_DOC}/header`, relTarget: null, verified: true },
    { id: 'footer', path: null, pathPattern: /^word\/footer\d+\.xml$/i, contentType: `${WORD_ML}.footer+xml`, legacyContentTypes: [], relType: `${REL_OFFICE_DOC}/footer`, relTarget: null, verified: true }
].map(spec => Object.freeze(spec)));

const BY_ID = new Map(PART_SPECS.map(spec => [spec.id, spec]));

/** Returns the spec for a registry id; throws for an unknown id so typos fail loudly. */
export function getPartSpec(id) {
    const spec = BY_ID.get(id);
    if (!spec) throw new Error(`Unknown package part: ${id}`);
    return spec;
}

/** All registered part specs. */
export function listPartSpecs() {
    return PART_SPECS;
}

/** Normalizes '/word/x.xml' or 'word/x.xml' to 'word/x.xml'. */
function toEntryName(partName) {
    return String(partName || '').replace(/^\/+/, '');
}

/** Finds the spec that owns a part name (with or without a leading slash), or null. */
export function findPartSpecByPartName(partName) {
    const name = toEntryName(partName);
    const lower = name.toLowerCase();
    return PART_SPECS.find(spec => spec.path ? spec.path.toLowerCase() === lower : spec.pathPattern?.test(name)) || null;
}

/**
 * When an override's content type is a value earlier versions of this package are known to have
 * written wrongly for that part, returns the correct one; otherwise null. Deliberately does not
 * "correct" types it has no evidence about (for example macro-enabled document types).
 */
export function correctedContentTypeFor(partName, actualContentType) {
    const spec = findPartSpecByPartName(partName);
    if (!spec || !spec.legacyContentTypes.includes(actualContentType)) return null;
    return spec.contentType;
}
