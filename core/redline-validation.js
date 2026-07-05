/**
 * Runtime structural validation for generated redline OOXML.
 *
 * Mirrors the invariants enforced by the test-suite round-trip harness so
 * downstream consumers can verify output before writing it into a package:
 * no nested revisions, deleted text uses w:delText, revision metadata is
 * complete, revision ids are unique, and boundary whitespace is preserved.
 */

import { parseXml } from '../adapters/xml-adapter.js';
import { NS_W } from './types.js';

const REVISION_ID_ELEMENTS = new Set(['ins', 'del', 'rPrChange', 'pPrChange']);
const REVISION_DATE_PATTERN = /^\d{4}-\d{2}-\d{2}T/;

function localNameOf(node) {
    return String(node?.localName || node?.nodeName || '').replace(/^.*:/, '');
}

function elementsByLocalName(root, name) {
    return Array.from(root.getElementsByTagName('*')).filter(el => localNameOf(el) === name);
}

function wordAttribute(node, name) {
    return node.getAttribute(`w:${name}`) || node.getAttribute(name) || '';
}

function xmlSpaceAttribute(node) {
    return node.getAttribute('xml:space') ||
        node.getAttribute('space') ||
        node.getAttributeNS?.('http://www.w3.org/XML/1998/namespace', 'space') ||
        '';
}

function isParagraphMarkRevision(node) {
    return localNameOf(node.parentNode) === 'rPr';
}

function parseOoxmlForValidation(oxml) {
    const attempt = xml => {
        const doc = parseXml(xml);
        const parseError = doc.getElementsByTagName('parsererror')[0];
        if (parseError) {
            throw new Error(parseError.textContent || 'XML parse error');
        }
        return doc;
    };

    try {
        return { doc: attempt(oxml) };
    } catch {
        try {
            return { doc: attempt(`<w:root xmlns:w="${NS_W}">${oxml}</w:root>`) };
        } catch (error) {
            return { error: error?.message || 'XML parse error' };
        }
    }
}

/**
 * Validates redline OOXML against the package's structural invariants.
 *
 * Issue severities: 'error' issues indicate output Word may repair or
 * mis-resolve; 'warning' issues are suspicious but tolerated by Word.
 *
 * @param {string} oxml - OOXML string (fragment, document, or package scope)
 * @returns {{ valid: boolean, issues: Array<{ code: string, severity: 'error'|'warning', message: string }> }}
 */
export function validateRedlineOoxml(oxml) {
    const issues = [];
    const addIssue = (code, severity, message) => issues.push({ code, severity, message });

    if (typeof oxml !== 'string' || oxml.trim() === '') {
        addIssue('PARSE_ERROR', 'error', 'Input is not a non-empty OOXML string.');
        return { valid: false, issues };
    }

    const { doc, error } = parseOoxmlForValidation(oxml);
    if (!doc) {
        addIssue('PARSE_ERROR', 'error', `OOXML does not parse as XML: ${error}`);
        return { valid: false, issues };
    }

    const insElements = elementsByLocalName(doc, 'ins');
    const delElements = elementsByLocalName(doc, 'del');
    const revisions = insElements.concat(delElements);

    // No w:ins/w:del nested inside another w:ins/w:del.
    for (const revision of revisions) {
        const nested = Array.from(revision.getElementsByTagName('*'))
            .filter(el => el !== revision && ['ins', 'del'].includes(localNameOf(el)));
        if (nested.length > 0) {
            addIssue('NESTED_REVISION', 'error',
                `<${revision.nodeName}> (w:id="${wordAttribute(revision, 'id')}") contains nested <${nested[0].nodeName}>.`);
        }
    }

    // Deleted runs must carry w:delText, never w:t.
    for (const del of delElements) {
        const plainTextNodes = elementsByLocalName(del, 't');
        if (plainTextNodes.length > 0) {
            addIssue('DEL_CONTAINS_T', 'error',
                `<w:del> (w:id="${wordAttribute(del, 'id')}") contains <w:t>; deleted text must use <w:delText>.`);
        }
    }

    // Every revision needs complete metadata.
    for (const revision of revisions) {
        const missing = [];
        if (!wordAttribute(revision, 'id')) missing.push('w:id');
        if (!wordAttribute(revision, 'author')) missing.push('w:author');
        if (!REVISION_DATE_PATTERN.test(wordAttribute(revision, 'date'))) missing.push('w:date');
        if (missing.length > 0) {
            addIssue('MISSING_REVISION_METADATA', 'error',
                `<${revision.nodeName}> is missing or has malformed ${missing.join(', ')}.`);
        }
    }

    // Revision ids must be unique among ins/del/rPrChange/pPrChange.
    const seenIds = new Set();
    for (const node of Array.from(doc.getElementsByTagName('*'))) {
        if (!REVISION_ID_ELEMENTS.has(localNameOf(node))) continue;
        const id = wordAttribute(node, 'id');
        if (!id) continue;
        if (seenIds.has(id)) {
            addIssue('DUPLICATE_REVISION_ID', 'error', `Revision id ${id} appears more than once.`);
        }
        seenIds.add(id);
    }

    // Boundary whitespace requires xml:space="preserve".
    const textNodes = elementsByLocalName(doc, 't').concat(elementsByLocalName(doc, 'delText'));
    for (const node of textNodes) {
        const text = node.textContent || '';
        if (/^\s|\s$/.test(text) && xmlSpaceAttribute(node) !== 'preserve') {
            addIssue('MISSING_SPACE_PRESERVE', 'error',
                `<${node.nodeName}> has boundary whitespace without xml:space="preserve".`);
        }
        if (text === '') {
            addIssue('EMPTY_TEXT_ELEMENT', 'warning', `<${node.nodeName}> is empty.`);
        }
    }

    // Empty w:ins/w:del wrappers (paragraph-mark revisions inside w:rPr are
    // legitimately empty and excluded).
    for (const revision of revisions) {
        if (isParagraphMarkRevision(revision)) continue;
        const hasElementChild = Array.from(revision.childNodes || []).some(child => child.nodeType === 1);
        if (!hasElementChild) {
            addIssue('EMPTY_REVISION_WRAPPER', 'warning',
                `<${revision.nodeName}> (w:id="${wordAttribute(revision, 'id')}") wraps no content.`);
        }
    }

    return { valid: !issues.some(issue => issue.severity === 'error'), issues };
}
