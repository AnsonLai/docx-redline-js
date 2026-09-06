import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { configureXmlProvider, parseOoxmlSafe } from '../../adapters/xml-adapter.js';
import { unzipDocx } from '../../node/zip-archive.js';
import { canonicalizeOoxml } from './canonical-ooxml.mjs';
import { assertPackageFidelity } from './package-fidelity.mjs';
import { getParagraphId, getParagraphText, findContainingWordElement } from '../../core/paragraph-targeting.js';
import { acceptTrackedChangesInOoxml, rejectTrackedChangesInOoxml } from '../../services/revision-comment-management.js';

configureXmlProvider({ DOMParser, XMLSerializer });

export const MUTATION_ENVELOPES = Object.freeze({
    surgical_text: Object.freeze({
        name: 'surgical_text',
        allowedPackageEntries: Object.freeze(['word/document.xml']),
        allowsArtifacts: Object.freeze({ comments: false, numbering: false })
    }),
    reconstruction: Object.freeze({
        name: 'reconstruction',
        allowedPackageEntries: Object.freeze(['word/document.xml']),
        allowsArtifacts: Object.freeze({ comments: false, numbering: false })
    }),
    comment: Object.freeze({
        name: 'comment',
        allowedPackageEntries: Object.freeze([
            'word/document.xml',
            'word/comments.xml',
            'word/_rels/document.xml.rels',
            '[Content_Types].xml'
        ]),
        allowsArtifacts: Object.freeze({ comments: true, numbering: false })
    }),
    highlight: Object.freeze({
        name: 'highlight',
        allowedPackageEntries: Object.freeze(['word/document.xml']),
        allowsArtifacts: Object.freeze({ comments: false, numbering: false })
    }),
    list: Object.freeze({
        name: 'list',
        allowedPackageEntries: Object.freeze([
            'word/document.xml',
            'word/numbering.xml',
            'word/_rels/document.xml.rels',
            '[Content_Types].xml'
        ]),
        allowsArtifacts: Object.freeze({ comments: false, numbering: true })
    }),
    structured_content: Object.freeze({
        name: 'structured_content',
        allowedPackageEntries: Object.freeze([
            'word/document.xml',
            'word/numbering.xml',
            'word/_rels/document.xml.rels',
            '[Content_Types].xml'
        ]),
        allowsArtifacts: Object.freeze({ comments: false, numbering: true })
    }),
    table_reconciliation: Object.freeze({
        name: 'table_reconciliation',
        allowedPackageEntries: Object.freeze(['word/document.xml']),
        allowsArtifacts: Object.freeze({ comments: false, numbering: false })
    }),
    accept_revisions: Object.freeze({
        name: 'accept_revisions',
        allowedPackageEntries: Object.freeze(['word/document.xml']),
        allowsArtifacts: Object.freeze({ comments: false, numbering: false })
    }),
    reject_revisions: Object.freeze({
        name: 'reject_revisions',
        allowedPackageEntries: Object.freeze(['word/document.xml']),
        allowsArtifacts: Object.freeze({ comments: false, numbering: false })
    })
});

/**
 * Inventories all block-level body subtrees (<w:p> and <w:tbl>) in a document XML string or DOCX buffer.
 *
 * @param {string|Buffer|Map} input
 * @returns {Array<{ index: number, tag: string, paraId: string|null, inTable: boolean, text: string, sha256: string, canonicalXml: string, element: Element }>}
 */
export function inventoryDocumentSubtrees(input) {
    let documentXml;
    if (typeof input === 'string') {
        documentXml = input;
    } else {
        const entries = Buffer.isBuffer(input) ? unzipDocx(input) : (input?.entries || input);
        const docEntry = entries.get('word/document.xml');
        if (!docEntry) throw new Error('Missing word/document.xml in package input');
        documentXml = docEntry.toString('utf8');
    }

    const parsed = parseOoxmlSafe(documentXml, 'application/xml');
    if (parsed.error || !parsed.doc) {
        throw new Error(`Failed to parse document XML: ${parsed.error?.message || 'invalid XML'}`);
    }

    const bodies = parsed.doc.getElementsByTagNameNS('*', 'body');
    const body = bodies.length > 0 ? bodies[0] : parsed.doc.documentElement;

    const subtrees = [];
    let elementIndex = 0;

    for (let child = body.firstChild; child; child = child.nextSibling) {
        if (child.nodeType !== 1) continue;
        const local = String(child.localName || child.nodeName || '').replace(/^.*:/, '');
        if (local !== 'p' && local !== 'tbl' && local !== 'sdt') continue;

        elementIndex++;
        const paraId = local === 'p' ? getParagraphId(child) : null;
        const text = local === 'p' ? getParagraphText(child) : '';
        const inTable = local === 'p' ? !!findContainingWordElement(child, 'tbl') : false;
        const canonical = canonicalizeOoxml(child);

        subtrees.push({
            index: elementIndex,
            tag: local,
            paraId,
            inTable,
            text,
            sha256: canonical.sha256,
            canonicalXml: canonical.canonicalXml,
            element: child
        });
    }

    return subtrees;
}

/**
 * Asserts full mutation fidelity between a before DOCX package and an after DOCX package.
 *
 * 1. Checks that untouched ZIP entries are bit-identical.
 * 2. Checks that untouched body subtrees in word/document.xml are canonical-XML identical.
 * 3. Checks that Accept All and Reject All execute cleanly on the modified document.
 *
 * @param {Buffer|Map} beforeDocx
 * @param {Buffer|Map} afterDocx
 * @param {object} options
 * @param {object} options.envelope - One of MUTATION_ENVELOPES
 * @param {number[]} [options.targetedIndexes=[]] - 1-based indexes of targeted body subtrees
 * @param {string[]} [options.targetedParagraphIds=[]] - paraIds of targeted paragraphs
 * @param {boolean} [options.lifecycleCheck=true] - whether to assert accept/reject roundtrip sanity
 * @returns {{ packageComparison: object, untouchedSubtreesChecked: number }}
 */
export function verifyMutationFidelity(beforeDocx, afterDocx, options = {}) {
    const envelope = options.envelope;
    if (!envelope || !Array.isArray(envelope.allowedPackageEntries)) {
        throw new TypeError('verifyMutationFidelity requires a valid mutation envelope in options.envelope');
    }

    // 1. Package entry fidelity
    const packageComparison = assertPackageFidelity(beforeDocx, afterDocx, envelope.allowedPackageEntries);

    // 2. Subtree canonical fidelity in word/document.xml
    const beforeSubtrees = inventoryDocumentSubtrees(beforeDocx);
    const afterSubtrees = inventoryDocumentSubtrees(afterDocx);

    const targetedIndexes = new Set(options.targetedIndexes || []);
    const targetedParaIds = new Set(options.targetedParagraphIds || []);

    const untouchedBefore = beforeSubtrees.filter(item =>
        !targetedIndexes.has(item.index) && (!item.paraId || !targetedParaIds.has(item.paraId))
    );

    let untouchedChecked = 0;
    for (const beforeItem of untouchedBefore) {
        let afterMatch = null;
        if (beforeItem.paraId) {
            afterMatch = afterSubtrees.find(candidate => candidate.paraId === beforeItem.paraId);
        }
        if (!afterMatch) {
            // Match by exact text and tag
            afterMatch = afterSubtrees.find(candidate =>
                candidate.tag === beforeItem.tag && candidate.text === beforeItem.text
            );
        }

        if (!afterMatch) {
            throw new Error(
                `Untouched subtree missing after mutation: Index ${beforeItem.index}, Tag <${beforeItem.tag}>, ` +
                `paraId=${beforeItem.paraId}, text="${beforeItem.text.slice(0, 40)}..."`
            );
        }

        if (beforeItem.sha256 !== afterMatch.sha256) {
            const error = new Error(
                `[SUBTREE_FIDELITY_VIOLATION] Untouched subtree modified during operation!\n` +
                `  Tag: <${beforeItem.tag}>, Index: ${beforeItem.index} -> ${afterMatch.index}\n` +
                `  paraId: ${beforeItem.paraId}\n` +
                `  Expected SHA-256: ${beforeItem.sha256}\n` +
                `  Actual SHA-256:   ${afterMatch.sha256}\n` +
                `  Expected Canonical XML: ${beforeItem.canonicalXml}\n` +
                `  Actual Canonical XML:   ${afterMatch.canonicalXml}`
            );
            error.code = 'SUBTREE_FIDELITY_VIOLATION';
            error.beforeItem = beforeItem;
            error.afterMatch = afterMatch;
            throw error;
        }

        untouchedChecked++;
    }

    // 3. Lifecycle check on modified document
    if (options.lifecycleCheck !== false) {
        const afterEntries = Buffer.isBuffer(afterDocx) ? unzipDocx(afterDocx) : (afterDocx?.entries || afterDocx);
        const docXml = afterEntries.get('word/document.xml')?.toString('utf8');
        if (docXml) {
            const accepted = acceptTrackedChangesInOoxml(docXml, { allAuthors: true });
            if (accepted.status === 'error') {
                throw new Error(`Accept All failed on mutated document: ${accepted.error?.message}`);
            }
            const rejected = rejectTrackedChangesInOoxml(docXml, { allAuthors: true });
            if (rejected.status === 'error') {
                throw new Error(`Reject All failed on mutated document: ${rejected.error?.message}`);
            }
        }
    }

    return {
        packageComparison,
        untouchedSubtreesChecked: untouchedChecked
    };
}
