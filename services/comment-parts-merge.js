/**
 * Parts-level comment merge: the comment half of writing a batch result back into a
 * package, for hosts that hold package parts as XML strings rather than a ZIP (a Flat OPC
 * package, a browser parts map). Strings in, strings out; no package, relationship or
 * content-type handling, which stays with the host.
 */

import { mergeCommentsPartXml, parseXmlStrictStandalone } from './standalone-docx-plumbing.js';
import { reconcileCommentSiblingParts } from './comment-thread-parts.js';

const PART_KEYS = Object.freeze({
    comments: 'commentsXml',
    commentsExtended: 'commentsExtendedXml',
    commentsIds: 'commentsIdsXml',
    commentsExtensible: 'commentsExtensibleXml'
});

const partsOf = parts => Object.fromEntries(Object.values(PART_KEYS).map(key => [key, parts?.[key] || null]));

/**
 * Merges a batch result's comment output into a package's comment parts.
 *
 * - `commentsXml` is appended to the existing comments part (`commentsXmlMode: 'merge'`, the default) or replaces it
 *   (`'replace'`). commentsIds.xml and commentsExtensible.xml, when the package has them, are then brought in line with
 *   comments.xml; they are never created.
 * - `commentsExtendedXml` replaces the thread-metadata part. A `'merge'` delta for a package that already has
 *   commentsExtended.xml cannot be merged safely and is refused.
 *
 * Never throws for document content; on error the input parts are returned unchanged.
 *
 * @param {{ commentsXml?: string|null, commentsExtendedXml?: string|null, commentsIdsXml?: string|null, commentsExtensibleXml?: string|null }} parts
 *   the package's current comment parts (absent parts null or omitted)
 * @param {{ commentsXml?: string|null, commentsXmlMode?: 'merge'|'replace', commentsExtendedXml?: string|null, commentsExtendedXmlMode?: 'merge'|'replace' }} output
 *   comment output of a batch, for example an `applyOperationsToDocumentXml` result
 * @returns {{ status: 'ok'|'error', hasChanges: boolean, parts: { commentsXml: string|null, commentsExtendedXml: string|null, commentsIdsXml: string|null, commentsExtensibleXml: string|null }, changedParts: string[], addedParts: string[], error?: { code: string, message: string, part?: string, commentId?: string } }}
 *   `changedParts` and `addedParts` name parts as `comments`, `commentsExtended`, `commentsIds`, `commentsExtensible`;
 *   `addedParts` lists parts the input did not have, which need a relationship and content type. It reflects part
 *   presence only: relationships are not visible here, so a host that owns them should also repair a missing
 *   relationship for any part in `changedParts`.
 */
export function mergeCommentParts(parts = {}, output = {}) {
    const before = partsOf(parts);
    const after = { ...before };
    const fail = error => ({ status: 'error', hasChanges: false, parts: before, changedParts: [], addedParts: [], error });
    const commentsMode = output?.commentsXmlMode === 'replace' ? 'replace' : 'merge';
    const extendedMode = output?.commentsExtendedXmlMode === 'replace' ? 'replace' : 'merge';

    for (const [label, value] of [
        ...Object.values(PART_KEYS).map(key => [`parts.${key}`, before[key]]),
        ['output.commentsXml', output?.commentsXml], ['output.commentsExtendedXml', output?.commentsExtendedXml]
    ]) {
        if (value != null && typeof value !== 'string') {
            return fail({ code: 'INVALID_COMMENT_PARTS', message: `${label} must be an XML string, null or omitted (got ${typeof value}).` });
        }
    }

    // Parses a part and checks its root element, so a wrong part is refused rather than merged as empty.
    const requirePart = (xml, part, rootName, label) => {
        let doc;
        try {
            doc = parseXmlStrictStandalone(xml, label);
        } catch (error) {
            throw Object.assign(error, { part });
        }
        if (doc.documentElement?.localName !== rootName) {
            throw Object.assign(new Error(`${label}: expected a <${rootName}> root element, found <${doc.documentElement?.localName || 'none'}>.`), { code: 'INVALID_COMMENT_PARTS', part });
        }
        return doc;
    };

    try {
        if (output?.commentsXml) {
            const incomingDoc = requirePart(output.commentsXml, 'comments', 'comments', 'word/comments.xml (incoming)');
            // Ids must be unique within the incoming part in every mode (the merge below also checks them
            // against the existing part).
            const seen = new Set();
            for (const node of Array.from(incomingDoc.getElementsByTagNameNS('*', 'comment'))) {
                const id = node.getAttribute('w:id') || node.getAttribute('id');
                if (!id) continue;
                if (seen.has(id)) {
                    throw Object.assign(new Error(`Duplicate comment id: ${id}`), { code: 'DUPLICATE_COMMENT_ID', commentId: id, part: 'comments' });
                }
                seen.add(id);
            }
            if (before.commentsXml) requirePart(before.commentsXml, 'comments', 'comments', 'word/comments.xml (existing)');
            // Sibling parts are reconciled below; a malformed one would otherwise be skipped silently.
            if (before.commentsIdsXml) requirePart(before.commentsIdsXml, 'commentsIds', 'commentsIds', 'word/commentsIds.xml (existing)');
            if (before.commentsExtensibleXml) requirePart(before.commentsExtensibleXml, 'commentsExtensible', 'commentsExtensible', 'word/commentsExtensible.xml (existing)');
            after.commentsXml = mergeCommentsPartXml(before.commentsXml, output.commentsXml, { replaceExisting: commentsMode === 'replace' });
            Object.assign(after, reconcileCommentSiblingParts({
                commentsXml: after.commentsXml,
                commentsIdsXml: after.commentsIdsXml,
                commentsExtensibleXml: after.commentsExtensibleXml
            }));
        }
        if (output?.commentsExtendedXml) {
            if (extendedMode !== 'replace' && before.commentsExtendedXml) {
                return fail({
                    code: 'UNSUPPORTED_EXISTING_COMMENT_THREAD_METADATA',
                    part: 'commentsExtended',
                    message: 'The output carries a commentsExtended delta (mode "merge") for parts that already have thread metadata; a delta cannot be merged into an existing commentsExtended.xml. Nothing was merged.'
                });
            }
            requirePart(output.commentsExtendedXml, 'commentsExtended', 'commentsEx', 'word/commentsExtended.xml (incoming)');
            after.commentsExtendedXml = output.commentsExtendedXml;
        }
    } catch (error) {
        return fail({
            code: error?.code || 'COMMENT_PARTS_MERGE_FAILED',
            message: error?.message || String(error),
            ...(error?.part || error?.code === 'DUPLICATE_COMMENT_ID' ? { part: error.part || 'comments' } : {}),
            ...(error?.commentId ? { commentId: error.commentId } : {})
        });
    }

    const changedParts = [];
    const addedParts = [];
    for (const [name, key] of Object.entries(PART_KEYS)) {
        if (after[key] === before[key]) continue;
        changedParts.push(name);
        if (!before[key]) addedParts.push(name);
    }
    return { status: 'ok', hasChanges: changedParts.length > 0, parts: after, changedParts, addedParts };
}
