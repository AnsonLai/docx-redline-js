/**
 * Helpers for the Word comment-threading parts (commentsExtended, commentsIds, commentsExtensible).
 *
 * Word joins those parts to comments.xml through w14:paraId, and it uses the paraId of the LAST
 * paragraph of a comment (verified against real Word output, see
 * tests/fixtures/word-authored/multi-paragraph-thread.docx). Comments with one paragraph make first and
 * last identical, which is why treating the first paragraph as the key went unnoticed.
 */

import { createSerializer, parseOoxmlSafe } from '../adapters/xml-adapter.js';
import { createCommentParaId, NS_W14, NS_W15 } from './comment-builders.js';

const attrOf = (node, qualified, local) => node?.getAttribute?.(qualified) || node?.getAttribute?.(local) || '';

/** The paragraph whose paraId identifies a w:comment in the threading parts, or null. */
export function commentThreadParagraph(commentElement) {
    const paragraphs = Array.from(commentElement?.getElementsByTagNameNS?.('*', 'p') || []);
    return paragraphs.length ? paragraphs[paragraphs.length - 1] : null;
}

/** The threading paraId (upper-cased) of a w:comment, or '' when its last paragraph has none. */
export function commentParaId(commentElement) {
    return attrOf(commentThreadParagraph(commentElement), 'w14:paraId', 'paraId').toUpperCase();
}

/** Every paraId (upper-cased) already used by comments.xml paragraphs and commentsExtended entries. */
export function usedParaIds(commentsDoc, extendedDoc) {
    const ids = new Set();
    for (const p of Array.from(commentsDoc?.getElementsByTagNameNS('*', 'p') || [])) {
        const id = attrOf(p, 'w14:paraId', 'paraId');
        if (id) ids.add(id.toUpperCase());
    }
    for (const ex of Array.from(extendedDoc?.getElementsByTagNameNS('*', 'commentEx') || [])) {
        const id = attrOf(ex, 'w15:paraId', 'paraId');
        if (id) ids.add(id.toUpperCase());
    }
    return ids;
}

/** Allocates a paraId derived from the comment id that is not in `occupied`, and marks it occupied. */
export function allocateParaId(commentId, occupied) {
    let candidate = createCommentParaId(commentId);
    let value = Number.parseInt(candidate, 16) >>> 0;
    while (occupied.has(candidate)) {
        value = (value + 1) >>> 0;
        candidate = value.toString(16).toUpperCase().padStart(8, '0');
    }
    occupied.add(candidate);
    return candidate;
}

const parseDoc = (xml, partName) => {
    const parsed = parseOoxmlSafe(xml, 'application/xml');
    if (!parsed.doc || parsed.error) return { error: { code: 'PARSE_ERROR', message: `Could not parse ${partName}: ${parsed.error?.message || 'invalid XML'}` } };
    return { doc: parsed.doc };
};

/**
 * Sets the resolved ("done") state of a whole comment thread.
 *
 * Word treats resolved as a property of the thread: resolving the root marks every reply done, and
 * resolving a single reply marks the root and its siblings done too (verified against real Word,
 * tests/fixtures/word-authored/resolved-threads.docx). So this always updates root + all replies.
 *
 * @returns {{ status: 'ok', hasChanges: boolean, commentsXml?: string, commentsExtendedXml?: string,
 *   threadRootId: string, commentIds: string[], resolved: boolean } | { status: 'error', error: object }}
 */
export function applyThreadResolutionToParts({ commentsXml, commentsExtendedXml = null, commentId, resolved = true }) {
    if (!commentsXml) return { status: 'error', error: { code: 'COMMENTS_PART_MISSING', message: 'Resolving a comment requires an existing word/comments.xml part.' } };
    const commentsParsed = parseDoc(commentsXml, 'word/comments.xml');
    if (commentsParsed.error) return { status: 'error', error: commentsParsed.error };
    const commentsDoc = commentsParsed.doc;
    const idOf = node => attrOf(node, 'w:id', 'id');
    const comments = Array.from(commentsDoc.getElementsByTagNameNS('*', 'comment'));
    const target = comments.find(node => idOf(node) === String(commentId));
    if (!target) return { status: 'error', error: { code: 'COMMENT_NOT_FOUND', message: `Comment '${commentId}' was not found.` } };

    let extendedDoc;
    if (commentsExtendedXml) {
        const parsed = parseDoc(commentsExtendedXml, 'word/commentsExtended.xml');
        if (parsed.error) return { status: 'error', error: parsed.error };
        extendedDoc = parsed.doc;
    } else {
        extendedDoc = parseDoc(`<w15:commentsEx xmlns:w15="${NS_W15}"/>`, 'word/commentsExtended.xml').doc;
    }
    const root = extendedDoc.documentElement;
    const entryFor = paraId => Array.from(root.getElementsByTagNameNS('*', 'commentEx')).find(node => attrOf(node, 'w15:paraId', 'paraId').toUpperCase() === paraId) || null;

    // Every comment needs a threading key. Legacy comments without one get a fresh paraId.
    const occupied = usedParaIds(commentsDoc, extendedDoc);
    let commentsChanged = false;
    const keyOf = node => {
        let paraId = commentParaId(node);
        if (!paraId) {
            paraId = allocateParaId(idOf(node), occupied);
            commentThreadParagraph(node)?.setAttributeNS(NS_W14, 'w14:paraId', paraId);
            commentsChanged = true;
        }
        return paraId;
    };

    const targetKey = keyOf(target);
    const targetEntry = entryFor(targetKey);
    const rootKey = (targetEntry && attrOf(targetEntry, 'w15:paraIdParent', 'paraIdParent').toUpperCase()) || targetKey;

    const threadKeys = new Set([rootKey, targetKey]);
    for (const entry of Array.from(root.getElementsByTagNameNS('*', 'commentEx'))) {
        if (attrOf(entry, 'w15:paraIdParent', 'paraIdParent').toUpperCase() === rootKey) threadKeys.add(attrOf(entry, 'w15:paraId', 'paraId').toUpperCase());
    }

    const doneValue = resolved ? '1' : '0';
    let extendedChanged = !commentsExtendedXml;
    for (const key of threadKeys) {
        let entry = entryFor(key);
        if (!entry) {
            entry = extendedDoc.createElementNS(NS_W15, 'w15:commentEx');
            entry.setAttributeNS(NS_W15, 'w15:paraId', key);
            if (key !== rootKey) entry.setAttributeNS(NS_W15, 'w15:paraIdParent', rootKey);
            entry.setAttributeNS(NS_W15, 'w15:done', doneValue);
            root.appendChild(entry);
            extendedChanged = true;
        } else if (attrOf(entry, 'w15:done', 'done') !== doneValue) {
            entry.setAttributeNS(NS_W15, 'w15:done', doneValue);
            extendedChanged = true;
        }
    }

    const byKey = new Map(comments.map(node => [commentParaId(node), node]));
    const commentIds = [...threadKeys].map(key => byKey.get(key)).filter(Boolean).map(idOf);
    const rootComment = byKey.get(rootKey);
    const serializer = createSerializer();
    const hasChanges = extendedChanged || commentsChanged;
    return {
        status: 'ok', hasChanges, resolved,
        threadRootId: rootComment ? idOf(rootComment) : String(commentId),
        commentIds,
        ...(hasChanges ? { commentsExtendedXml: serializer.serializeToString(extendedDoc) } : {}),
        ...(commentsChanged ? { commentsXml: serializer.serializeToString(commentsDoc) } : {})
    };
}

const NS_W_MAIN = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

/**
 * Anchors a reply in the document body the way Word does.
 *
 * Word gives every reply its own commentRangeStart / commentRangeEnd / commentReference around the same
 * range as its thread (verified in tests/fixtures/word-authored/threaded-comments.docx). A reply defined only
 * in comments.xml/commentsExtended.xml is not shown by Word at all, so the markers are required.
 *
 * The new start goes right after the last existing thread member's start; the new end and reference run go
 * right after the last existing thread member's reference run, mirroring Word's ordering.
 *
 * @param {Document} xmlDoc live document.xml DOM (mutated)
 * @param {string[]} memberIds comment ids already in the thread (root first)
 * @param {string|number} newId id of the reply comment
 * @returns {boolean} false when the thread has no anchors in the body (nothing is changed)
 */
export function anchorReplyInDocument(xmlDoc, memberIds, newId) {
    const wanted = new Set(memberIds.map(String));
    const idOf = node => attrOf(node, 'w:id', 'id');
    const all = name => Array.from(xmlDoc.getElementsByTagNameNS('*', name)).filter(node => wanted.has(idOf(node)));
    const starts = all('commentRangeStart');
    const references = all('commentReference');
    if (!starts.length || !references.length) return false;

    const lastStart = starts[starts.length - 1];
    const lastReference = references[references.length - 1];
    const referenceRun = lastReference.parentNode?.localName === 'r' ? lastReference.parentNode : lastReference;

    const make = (name, parent) => {
        const element = xmlDoc.createElementNS(NS_W_MAIN, `w:${name}`);
        element.setAttributeNS(NS_W_MAIN, 'w:id', String(newId));
        if (parent) parent.appendChild(element);
        return element;
    };
    const insertAfter = (anchor, node) => anchor.parentNode.insertBefore(node, anchor.nextSibling);

    insertAfter(lastStart, make('commentRangeStart'));
    const end = make('commentRangeEnd');
    const run = xmlDoc.createElementNS(NS_W_MAIN, 'w:r');
    run.appendChild(xmlDoc.createElementNS(NS_W_MAIN, 'w:rPr'));
    make('commentReference', run);
    insertAfter(referenceRun, end);
    insertAfter(end, run);
    return true;
}

const NS_W16CID = 'http://schemas.microsoft.com/office/word/2016/wordml/cid';
const NS_W16CEX = 'http://schemas.microsoft.com/office/word/2018/wordml/cex';

const durableIdFor = (paraId, taken) => {
    // Word's durable ids are 8 hex digits below 0x7FFFFFFF. Derive from the paraId so output is deterministic.
    let value = (Number.parseInt(paraId, 16) ^ 0x2545F491) & 0x7FFFFFFE;
    let candidate = value.toString(16).toUpperCase().padStart(8, '0');
    while (taken.has(candidate)) {
        value = (value + 2) & 0x7FFFFFFE;
        candidate = value.toString(16).toUpperCase().padStart(8, '0');
    }
    taken.add(candidate);
    return candidate;
};

/**
 * Brings commentsIds.xml and commentsExtensible.xml in line with comments.xml. Only parts that already
 * exist are touched (Word does not require them, so they are never created). Entries for comments that no
 * longer exist are removed and comments that lack an entry get one.
 *
 * @returns {{ commentsIdsXml?: string, commentsExtensibleXml?: string }} only the parts that changed
 */
export function reconcileCommentSiblingParts({ commentsXml, commentsIdsXml = null, commentsExtensibleXml = null }) {
    const out = {};
    if (!commentsXml || (!commentsIdsXml && !commentsExtensibleXml)) return out;
    const commentsParsed = parseOoxmlSafe(commentsXml, 'application/xml');
    if (!commentsParsed.doc || commentsParsed.error) return out;
    const serializer = createSerializer();

    const byKey = new Map();
    for (const node of Array.from(commentsParsed.doc.getElementsByTagNameNS('*', 'comment'))) {
        const key = commentParaId(node);
        if (key) byKey.set(key, node);
    }

    let idsDoc = null;
    const durableByKey = new Map();
    const taken = new Set();
    let idsChanged = false;
    if (commentsIdsXml) {
        const parsed = parseOoxmlSafe(commentsIdsXml, 'application/xml');
        if (!parsed.doc || parsed.error) return out;
        idsDoc = parsed.doc;
        const root = idsDoc.documentElement;
        for (const entry of Array.from(root.getElementsByTagNameNS('*', 'commentId'))) {
            const key = attrOf(entry, 'w16cid:paraId', 'paraId').toUpperCase();
            const durable = attrOf(entry, 'w16cid:durableId', 'durableId').toUpperCase();
            if (durable) taken.add(durable);
            if (!byKey.has(key)) { root.removeChild(entry); idsChanged = true; continue; }
            durableByKey.set(key, durable);
        }
        for (const key of byKey.keys()) {
            if (durableByKey.has(key)) continue;
            const durable = durableIdFor(key, taken);
            const entry = idsDoc.createElementNS(NS_W16CID, 'w16cid:commentId');
            entry.setAttributeNS(NS_W16CID, 'w16cid:paraId', key);
            entry.setAttributeNS(NS_W16CID, 'w16cid:durableId', durable);
            root.appendChild(entry);
            durableByKey.set(key, durable);
            idsChanged = true;
        }
    }

    let extDoc = null;
    let extChanged = false;
    if (commentsExtensibleXml && idsDoc) {
        const parsed = parseOoxmlSafe(commentsExtensibleXml, 'application/xml');
        if (parsed.doc && !parsed.error) {
            extDoc = parsed.doc;
            const root = extDoc.documentElement;
            const live = new Set(durableByKey.values());
            const present = new Set();
            for (const entry of Array.from(root.getElementsByTagNameNS('*', 'commentExtensible'))) {
                const durable = attrOf(entry, 'w16cex:durableId', 'durableId').toUpperCase();
                if (!live.has(durable)) { root.removeChild(entry); extChanged = true; } else present.add(durable);
            }
            for (const [key, durable] of durableByKey) {
                if (present.has(durable)) continue;
                const date = attrOf(byKey.get(key), 'w:date', 'date');
                const entry = extDoc.createElementNS(NS_W16CEX, 'w16cex:commentExtensible');
                entry.setAttributeNS(NS_W16CEX, 'w16cex:durableId', durable);
                if (date) entry.setAttributeNS(NS_W16CEX, 'w16cex:dateUtc', date);
                root.appendChild(entry);
                extChanged = true;
            }
        }
    }

    if (idsChanged) out.commentsIdsXml = serializer.serializeToString(idsDoc);
    if (extChanged) out.commentsExtensibleXml = serializer.serializeToString(extDoc);
    return out;
}
