import { createSerializer, parseOoxmlSafe } from '../adapters/xml-adapter.js';
import { allocateParaId, commentParaId, commentThreadParagraph, usedParaIds } from './comment-thread-parts.js';
import { buildCommentElement, buildCommentsExtendedPartXml, NS_W14, NS_W15 } from './comment-builders.js';

const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

function attr(node, qualified, local) {
    return node?.getAttribute?.(qualified) || node?.getAttribute?.(local) || '';
}

function parseRequired(xml, partName) {
    const parsed = parseOoxmlSafe(xml, 'application/xml');
    if (!parsed.doc || parsed.error) {
        return { error: { code: 'PARSE_ERROR', message: `Could not parse ${partName}: ${parsed.error?.message || 'invalid XML'}` } };
    }
    return { doc: parsed.doc };
}

export function applyCommentReplyToParts({ commentsXml, commentsExtendedXml = null, documentXml = null, parentCommentId, commentId, commentContent, author, date = new Date().toISOString() }) {
    if (!commentsXml) return { status: 'error', error: { code: 'COMMENTS_PART_MISSING', message: 'A comment reply requires an existing word/comments.xml part.' } };
    const commentsParsed = parseRequired(commentsXml, 'word/comments.xml');
    if (commentsParsed.error) return { status: 'error', error: commentsParsed.error };
    const commentsDoc = commentsParsed.doc;
    const parent = Array.from(commentsDoc.getElementsByTagNameNS('*', 'comment')).find(node => attr(node, 'w:id', 'id') === String(parentCommentId));
    if (!parent) return { status: 'error', error: { code: 'PARENT_COMMENT_NOT_FOUND', message: `Parent comment '${parentCommentId}' was not found.` } };

    let extendedDoc = null;
    if (commentsExtendedXml) {
        const parsed = parseRequired(commentsExtendedXml, 'word/commentsExtended.xml');
        if (parsed.error) return { status: 'error', error: parsed.error };
        extendedDoc = parsed.doc;
    }
    const occupied = usedParaIds(commentsDoc, extendedDoc);
    // paraIds are unique across the whole package; body paragraphs can carry them too.
    if (documentXml) for (const match of String(documentXml).matchAll(/w14:paraId="([0-9A-Fa-f]{8})"/g)) occupied.add(match[1].toUpperCase());
    const parentParagraph = commentThreadParagraph(parent);
    if (!parentParagraph) return { status: 'error', error: { code: 'PARENT_COMMENT_INVALID', message: `Parent comment '${parentCommentId}' has no paragraph.` } };
    let parentParaId = commentParaId(parent);
    if (!parentParaId) {
        parentParaId = allocateParaId(parentCommentId, occupied);
        parentParagraph.setAttributeNS(NS_W14, 'w14:paraId', parentParaId);
    } else {
        parentParaId = parentParaId.toUpperCase();
    }
    // Word threads are flat: a reply to a reply belongs to the same root thread.
    for (const node of Array.from(extendedDoc?.getElementsByTagNameNS('*', 'commentEx') || [])) {
        if (attr(node, 'w15:paraId', 'paraId').toUpperCase() === parentParaId) {
            const root = attr(node, 'w15:paraIdParent', 'paraIdParent').toUpperCase();
            if (root) parentParaId = root;
            break;
        }
    }
    // Comments already in this thread (root first), needed to anchor the reply in the body.
    const threadCommentIds = [];
    const commentNodes = Array.from(commentsDoc.getElementsByTagNameNS('*', 'comment'));
    const parentOf = new Map(Array.from(extendedDoc?.getElementsByTagNameNS('*', 'commentEx') || [])
        .map(node => [attr(node, 'w15:paraId', 'paraId').toUpperCase(), attr(node, 'w15:paraIdParent', 'paraIdParent').toUpperCase()]));
    for (const node of commentNodes) {
        const key = commentParaId(node);
        if (key === parentParaId || parentOf.get(key) === parentParaId) threadCommentIds.push(attr(node, 'w:id', 'id'));
    }
    const threadRootCommentId = attr(commentNodes.find(node => commentParaId(node) === parentParaId), 'w:id', 'id') || String(parentCommentId);
    const replyParaId = allocateParaId(commentId, occupied);
    const replyParsed = parseRequired(`<w:comments xmlns:w="${NS_W}" xmlns:w14="${NS_W14}">${buildCommentElement(commentId, author, commentContent, date, replyParaId)}</w:comments>`, 'reply comment');
    commentsDoc.documentElement.appendChild(commentsDoc.importNode(replyParsed.doc.documentElement.firstChild, true));

    if (!extendedDoc) {
        extendedDoc = parseRequired(buildCommentsExtendedPartXml([]), 'word/commentsExtended.xml').doc;
    }
    const root = extendedDoc.documentElement;
    const entries = Array.from(root.getElementsByTagNameNS('*', 'commentEx'));
    const rootEntry = entries.find(node => attr(node, 'w15:paraId', 'paraId').toUpperCase() === parentParaId);
    if (!rootEntry) {
        const parentEx = extendedDoc.createElementNS(NS_W15, 'w15:commentEx');
        parentEx.setAttributeNS(NS_W15, 'w15:paraId', parentParaId);
        parentEx.setAttributeNS(NS_W15, 'w15:done', '0');
        root.appendChild(parentEx);
    }
    // Resolved is thread-level: Word gives a reply added to a resolved thread done="1" too.
    const threadDone = attr(rootEntry, 'w15:done', 'done') === '1' ? '1' : '0';
    const replyEx = extendedDoc.createElementNS(NS_W15, 'w15:commentEx');
    replyEx.setAttributeNS(NS_W15, 'w15:paraId', replyParaId);
    replyEx.setAttributeNS(NS_W15, 'w15:paraIdParent', parentParaId);
    replyEx.setAttributeNS(NS_W15, 'w15:done', threadDone);
    root.appendChild(replyEx);

    const serializer = createSerializer();
    return {
        status: 'ok', hasChanges: true,
        commentsXml: serializer.serializeToString(commentsDoc),
        commentsExtendedXml: serializer.serializeToString(extendedDoc),
        commentsXmlMode: 'replace', commentsExtendedXmlMode: 'replace',
        commentId, parentCommentId: String(parentCommentId), paraId: replyParaId, parentParaId,
        threadRootCommentId, threadCommentIds: threadCommentIds.length ? threadCommentIds : [threadRootCommentId]
    };
}
