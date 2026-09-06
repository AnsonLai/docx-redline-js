import { createSerializer, parseOoxmlSafe } from '../adapters/xml-adapter.js';
import { buildCommentElement, buildCommentsExtendedPartXml, createCommentParaId, NS_W14, NS_W15 } from './comment-builders.js';

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

function usedParaIds(commentsDoc, extendedDoc) {
    const ids = new Set();
    for (const p of Array.from(commentsDoc?.getElementsByTagNameNS('*', 'p') || [])) {
        const id = attr(p, 'w14:paraId', 'paraId');
        if (id) ids.add(id.toUpperCase());
    }
    for (const ex of Array.from(extendedDoc?.getElementsByTagNameNS('*', 'commentEx') || [])) {
        const id = attr(ex, 'w15:paraId', 'paraId');
        if (id) ids.add(id.toUpperCase());
    }
    return ids;
}

function allocateParaId(commentId, occupied) {
    let candidate = createCommentParaId(commentId);
    let value = Number.parseInt(candidate, 16) >>> 0;
    while (occupied.has(candidate)) {
        value = (value + 1) >>> 0;
        candidate = value.toString(16).toUpperCase().padStart(8, '0');
    }
    occupied.add(candidate);
    return candidate;
}

export function applyCommentReplyToParts({ commentsXml, commentsExtendedXml = null, parentCommentId, commentId, commentContent, author, date = new Date().toISOString() }) {
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
    const parentParagraph = Array.from(parent.getElementsByTagNameNS(NS_W, 'p'))[0] || Array.from(parent.getElementsByTagNameNS('*', 'p'))[0];
    if (!parentParagraph) return { status: 'error', error: { code: 'PARENT_COMMENT_INVALID', message: `Parent comment '${parentCommentId}' has no paragraph.` } };
    let parentParaId = attr(parentParagraph, 'w14:paraId', 'paraId');
    if (!parentParaId) {
        parentParaId = allocateParaId(parentCommentId, occupied);
        parentParagraph.setAttributeNS(NS_W14, 'w14:paraId', parentParaId);
    } else {
        parentParaId = parentParaId.toUpperCase();
    }
    const replyParaId = allocateParaId(commentId, occupied);
    const replyParsed = parseRequired(`<w:comments xmlns:w="${NS_W}" xmlns:w14="${NS_W14}">${buildCommentElement(commentId, author, commentContent, date, replyParaId)}</w:comments>`, 'reply comment');
    commentsDoc.documentElement.appendChild(commentsDoc.importNode(replyParsed.doc.documentElement.firstChild, true));

    if (!extendedDoc) {
        extendedDoc = parseRequired(buildCommentsExtendedPartXml([]), 'word/commentsExtended.xml').doc;
    }
    const root = extendedDoc.documentElement;
    const entries = Array.from(root.getElementsByTagNameNS('*', 'commentEx'));
    if (!entries.some(node => attr(node, 'w15:paraId', 'paraId').toUpperCase() === parentParaId)) {
        const parentEx = extendedDoc.createElementNS(NS_W15, 'w15:commentEx');
        parentEx.setAttributeNS(NS_W15, 'w15:paraId', parentParaId);
        parentEx.setAttributeNS(NS_W15, 'w15:done', '0');
        root.appendChild(parentEx);
    }
    const replyEx = extendedDoc.createElementNS(NS_W15, 'w15:commentEx');
    replyEx.setAttributeNS(NS_W15, 'w15:paraId', replyParaId);
    replyEx.setAttributeNS(NS_W15, 'w15:paraIdParent', parentParaId);
    replyEx.setAttributeNS(NS_W15, 'w15:done', '0');
    root.appendChild(replyEx);

    const serializer = createSerializer();
    return {
        status: 'ok', hasChanges: true,
        commentsXml: serializer.serializeToString(commentsDoc),
        commentsExtendedXml: serializer.serializeToString(extendedDoc),
        commentsXmlMode: 'replace', commentsExtendedXmlMode: 'replace',
        commentId, parentCommentId: String(parentCommentId), paraId: replyParaId, parentParaId
    };
}
