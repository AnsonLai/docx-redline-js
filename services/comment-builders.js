/**
 * Comment XML builders.
 */

import { NS_W, escapeXml } from '../core/types.js';

export const NS_W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
export const NS_W15 = 'http://schemas.microsoft.com/office/word/2012/wordml';

export function createCommentParaId(commentId) {
    const numeric = Number.parseInt(String(commentId), 10);
    const value = Number.isFinite(numeric) ? (0x70000000 + (numeric >>> 0)) >>> 0 : 0x70000000;
    return value.toString(16).toUpperCase().padStart(8, '0').slice(-8);
}

/**
 * Builds a single w:comment element.
 *
 * @param {number} commentId - Unique comment ID
 * @param {string} author - Author name
 * @param {string} content - Comment text content
 * @param {string} date - ISO date string
 * @returns {string}
 */
export function buildCommentElement(commentId, author, content, date, paraId = createCommentParaId(commentId)) {
    const initials = author.split(' ').map(word => word[0]).join('').toUpperCase() || 'AI';
    const escapedContent = escapeXml(content);
    const escapedAuthor = escapeXml(author);

    return `<w:comment w:id="${commentId}" w:author="${escapedAuthor}" w:date="${date}" w:initials="${initials}">
      <w:p w14:paraId="${escapeXml(paraId)}" xmlns:w14="${NS_W14}">
        <w:r><w:t>${escapedContent}</w:t></w:r>
      </w:p>
    </w:comment>`;
}

/**
 * Builds the complete comments.xml part.
 *
 * @param {Array<{id:number,content:string,author:string,date:string}>} comments - Placed comments
 * @returns {string}
 */
export function buildCommentsPartXml(comments) {
    if (!comments || comments.length === 0) {
        return `<w:comments xmlns:w="${NS_W}"></w:comments>`;
    }

    const commentElements = comments.map(comment =>
        buildCommentElement(comment.id, comment.author, comment.content, comment.date, comment.paraId)
    ).join('\n    ');

    return `<w:comments xmlns:w="${NS_W}">
    ${commentElements}
  </w:comments>`;
}

export function buildCommentsExtendedPartXml(entries) {
    const body = (entries || []).map(entry => {
        const parent = entry.paraIdParent ? ` w15:paraIdParent="${escapeXml(entry.paraIdParent)}"` : '';
        return `<w15:commentEx w15:paraId="${escapeXml(entry.paraId)}"${parent} w15:done="${entry.done ? '1' : '0'}"/>`;
    }).join('');
    return `<w15:commentsEx xmlns:w15="${NS_W15}">${body}</w15:commentsEx>`;
}

/**
 * Builds inline range/reference markers for a comment id.
 *
 * @param {number} commentId - The comment ID
 * @returns {{ start: string, end: string, reference: string }}
 */
export function buildCommentMarkers(commentId) {
    return {
        start: `<w:commentRangeStart w:id="${commentId}"/>`,
        end: `<w:commentRangeEnd w:id="${commentId}"/>`,
        reference: `<w:r><w:rPr></w:rPr><w:commentReference w:id="${commentId}"/></w:r>`
    };
}
