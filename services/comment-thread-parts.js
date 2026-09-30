/**
 * Helpers for the Word comment-threading parts (commentsExtended, commentsIds, commentsExtensible).
 *
 * Word joins those parts to comments.xml through w14:paraId, and it uses the paraId of the LAST
 * paragraph of a comment (verified against real Word output, see
 * tests/fixtures/word-authored/multi-paragraph-thread.docx). Comments with one paragraph make first and
 * last identical, which is why treating the first paragraph as the key went unnoticed.
 */

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
