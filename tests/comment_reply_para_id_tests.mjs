import './setup-xml-provider.mjs';

// A comment reply's w14:paraId must be unique across the package, including body
// paragraphs. Regression: the body scan used a regex with a literal backspace in
// place of \b, so body paraIds were never collected and a reply could reuse one.
import assert from 'assert/strict';
import { applyCommentReplyToParts } from '../index.js';
import { createCommentParaId } from '../services/comment-builders.js';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"';
const commentsXml = `<w:comments ${W}><w:comment w:id="0" w:author="Reviewer" w:date="2026-01-01T00:00:00Z">`
    + '<w:p w14:paraId="00000ABC"><w:r><w:t>Root comment</w:t></w:r></w:p></w:comment></w:comments>';
const reply = documentXml => applyCommentReplyToParts({
    commentsXml, documentXml, parentCommentId: 0, commentId: 1,
    commentContent: 'Reply', author: 'Me', date: '2026-01-02T00:00:00Z'
});

// The reply's preferred paraId is already taken by a body paragraph.
{
    const preferred = createCommentParaId(1).toUpperCase();
    const documentXml = `<w:document ${W}><w:body><w:p w14:paraId="${preferred}"><w:r><w:t>Body</w:t></w:r></w:p></w:body></w:document>`;
    const result = reply(documentXml);
    assert.equal(result.status, 'ok', JSON.stringify(result.error));
    assert.notEqual(result.paraId, preferred, 'reply paraId must not collide with a body paragraph');
    assert.match(result.paraId, /^[0-9A-F]{8}$/);
    assert.notEqual(result.paraId, '00000ABC', 'reply paraId must not collide with the parent comment');
}

// Lower-case body ids collide too (paraIds are compared case-insensitively).
{
    const preferred = createCommentParaId(1).toUpperCase();
    const documentXml = `<w:document ${W}><w:body><w:p w14:paraId="${preferred.toLowerCase()}"/></w:body></w:document>`;
    assert.notEqual(reply(documentXml).paraId, preferred);
}

// Without a collision the preferred id is used, so allocation stays deterministic.
{
    const documentXml = `<w:document ${W}><w:body><w:p w14:paraId="12345678"/></w:body></w:document>`;
    assert.equal(reply(documentXml).paraId, createCommentParaId(1).toUpperCase());
    assert.equal(reply(null).paraId, createCommentParaId(1).toUpperCase());
}

console.log('comment_reply_para_id_tests: all passed');
