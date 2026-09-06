import './setup-xml-provider.mjs';
import assert from 'node:assert/strict';
import { buildZip } from '../scripts/lib/minimal-zip.mjs';
import { openDocx } from '../node/index.js';
import { unzipDocx } from '../node/zip-archive.js';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const contentTypes = '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    + '<Default Extension="xml" ContentType="application/xml"/>'
    + '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
    + '<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>'
    + '</Types>';
const rels = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    + '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>'
    + '</Relationships>';
const documentXml = `<w:document xmlns:w="${W}"><w:body><w:p><w:commentRangeStart w:id="8"/><w:r><w:t>Clause 2.8</w:t></w:r><w:commentRangeEnd w:id="8"/><w:r><w:commentReference w:id="8"/></w:r></w:p><w:sectPr/></w:body></w:document>`;
const commentsXml = `<w:comments xmlns:w="${W}"><w:comment w:id="8" w:author="Emma"><w:p><w:r><w:t>What does this mean?</w:t></w:r></w:p></w:comment></w:comments>`;
const input = buildZip([
    { name: '[Content_Types].xml', data: contentTypes },
    { name: 'word/document.xml', data: documentXml },
    { name: 'word/_rels/document.xml.rels', data: rels },
    { name: 'word/comments.xml', data: commentsXml }
]);

const doc = openDocx(input);
const preflight = doc.preflight([{ type: 'comment_reply', parentCommentId: 8, commentContent: 'It means the inputs are retained.' }], 'Anson');
assert.equal(preflight.valid, true);
assert.equal(preflight.results[0].resolvedBy, 'parent_comment');

const result = await doc.applyOperations([{ type: 'comment_reply', parentCommentId: 8, commentContent: 'It means the inputs are retained.', author: 'Anson' }], { validate: true });
assert.equal(result.status, 'ok', result.error?.message);
assert.equal(result.written, true);
assert.deepEqual(result.receipts[0].commentIds, ['9']);
const entries = unzipDocx(result.toBuffer());
const outputDocument = entries.get('word/document.xml').toString('utf8');
const outputComments = entries.get('word/comments.xml').toString('utf8');
const outputExtended = entries.get('word/commentsExtended.xml').toString('utf8');
assert.equal(outputDocument, documentXml, 'reply-only operation must leave document.xml byte text unchanged');
assert.equal((outputDocument.match(/commentReference/g) || []).length, 1, 'reply must not add a body anchor');
assert.match(outputComments, /w:id="9"/);
assert.match(outputComments, /It means the inputs are retained\./);
assert.match(outputComments, /w14:paraId="[0-9A-F]{8}"/, 'legacy parent should be upgraded with a paraId');
assert.match(outputExtended, /w15:paraIdParent="[0-9A-F]{8}"/);
assert.match(entries.get('[Content_Types].xml').toString('utf8'), /commentsExtended\+xml/);
assert.match(entries.get('word\/_rels\/document.xml.rels').toString('utf8'), /relationships\/commentsExtended/);

const inspection = doc.inspect();
const reply = inspection.comments.find(comment => comment.id === '9');
assert.equal(reply.parentCommentId, '8');
assert.equal(reply.author, 'Anson');

const missing = await doc.applyOperations([{ type: 'comment_reply', parentCommentId: 999, commentContent: 'No parent.' }], { author: 'Anson', atomic: true });
assert.equal(missing.status, 'error');
assert.equal(missing.rolledBack, true);
assert.equal(missing.results[0].error.code, 'PARENT_COMMENT_NOT_FOUND');

const deleted = await doc.deleteComments({ author: 'Anson', validate: true });
assert.equal(deleted.status, 'ok');
assert.equal(deleted.commentsRemoved, 1);
const afterDelete = unzipDocx(deleted.toBuffer());
assert.doesNotMatch(afterDelete.get('word/comments.xml').toString('utf8'), /It means the inputs are retained\./);
assert.doesNotMatch(afterDelete.get('word/commentsExtended.xml').toString('utf8'), /paraIdParent/);

console.log('PASS: threaded comment reply operations');
