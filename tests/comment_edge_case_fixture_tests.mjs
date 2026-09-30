/**
 * Edits, new comments and replies against comment shapes written by real Word
 * (tests/fixtures/word-authored/comment-edge-cases.docx, scripts/generate-word-comment-edge-case-fixture.ps1):
 * a comment on middle words, a comment over another author's tracked insertion, one comment spanning two
 * paragraphs, a comment on a paragraph with a hyperlink, and a comment inside a table cell.
 * Every operation must succeed, keep every existing comment anchored, and leave a valid package.
 */
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { unzipEntries } from '../scripts/lib/zip-reader.mjs';
import { openDocx } from '../node/index.js';

const bytes = readFileSync(new URL('./fixtures/word-authored/comment-edge-cases.docx', import.meta.url));
const markersOf = xml => [...xml.matchAll(/<w:(commentRangeStart|commentRangeEnd|commentReference) w:id="(\d+)"/g)].map(m => `${m[1]}:${m[2]}`);
const documentOf = buffer => unzipEntries(buffer).get('word/document.xml').toString('utf8');
const before = markersOf(documentOf(bytes));
const paragraphs = openDocx(bytes).inspect().paragraphs.map(p => p.text);
const commentIds = openDocx(bytes).inspect().comments.map(c => c.id);
assert.deepEqual(commentIds, ['0', '1', '3', '4', '5'], 'fixture carries the five Word comments');

async function check(label, op, options = {}) {
    const result = await openDocx(bytes).applyOperations([op], { author: 'Agent', atomic: true, ...options });
    assert.equal(result.status, 'ok', `${label}: ${result.error?.code} ${result.error?.message}`);
    assert.equal(result.written, true, label);
    const after = markersOf(documentOf(result.toBuffer()));
    for (const marker of before) assert.ok(after.includes(marker), `${label}: lost ${marker}`);
    for (const id of commentIds) {
        const xml = documentOf(result.toBuffer());
        const start = xml.indexOf(`<w:commentRangeStart w:id="${id}"`);
        const end = xml.indexOf(`<w:commentRangeEnd w:id="${id}"`);
        const reference = xml.indexOf(`<w:commentReference w:id="${id}"`);
        assert.ok(start >= 0 && start < end && end < reference, `${label}: comment ${id} markers out of order`);
    }
    return result;
}
const replace = (index, to) => ({ type: 'replace', target: { exactText: paragraphs[index - 1] }, modified: to });

await check('middle words: reword inside the comment', replace(1, 'Alpha BETA GAMMA delta epsilon.'));
await check('middle words: delete the commented words', replace(1, 'Alpha delta epsilon.'));
await check('middle words: append', replace(1, 'Alpha beta gamma delta epsilon and zeta.'));
// Another author's tracked insertion inside the paragraph is protected unless the caller opts into slicing.
const refused = await openDocx(bytes).applyOperations([replace(2, 'The counterparty proposed short-term tenure.')], { author: 'Agent', atomic: true });
assert.equal(refused.status, 'error');
assert.equal(refused.results[0].error.code, 'EXISTING_REVISIONS');
assert.deepEqual(refused.toBuffer(), bytes, 'a refused edit leaves the document byte-identical');
await check('tracked insertion: reword', replace(2, 'The counterparty proposed short-term tenure.'), { existingRevisions: 'slice-cross-author' });
await check('tracked insertion: append', replace(2, 'The counterparty proposed long-term tenure and rent.'), { existingRevisions: 'slice-cross-author' });
await check('spanning comment: edit first paragraph', replace(3, 'Spanning FIRST paragraph.'));
await check('spanning comment: edit last paragraph', replace(4, 'Spanning SECOND paragraph.'));
await check('hyperlink: change link text', replace(5, 'Link here: example.org now.'));
await check('hyperlink: change text before the link', replace(5, 'Link there: example.com now.'));
await check('table cell: edit commented cell', replace(8, 'Cell TWO'));
for (const id of commentIds) await check(`reply to comment ${id}`, { type: 'comment_reply', parentCommentId: id, commentContent: 'Reply', author: 'Agent' });
for (const index of [1, 2, 3, 4, 5, 8]) await check(`new comment on paragraph ${index}`, { type: 'comment', target: { exactText: paragraphs[index - 1] }, commentContent: 'New', author: 'Agent' });

console.log('PASS: comment edge cases from Word');
