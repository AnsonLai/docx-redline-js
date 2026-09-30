/**
 * Comment thread behavior against documents saved by real Word.
 *
 * Fixtures (scripts/generate-word-thread-state-fixtures.ps1):
 *  - multi-paragraph-thread.docx: a 3-paragraph comment with a reply. Word keys commentsExtended on the
 *    LAST paragraph of a comment; treating the first paragraph as the key broke inspection, validation,
 *    replies and delete-cascade for any multi-paragraph comment.
 *  - resolved-threads.docx: Word marks a whole thread done whichever comment is resolved.
 */
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { unzipEntries } from '../scripts/lib/zip-reader.mjs';
import { openDocx } from '../node/index.js';
import { MemoryZip, unzipDocx } from '../document/zip-archive.js';
import { validateDocxPackage } from '../services/standalone-docx-plumbing.js';

const load = name => readFileSync(new URL(`./fixtures/word-authored/${name}`, import.meta.url));
const text = (entries, name) => entries.get(name)?.toString('utf8') || '';
const commentExes = xml => [...xml.matchAll(/<w15:commentEx\b[^>]*>/g)].map(m => Object.fromEntries([...m[0].matchAll(/w15:(\w+)="([^"]*)"/g)].map(a => [a[1], a[2]])));

// --- multi-paragraph thread ---------------------------------------------------------------
{
    const bytes = load('multi-paragraph-thread.docx');
    const entries = unzipEntries(bytes);
    const [rootEx, replyEx] = commentExes(text(entries, 'word/commentsExtended.xml'));
    assert.equal(commentExes(text(entries, 'word/commentsExtended.xml')).length, 2);

    // The package as Word wrote it must validate.
    await validateDocxPackage(new MemoryZip(unzipDocx(bytes)));

    // Inspection joins the thread through the last paragraph.
    const comments = openDocx(bytes).inspect().comments;
    const root = comments.find(c => c.id === '0');
    const reply = comments.find(c => c.id === '1');
    assert.equal(root.paraId, rootEx.paraId, 'root paraId is the last paragraph of the comment');
    assert.equal(root.done, false);
    assert.equal(reply.parentCommentId, '0');
    assert.equal(reply.parentParaId, rootEx.paraId);
    assert.equal(reply.paraId, replyEx.paraId);

    // A reply attaches to the root's Word key: no duplicate parent entry, parent points at the root.
    const replied = await openDocx(bytes).applyOperations([{ type: 'comment_reply', parentCommentId: 0, commentContent: 'Second reply', author: 'Agent' }], { author: 'Agent', atomic: true });
    assert.equal(replied.status, 'ok', replied.error?.message);
    const after = commentExes(text(unzipEntries(replied.toBuffer()), 'word/commentsExtended.xml'));
    assert.equal(after.length, 3, 'exactly one new entry (no duplicate parent entry)');
    assert.equal(after.filter(e => e.paraId === rootEx.paraId).length, 1);
    assert.equal(after.filter(e => e.paraIdParent === rootEx.paraId).length, 2);
    assert.equal(new Set(after.map(e => e.paraId)).size, 3, 'paraIds are unique');

    // A reply to a reply joins the same root thread (Word threads are flat).
    const flat = await openDocx(bytes).applyOperations([{ type: 'comment_reply', parentCommentId: 1, commentContent: 'Reply to reply', author: 'Agent' }], { author: 'Agent', atomic: true });
    assert.equal(flat.status, 'ok', flat.error?.message);
    const flatEx = commentExes(text(unzipEntries(flat.toBuffer()), 'word/commentsExtended.xml'));
    assert.equal(flatEx.filter(e => e.paraIdParent === rootEx.paraId).length, 2, 'reply-to-reply parents to the root');
    assert.ok(!flatEx.some(e => e.paraIdParent === replyEx.paraId), 'no nested reply chain');

    // Deleting the root cascades to its reply and empties the threading entries.
    const deleted = await openDocx(bytes).deleteComments({ author: 'Counterparty' });
    assert.equal(deleted.status, 'ok', deleted.error?.message);
    assert.equal(deleted.commentsRemoved, 2, 'root and its reply');
    assert.equal(commentExes(text(unzipEntries(deleted.toBuffer()), 'word/commentsExtended.xml')).length, 0);
}

// --- resolved threads (Word's own thread-level done semantics) -------------------------------
{
    const bytes = load('resolved-threads.docx');
    await validateDocxPackage(new MemoryZip(unzipDocx(bytes)));
    const comments = openDocx(bytes).inspect().comments;
    assert.equal(comments.length, 5);
    assert.ok(comments.every(c => c.done === true), 'Word marks every comment of a resolved thread done');
    assert.deepEqual(comments.filter(c => c.parentCommentId).map(c => [c.id, c.parentCommentId]), [['1', '0'], ['2', '0'], ['4', '3']]);
}

console.log('PASS: comment thread Word fixtures');
