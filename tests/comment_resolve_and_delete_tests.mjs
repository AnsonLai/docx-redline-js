/**
 * Batch `comment_resolve` operations, replies to resolved threads, and deleting comments by id.
 *
 * Oracles come from documents saved by real Word (scripts/generate-word-thread-state-fixtures.ps1 and
 * scripts/generate-word-comment-thread-fixture.ps1):
 *  - threaded-comments.docx: comment 0 with reply 1 (open), comment 2 resolved; carries commentsIds and
 *    commentsExtensible.
 *  - resolved-threads.docx: resolved is thread-level whichever comment is marked done.
 *  - reply-to-resolved.docx: Word writes a reply added to a resolved thread as done too.
 */
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { unzipEntries } from '../scripts/lib/zip-reader.mjs';
import { openDocx } from '../node/index.js';
import { executeCli } from '../node/cli.js';

const load = name => readFileSync(new URL(`./fixtures/word-authored/${name}`, import.meta.url));
const text = (bytes, name) => unzipEntries(bytes).get(name)?.toString('utf8') || '';
const commentExes = xml => [...xml.matchAll(/<w15:commentEx\b[^>]*>/g)].map(m => Object.fromEntries([...m[0].matchAll(/w15:(\w+)="([^"]*)"/g)].map(a => [a[1], a[2]])));
const doneById = bytes => Object.fromEntries(openDocx(bytes).inspect().comments.map(c => [c.id, c.done === true]));
const sameEntry = (a, b, name) => assert.ok(unzipEntries(a).get(name).equals(unzipEntries(b).get(name)), `${name} must be byte-identical`);

const threaded = load('threaded-comments.docx');

// --- resolve through applyOperations ---------------------------------------------------------
{
    assert.deepEqual(doneById(threaded), { 0: false, 1: false, 2: true }, 'fixture baseline');

    const result = await openDocx(threaded).applyOperations([{ type: 'comment_resolve', commentId: 0 }]);
    assert.equal(result.status, 'ok', result.error?.message);
    assert.equal(result.written, true);
    assert.equal(result.results[0].status, 'applied');
    assert.deepEqual(result.results[0].commentIds.sort(), ['0', '1']);
    assert.equal(result.results[0].threadRootId, '0');
    assert.deepEqual(doneById(result.toBuffer()), { 0: true, 1: true, 2: true });
    // Resolving only changes the threading part.
    assert.deepEqual(result.artifactsChanged, ['word/commentsExtended.xml']);
    for (const name of ['word/document.xml', 'word/comments.xml', 'word/commentsIds.xml', 'word/commentsExtensible.xml', '[Content_Types].xml']) {
        sameEntry(threaded, result.toBuffer(), name);
    }

    // Resolving a reply resolves the whole thread (Word behavior).
    const viaReply = await openDocx(threaded).applyOperations([{ type: 'comment_resolve', commentId: '1' }]);
    assert.deepEqual(doneById(viaReply.toBuffer()), { 0: true, 1: true, 2: true });

    // Idempotent: already resolved is a no_change, nothing written.
    const again = await openDocx(threaded).applyOperations([{ type: 'comment_resolve', commentId: 2 }]);
    assert.equal(again.status, 'ok');
    assert.equal(again.written, false);
    assert.equal(again.results[0].status, 'no_change');

    // Reopen.
    const reopened = await openDocx(threaded).applyOperations([{ type: 'comment_resolve', commentId: 2, resolved: false }]);
    assert.equal(reopened.written, true);
    assert.equal(doneById(reopened.toBuffer())['2'], false);
}

// --- reopen a Word-resolved thread from a reply ---------------------------------------------
{
    const bytes = load('resolved-threads.docx');
    const before = doneById(bytes);
    assert.ok(Object.values(before).every(Boolean), 'every comment in the fixture is resolved');
    const reply = openDocx(bytes).inspect().comments.find(c => c.parentCommentId);
    const result = await openDocx(bytes).applyOperations([{ type: 'comment_resolve', commentId: reply.id, resolved: false }]);
    assert.equal(result.status, 'ok', result.error?.message);
    const after = openDocx(result.toBuffer()).inspect().comments;
    const threadIds = new Set([reply.parentCommentId, ...after.filter(c => c.parentCommentId === reply.parentCommentId).map(c => c.id)]);
    for (const comment of after) assert.equal(comment.done, !threadIds.has(comment.id), `comment ${comment.id} done state`);
}

// --- replies and resolve in one batch ----------------------------------------------------------
{
    // Reply then resolve: the new reply is part of the resolved thread.
    const replyThenResolve = await openDocx(threaded).applyOperations([
        { type: 'comment_reply', parentCommentId: 0, commentContent: 'Agreed.', author: 'Agent' },
        { type: 'comment_resolve', commentId: 0 }
    ]);
    assert.equal(replyThenResolve.status, 'ok', replyThenResolve.error?.message);
    const comments = openDocx(replyThenResolve.toBuffer()).inspect().comments;
    assert.equal(comments.length, 4);
    assert.ok(comments.filter(c => c.id === '0' || c.parentCommentId === '0').every(c => c.done), 'whole thread including the new reply is done');

    // Reply to an already-resolved thread: Word writes the reply as done too.
    const wordOracle = commentExes(text(load('reply-to-resolved.docx'), 'word/commentsExtended.xml'));
    assert.deepEqual(wordOracle.map(e => e.done), ['1', '1'], 'Word oracle: reply to a resolved thread is done');
    const late = await openDocx(threaded).applyOperations([{ type: 'comment_reply', parentCommentId: 2, commentContent: 'Late reply', author: 'Agent' }]);
    assert.equal(late.status, 'ok', late.error?.message);
    const lateReply = openDocx(late.toBuffer()).inspect().comments.find(c => c.parentCommentId === '2');
    assert.equal(lateReply.done, true, 'reply inherits the resolved thread state');
    // An open thread still gets an open reply.
    const open = await openDocx(threaded).applyOperations([{ type: 'comment_reply', parentCommentId: 0, commentContent: 'Open reply', author: 'Agent' }]);
    assert.ok(openDocx(open.toBuffer()).inspect().comments.filter(c => c.parentCommentId === '0').every(c => c.done === false));
}

// --- failures are closed ------------------------------------------------------------------------
{
    const unknown = await openDocx(threaded).applyOperations([{ type: 'comment_resolve', commentId: 99 }]);
    assert.equal(unknown.written, false);
    assert.equal(unknown.results[0].error.code, 'COMMENT_NOT_FOUND');

    // Atomic batch: a failing resolve rolls back a successful edit.
    const bodyText = openDocx(threaded).inspect().paragraphs.find(p => p.text.trim()).text;
    const atomic = await openDocx(threaded).applyOperations([
        { type: 'replace', target: { exactText: bodyText }, modified: `${bodyText} Amended.` },
        { type: 'comment_resolve', commentId: 99 }
    ], { atomic: true });
    assert.equal(atomic.written, false);
    assert.ok(Buffer.from(atomic.toUint8Array()).equals(threaded), 'original bytes returned');

    const missingId = await openDocx(threaded).applyOperations([{ type: 'comment_resolve' }]);
    assert.equal(missingId.results[0].error.code, 'INVALID_OPERATION');
    const badFlag = await openDocx(threaded).applyOperations([{ type: 'comment_resolve', commentId: 0, resolved: 'yes' }]);
    assert.equal(badFlag.results[0].error.code, 'INVALID_OPERATION');

    // Headers and footers cannot hold comments.
    const inHeader = await openDocx(load('header-footer.docx')).applyOperations([{ type: 'comment_resolve', commentId: 0, part: { kind: 'header' } }]);
    assert.equal(inHeader.written, false);
    assert.equal(inHeader.error.code, 'COMMENT_IN_HEADER_FOOTER');
}

// --- preflight -----------------------------------------------------------------------------------
{
    const preflight = openDocx(threaded).preflight([
        { type: 'comment_resolve', commentId: 1 },
        { type: 'comment_resolve', commentId: 'nope' }
    ]);
    assert.equal(preflight.results[0].status, 'ready');
    assert.equal(preflight.results[1].status, 'error');
    assert.equal(preflight.results[1].error.code, 'COMMENT_NOT_FOUND');
    assert.equal(preflight.valid, false);

    // Replies to a multi-paragraph comment are recognised (keyed on the last paragraph).
    const multi = openDocx(load('multi-paragraph-thread.docx')).preflight([{ type: 'comment_reply', parentCommentId: 0, commentContent: 'x' }]);
    assert.equal(multi.results[0].status, 'ready');
}

// --- deleteComments by id ------------------------------------------------------------------------
{
    // Deleting a thread root takes its replies with it; sibling parts are kept in sync.
    const root = await openDocx(threaded).deleteComments({ ids: [0] });
    assert.equal(root.status, 'ok', root.error?.message);
    assert.equal(root.commentsRemoved, 2);
    assert.deepEqual(openDocx(root.toBuffer()).inspect().comments.map(c => c.id), ['2']);
    for (const name of ['word/document.xml', 'word/comments.xml', 'word/commentsExtended.xml', 'word/commentsIds.xml', 'word/commentsExtensible.xml']) {
        assert.ok(root.artifactsChanged.includes(name), `artifactsChanged lists ${name}`);
    }
    assert.equal(commentExes(text(root.toBuffer(), 'word/commentsExtended.xml')).length, 1);
    assert.equal((text(root.toBuffer(), 'word/commentsIds.xml').match(/<w16cid:commentId\b/g) || []).length, 1);

    // Deleting only a reply keeps the root and its anchors.
    const reply = await openDocx(threaded).deleteComments({ ids: ['1'] });
    assert.equal(reply.commentsRemoved, 1);
    const remaining = openDocx(reply.toBuffer()).inspect().comments;
    assert.deepEqual(remaining.map(c => c.id).sort(), ['0', '2']);
    assert.match(text(reply.toBuffer(), 'word/document.xml'), /<w:commentRangeStart w:id="0"\/>/);
    assert.doesNotMatch(text(reply.toBuffer(), 'word/document.xml'), /w:id="1"\/>/);

    // Unknown ids fail closed.
    const unknown = await openDocx(threaded).deleteComments({ ids: [0, 42] });
    assert.equal(unknown.status, 'error');
    assert.equal(unknown.error.code, 'COMMENT_NOT_FOUND');
    assert.equal(unknown.written, false);
}

// --- CLI ------------------------------------------------------------------------------------------
{
    const directory = await mkdtemp(path.join(tmpdir(), 'docx-redline-resolve-'));
    try {
        const input = path.join(directory, 'threaded.docx');
        await writeFile(input, threaded);

        const version = await executeCli(['version']);
        for (const capability of ['comment-resolve-operation-v1', 'delete-comments-by-id-v1', 'header-footer-parts-v1']) {
            assert.ok(version.capabilities.includes(capability), capability);
        }

        const ops = path.join(directory, 'resolve.json');
        await writeFile(ops, JSON.stringify([{ type: 'comment_resolve', commentId: 0 }]));
        const applied = await executeCli(['apply', input, '--operations', ops, '--output', path.join(directory, 'resolved.docx')]);
        assert.equal(applied.status, 'ok', JSON.stringify(applied.error));
        assert.equal(applied.written, true);
        assert.deepEqual(doneById(readFileSync(applied.outputPath)), { 0: true, 1: true, 2: true });

        const deleted = await executeCli(['delete-comments', input, '--comment-id', '1,2', '--output', path.join(directory, 'deleted.docx')]);
        assert.equal(deleted.status, 'ok', JSON.stringify(deleted.error));
        assert.equal(deleted.commentsRemoved, 2);
        assert.deepEqual(openDocx(readFileSync(deleted.outputPath)).inspect().comments.map(c => c.id), ['0']);

        const missing = await executeCli(['delete-comments', input, '--comment-id', '77', '--output', path.join(directory, 'missing.docx')]);
        assert.equal(missing.error.code, 'COMMENT_NOT_FOUND');

        const noFilter = await executeCli(['delete-comments', input]);
        assert.equal(noFilter.error.code, 'AUTHOR_REQUIRED');
    } finally {
        await rm(directory, { recursive: true, force: true });
    }
}

console.log('comment_resolve_and_delete_tests: PASS');
