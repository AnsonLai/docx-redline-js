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
import { buildZip } from '../scripts/lib/minimal-zip.mjs';

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

    // Word shows a reply only if it has its own body markers; they follow Word's own ordering
    // (compare word/document.xml of threaded-comments.docx: start0 start1 ... end0 ref0 end1 ref1).
    const bodyOrder = xml => [...xml.matchAll(/<w:(commentRangeStart|commentRangeEnd|commentReference) w:id="([0-9]+)"/g)].map(m => `${m[1].replace('comment', '').replace('Range', '')}${m[2]}`);
    assert.deepEqual(bodyOrder(text(entries, 'word/document.xml')), ['Start0', 'Start1', 'End0', 'Reference0', 'End1', 'Reference1'], 'Word fixture ordering');
    assert.deepEqual(bodyOrder(text(unzipEntries(replied.toBuffer()), 'word/document.xml')), ['Start0', 'Start1', 'Start2', 'End0', 'Reference0', 'End1', 'Reference1', 'End2', 'Reference2'], 'new reply anchored after the last thread member');

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

// --- resolve / reopen a thread ---------------------------------------------------------------
{
    const bytes = load('threaded-comments.docx'); // thread 0 (+ reply 1) open, comment 2 resolved
    const before = unzipEntries(bytes);
    const doneByParaId = entries => Object.fromEntries(commentExes(text(entries, 'word/commentsExtended.xml')).map(e => [e.paraId, e.done]));
    const beforeDone = doneByParaId(before);

    // Resolving the REPLY resolves the whole thread (root + reply) and nothing else.
    const doc = openDocx(bytes);
    const resolved = await doc.resolveComment(1);
    assert.equal(resolved.status, 'ok', resolved.error?.message);
    assert.equal(resolved.written, true);
    assert.equal(resolved.threadRootId, '0');
    assert.deepEqual(resolved.commentIds.sort(), ['0', '1']);
    assert.deepEqual(resolved.artifactsChanged, ['word/commentsExtended.xml'], 'only the threading part changes');
    const after = unzipEntries(resolved.toBuffer());
    for (const name of ['[Content_Types].xml', 'word/_rels/document.xml.rels', 'word/comments.xml', 'word/document.xml', 'word/commentsIds.xml', 'word/commentsExtensible.xml']) {
        assert.deepEqual(after.get(name), before.get(name), `${name} must be untouched by resolve`);
    }
    const afterDone = doneByParaId(after);
    assert.equal(afterDone['18AFD323'], '1');
    assert.equal(afterDone['34E87DA3'], '1');
    assert.equal(afterDone['280B7729'], beforeDone['280B7729'], 'other threads are untouched');
    assert.ok(doc.inspect().comments.filter(c => ['0', '1'].includes(c.id)).every(c => c.done === true));

    // Idempotent: resolving an already-resolved thread writes nothing.
    const again = await doc.resolveComment(0);
    assert.equal(again.status, 'ok');
    assert.equal(again.written, false);
    assert.equal(again.hasChanges, false);

    // Reopen from the root reopens the reply too.
    const reopened = await doc.resolveComment(0, { resolved: false });
    assert.equal(reopened.written, true);
    const reopenedDone = doneByParaId(unzipEntries(reopened.toBuffer()));
    assert.equal(reopenedDone['18AFD323'], '0');
    assert.equal(reopenedDone['34E87DA3'], '0');

    // Unknown comment fails closed and leaves the document unchanged.
    const missing = await openDocx(bytes).resolveComment(99);
    assert.equal(missing.status, 'error');
    assert.equal(missing.error.code, 'COMMENT_NOT_FOUND');
    assert.equal(missing.written, false);
    assert.deepEqual(missing.toBuffer(), bytes);

    // Multi-paragraph thread: keyed on the last paragraph, package still validates.
    const multi = await openDocx(load('multi-paragraph-thread.docx')).resolveComment(0);
    assert.equal(multi.status, 'ok', multi.error?.message);
    assert.ok(commentExes(text(unzipEntries(multi.toBuffer()), 'word/commentsExtended.xml')).every(e => e.done === '1'));

    // A reopened thread from Word's own resolved fixture.
    const wordResolved = await openDocx(load('resolved-threads.docx')).resolveComment(3, { resolved: false });
    assert.equal(wordResolved.status, 'ok', wordResolved.error?.message);
    const states = openDocx(wordResolved.toBuffer()).inspect().comments.map(c => [c.id, c.done]);
    assert.deepEqual(states, [['0', true], ['1', true], ['2', true], ['3', false], ['4', false]]);
}

// --- resolving on a legacy document with no commentsExtended part creates it correctly ---------
{
    const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
    const legacy = buildZip([
        { name: '[Content_Types].xml', data: '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/></Types>' },
        { name: 'word/document.xml', data: `<w:document xmlns:w="${W}"><w:body><w:p><w:commentRangeStart w:id="4"/><w:r><w:t>Anchored</w:t></w:r><w:commentRangeEnd w:id="4"/><w:r><w:commentReference w:id="4"/></w:r></w:p><w:sectPr/></w:body></w:document>` },
        { name: 'word/_rels/document.xml.rels', data: '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/></Relationships>' },
        { name: 'word/comments.xml', data: `<w:comments xmlns:w="${W}"><w:comment w:id="4" w:author="Old"><w:p><w:r><w:t>No paraId, no threading part</w:t></w:r></w:p></w:comment></w:comments>` }
    ]);
    const result = await openDocx(legacy).resolveComment(4);
    assert.equal(result.status, 'ok', result.error?.message);
    const entries = unzipEntries(result.toBuffer());
    assert.match(text(entries, 'word/comments.xml'), /w14:paraId="[0-9A-F]{8}"/, 'legacy comment gets a paraId');
    assert.deepEqual(commentExes(text(entries, 'word/commentsExtended.xml')).map(e => e.done), ['1']);
    assert.ok(text(entries, '[Content_Types].xml').includes('wordprocessingml.commentsExtended+xml'));
    assert.ok(!text(entries, '[Content_Types].xml').includes('vnd.ms-word.commentsExtended'));
    assert.ok(text(entries, 'word/_rels/document.xml.rels').includes('relationships/commentsExtended'));
    assert.equal(openDocx(result.toBuffer()).inspect().comments[0].done, true);
}

// --- commentsIds / commentsExtensible stay consistent with comments.xml ---------------------------
{
    const bytes = load('threaded-comments.docx');
    const idsOf = entries => [...text(entries, 'word/commentsIds.xml').matchAll(/<w16cid:commentId w16cid:paraId="([0-9A-F]+)" w16cid:durableId="([0-9A-F]+)"/g)].map(m => ({ paraId: m[1], durableId: m[2] }));
    const extensibleOf = entries => [...text(entries, 'word/commentsExtensible.xml').matchAll(/<w16cex:commentExtensible w16cex:durableId="([0-9A-F]+)"( w16cex:dateUtc="[^"]*")?/g)].map(m => ({ durableId: m[1], hasDate: !!m[2] }));
    const commentParaIds = entries => [...text(entries, 'word/comments.xml').matchAll(/<w:p [^>]*w14:paraId="([0-9A-F]+)"/g)].map(m => m[1]);
    const before = unzipEntries(bytes);
    assert.equal(idsOf(before).length, 3);

    // Adding a reply adds exactly one entry to each sibling part, with a fresh durable id and a date.
    const replied = await openDocx(bytes).applyOperations([{ type: 'comment_reply', parentCommentId: 0, commentContent: 'Sibling check', author: 'Agent' }], { author: 'Agent', atomic: true });
    assert.equal(replied.status, 'ok', replied.error?.message);
    const afterReply = unzipEntries(replied.toBuffer());
    const ids = idsOf(afterReply);
    assert.equal(ids.length, 4);
    assert.deepEqual(new Set(ids.map(e => e.paraId)), new Set(commentParaIds(afterReply)), 'one commentsIds entry per comment paragraph key');
    assert.equal(new Set(ids.map(e => e.durableId)).size, 4, 'durable ids are unique');
    const extensible = extensibleOf(afterReply);
    assert.deepEqual(new Set(extensible.map(e => e.durableId)), new Set(ids.map(e => e.durableId)), 'one commentsExtensible entry per durable id');
    assert.ok(extensible.every(e => e.hasDate));
    for (const e of idsOf(before)) assert.ok(ids.some(x => x.paraId === e.paraId && x.durableId === e.durableId), 'existing durable ids are preserved');

    // Deleting only the reply's author removes just that reply's entries.
    const internalOnly = await openDocx(bytes).deleteComments({ author: 'Internal' });
    assert.equal(internalOnly.status, 'ok', internalOnly.error?.message);
    const afterInternal = unzipEntries(internalOnly.toBuffer());
    assert.equal(idsOf(afterInternal).length, 2);
    assert.equal(extensibleOf(afterInternal).length, 2);
    assert.deepEqual(new Set(idsOf(afterInternal).map(e => e.paraId)), new Set(commentParaIds(afterInternal)));

    // Deleting a thread's root cascades to its reply and to every sibling entry.
    const rootDeleted = await openDocx(bytes).deleteComments({ author: 'Counterparty' });
    assert.equal(rootDeleted.status, 'ok', rootDeleted.error?.message);
    const afterRoot = unzipEntries(rootDeleted.toBuffer());
    assert.equal(idsOf(afterRoot).length, 0);
    assert.equal(extensibleOf(afterRoot).length, 0);
    assert.equal(commentExes(text(afterRoot, 'word/commentsExtended.xml')).length, 0);
}

console.log('PASS: comment thread Word fixtures');
