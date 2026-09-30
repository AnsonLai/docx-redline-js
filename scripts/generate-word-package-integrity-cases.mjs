/**
 * Builds .docx outputs for the Word open-check (scripts/word-com-package-integrity.ps1).
 *
 * Starting from a document Word itself saved (tests/fixtures/word-authored/threaded-comments.docx),
 * runs each write path of the package and writes the result to tmp/word-package-integrity/.
 * `expect-open` cases must open in Word with repair disabled. `expect-fail-*` cases are controls
 * that reproduce the original bug and prove the Word check can actually detect it.
 */
import { mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { openDocx } from '../node/index.js';
import { unzipDocx, zipDocx } from '../document/zip-archive.js';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const outDir = join(root, 'tmp', 'word-package-integrity');
const fixture = readFileSync(join(root, 'tests', 'fixtures', 'word-authored', 'threaded-comments.docx'));
const WORD_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml';
const BAD_TYPE = 'application/vnd.ms-word.commentsExtended+xml';

rmSync(outDir, { recursive: true, force: true });
mkdirSync(outDir, { recursive: true });

const write = (name, bytes) => writeFileSync(join(outDir, name), Buffer.from(bytes));

function withContentType(bytes, from, to) {
    const entries = unzipDocx(bytes);
    const xml = new TextDecoder().decode(entries.get('[Content_Types].xml'));
    if (!xml.includes(from)) throw new Error(`fixture does not contain ${from}`);
    entries.set('[Content_Types].xml', new TextEncoder().encode(xml.replace(from, to)));
    return zipDocx(entries);
}

// `expectDone`: the resolved state Word must report for each comment (document order). Checked through Word's
// own object model, so it proves Word reads our threading parts the way we wrote them.
// `expect` is either an array of Done states or { commentsDone?, ancestors?, texts? } where `ancestors[i]` is the
// 1-based index of comment i's thread parent (0 = top level), read from Word's Comment.Ancestor.
async function run(name, source, fn, expect) {
    const doc = openDocx(source);
    const result = await fn(doc);
    if (result.status !== 'ok' || !result.written) throw new Error(`${name}: ${result.error?.message || 'not written'}`);
    write(`expect-open-${name}.docx`, result.toBuffer());
    if (expect) writeFileSync(join(outDir, `expect-open-${name}.expect.json`), JSON.stringify(Array.isArray(expect) ? { commentsDone: expect } : expect));
}
const load = name => readFileSync(join(root, 'tests', 'fixtures', 'word-authored', name));

const edit = doc => doc.applyOperations([{ type: 'replace', target: { exactText: 'This agreement is governed by local law.' }, modified: 'This agreement is governed by New York law.' }], { author: 'Agent', atomic: true });
const reply = doc => doc.applyOperations([{ type: 'comment_reply', parentCommentId: 0, commentContent: 'Agreed, twenty days works.', author: 'Agent' }], { author: 'Agent', atomic: true });

write('expect-open-00-word-original.docx', fixture);
await run('01-edit-unrelated-to-comments', fixture, edit);
await run('02-reply', fixture, reply, [false, false, false, true]);
await run('03-add-comment', fixture, doc => doc.applyOperations([{ type: 'comment', target: { exactText: 'This agreement is governed by local law.' }, commentContent: 'Check governing law.', author: 'Agent' }], { author: 'Agent', atomic: true }));
await run('04-edit-then-accept', fixture, async doc => { await edit(doc); return doc.resolveRevisions('accept', { allAuthors: true }); });
await run('05-delete-comments', fixture, doc => doc.deleteComments({ author: 'Counterparty' }));

await run('07-resolve-thread-via-reply', fixture, doc => doc.resolveComment(1), [true, true, true]);
await run('08-reopen-word-resolved-thread', load('resolved-threads.docx'), doc => doc.resolveComment(3, { resolved: false }), [true, true, true, false, false]);
await run('09-reply-to-multi-paragraph-comment', load('multi-paragraph-thread.docx'), doc => doc.applyOperations([{ type: 'comment_reply', parentCommentId: 0, commentContent: 'Another reply', author: 'Agent' }], { author: 'Agent', atomic: true }), [false, false, false]);
await run('10-resolve-multi-paragraph-thread', load('multi-paragraph-thread.docx'), doc => doc.resolveComment(0), [true, true]);

// --- additional Word probes: scenarios likelier to expose bugs than the happy paths above ---
const stripThreadingParts = bytes => {
    const entries = unzipDocx(bytes);
    for (const name of ['word/commentsExtended.xml', 'word/commentsIds.xml', 'word/commentsExtensible.xml']) entries.delete(name);
    const dec = new TextDecoder(); const enc = new TextEncoder();
    entries.set('[Content_Types].xml', enc.encode(dec.decode(entries.get('[Content_Types].xml')).replace(/<Override PartName="\/word\/comments(Extended|Ids|Extensible)\.xml"[^>]*\/>/g, '')));
    entries.set('word/_rels/document.xml.rels', enc.encode(dec.decode(entries.get('word/_rels/document.xml.rels')).replace(/<Relationship [^>]*Target="comments(Extended|Ids|Extensible)\.xml"[^>]*\/>/g, '')));
    return zipDocx(entries);
};
const replyTo = (parent, text) => ({ type: 'comment_reply', parentCommentId: parent, commentContent: text, author: 'Agent' });
const apply = (doc, ops) => doc.applyOperations(ops, { author: 'Agent', atomic: true });

await run('11-reply-to-a-reply', fixture, doc => apply(doc, [replyTo(1, 'Reply to the reply')]), {ancestors: [0, 1, 1, 0], commentsDone: [false, false, false, true]});
await run('12-two-replies-same-root', fixture, doc => apply(doc, [replyTo(0, 'First new'), replyTo(0, 'Second new')]), {ancestors: [0, 1, 1, 1, 0], commentsDone: [false, false, false, false, true]});
await run('13-comment-then-reply-to-own-comment', fixture, async doc => {
    const added = await apply(doc, [{ type: 'comment', target: { exactText: 'This agreement is governed by local law.' }, commentContent: 'Governing law?', author: 'Agent' }]);
    if (added.status !== 'ok') return added;
    return apply(doc, [replyTo(3, 'Reply to my own comment')]);
}, {ancestors: [0, 1, 0, 0, 4], commentsDone: [false, false, true, false, false]});
await run('14-edit-inside-commented-paragraph', fixture, doc => apply(doc, [{ type: 'replace', target: { exactText: 'The Supplier shall deliver the goods within thirty days.' }, modified: 'The Supplier shall deliver the goods within twenty days.' }]), {ancestors: [0, 1, 0], commentsDone: [false, false, true]});
await run('15-edit-and-reply-in-one-batch', fixture, doc => apply(doc, [replyTo(0, 'Batch reply'), { type: 'replace', target: { exactText: 'Payment is due upon receipt of invoice.' }, modified: 'Payment is due within thirty days of invoice.' }]), {ancestors: [0, 1, 1, 0], commentsDone: [false, false, false, true]});
await run('16-reply-accept-resolve', fixture, async doc => {
    await apply(doc, [replyTo(0, 'Then accept')]);
    await doc.resolveRevisions('accept', { allAuthors: true });
    return doc.resolveComment(0);
}, {ancestors: [0, 1, 1, 0], commentsDone: [true, true, true, true]});
await run('17-legacy-no-threading-parts-reply', stripThreadingParts(fixture), doc => apply(doc, [replyTo(0, 'Reply on legacy file')]), {ancestors: [0, 1, 0, 0], commentsDone: [false, false, false, false]});
await run('18-legacy-no-threading-parts-resolve', stripThreadingParts(fixture), doc => doc.resolveComment(0), {ancestors: [0, 0, 0], commentsDone: [true, false, false]});
await run('19-edit-commented-text-multi-paragraph', load('multi-paragraph-thread.docx'), doc => apply(doc, [{ type: 'replace', target: { exactText: 'Alpha paragraph.' }, modified: 'Alpha paragraph, revised.' }]), {ancestors: [0, 1], commentsDone: [false, false]});
await run('20-delete-reply-only', fixture, doc => doc.deleteComments({ author: 'Internal' }), {ancestors: [0, 0], commentsDone: [false, true]});

// Comment shapes Word wrote (comment-edge-cases.docx): every edit must leave Word seeing the same comments.
const edgeCases = load('comment-edge-cases.docx');
const edgeText = doc => doc.inspect().paragraphs.map(p => p.text);
const edgeReplace = (index, to) => doc => apply(doc, [{ type: 'replace', target: { exactText: edgeText(doc)[index - 1] }, modified: to }]);
const five = { ancestors: [0, 0, 0, 0, 0], commentsDone: [false, false, false, false, false] };
await run('21-edge-edit-middle-words', edgeCases, edgeReplace(1, 'Alpha BETA GAMMA delta epsilon.'), five);
await run('22-edge-edit-over-tracked-insertion', edgeCases, doc => doc.applyOperations([{ type: 'replace', target: { exactText: edgeText(doc)[1] }, modified: 'The counterparty proposed short-term tenure.' }], { author: 'Agent', atomic: true, existingRevisions: 'slice-cross-author' }), five);
await run('23-edge-edit-spanning-comment-paragraph', edgeCases, edgeReplace(4, 'Spanning SECOND paragraph.'), five);
await run('24-edge-edit-hyperlink-paragraph', edgeCases, edgeReplace(5, 'Link here: example.org now.'), five);
await run('25-edge-edit-table-cell', edgeCases, edgeReplace(8, 'Cell TWO'), five);
await run('26-edge-reply-to-spanning-comment', edgeCases, doc => apply(doc, [replyTo(3, 'Reply on a two-paragraph range')]), { ancestors: [0, 0, 0, 3, 0, 0], commentsDone: [false, false, false, false, false, false] });
await run('27-edge-reply-in-table-cell', edgeCases, doc => apply(doc, [replyTo(5, 'Reply in a cell')]), { ancestors: [0, 0, 0, 0, 0, 5], commentsDone: [false, false, false, false, false, false] });

// Files damaged by earlier versions of this package: the next save must repair them.
const damaged = withContentType(fixture, WORD_TYPE, BAD_TYPE);
await run('06-repair-damaged-on-edit', damaged, edit);

// Controls: reproduce the original bug. Word must refuse these, or this check proves nothing.
write('expect-fail-control-bad-extended-type.docx', damaged);

console.log(`Wrote cases to ${outDir}`);
