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
async function run(name, source, fn, expectDone) {
    const doc = openDocx(source);
    const result = await fn(doc);
    if (result.status !== 'ok' || !result.written) throw new Error(`${name}: ${result.error?.message || 'not written'}`);
    write(`expect-open-${name}.docx`, result.toBuffer());
    if (expectDone) writeFileSync(join(outDir, `expect-open-${name}.expect.json`), JSON.stringify({ commentsDone: expectDone }));
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

// Files damaged by earlier versions of this package: the next save must repair them.
const damaged = withContentType(fixture, WORD_TYPE, BAD_TYPE);
await run('06-repair-damaged-on-edit', damaged, edit);

// Controls: reproduce the original bug. Word must refuse these, or this check proves nothing.
write('expect-fail-control-bad-extended-type.docx', damaged);

console.log(`Wrote cases to ${outDir}`);
