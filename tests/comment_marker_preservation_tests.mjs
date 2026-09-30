/**
 * Comment anchors must survive edits to the text they annotate.
 *
 * Regression: a mid-paragraph edit of a commented paragraph dropped the trailing commentRangeEnd and the
 * commentReference run (markers at the very end of the paragraph text were only emitted when the final diff
 * part was an insert), leaving the comment unanchored and the package invalid.
 */
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { buildZip } from '../scripts/lib/minimal-zip.mjs';
import { unzipEntries } from '../scripts/lib/zip-reader.mjs';
import { openDocx } from '../node/index.js';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const markersOf = xml => [...xml.matchAll(/<w:(commentRangeStart|commentRangeEnd|commentReference) w:id="(\d+)"/g)].map(m => `${m[1]}:${m[2]}`).sort();
const documentOf = result => unzipEntries(result.toBuffer()).get('word/document.xml').toString('utf8');
const replace = (from, to) => ({ type: 'replace', target: { exactText: from }, modified: to });

async function assertPreserved(label, bytes, op) {
    const before = markersOf(unzipEntries(bytes).get('word/document.xml').toString('utf8'));
    const result = await openDocx(bytes).applyOperations([op], { author: 'Agent', atomic: true });
    assert.equal(result.status, 'ok', `${label}: ${result.error?.message}`);
    assert.equal(result.written, true, label);
    const xml = documentOf(result);
    assert.deepEqual(markersOf(xml), before, `${label}: comment markers changed`);
    // Every range end comes after its start.
    for (const id of new Set(before.map(m => m.split(':')[1]))) {
        assert.ok(xml.indexOf(`<w:commentRangeStart w:id="${id}"`) < xml.indexOf(`<w:commentRangeEnd w:id="${id}"`), `${label}: comment ${id} end precedes start`);
        assert.ok(xml.indexOf(`<w:commentRangeEnd w:id="${id}"`) < xml.indexOf(`<w:commentReference w:id="${id}"`), `${label}: comment ${id} reference precedes end`);
    }
}

// --- Word-authored documents ---------------------------------------------------------------
const wordThreads = readFileSync(new URL('./fixtures/word-authored/threaded-comments.docx', import.meta.url));
const wordMulti = readFileSync(new URL('./fixtures/word-authored/multi-paragraph-thread.docx', import.meta.url));
const p1 = 'The Supplier shall deliver the goods within thirty days.';
const p2 = 'Payment is due upon receipt of invoice.';
await assertPreserved('thread paragraph, middle edit', wordThreads, replace(p1, 'The Supplier shall deliver the goods within twenty days.'));
await assertPreserved('thread paragraph, tail edit', wordThreads, replace(p1, `${p1.slice(0, -1)}, time being of the essence.`));
await assertPreserved('thread paragraph, head edit', wordThreads, replace(p1, 'The Vendor shall deliver the goods within thirty days.'));
await assertPreserved('single comment paragraph, middle edit', wordThreads, replace(p2, 'Payment is due within thirty days of invoice.'));
await assertPreserved('single comment paragraph, tail deletion', wordThreads, replace(p2, 'Payment is due.'));
await assertPreserved('single comment paragraph, full rewrite', wordThreads, replace(p2, 'Completely different payment sentence.'));
await assertPreserved('multi-paragraph thread, middle edit', wordMulti, replace('Alpha paragraph.', 'Alpha revised paragraph.'));

// --- hand-built shapes -----------------------------------------------------------------------
const contentTypes = '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/></Types>';
const rels = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/></Relationships>';
const comments = `<w:comments xmlns:w="${W}">${[7, 8].map(id => `<w:comment w:id="${id}" w:author="X"><w:p><w:r><w:t>c${id}</w:t></w:r></w:p></w:comment>`).join('')}</w:comments>`;
const build = paragraph => buildZip([
    { name: '[Content_Types].xml', data: contentTypes },
    { name: 'word/document.xml', data: `<w:document xmlns:w="${W}"><w:body>${paragraph}<w:sectPr/></w:body></w:document>` },
    { name: 'word/_rels/document.xml.rels', data: rels },
    { name: 'word/comments.xml', data: comments }
]);
const ref = id => `<w:r><w:commentReference w:id="${id}"/></w:r>`;
const T = 'Payment is due upon receipt of invoice.';
const M = 'Payment is due within thirty days of invoice.';

await assertPreserved('plain comment, middle edit', build(`<w:p><w:commentRangeStart w:id="7"/><w:r><w:t>${T}</w:t></w:r><w:commentRangeEnd w:id="7"/>${ref(7)}</w:p>`), replace(T, M));
await assertPreserved('two runs, middle edit', build(`<w:p><w:commentRangeStart w:id="7"/><w:r><w:t>Payment is due </w:t></w:r><w:r><w:t>upon receipt of invoice.</w:t></w:r><w:commentRangeEnd w:id="7"/>${ref(7)}</w:p>`), replace(T, M));
await assertPreserved('two comments ending at the same point', build(`<w:p><w:commentRangeStart w:id="7"/><w:commentRangeStart w:id="8"/><w:r><w:t>${T}</w:t></w:r><w:commentRangeEnd w:id="7"/>${ref(7)}<w:commentRangeEnd w:id="8"/>${ref(8)}</w:p>`), replace(T, M));
await assertPreserved('comment covering only the middle words', build(`<w:p><w:r><w:t>Payment is </w:t></w:r><w:commentRangeStart w:id="7"/><w:r><w:t>due upon receipt</w:t></w:r><w:commentRangeEnd w:id="7"/>${ref(7)}<w:r><w:t> of invoice.</w:t></w:r></w:p>`), replace(T, M));
await assertPreserved('comment on the trailing words, head edit', build(`<w:p><w:r><w:t>Payment is due </w:t></w:r><w:commentRangeStart w:id="7"/><w:r><w:t>upon receipt of invoice.</w:t></w:r><w:commentRangeEnd w:id="7"/>${ref(7)}</w:p>`), replace(T, 'Fees are due upon receipt of invoice.'));

console.log('PASS: comment markers survive edits to commented text');
