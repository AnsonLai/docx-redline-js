/**
 * Package content-type contract.
 *
 * Word rejects ("unreadable content" / repair prompt) a package whose [Content_Types].xml
 * labels a part with a type Word does not recognize. Our Word COM tests only run on Windows
 * and Word silently repairs on other paths, so this platform-independent test pins the
 * values Word itself writes. The oracle table below is deliberately NOT imported from
 * source: if the source constant drifts, this test must fail.
 *
 * Regression: commentsExtended.xml was relabeled application/vnd.ms-word.commentsExtended+xml
 * on every save of any document that already had comments.
 */
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { buildZip } from '../scripts/lib/minimal-zip.mjs';
import { unzipEntries } from '../scripts/lib/zip-reader.mjs';
import { openDocx } from '../node/index.js';
import { MemoryZip, unzipDocx } from '../document/zip-archive.js';
import { repairKnownContentTypes, validateDocxPackage } from '../services/standalone-docx-plumbing.js';

const WORD_OFFICE_ML = 'application/vnd.openxmlformats-officedocument.wordprocessingml';
// Word-authored values, keyed by lower-cased part name.
const WORD_CONTENT_TYPES = new Map([
    ['/word/document.xml', `${WORD_OFFICE_ML}.document.main+xml`],
    ['/word/comments.xml', `${WORD_OFFICE_ML}.comments+xml`],
    ['/word/commentsextended.xml', `${WORD_OFFICE_ML}.commentsExtended+xml`],
    ['/word/commentsids.xml', `${WORD_OFFICE_ML}.commentsIds+xml`],
    ['/word/commentsextensible.xml', `${WORD_OFFICE_ML}.commentsExtensible+xml`],
    ['/word/numbering.xml', `${WORD_OFFICE_ML}.numbering+xml`]
]);
const BAD_EXTENDED_TYPE = 'application/vnd.ms-word.commentsExtended+xml';
const GOOD_EXTENDED_TYPE = WORD_CONTENT_TYPES.get('/word/commentsextended.xml');

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const W15 = 'http://schemas.microsoft.com/office/word/2012/wordml';

const documentXml = `<w:document xmlns:w="${W}"><w:body><w:p><w:r><w:t>Hello world</w:t></w:r></w:p><w:p><w:commentRangeStart w:id="8"/><w:r><w:t>Anchored text</w:t></w:r><w:commentRangeEnd w:id="8"/><w:r><w:commentReference w:id="8"/></w:r></w:p><w:p><w:ins w:id="90" w:author="Counterparty" w:date="2026-01-01T00:00:00Z"><w:r><w:t>Proposed</w:t></w:r></w:ins></w:p><w:sectPr/></w:body></w:document>`;
const commentsXml = `<w:comments xmlns:w="${W}" xmlns:w14="${W14}"><w:comment w:id="8" w:author="Counterparty"><w:p w14:paraId="1A2B3C4D"><w:r><w:t>Please review</w:t></w:r></w:p></w:comment></w:comments>`;
const extendedXml = `<w15:commentsEx xmlns:w15="${W15}"><w15:commentEx w15:paraId="1A2B3C4D" w15:done="0"/></w15:commentsEx>`;
const relsXml = `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/><Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2011/relationships/commentsExtended" Target="commentsExtended.xml"/></Relationships>`;

function contentTypesXml(extendedType) {
    return `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${WORD_CONTENT_TYPES.get('/word/document.xml')}"/><Override PartName="/word/comments.xml" ContentType="${WORD_CONTENT_TYPES.get('/word/comments.xml')}"/><Override PartName="/word/commentsExtended.xml" ContentType="${extendedType}"/></Types>`;
}

function buildInput(extendedType = GOOD_EXTENDED_TYPE) {
    return buildZip([
        { name: '[Content_Types].xml', data: contentTypesXml(extendedType) },
        { name: 'word/document.xml', data: documentXml },
        { name: 'word/_rels/document.xml.rels', data: relsXml },
        { name: 'word/comments.xml', data: commentsXml },
        { name: 'word/commentsExtended.xml', data: extendedXml }
    ]);
}

function overrides(entries) {
    const xml = entries.get('[Content_Types].xml').toString('utf8');
    return [...xml.matchAll(/<Override\s+[^>]*?PartName="([^"]+)"[^>]*?ContentType="([^"]+)"/g)]
        .map(match => ({ partName: match[1], contentType: match[2] }));
}

/** Every override we know Word's value for must match it; no ms-word-prefixed labels allowed. */
function assertContentTypeContract(entries, label) {
    for (const { partName, contentType } of overrides(entries)) {
        const expected = WORD_CONTENT_TYPES.get(partName.toLowerCase());
        if (expected) assert.equal(contentType, expected, `${label}: ${partName} content type`);
        assert.ok(!contentType.startsWith('application/vnd.ms-word.'), `${label}: ${partName} uses unrecognized type ${contentType}`);
    }
}

const commentRelatedParts = ['[Content_Types].xml', 'word/_rels/document.xml.rels', 'word/comments.xml', 'word/commentsExtended.xml'];

// 1. An edit that never touches comments leaves every comment-related part byte-identical.
{
    const input = buildInput();
    const before = unzipEntries(input);
    const result = await openDocx(input).applyOperations(
        [{ type: 'replace', target: { exactText: 'Hello world' }, modified: 'Hello reliable world' }],
        { author: 'Agent', atomic: true }
    );
    assert.equal(result.status, 'ok', result.error?.message);
    assert.equal(result.written, true);
    const after = unzipEntries(result.toBuffer());
    assertContentTypeContract(after, 'unrelated edit');
    for (const name of commentRelatedParts) {
        assert.deepEqual(after.get(name), before.get(name), `${name} must be byte-identical after an edit that does not touch comments`);
    }
    assert.ok(!result.artifactsChanged.some(name => commentRelatedParts.includes(name)), `artifactsChanged reported comment parts: ${result.artifactsChanged}`);
}

// 2. Adding a comment (writes comments.xml, not the extended part) still keeps Word's type.
{
    const result = await openDocx(buildInput()).applyOperations(
        [{ type: 'comment', target: { exactText: 'Hello world' }, commentContent: 'Note', author: 'Agent' }],
        { author: 'Agent', atomic: true }
    );
    assert.equal(result.status, 'ok', result.error?.message);
    assertContentTypeContract(unzipEntries(result.toBuffer()), 'add comment');
}

// 3. A reply writes the extended part and must label it with Word's type.
{
    const result = await openDocx(buildInput()).applyOperations(
        [{ type: 'comment_reply', parentCommentId: 8, commentContent: 'Reply', author: 'Agent' }],
        { author: 'Agent', atomic: true }
    );
    assert.equal(result.status, 'ok', result.error?.message);
    assertContentTypeContract(unzipEntries(result.toBuffer()), 'comment reply');
}

// 3b. A reply on a document with no extended part at all creates it with Word's type.
{
    const bare = buildZip([
        { name: '[Content_Types].xml', data: contentTypesXml(GOOD_EXTENDED_TYPE).replace(/<Override PartName="\/word\/commentsExtended\.xml"[^>]*\/>/, '') },
        { name: 'word/document.xml', data: documentXml },
        { name: 'word/_rels/document.xml.rels', data: relsXml.replace(/<Relationship Id="rId2"[^>]*\/>/, '') },
        { name: 'word/comments.xml', data: commentsXml }
    ]);
    const result = await openDocx(bare).applyOperations(
        [{ type: 'comment_reply', parentCommentId: 8, commentContent: 'Reply', author: 'Agent' }],
        { author: 'Agent', atomic: true }
    );
    assert.equal(result.status, 'ok', result.error?.message);
    const after = unzipEntries(result.toBuffer());
    assert.ok(overrides(after).some(o => o.partName === '/word/commentsExtended.xml'), 'extended override was created');
    assertContentTypeContract(after, 'reply creating extended part');
}

// 4. Files already damaged by earlier versions are repaired on the next save, on every write path.
{
    const damaged = buildInput(BAD_EXTENDED_TYPE);
    const edit = await openDocx(damaged).applyOperations(
        [{ type: 'replace', target: { exactText: 'Hello world' }, modified: 'Hello reliable world' }],
        { author: 'Agent', atomic: true }
    );
    assert.equal(edit.status, 'ok', edit.error?.message);
    assert.equal(edit.written, true);
    assertContentTypeContract(unzipEntries(edit.toBuffer()), 'repair via applyOperations');
    assert.ok(edit.artifactsChanged.includes('[Content_Types].xml'));
    assert.deepEqual(edit.validation.generatedIssues, []);

    const accepted = await openDocx(damaged).resolveRevisions('accept', { allAuthors: true });
    assert.equal(accepted.status, 'ok', accepted.error?.message);
    assert.equal(accepted.written, true);
    assertContentTypeContract(unzipEntries(accepted.toBuffer()), 'repair via resolveRevisions');

    const deleted = await openDocx(damaged).deleteComments({ author: 'Counterparty' });
    assert.equal(deleted.status, 'ok', deleted.error?.message);
    assert.equal(deleted.written, true);
    assertContentTypeContract(unzipEntries(deleted.toBuffer()), 'repair via deleteComments');
}

// 5. The validator rejects the unrecognized type instead of blessing it.
{
    await validateDocxPackage(new MemoryZip(unzipDocx(buildInput())));
    await assert.rejects(
        validateDocxPackage(new MemoryZip(unzipDocx(buildInput(BAD_EXTENDED_TYPE)))),
        /commentsExtended CT override has wrong content type/
    );
}

// 6. repairKnownContentTypes is a no-op (no rewrite) on a correct package, and reports what it fixed otherwise.
{
    const good = new MemoryZip(unzipDocx(buildInput()));
    const goodBefore = Buffer.from(good.entries?.get?.('[Content_Types].xml') || await good.file('[Content_Types].xml').async('string'));
    assert.deepEqual(await repairKnownContentTypes(good), []);
    const goodAfter = Buffer.from(good.entries?.get?.('[Content_Types].xml') || await good.file('[Content_Types].xml').async('string'));
    assert.deepEqual(goodAfter, goodBefore);

    const bad = new MemoryZip(unzipDocx(buildInput(BAD_EXTENDED_TYPE)));
    const repairs = await repairKnownContentTypes(bad);
    assert.equal(repairs.length, 1);
    assert.equal(repairs[0].from, BAD_EXTENDED_TYPE);
    assert.equal(repairs[0].to, GOOD_EXTENDED_TYPE);
}

// 7. Real Word oracle. tests/fixtures/word-authored/threaded-comments.docx was saved by Word itself
// (scripts/generate-word-comment-thread-fixture.ps1): a comment with a reply plus a resolved comment.
// Its [Content_Types].xml is ground truth, so this cannot drift with our own constants.
{
    const wordBytes = readFileSync(new URL('./fixtures/word-authored/threaded-comments.docx', import.meta.url));
    const wordEntries = unzipEntries(wordBytes);
    const wordTypes = new Map(overrides(wordEntries).map(o => [o.partName, o.contentType]));

    // The hand-written table above must agree with what Word actually wrote.
    for (const [partName, contentType] of wordTypes) {
        const expected = WORD_CONTENT_TYPES.get(partName.toLowerCase());
        if (expected) assert.equal(contentType, expected, `oracle table disagrees with Word for ${partName}`);
    }
    assert.ok(wordTypes.has('/word/commentsExtended.xml'), 'Word fixture carries a commentsExtended part');
    assert.ok(wordTypes.has('/word/commentsIds.xml') && wordTypes.has('/word/commentsExtensible.xml'), 'Word fixture carries the sibling comment parts');

    const scenarios = {
        edit: doc => doc.applyOperations([{ type: 'replace', target: { exactText: 'This agreement is governed by local law.' }, modified: 'This agreement is governed by New York law.' }], { author: 'Agent', atomic: true }),
        reply: doc => doc.applyOperations([{ type: 'comment_reply', parentCommentId: 0, commentContent: 'Agreed.', author: 'Agent' }], { author: 'Agent', atomic: true }),
        deleteComments: doc => doc.deleteComments({ author: 'Counterparty' })
    };
    for (const [name, run] of Object.entries(scenarios)) {
        const result = await run(openDocx(wordBytes));
        assert.equal(result.status, 'ok', `${name}: ${result.error?.message}`);
        const after = unzipEntries(result.toBuffer());
        const afterTypes = new Map(overrides(after).map(o => [o.partName, o.contentType]));
        for (const [partName, contentType] of afterTypes) {
            if (wordTypes.has(partName)) assert.equal(contentType, wordTypes.get(partName), `${name}: ${partName} differs from what Word wrote`);
        }
        assertContentTypeContract(after, `word fixture / ${name}`);
    }

    // An edit that does not touch comments must leave every comment-related part Word wrote untouched,
    // including the sibling parts we do not otherwise understand yet.
    const edited = unzipEntries((await scenarios.edit(openDocx(wordBytes))).toBuffer());
    for (const name of ['[Content_Types].xml', 'word/_rels/document.xml.rels', 'word/comments.xml', 'word/commentsExtended.xml', 'word/commentsIds.xml', 'word/commentsExtensible.xml']) {
        assert.deepEqual(edited.get(name), wordEntries.get(name), `${name} must be byte-identical after an unrelated edit`);
    }
}

console.log('PASS: package content-type contract');
