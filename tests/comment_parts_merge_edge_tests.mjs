import './setup-xml-provider.mjs';

// mergeCommentParts edge cases: input hardening, comment shapes, idempotence, sibling-part
// integrity, and part-for-part equivalence with openDocx on more package shapes (no comment
// parts, comments.xml only, sequential batches, a mixed batch).
import assert from 'assert/strict';
import { readFileSync } from 'node:fs';
import { inspectDocumentParts, mergeCommentParts, openDocx } from '../index.js';
import { applyOperationsToDocumentXml } from '../services/standalone-operation-runner.js';
import { unzipDocx } from '../document/zip-archive.js';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"';
const para = (text, paraId) => `<w:p${paraId ? ` w14:paraId="${paraId}"` : ''}><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
const comment = (id, body) => `<w:comment w:id="${id}" w:author="Reviewer" w:date="2026-01-01T00:00:00Z">${body}</w:comment>`;
const comments = (...items) => `<w:comments ${W}>${items.join('')}</w:comments>`;
const ids = (...pairs) => `<w16cid:commentsIds xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid">${pairs.map(([paraId, durable]) => `<w16cid:commentId w16cid:paraId="${paraId}" w16cid:durableId="${durable}"/>`).join('')}</w16cid:commentsIds>`;
const extensible = (...durables) => `<w16cex:commentsExtensible xmlns:w16cex="http://schemas.microsoft.com/office/word/2018/wordml/cex">${durables.map(durable => `<w16cex:commentExtensible w16cex:durableId="${durable}" w16cex:dateUtc="2026-01-01T00:00:00Z"/>`).join('')}</w16cex:commentsExtensible>`;
const attrValues = (xml, attr) => [...(xml || '').matchAll(new RegExp(`${attr}="([^"]*)"`, 'g'))].map(m => m[1]);
const base = () => ({
    commentsXml: comments(comment(0, para('Old', 'AAAA0001'))),
    commentsIdsXml: ids(['AAAA0001', '11111111']),
    commentsExtensibleXml: extensible('11111111')
});
const incoming = (...items) => ({ commentsXml: comments(...items) });

// ---------------------------------------------------------------------------------------------------------
// 1. Input hardening: refused with a code and the input unchanged, never thrown
// ---------------------------------------------------------------------------------------------------------
{
    const source = base();
    const refused = (result, code, part, given = source) => {
        assert.equal(result.status, 'error', JSON.stringify(result));
        assert.equal(result.error.code, code, JSON.stringify(result.error));
        if (part !== undefined) assert.equal(result.error.part, part, JSON.stringify(result.error));
        assert.equal(result.hasChanges, false);
        assert.deepEqual(result.changedParts, []);
        assert.equal(result.parts.commentsXml, given.commentsXml, 'input parts come back unchanged');
    };
    refused(mergeCommentParts({ ...source, commentsIdsXml: new Uint8Array([60]) }, incoming(comment(1, para('x', 'BBBB0001')))), 'INVALID_COMMENT_PARTS');
    refused(mergeCommentParts(source, { commentsXml: 42 }), 'INVALID_COMMENT_PARTS');
    refused(mergeCommentParts(source, { commentsXml: `<w:document ${W}/>` }), 'INVALID_COMMENT_PARTS', 'comments');
    refused(mergeCommentParts(source, { commentsExtendedXml: `<w:comments ${W}/>`, commentsExtendedXmlMode: 'replace' }), 'INVALID_COMMENT_PARTS', 'commentsExtended');
    const wrongExisting = { ...source, commentsXml: `<w:document ${W}/>` };
    refused(mergeCommentParts(wrongExisting, incoming(comment(1, para('x', 'BBBB0001')))), 'INVALID_COMMENT_PARTS', 'comments', wrongExisting);
    refused(mergeCommentParts({ ...source, commentsIdsXml: '<w16cid:commentsIds' }, incoming(comment(1, para('x', 'BBBB0001')))), 'PARSE_ERROR', 'commentsIds');
    refused(mergeCommentParts({ ...source, commentsExtensibleXml: '<nope' }, incoming(comment(1, para('x', 'BBBB0001')))), 'PARSE_ERROR', 'commentsExtensible');
    refused(mergeCommentParts(source, incoming(comment(5, para('a', 'BBBB0001')), comment(5, para('b', 'BBBB0002')))), 'DUPLICATE_COMMENT_ID', 'comments');
    // Duplicates inside the incoming part are refused in every mode, not only when merging into an existing part.
    refused(mergeCommentParts(source, { ...incoming(comment(5, para('a', 'BBBB0001')), comment(5, para('b', 'BBBB0002'))), commentsXmlMode: 'replace' }), 'DUPLICATE_COMMENT_ID', 'comments');
    const noComments = {};
    refused(mergeCommentParts(noComments, incoming(comment(5, para('a', 'BBBB0001')), comment(5, para('b', 'BBBB0002')))), 'DUPLICATE_COMMENT_ID', 'comments', { commentsXml: null });
    // Sibling parts are only validated when they would be touched.
    assert.equal(mergeCommentParts({ ...source, commentsIdsXml: '<broken' }, {}).status, 'ok');
}

// ---------------------------------------------------------------------------------------------------------
// 2. No-ops, modes and inputs
// ---------------------------------------------------------------------------------------------------------
{
    // A Word-style part (declaration, line breaks) would not survive a reserialization byte for byte.
    const declaration = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';
    const source = { ...base(), commentsXml: declaration + base().commentsXml.replace('><w:comment ', '>\r\n  <w:comment ') };
    const empty = mergeCommentParts(source, { commentsXml: comments() });
    assert.equal(empty.status, 'ok');
    assert.equal(empty.hasChanges, false, 'an incoming part with no comments changes nothing');
    assert.equal(empty.parts.commentsXml, source.commentsXml, 'and keeps the existing string byte for byte');
    const emptied = mergeCommentParts(source, { commentsXml: comments(), commentsXmlMode: 'replace' });
    assert.deepEqual(emptied.changedParts, ['comments', 'commentsIds', 'commentsExtensible'], 'replace with no comments empties the siblings too');
    assert.deepEqual(attrValues(emptied.parts.commentsIdsXml, 'w16cid:paraId'), []);
    // Unknown modes fall back to merge.
    const odd = mergeCommentParts(source, { commentsXml: comments(comment(1, para('New', 'BBBB0001'))), commentsXmlMode: 'Replace' });
    assert.match(odd.parts.commentsXml, /Old[\s\S]*New/);
    // Empty strings count as absent.
    assert.deepEqual(mergeCommentParts({ commentsXml: '' }, { commentsXml: '' }).parts.commentsXml, null);
}

// ---------------------------------------------------------------------------------------------------------
// 3. Comment shapes
// ---------------------------------------------------------------------------------------------------------
{
    // A multi-paragraph comment is keyed on its last paragraph's paraId, as Word does.
    const multi = mergeCommentParts(base(), incoming(comment(1, para('First', 'BBBB0001') + para('Last', 'BBBB0002'))));
    assert.equal(multi.status, 'ok', JSON.stringify(multi.error));
    assert.deepEqual(attrValues(multi.parts.commentsIdsXml, 'w16cid:paraId'), ['AAAA0001', 'BBBB0002']);
    // A comment without a paraId gets no sibling entry and existing entries are kept.
    const bare = mergeCommentParts(base(), incoming(comment(1, para('No id'))));
    assert.equal(bare.status, 'ok');
    assert.deepEqual(attrValues(bare.parts.commentsIdsXml, 'w16cid:paraId'), ['AAAA0001']);
    // Special characters survive the merge.
    const text = 'Fees & <costs> "quoted" — ünïcödé 😀';
    const escaped = text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
    const special = mergeCommentParts(base(), incoming(comment(1, para(escaped, 'BBBB0001'))));
    const documentXml = `<w:document ${W}><w:body>${para('Body')}<w:sectPr/></w:body></w:document>`;
    const texts = inspectDocumentParts({ documentXml, commentsXml: special.parts.commentsXml }).comments.map(c => c.text);
    assert.deepEqual(texts, ['Old', text]);
}

// ---------------------------------------------------------------------------------------------------------
// 4. Idempotence and sibling-part integrity
// ---------------------------------------------------------------------------------------------------------
{
    const output = incoming(comment(1, para('New', 'BBBB0001')), comment(2, para('Newer', 'BBBB0002')));
    const once = mergeCommentParts(base(), output);
    assert.equal(once.status, 'ok');
    const twice = mergeCommentParts(once.parts, output);
    assert.equal(twice.error.code, 'DUPLICATE_COMMENT_ID', 'merging the same output twice is refused');
    // Every comment has exactly one durable id, every durable id one extensible entry, no orphans.
    const paraIds = attrValues(once.parts.commentsIdsXml, 'w16cid:paraId');
    const durables = attrValues(once.parts.commentsIdsXml, 'w16cid:durableId');
    assert.deepEqual(paraIds, ['AAAA0001', 'BBBB0001', 'BBBB0002']);
    assert.equal(new Set(durables).size, durables.length, 'durable ids are unique');
    assert.deepEqual(attrValues(once.parts.commentsExtensibleXml, 'w16cex:durableId').sort(), [...durables].sort());
    // Merging an already-merged result's own comments.xml by replace is stable.
    const stable = mergeCommentParts(once.parts, { commentsXml: once.parts.commentsXml, commentsXmlMode: 'replace' });
    assert.equal(stable.hasChanges, false, 'replace with identical content changes nothing');
}

// ---------------------------------------------------------------------------------------------------------
// 5. Equivalence with openDocx on more package shapes
// ---------------------------------------------------------------------------------------------------------
const decoder = new TextDecoder();
const PART_PATHS = {
    commentsXml: 'word/comments.xml',
    commentsExtendedXml: 'word/commentsExtended.xml',
    commentsIdsXml: 'word/commentsIds.xml',
    commentsExtensibleXml: 'word/commentsExtensible.xml'
};
const partsOfPackage = entries => Object.fromEntries(Object.entries(PART_PATHS)
    .map(([key, path]) => [key, entries.get(path) ? decoder.decode(entries.get(path)) : null]));
const undated = xml => xml?.replace(/(w:date|w16cex:dateUtc)="[^"]*"/g, '$1=""') ?? null;
const allocatorFor = (documentXml, commentsXml) => {
    const found = `${documentXml} ${commentsXml || ''}`.match(/(?:w:)?id=["'](\d+)["']/g) || [];
    let next = found.reduce((max, token) => Math.max(max, Number(token.match(/\d+/)[0])), 0) + 1;
    return () => next++;
};

/** Runs one batch through openDocx and through the runner + mergeCommentParts; asserts the comment parts match. */
async function assertEquivalent(name, bytes, operations, expectedAdded) {
    const facade = await openDocx(bytes).applyOperations(operations, { author: 'Agent' });
    assert.equal(facade.status, 'ok', `${name}: ${JSON.stringify(facade.error ?? facade.results)}`);
    const written = unzipDocx(facade.toUint8Array());

    const entries = unzipDocx(bytes);
    const documentXml = decoder.decode(entries.get('word/document.xml'));
    const parts = partsOfPackage(entries);
    const runner = await applyOperationsToDocumentXml(documentXml, operations, 'Agent',
        { commentsXml: parts.commentsXml, commentsExtendedXml: parts.commentsExtendedXml },
        { strictTargets: true, commentIdAllocator: allocatorFor(documentXml, parts.commentsXml) });
    assert.equal(runner.status, 'ok', `${name}: ${JSON.stringify(runner.error ?? runner.results)}`);
    const merged = mergeCommentParts(parts, runner);
    assert.equal(merged.status, 'ok', `${name}: ${JSON.stringify(merged.error)}`);
    if (expectedAdded) assert.deepEqual(merged.addedParts, expectedAdded, `${name}: addedParts`);
    const facadeParts = partsOfPackage(written);
    for (const key of Object.keys(PART_PATHS)) {
        assert.equal(undated(merged.parts[key]), undated(facadeParts[key]), `${name}: ${PART_PATHS[key]} matches openDocx`);
    }
    assert.deepEqual(Object.keys(PART_PATHS).filter(key => merged.parts[key]), Object.keys(PART_PATHS).filter(key => facadeParts[key]),
        `${name}: same parts present`);
    return facade.toUint8Array();
}

const sample = new Uint8Array(readFileSync(new URL('./fixtures/sample_doc_test.docx', import.meta.url)));
const threaded = new Uint8Array(readFileSync(new URL('./fixtures/word-authored/threaded-comments.docx', import.meta.url)));
const handleFor = (bytes, nth = 0) => {
    const p = openDocx(bytes).inspect().paragraphs.filter(item => item.exactText.trim().split(/\s+/).length > 1)[nth];
    return { target: { index: p.index, exactText: p.exactText, paragraphId: p.paragraphId, fingerprint: p.fingerprint }, word: p.exactText.trim().split(/\s+/)[0] };
};
{
    assert.equal(unzipDocx(sample).has('word/comments.xml'), false, 'fixture: sample has no comments');
    // First comment on a package without comment parts.
    const first = handleFor(sample);
    const commented = await assertEquivalent('first comment', sample,
        [{ type: 'comment', target: first.target, textToComment: first.word, commentContent: 'First note.' }], ['comments']);
    // comments.xml only: a reply adds thread metadata.
    await assertEquivalent('reply on comments-only package', commented,
        [{ type: 'comment_reply', parentCommentId: 1, commentContent: 'Reply.', author: 'Agent' }]);
    // Sequential batches: the merged parts of batch 1 feed batch 2 exactly as openDocx's written package does.
    const second = handleFor(commented, 1);
    await assertEquivalent('second batch', commented,
        [{ type: 'comment', target: second.target, textToComment: second.word, commentContent: 'Second note.' }], []);
}
{
    // A mixed batch on the Word-authored threaded package: text edit, new comment, reply and resolve.
    const edit = handleFor(threaded);
    const other = handleFor(threaded, 1);
    await assertEquivalent('mixed batch', threaded, [
        { type: 'redline', target: edit.target, modified: `${edit.target.exactText} (amended)` },
        { type: 'comment', target: other.target, textToComment: other.word, commentContent: 'New thread.' },
        { type: 'comment_reply', parentCommentId: 0, commentContent: 'Agreed.', author: 'Agent' },
        { type: 'comment_resolve', commentId: 2, resolved: false }
    ], []);
}

console.log('comment_parts_merge_edge_tests: all passed');
