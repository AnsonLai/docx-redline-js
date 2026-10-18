import './setup-xml-provider.mjs';

// mergeCommentParts: the parts-level (string in, string out) comment merge for hosts
// that hold package parts outside a ZIP. Unit cases, then equivalence with the parts
// openDocx writes for the same batch on a Word-authored threaded-comments package.
import assert from 'assert/strict';
import { readFileSync } from 'node:fs';
import { inspectDocumentParts, mergeCommentParts, openDocx } from '../index.js';
import { applyOperationsToDocumentXml } from '../services/standalone-operation-runner.js';
import { unzipDocx } from '../document/zip-archive.js';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"';
const comment = (id, text, paraId) => `<w:comment w:id="${id}" w:author="Reviewer" w:date="2026-01-01T00:00:00Z"><w:p w14:paraId="${paraId}"><w:r><w:t>${text}</w:t></w:r></w:p></w:comment>`;
const comments = (...items) => `<w:comments ${W}>${items.join('')}</w:comments>`;
const ids = (...pairs) => `<w16cid:commentsIds xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid">${pairs.map(([paraId, durable]) => `<w16cid:commentId w16cid:paraId="${paraId}" w16cid:durableId="${durable}"/>`).join('')}</w16cid:commentsIds>`;
const extensible = (...durables) => `<w16cex:commentsExtensible xmlns:w16cex="http://schemas.microsoft.com/office/word/2018/wordml/cex">${durables.map(durable => `<w16cex:commentExtensible w16cex:durableId="${durable}" w16cex:dateUtc="2026-01-01T00:00:00Z"/>`).join('')}</w16cex:commentsExtensible>`;
const extended = (...paraIds) => `<w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">${paraIds.map(paraId => `<w15:commentEx w15:paraId="${paraId}" w15:done="0"/>`).join('')}</w15:commentsEx>`;
const count = (xml, tag) => (xml?.match(new RegExp(`<[\\w]+:${tag}\\b`, 'g')) || []).length;
const NULL_PARTS = { commentsXml: null, commentsExtendedXml: null, commentsIdsXml: null, commentsExtensibleXml: null };

// ---------------------------------------------------------------------------------------------------------
// 1. Unit cases
// ---------------------------------------------------------------------------------------------------------
{
    // Nothing to merge.
    assert.deepEqual(mergeCommentParts(), { status: 'ok', hasChanges: false, parts: NULL_PARTS, changedParts: [], addedParts: [] });
    const source = { commentsXml: comments(comment(0, 'Old', 'AAAA0001')) };
    const untouched = mergeCommentParts(source, { commentsXml: null, commentsExtendedXml: null });
    assert.equal(untouched.hasChanges, false);
    assert.equal(untouched.parts.commentsXml, source.commentsXml, 'unchanged parts are returned as given');
}
{
    // A package with no comments gains comments.xml; sibling parts are never created.
    const result = mergeCommentParts({}, { commentsXml: comments(comment(1, 'New', 'BBBB0001')) });
    assert.equal(result.status, 'ok');
    assert.deepEqual(result.changedParts, ['comments']);
    assert.deepEqual(result.addedParts, ['comments']);
    assert.equal(result.parts.commentsIdsXml, null);
    assert.equal(result.parts.commentsExtensibleXml, null);
}
{
    // Merge appends, and commentsIds / commentsExtensible gain entries for the new comment.
    const source = {
        commentsXml: comments(comment(0, 'Old', 'AAAA0001')),
        commentsIdsXml: ids(['AAAA0001', '11111111']),
        commentsExtensibleXml: extensible('11111111')
    };
    const frozen = JSON.stringify(source);
    const result = mergeCommentParts(source, { commentsXml: comments(comment(1, 'New', 'BBBB0001')), commentsXmlMode: 'merge' });
    assert.equal(result.status, 'ok', JSON.stringify(result.error));
    assert.deepEqual(result.changedParts, ['comments', 'commentsIds', 'commentsExtensible']);
    assert.deepEqual(result.addedParts, []);
    assert.equal(count(result.parts.commentsXml, 'comment'), 2);
    assert.match(result.parts.commentsXml, /Old[\s\S]*New/, 'existing comments stay first');
    assert.equal(count(result.parts.commentsIdsXml, 'commentId'), 2);
    assert.match(result.parts.commentsIdsXml, /w16cid:paraId="BBBB0001"/);
    assert.equal(count(result.parts.commentsExtensibleXml, 'commentExtensible'), 2);
    assert.equal(JSON.stringify(source), frozen, 'inputs are not mutated');
    // The default mode is merge.
    assert.equal(count(mergeCommentParts(source, { commentsXml: comments(comment(1, 'New', 'BBBB0001')) }).parts.commentsXml, 'comment'), 2);
}
{
    // Replace swaps the part, and sibling entries for comments that are gone are removed.
    const source = {
        commentsXml: comments(comment(0, 'Old', 'AAAA0001'), comment(1, 'Gone', 'AAAA0002')),
        commentsIdsXml: ids(['AAAA0001', '11111111'], ['AAAA0002', '22222222']),
        commentsExtensibleXml: extensible('11111111', '22222222')
    };
    const result = mergeCommentParts(source, { commentsXml: comments(comment(0, 'Old', 'AAAA0001')), commentsXmlMode: 'replace' });
    assert.equal(result.status, 'ok', JSON.stringify(result.error));
    assert.equal(count(result.parts.commentsXml, 'comment'), 1);
    assert.equal(count(result.parts.commentsIdsXml, 'commentId'), 1);
    assert.doesNotMatch(result.parts.commentsIdsXml, /AAAA0002/);
    assert.equal(count(result.parts.commentsExtensibleXml, 'commentExtensible'), 1);
}
{
    // Refusals return the input parts unchanged and never throw.
    const source = { commentsXml: comments(comment(0, 'Old', 'AAAA0001')), commentsExtendedXml: extended('AAAA0001') };
    const duplicate = mergeCommentParts(source, { commentsXml: comments(comment(0, 'Clash', 'CCCC0001')) });
    assert.equal(duplicate.status, 'error');
    assert.equal(duplicate.error.code, 'DUPLICATE_COMMENT_ID');
    assert.equal(duplicate.error.commentId, '0');
    assert.equal(duplicate.hasChanges, false);
    assert.equal(duplicate.parts.commentsXml, source.commentsXml);

    assert.equal(mergeCommentParts(source, { commentsXml: '<w:comments' }).error.code, 'PARSE_ERROR', 'malformed incoming comments');
    assert.equal(mergeCommentParts(source, { commentsXml: '<w:comments', commentsXmlMode: 'replace' }).error.code, 'PARSE_ERROR', 'malformed replacement');
    assert.equal(mergeCommentParts({ commentsXml: '<broken' }, { commentsXml: comments(comment(1, 'New', 'BBBB0001')) }).error.code, 'PARSE_ERROR', 'malformed existing comments');
    assert.equal(mergeCommentParts({}, { commentsExtendedXml: '<w15:commentsEx' }).error.code, 'PARSE_ERROR', 'malformed thread metadata');

    // A commentsExtended delta cannot be merged into existing thread metadata.
    const delta = mergeCommentParts(source, { commentsXml: comments(comment(1, 'New', 'BBBB0001')), commentsExtendedXml: extended('BBBB0001') });
    assert.equal(delta.status, 'error');
    assert.equal(delta.error.code, 'UNSUPPORTED_EXISTING_COMMENT_THREAD_METADATA');
    assert.equal(delta.error.part, 'commentsExtended');
    assert.deepEqual(delta.parts, { ...NULL_PARTS, ...source }, 'nothing is merged, not even comments.xml');
}
{
    // commentsExtended: replace swaps it; a delta is fine when there is none yet.
    const source = { commentsXml: comments(comment(0, 'Old', 'AAAA0001')), commentsExtendedXml: extended('AAAA0001') };
    const replaced = mergeCommentParts(source, { commentsExtendedXml: extended('AAAA0001', 'BBBB0001'), commentsExtendedXmlMode: 'replace' });
    assert.equal(replaced.status, 'ok');
    assert.deepEqual(replaced.changedParts, ['commentsExtended']);
    assert.deepEqual(replaced.addedParts, []);
    const added = mergeCommentParts({ commentsXml: source.commentsXml }, { commentsExtendedXml: extended('AAAA0001') });
    assert.equal(added.status, 'ok');
    assert.deepEqual(added.addedParts, ['commentsExtended']);
}

// ---------------------------------------------------------------------------------------------------------
// 2. Equivalence with openDocx on a Word-authored threaded package
// ---------------------------------------------------------------------------------------------------------
const bytes = new Uint8Array(readFileSync(new URL('./fixtures/word-authored/threaded-comments.docx', import.meta.url)));
const entries = unzipDocx(bytes);
const decoder = new TextDecoder();
const part = (map, name) => (map.get(name) ? decoder.decode(map.get(name)) : null);
const sourceParts = {
    commentsXml: part(entries, 'word/comments.xml'),
    commentsExtendedXml: part(entries, 'word/commentsExtended.xml'),
    commentsIdsXml: part(entries, 'word/commentsIds.xml'),
    commentsExtensibleXml: part(entries, 'word/commentsExtensible.xml')
};
const documentXml = part(entries, 'word/document.xml');
// Generated dates differ between the two runs; everything else must match.
const undated = xml => xml?.replace(/(w:date|w16cex:dateUtc)="[^"]*"/g, '$1=""') ?? null;
const nextCommentId = () => {
    const found = `${documentXml} ${sourceParts.commentsXml}`.match(/(?:w:)?id=["'](\d+)["']/g) || [];
    let next = found.reduce((max, token) => Math.max(max, Number(token.match(/\d+/)[0])), 0) + 1;
    return () => next++;
};
const target = inspectDocumentParts({ documentXml }).paragraphs.find(p => p.exactText.trim().length > 3);
const handle = { index: target.index, exactText: target.exactText, paragraphId: target.paragraphId, fingerprint: target.fingerprint };
const batches = {
    'new comment': [{ type: 'comment', target: handle, textToComment: target.exactText.trim().split(/\s+/)[0], commentContent: 'Please review.' }],
    'reply': [{ type: 'comment_reply', parentCommentId: 0, commentContent: 'Agreed.', author: 'Agent' }],
    'resolve': [{ type: 'comment_resolve', commentId: 0, resolved: true }]
};
for (const [name, operations] of Object.entries(batches)) {
    const facade = await openDocx(bytes).applyOperations(operations, { author: 'Agent' });
    assert.equal(facade.status, 'ok', `${name}: ${JSON.stringify(facade.error ?? facade.results)}`);
    const written = unzipDocx(facade.toUint8Array());

    const runner = await applyOperationsToDocumentXml(documentXml, operations, 'Agent',
        { commentsXml: sourceParts.commentsXml, commentsExtendedXml: sourceParts.commentsExtendedXml },
        { strictTargets: true, commentIdAllocator: nextCommentId() });
    assert.equal(runner.status, 'ok', `${name}: ${JSON.stringify(runner.error ?? runner.results)}`);
    const merged = mergeCommentParts(sourceParts, runner);
    assert.equal(merged.status, 'ok', `${name}: ${JSON.stringify(merged.error)}`);
    assert.deepEqual(merged.addedParts, [], `${name}: the package already has every comment part`);
    assert.deepEqual(merged.changedParts, { 'new comment': ['comments', 'commentsIds', 'commentsExtensible'], reply: ['comments', 'commentsExtended', 'commentsIds', 'commentsExtensible'], resolve: ['commentsExtended'] }[name], `${name}: changed parts`);
    for (const [key, path] of [['commentsXml', 'word/comments.xml'], ['commentsExtendedXml', 'word/commentsExtended.xml'],
        ['commentsIdsXml', 'word/commentsIds.xml'], ['commentsExtensibleXml', 'word/commentsExtensible.xml']]) {
        assert.equal(undated(merged.parts[key]), undated(part(written, path)), `${name}: ${path} matches openDocx`);
    }
    const inspected = inspectDocumentParts({ documentXml: runner.documentXml, commentsXml: merged.parts.commentsXml, commentsExtendedXml: merged.parts.commentsExtendedXml });
    const facadeInspected = openDocx(facade.toUint8Array()).inspect();
    assert.deepEqual(inspected.comments.map(c => [c.author, c.text, c.done ?? false, c.parentCommentId ?? null]),
        facadeInspected.comments.map(c => [c.author, c.text, c.done ?? false, c.parentCommentId ?? null]), `${name}: same comments and threads`);
}

console.log('comment_parts_merge_tests: all passed');
