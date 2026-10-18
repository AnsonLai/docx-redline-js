import './setup-xml-provider.mjs';

// compareDocumentInspections ignores host-regenerated identity (paraIds, rsids,
// revision ids/dates, run splits) and reports visible or review differences.
import assert from 'assert/strict';
import { compareDocumentInspections, inspectDocumentParts } from '../index.js';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"';
const doc = body => `<w:document ${W}><w:body>${body}<w:sectPr/></w:body></w:document>`;
const inspect = body => inspectDocumentParts({ documentXml: doc(body) });

const prepared = inspect('<w:p w14:paraId="1A2B3C4D"><w:r><w:t xml:space="preserve">Hello </w:t></w:r>'
    + '<w:ins w:id="1" w:author="Me" w:date="2026-01-01T00:00:00Z"><w:r><w:t>new world</w:t></w:r></w:ins></w:p>');
// Same content as Word writes it back: new paraId, rsids, revision id/date, and split runs.
const committed = inspect('<w:p w14:paraId="7F7F7F7F" w:rsidR="00AB12CD"><w:r w:rsidR="00AB12CD"><w:t xml:space="preserve">Hel</w:t></w:r><w:r><w:t xml:space="preserve">lo </w:t></w:r>'
    + '<w:ins w:id="99" w:author="Me" w:date="2026-10-17T08:00:00Z"><w:r><w:t xml:space="preserve">new </w:t></w:r></w:ins>'
    + '<w:ins w:id="100" w:author="Me" w:date="2026-10-17T08:00:00Z"><w:r><w:t>world</w:t></w:r></w:ins></w:p>');
assert.deepEqual(compareDocumentInspections(prepared, committed), { matches: true, differences: [] });

const untracked = inspect('<w:p><w:r><w:t>Hello new world</w:t></w:r></w:p>');
const lostReview = compareDocumentInspections(prepared, untracked);
assert.equal(lostReview.matches, false, 'losing the tracked insertion is a difference');
assert.ok(lostReview.differences.some(item => item.path === 'paragraphs[0].review'));

const extra = compareDocumentInspections(prepared, inspect('<w:p><w:r><w:t>Hello </w:t></w:r><w:ins w:id="1" w:author="Me" w:date="2026-01-01T00:00:00Z"><w:r><w:t>new world</w:t></w:r></w:ins></w:p><w:p/>'));
assert.equal(extra.matches, false);
assert.deepEqual(extra.differences[0], { path: 'paragraphs.length', expected: 1, actual: 2 });

const paths = result => result.differences.map(item => item.path);
const insBy = (author, text, id = 1) => `<w:ins w:id="${id}" w:author="${author}" w:date="2026-01-01T00:00:00Z"><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:ins>`;
const delBy = (author, text, id = 1) => `<w:del w:id="${id}" w:author="${author}" w:date="2026-01-01T00:00:00Z"><w:r><w:delText xml:space="preserve">${text}</w:delText></w:r></w:del>`;
const plain = text => `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;

// Identical inspections match, and comparison is symmetric in what it reports.
{
    const body = plain('One') + plain('Two');
    assert.deepEqual(compareDocumentInspections(inspect(body), inspect(body)), { matches: true, differences: [] });
    const changed = compareDocumentInspections(inspect(plain('One')), inspect(plain('Uno')));
    assert.deepEqual(changed.differences, [{ path: 'paragraphs[0].text', expected: 'One', actual: 'Uno' }]);
    assert.deepEqual(compareDocumentInspections(inspect(plain('Uno')), inspect(plain('One'))).differences,
        [{ path: 'paragraphs[0].text', expected: 'Uno', actual: 'One' }]);
}

// Review segments: author, kind and adjacency matter; splitting does not.
{
    const base = '<w:p><w:r><w:t xml:space="preserve">A </w:t></w:r>';
    const byMe = inspect(`${base}${insBy('Me', 'new')}</w:p>`);
    assert.equal(compareDocumentInspections(byMe, inspect(`${base}${insBy('Other', 'new')}</w:p>`)).matches, false, 'author change');
    assert.equal(compareDocumentInspections(byMe, inspect(`${base}${delBy('Me', 'new')}</w:p>`)).matches, false, 'insertion vs deletion');
    // Two adjacent insertions by different authors are two changes, not one.
    const split = inspect(`${base}${insBy('Me', 'ne', 1)}${insBy('Other', 'w', 2)}</w:p>`);
    assert.ok(paths(compareDocumentInspections(byMe, split)).includes('paragraphs[0].review'));
    // Deletions split by the host merge like insertions.
    const del1 = inspect(`${base}${delBy('Me', 'old text')}</w:p>`);
    const del2 = inspect(`${base}${delBy('Me', 'old ', 5)}${delBy('Me', 'text', 6)}</w:p>`);
    assert.equal(compareDocumentInspections(del1, del2).matches, true, 'adjacent same-author deletions merge');
    // Same text and author at a different position is a different change.
    const moved = inspect(`<w:p>${insBy('Me', 'new')}<w:r><w:t xml:space="preserve">A </w:t></w:r></w:p>`);
    assert.equal(compareDocumentInspections(byMe, moved).matches, false, 'segment position');
}

// Table position and list labels.
{
    const cell = text => `<w:tc>${plain(text)}</w:tc>`;
    const table = rows => `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol/><w:gridCol/></w:tblGrid>${rows.map(row => `<w:tr>${row.map(cell).join('')}</w:tr>`).join('')}</w:tbl>`;
    const original = inspect(table([['a', 'b'], ['c', 'd']]));
    assert.equal(compareDocumentInspections(original, inspect(table([['a', 'b'], ['c', 'd']]))).matches, true);
    // Same text order, different cell layout.
    const reshaped = compareDocumentInspections(original, inspect(table([['a'], ['b'], ['c'], ['d']])));
    assert.ok(paths(reshaped).some(path => path.endsWith('.table')), JSON.stringify(reshaped));
    assert.ok(paths(compareDocumentInspections(original, inspect(plain('a') + plain('b') + plain('c') + plain('d')))).includes('paragraphs[0].inTable'));

    const numberingXml = `<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">`
        + '<w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum>'
        + '<w:abstractNum w:abstractNumId="2"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="upperLetter"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum>'
        + '<w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="2"/></w:num></w:numbering>';
    const item = (numId, text) => `<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="${numId}"/></w:numPr></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`;
    const listed = numId => inspectDocumentParts({ documentXml: doc(item(numId, 'First') + item(numId, 'Second')), numberingXml });
    assert.equal(compareDocumentInspections(listed(1), listed(1)).matches, true);
    const relabeled = compareDocumentInspections(listed(1), listed(2));
    assert.deepEqual(paths(relabeled), ['paragraphs[0].list', 'paragraphs[1].list'], 'a list format change shows in the labels');
    assert.ok(paths(compareDocumentInspections(listed(1), inspect(plain('First') + plain('Second')))).includes('paragraphs[0].list'));
}

// Comments: text, author, anchor, resolved state and thread parent; order and ids are ignored.
{
    const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"';
    const comment = (id, author, text, paraId) => `<w:comment w:id="${id}" w:author="${author}" w:date="2026-01-01T00:00:00Z"><w:p w14:paraId="${paraId}"><w:r><w:t>${text}</w:t></w:r></w:p></w:comment>`;
    const anchored = (ids, text) => `<w:p>${ids.map(id => `<w:commentRangeStart w:id="${id}"/>`).join('')}<w:r><w:t>${text}</w:t></w:r>${ids.map(id => `<w:commentRangeEnd w:id="${id}"/><w:r><w:commentReference w:id="${id}"/></w:r>`).join('')}</w:p>`;
    const extended = entries => `<w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">${entries.map(([paraId, done, parent]) => `<w15:commentEx w15:paraId="${paraId}" w15:done="${done ? 1 : 0}"${parent ? ` w15:paraIdParent="${parent}"` : ''}/>`).join('')}</w15:commentsEx>`;
    const build = ({ ids = ['0', '1'], comments, ex, text = 'Clause' }) => inspectDocumentParts({
        documentXml: doc(anchored(ids, text)),
        commentsXml: `<w:comments ${NS}>${comments.join('')}</w:comments>`,
        commentsExtendedXml: ex ? extended(ex) : null
    });
    const thread = (overrides = {}) => build({
        comments: overrides.comments || [comment('0', 'Reviewer', 'Why?', 'AAAA0001'), comment('1', 'Me', 'Because.', 'AAAA0002')],
        ex: overrides.ex || [['AAAA0001', false], ['AAAA0002', false, 'AAAA0001']],
        ids: overrides.ids, text: overrides.text
    });
    const reference = thread();
    assert.equal(reference.comments.length, 2);
    // New comment ids and paraIds, reordered definitions: still the same thread.
    const renumbered = thread({
        ids: ['8', '9'],
        comments: [comment('9', 'Me', 'Because.', 'BBBB0002'), comment('8', 'Reviewer', 'Why?', 'BBBB0001')],
        ex: [['BBBB0001', false], ['BBBB0002', false, 'BBBB0001']]
    });
    assert.deepEqual(compareDocumentInspections(reference, renumbered), { matches: true, differences: [] });
    assert.ok(paths(compareDocumentInspections(reference, thread({ comments: [comment('0', 'Reviewer', 'Why not?', 'AAAA0001'), comment('1', 'Me', 'Because.', 'AAAA0002')] }))).includes('comments'), 'comment text');
    assert.ok(paths(compareDocumentInspections(reference, thread({ comments: [comment('0', 'Someone', 'Why?', 'AAAA0001'), comment('1', 'Me', 'Because.', 'AAAA0002')] }))).includes('comments'), 'comment author');
    assert.ok(paths(compareDocumentInspections(reference, thread({ ex: [['AAAA0001', true], ['AAAA0002', true, 'AAAA0001']] }))).includes('comments'), 'resolved state');
    assert.ok(paths(compareDocumentInspections(reference, thread({ ex: [['AAAA0001', false], ['AAAA0002', false]] }))).includes('comments'), 'reply detached from its thread');
    assert.ok(paths(compareDocumentInspections(reference, thread({ ex: [['AAAA0001', false], ['AAAA0002', false, 'FFFF9999']] }))).includes('comments'), 'reply to a missing parent');
    const lost = compareDocumentInspections(reference, thread({ ids: ['0'], comments: [comment('0', 'Reviewer', 'Why?', 'AAAA0001')], ex: [['AAAA0001', false]] }));
    assert.deepEqual(paths(lost), ['comments'], 'a dropped reply');
    // Anchor text is part of a comment's identity.
    assert.ok(paths(compareDocumentInspections(reference, thread({ text: 'Other clause' }))).includes('comments'));
}

// At most ten differences, in document order; empty inspections compare cleanly.
{
    const many = Array.from({ length: 15 }, (_, index) => plain(`P${index}`)).join('');
    const changed = Array.from({ length: 15 }, (_, index) => plain(`Q${index}`)).join('');
    const result = compareDocumentInspections(inspect(many), inspect(changed));
    assert.equal(result.matches, false);
    assert.equal(result.differences.length, 10);
    assert.deepEqual(paths(result).slice(0, 2), ['paragraphs[0].text', 'paragraphs[1].text']);
    assert.deepEqual(compareDocumentInspections({ paragraphs: [] }, { paragraphs: [] }), { matches: true, differences: [] });
    assert.deepEqual(compareDocumentInspections(null, undefined), { matches: true, differences: [] });
    assert.equal(compareDocumentInspections({ paragraphs: [] }, inspect(plain('x'))).matches, false);
}

// Works on openDocx inspections, the shape a package host has.
{
    const { openDocx } = await import('../index.js');
    const { unzipDocx, zipDocx } = await import('../document/zip-archive.js');
    const { readFileSync } = await import('node:fs');
    const template = unzipDocx(new Uint8Array(readFileSync(new URL('./fixtures/sample_doc_test.docx', import.meta.url))));
    const source = openDocx(zipDocx(template));
    const before = source.inspect();
    const target = before.paragraphs.find(p => p.exactText.trim().length > 0);
    const result = await source.applyOperations([{ type: 'redline', target: { index: target.index, exactText: target.exactText, paragraphId: target.paragraphId, fingerprint: target.fingerprint }, modified: `${target.exactText} (amended)` }], { author: 'Me' });
    assert.equal(result.status, 'ok', JSON.stringify(result.error));
    const after = openDocx(result.toUint8Array()).inspect();
    assert.deepEqual(compareDocumentInspections(after, openDocx(result.toUint8Array()).inspect()), { matches: true, differences: [] });
    const diff = compareDocumentInspections(before, after);
    assert.equal(diff.matches, false);
    assert.ok(paths(diff).every(path => path.startsWith(`paragraphs[${before.paragraphs.indexOf(target)}].`)), JSON.stringify(diff));
}

console.log('inspection_comparison_tests: all passed');
