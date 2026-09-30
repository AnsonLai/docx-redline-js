/**
 * Header and footer editing, against documents saved by real Word:
 *  - header-footer.docx: default + first + even headers/footers, titlePg, a PAGE field in the default footer
 *  - multi-section-headers.docx: section 2 has its own header and inherits section 1's footer
 * (scripts/generate-word-header-footer-fixture.ps1, scripts/generate-word-multi-section-fixture.ps1)
 */
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { unzipEntries } from '../scripts/lib/zip-reader.mjs';
import { openDocx } from '../node/index.js';
import { MemoryZip, unzipDocx, zipDocx } from '../document/zip-archive.js';
import { validateDocxPackage } from '../services/standalone-docx-plumbing.js';

const load = name => readFileSync(new URL(`./fixtures/word-authored/${name}`, import.meta.url));
const single = load('header-footer.docx');
const multi = load('multi-section-headers.docx');
const HDR = { kind: 'header', type: 'default' };
const HDR_TEXT = 'CONFIDENTIAL - Acme Master Agreement';
const replaceIn = (part, from, to) => ({ type: 'replace', part, target: { exactText: from }, modified: to });
const run = (bytes, operations, options = {}) => openDocx(bytes).applyOperations(operations, { author: 'Agent', atomic: true, ...options });
const byName = parts => Object.fromEntries(parts.map(part => [part.path, part]));

// --- discovery -----------------------------------------------------------------------------------------
{
    const parts = byName(openDocx(single).inspect().headersFooters);
    assert.equal(Object.keys(parts).length, 6, 'Word wrote even, default and first for both header and footer');
    assert.deepEqual([parts['word/header2.xml'].kind, parts['word/header2.xml'].type, parts['word/header2.xml'].active], ['header', 'default', true]);
    assert.equal(parts['word/header1.xml'].type, 'even');
    assert.equal(parts['word/header1.xml'].active, false, 'even parts are inactive without evenAndOddHeaders');
    assert.equal(parts['word/header3.xml'].type, 'first');
    assert.equal(parts['word/header3.xml'].active, true, 'first-page parts are active with titlePg');
    assert.equal(parts['word/footer2.xml'].hasFields, true, 'the default footer holds a PAGE field');
    assert.equal(parts['word/header2.xml'].hasFields, false);
    assert.deepEqual(parts['word/header2.xml'].paragraphs, [{ index: 1, text: HDR_TEXT }]);

    const sections = byName(openDocx(multi).inspect().headersFooters);
    assert.deepEqual(sections['word/header2.xml'].appliesToSections, [0]);
    assert.deepEqual(sections['word/header4.xml'].appliesToSections, [1]);
    assert.deepEqual(sections['word/footer2.xml'].sections, [0], 'only section 1 references the footer');
    assert.deepEqual(sections['word/footer2.xml'].appliesToSections, [0, 1], 'section 2 inherits it');
    assert.equal(sections['word/footer2.xml'].sharedBySections, true);
}

// --- editing a header ---------------------------------------------------------------------------------
{
    const before = unzipEntries(single);
    const result = await run(single, [replaceIn(HDR, HDR_TEXT, `STRICTLY ${HDR_TEXT}`)]);
    assert.equal(result.status, 'ok', result.error?.message);
    assert.equal(result.written, true);
    assert.deepEqual(result.artifactsChanged, ['word/header2.xml'], 'only the edited part changes');
    assert.equal(result.results[0].part, 'word/header2.xml');
    const after = unzipEntries(result.toBuffer());
    for (const [name, data] of before) {
        if (name !== 'word/header2.xml') assert.deepEqual(after.get(name), data, `${name} must be untouched`);
    }
    assert.match(after.get('word/header2.xml').toString(), /<w:ins [^>]*w:author="Agent"[^>]*>.*STRICTLY /s);
    await validateDocxPackage(new MemoryZip(unzipDocx(result.toBuffer())));

    // Selecting by path is equivalent.
    const byPath = await run(single, [replaceIn('word/header2.xml', HDR_TEXT, `STRICTLY ${HDR_TEXT}`)]);
    assert.equal(byPath.status, 'ok');
    assert.deepEqual(byPath.artifactsChanged, ['word/header2.xml']);

    // The first-page header is a different part.
    const first = await run(single, [replaceIn({ kind: 'header', type: 'first' }, 'First page header', 'First page header (rev)')]);
    assert.deepEqual(first.artifactsChanged, ['word/header3.xml']);

    // Body-only operations never touch header/footer parts.
    const bodyOnly = await run(single, [{ type: 'replace', target: { exactText: 'Body paragraph one.' }, modified: 'Body paragraph ONE.' }]);
    assert.deepEqual(bodyOnly.artifactsChanged, ['word/document.xml']);
}

// --- selectors fail closed --------------------------------------------------------------------------
{
    const cases = [
        [{ kind: 'footer', type: 'weird' }, 'INVALID_OPERATION'],
        [{ kind: 'sidebar' }, 'INVALID_OPERATION'],
        ['word/header99.xml', 'PART_NOT_FOUND'],
        [{ kind: 'header', type: 'default', section: 5 }, 'PART_NOT_FOUND']
    ];
    for (const [part, code] of cases) {
        const result = await run(single, [replaceIn(part, HDR_TEXT, 'x')]);
        assert.equal(result.status, 'error', JSON.stringify(part));
        assert.equal(result.error.code, code, JSON.stringify(part));
        assert.equal(result.written, false);
        assert.deepEqual(result.toBuffer(), single, 'a refused request leaves the document byte-identical');
    }
    // Two sections have distinct default headers: a selector without "section" is ambiguous, with it is exact.
    const ambiguous = await run(multi, [replaceIn(HDR, 'Section one header', 'x')]);
    assert.equal(ambiguous.error.code, 'PART_AMBIGUOUS');
    assert.equal(ambiguous.error.candidates?.length ?? 2, 2);
    const second = await run(multi, [replaceIn({ ...HDR, section: 1 }, 'Section two header', 'Section two header (rev)')]);
    assert.equal(second.status, 'ok', second.error?.message);
    assert.deepEqual(second.artifactsChanged, ['word/header4.xml']);
    const wrongSection = await run(multi, [replaceIn({ ...HDR, section: 1 }, 'Section one header', 'x')]);
    assert.equal(wrongSection.status, 'error', 'text from section 1 is not in section 2 header');
    // An inherited footer is reachable from the section that inherits it, and the result says who is affected.
    const inherited = await run(multi, [replaceIn({ kind: 'footer', section: 1 }, 'Shared footer', 'Shared footer (rev)')]);
    assert.equal(inherited.status, 'ok', inherited.error?.message);
    assert.deepEqual(inherited.artifactsChanged, ['word/footer2.xml']);
    assert.deepEqual(inherited.results[0].partSections, [0, 1]);
}

// --- fields are atomic ---------------------------------------------------------------------------------
{
    const footer = { kind: 'footer' };
    const around = await run(single, [replaceIn(footer, 'Page 1', 'Sheet 1')]);
    assert.equal(around.status, 'ok', around.error?.message);
    const xml = unzipEntries(around.toBuffer()).get('word/footer2.xml').toString();
    assert.equal((xml.match(/w:fldCharType="begin"/g) || []).length, 1);
    assert.equal((xml.match(/w:fldCharType="end"/g) || []).length, 1);
    assert.ok(xml.includes('PAGE'), 'field instruction is intact');

    const inside = await run(single, [replaceIn(footer, 'Page 1', 'Page 2')]);
    assert.equal(inside.status, 'error');
    assert.equal(inside.results[0].error.code, 'FIELD_EDIT_REFUSED');
    assert.deepEqual(inside.toBuffer(), single, 'atomic: nothing is written');

    // Non-atomic: the refused edit is skipped and the others are applied.
    const partial = await run(single, [replaceIn(footer, 'Page 1', 'Page 2'), replaceIn(HDR, HDR_TEXT, `STRICTLY ${HDR_TEXT}`)], { atomic: false });
    assert.equal(partial.status, 'partial');
    assert.equal(partial.written, true);
    assert.deepEqual(partial.results.map(r => r.status), ['error', 'applied']);
    assert.deepEqual(partial.artifactsChanged, ['word/header2.xml']);
}

// --- comments are refused --------------------------------------------------------------------------------
for (const operation of [
    { type: 'comment', part: HDR, target: { exactText: HDR_TEXT }, commentContent: 'x' },
    { type: 'comment_reply', part: HDR, parentCommentId: 0, commentContent: 'x' }
]) {
    const result = await run(single, [operation]);
    assert.equal(result.status, 'error');
    assert.equal(result.error.code, 'COMMENT_IN_HEADER_FOOTER');
    assert.deepEqual(result.toBuffer(), single);
}

// --- mixed batches ----------------------------------------------------------------------------------------
{
    const operations = [
        { type: 'replace', target: { exactText: 'Body paragraph one.' }, modified: 'Body paragraph ONE.' },
        replaceIn(HDR, HDR_TEXT, `STRICTLY ${HDR_TEXT}`),
        replaceIn('word/header3.xml', 'First page header', 'First page header (rev)'),
        { type: 'replace', target: { exactText: 'Body paragraph two.' }, modified: 'Body paragraph TWO.' }
    ];
    const result = await run(single, operations);
    assert.equal(result.status, 'ok', result.error?.message);
    assert.deepEqual(result.results.map(r => [r.index, r.status, r.part ?? null]), [[1, 'applied', null], [2, 'applied', 'word/header2.xml'], [3, 'applied', 'word/header3.xml'], [4, 'applied', null]]);
    assert.deepEqual(result.receipts.map(r => r.operationIndex), [1, 2, 3, 4]);
    assert.deepEqual([...result.artifactsChanged].sort(), ['word/document.xml', 'word/header2.xml', 'word/header3.xml']);

    // Revision ids are unique across document.xml and every edited part.
    const out = unzipEntries(result.toBuffer());
    const ids = ['word/document.xml', 'word/header2.xml', 'word/header3.xml']
        .flatMap(name => [...out.get(name).toString().matchAll(/<w:(?:ins|del) [^>]*w:id="(\d+)"/g)].map(m => m[1]));
    assert.ok(ids.length >= 4);
    assert.equal(new Set(ids).size, ids.length, `duplicate revision ids: ${ids}`);

    // Receipts describe what is in the file: the part an edit lives in and its final (renumbered) revision id.
    const headerReceipt = result.receipts.find(r => r.operationIndex === 2);
    assert.equal(headerReceipt.revisionItems[0].partName, 'word/header2.xml');
    const headerIds = [...out.get('word/header2.xml').toString().matchAll(/<w:ins [^>]*w:id="(\d+)"/g)].map(m => m[1]);
    assert.ok(headerIds.includes(headerReceipt.revisionItems[0].id), 'receipt id matches the id written to the header');
    assert.equal(result.results.find(r => r.index === 2).receipt.revisionItems[0].partName, 'word/header2.xml');

    // One failing operation rolls the whole batch back (atomic), across body and parts.
    const failing = await run(single, [operations[0], replaceIn(HDR, 'not in the header', 'x')]);
    assert.equal(failing.status, 'error');
    assert.equal(failing.written, false);
    assert.deepEqual(failing.toBuffer(), single, 'the body edit is rolled back too');
    assert.equal(failing.results.find(r => r.index === 2).error.code, 'TARGET_NOT_FOUND');

    // Revision preconditions apply to part operations.
    const doc = openDocx(single);
    const token = doc.getRevisionToken();
    const ok = await doc.applyOperations([replaceIn(HDR, HDR_TEXT, `STRICTLY ${HDR_TEXT}`)], { author: 'Agent', atomic: true, expectedRevision: token });
    assert.equal(ok.status, 'ok', ok.error?.message);
    const stale = await doc.applyOperations([replaceIn(HDR, HDR_TEXT, 'x')], { author: 'Agent', atomic: true, expectedRevision: token });
    assert.equal(stale.status, 'error');
    assert.equal(stale.error.code, 'REVISION_MISMATCH');
}

// --- accept / reject reach headers and footers ---------------------------------------------------------
{
    const edited = (await run(single, [replaceIn(HDR, HDR_TEXT, `STRICTLY ${HDR_TEXT}`), replaceIn('word/footer3.xml', 'First page footer', 'First page footer v2')])).toBuffer();
    const accepted = await openDocx(edited).resolveRevisions('accept', { allAuthors: true });
    assert.equal(accepted.status, 'ok', accepted.error?.message);
    assert.deepEqual([...accepted.artifactsChanged].sort(), ['word/footer3.xml', 'word/header2.xml']);
    const acceptedText = byName(openDocx(accepted.toBuffer()).inspect().headersFooters);
    assert.equal(acceptedText['word/header2.xml'].paragraphs[0].text, `STRICTLY ${HDR_TEXT}`);
    assert.equal(acceptedText['word/footer3.xml'].paragraphs[0].text, 'First page footer v2');
    assert.ok(!unzipEntries(accepted.toBuffer()).get('word/header2.xml').toString().includes('<w:ins '));

    const rejected = await openDocx(edited).resolveRevisions('reject', { allAuthors: true });
    const rejectedText = byName(openDocx(rejected.toBuffer()).inspect().headersFooters);
    assert.equal(rejectedText['word/header2.xml'].paragraphs[0].text, HDR_TEXT);
    assert.equal(rejectedText['word/footer3.xml'].paragraphs[0].text, 'First page footer');

    // Only the matching author's revisions in headers are resolved.
    const other = await openDocx(edited).resolveRevisions('accept', { author: 'Someone Else' });
    assert.equal(other.written, false);
}

// --- preflight ---------------------------------------------------------------------------------------------
{
    const doc = openDocx(single);
    const ok = doc.preflight([replaceIn(HDR, HDR_TEXT, 'x'), { type: 'replace', target: { exactText: 'Body paragraph two.' }, modified: 'y' }], 'Agent');
    assert.equal(ok.valid, true);
    assert.deepEqual(ok.results.map(r => [r.index, r.status]), [[1, 'ready'], [2, 'ready']]);
    const bad = doc.preflight([replaceIn(HDR, 'missing text', 'x'), { type: 'comment', part: HDR, target: { exactText: HDR_TEXT }, commentContent: 'c' }, replaceIn({ kind: 'footer', type: 'weird' }, 'a', 'b')], 'Agent');
    assert.equal(bad.valid, false);
    assert.deepEqual(bad.results.map(r => r.status), ['error', 'error', 'error']);
    assert.equal(bad.results[1].error.code, 'COMMENT_IN_HEADER_FOOTER');
}

// --- package validation of header/footer wiring --------------------------------------------------------
{
    const mutate = fn => { const entries = unzipDocx(single); fn(entries, s => new TextDecoder().decode(s), s => new TextEncoder().encode(s)); return new MemoryZip(entries); };
    await validateDocxPackage(new MemoryZip(unzipDocx(single)));
    await assert.rejects(validateDocxPackage(mutate((e, d, enc) => e.set('word/_rels/document.xml.rels', enc(d(e.get('word/_rels/document.xml.rels')).replace(/<Relationship [^>]*Target="header2\.xml"[^>]*\/>/, ''))))), /has no document relationship/);
    await assert.rejects(validateDocxPackage(mutate(e => e.delete('word/footer2.xml'))), /targets a missing part/);
    await assert.rejects(validateDocxPackage(mutate((e, d, enc) => e.set('[Content_Types].xml', enc(d(e.get('[Content_Types].xml')).replace(/<Override PartName="\/word\/header2\.xml"[^>]*\/>/, ''))))), /content type is missing/);
    await assert.rejects(validateDocxPackage(mutate((e, d, enc) => e.set('[Content_Types].xml', enc(d(e.get('[Content_Types].xml')).replace('wordprocessingml.header+xml"/><Override PartName="/word/header2', 'wordprocessingml.footer+xml"/><Override PartName="/word/header2'))))), /content type is/);
    assert.ok(zipDocx(unzipDocx(single)).length > 0);
}

console.log('PASS: header and footer editing');
