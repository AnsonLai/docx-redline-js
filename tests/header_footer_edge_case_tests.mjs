/**
 * Header/footer edge cases beyond tests/header_footer_tests.mjs: tables in headers, edits next to fields,
 * accept/reject round-trips, preflight, atomic vs partial mixed batches, revision id uniqueness, inherited parts.
 */
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { unzipEntries } from '../scripts/lib/zip-reader.mjs';
import { openDocx } from '../node/index.js';
import { zipDocx, unzipDocx } from '../document/zip-archive.js';

const load = name => readFileSync(new URL(`./fixtures/word-authored/${name}`, import.meta.url));
const single = load('header-footer.docx');
const multi = load('multi-section-headers.docx');
const HDR = { kind: 'header', type: 'default' };
const HDR_TEXT = 'CONFIDENTIAL - Acme Master Agreement';
const replaceIn = (part, from, to) => ({ type: 'replace', part, target: { exactText: from }, modified: to });
const run = (bytes, operations, options = {}) => openDocx(bytes).applyOperations(operations, { author: 'Agent', atomic: true, ...options });
const byName = parts => Object.fromEntries(parts.map(part => [part.path, part]));
const xmlOf = (bytes, name) => unzipEntries(bytes).get(name).toString();
const inspectParts = bytes => byName(openDocx(bytes).inspect().headersFooters);

/** Rewrites one part of a docx: fn(xml) -> xml. */
function withPart(bytes, name, fn) {
    const entries = unzipDocx(bytes);
    entries.set(name, new TextEncoder().encode(fn(new TextDecoder().decode(entries.get(name)))));
    return Buffer.from(zipDocx(entries));
}

// --- 1. table inside a header ----------------------------------------------------------------------------
{
    const table = '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>'
        + '<w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>Client: Acme Corp</w:t></w:r></w:p></w:tc>'
        + '<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>Matter 1234</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p/>';
    const tableDoc = withPart(single, 'word/header2.xml', xml => xml.replace(/<w:p [\s\S]*<\/w:p>(?=<\/w:hdr>)/, table));
    assert.ok(xmlOf(tableDoc, 'word/header2.xml').includes('<w:tbl>'), 'fixture mutation produced a table');

    const result = await run(tableDoc, [replaceIn(HDR, 'Client: Acme Corp', 'Client: Acme Holdings')]);
    assert.equal(result.status, 'ok', result.error?.message);
    assert.deepEqual(result.artifactsChanged, ['word/header2.xml']);
    const xml = xmlOf(result.toBuffer(), 'word/header2.xml');
    assert.match(xml, /<w:ins [^>]*w:author="Agent"/);
    assert.match(xml, /<w:del [^>]*w:author="Agent"/);
    assert.ok(xml.includes('<w:tbl>') && xml.includes('Matter 1234'), 'table structure and sibling cell survive');
    // openDocx validates the package on open.
    const reopened = openDocx(result.toBuffer());
    const texts = byName(reopened.inspect().headersFooters)['word/header2.xml'].paragraphs.map(p => p.text).join('|');
    assert.match(texts, /Acme Holdings/);
    assert.match(texts, /Matter 1234/);
}

// --- 2. edit beside a field keeps the field runs intact --------------------------------------------------
// Targets match whole paragraph text (field results included), so the text is "Page 1 of 9".
{
    const fieldFooter = withPart(single, 'word/footer2.xml', xml => xml.replace(/<w:p [\s\S]*<\/w:p>(?=<\/w:ftr>)/,
        '<w:p><w:pPr><w:pStyle w:val="Footer"/></w:pPr><w:r><w:t xml:space="preserve">Page </w:t></w:r>'
        + '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>'
        + '<w:r><w:t xml:space="preserve"> of </w:t></w:r>'
        + '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> NUMPAGES </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>9</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'));
    const footer = { kind: 'footer' };
    for (const [from, to] of [['Page 1 of 9', 'Sheet 1 of 9'], ['Page 1 of 9', 'Page 1 out of 9']]) {
        const result = await run(fieldFooter, [replaceIn(footer, from, to)]);
        assert.equal(result.status, 'ok', `${JSON.stringify(from)}: ${result.error?.message}`);
        const xml = xmlOf(result.toBuffer(), 'word/footer2.xml');
        const count = re => (xml.match(re) || []).length;
        assert.equal(count(/w:fldCharType="begin"/g), 2, 'both fields keep begin');
        assert.equal(count(/w:fldCharType="separate"/g), 2, 'both fields keep separate');
        assert.equal(count(/w:fldCharType="end"/g), 2, 'both fields keep end');
        assert.ok(xml.includes('> PAGE <') && xml.includes('> NUMPAGES <'), 'instrText intact');
        assert.match(xml, /<w:ins /);
    }
}

// --- 3. accept / reject round-trip -----------------------------------------------------------------------
{
    const edited = (await run(single, [replaceIn(HDR, HDR_TEXT, `STRICTLY ${HDR_TEXT}`)])).toBuffer();
    const accepted = await openDocx(edited).resolveRevisions('accept', { allAuthors: true });
    assert.equal(accepted.status, 'ok', accepted.error?.message);
    assert.ok(accepted.artifactsChanged.includes('word/header2.xml'));
    assert.equal(inspectParts(accepted.toBuffer())['word/header2.xml'].paragraphs[0].text, `STRICTLY ${HDR_TEXT}`);
    const rejected = await openDocx(edited).resolveRevisions('reject', { allAuthors: true });
    assert.equal(rejected.status, 'ok', rejected.error?.message);
    assert.ok(rejected.artifactsChanged.includes('word/header2.xml'));
    assert.equal(inspectParts(rejected.toBuffer())['word/header2.xml'].paragraphs[0].text, HDR_TEXT);
    for (const r of [accepted, rejected]) {
        assert.ok(!/<w:(ins|del) /.test(xmlOf(r.toBuffer(), 'word/header2.xml')), 'no revision markup remains');
    }
}

// --- 4. preflight with part ----------------------------------------------------------------------------------
{
    const doc = openDocx(single);
    const ok = doc.preflight([replaceIn(HDR, HDR_TEXT, 'x')], 'Agent');
    assert.equal(ok.valid, true);
    assert.equal(ok.results[0].status, 'ready');
    assert.equal(ok.results[0].part, 'word/header2.xml', 'preflight reports the resolved part path');
    const bad = doc.preflight([replaceIn('word/header99.xml', HDR_TEXT, 'x')], 'Agent');
    assert.equal(bad.valid, false);
    assert.equal(bad.results[0].error.code, 'PART_NOT_FOUND');
    // toBuffer() re-zips (bytes differ from Word's), so compare the parts, not the archive.
    const after = unzipEntries(Buffer.from(doc.toBuffer()));
    for (const [name, data] of unzipEntries(single)) assert.deepEqual(after.get(name), data, `preflight must not change ${name}`);
}

// --- 5. mixed batch: atomic vs non-atomic ---------------------------------------------------------------------
{
    const ops = [
        { type: 'replace', target: { exactText: 'Body paragraph one.' }, modified: 'Body paragraph ONE.' },
        replaceIn(HDR, HDR_TEXT, `STRICTLY ${HDR_TEXT}`),
        replaceIn('word/footer3.xml', 'First page footer', 'First page footer v2'),
        replaceIn(HDR, 'no such header text', 'x')
    ];
    const atomic = await run(single, ops);
    assert.equal(atomic.status, 'error');
    assert.equal(atomic.written, false);
    assert.deepEqual(atomic.toBuffer(), single, 'atomic: original bytes returned');
    assert.equal(atomic.results.find(r => r.index === 4).error.code, 'TARGET_NOT_FOUND');

    const partial = await run(single, ops, { atomic: false });
    assert.equal(partial.status, 'partial');
    assert.equal(partial.written, true);
    assert.deepEqual(partial.results.map(r => [r.index, r.status]), [[1, 'applied'], [2, 'applied'], [3, 'applied'], [4, 'error']]);
    assert.deepEqual([...partial.artifactsChanged].sort(), ['word/document.xml', 'word/footer3.xml', 'word/header2.xml']);
}

// --- 6. revision ids unique across document.xml and two different parts ---------------------------------------
{
    const result = await run(single, [
        replaceIn(HDR, HDR_TEXT, `STRICTLY ${HDR_TEXT}`),
        replaceIn('word/footer3.xml', 'First page footer', 'First page footer v2'),
        { type: 'replace', target: { exactText: 'Body paragraph one.' }, modified: 'Body paragraph ONE.' }
    ]);
    assert.equal(result.status, 'ok', result.error?.message);
    const out = unzipEntries(result.toBuffer());
    const ids = ['word/document.xml', 'word/header2.xml', 'word/footer3.xml']
        .flatMap(name => [...out.get(name).toString().matchAll(/<w:(?:ins|del) [^>]*w:id="(\d+)"/g)].map(m => m[1]));
    assert.ok(ids.length >= 3, `expected revisions in three parts, got ${ids.length}`);
    for (const name of ['word/document.xml', 'word/header2.xml', 'word/footer3.xml']) assert.match(out.get(name).toString(), /<w:(ins|del) /, name);
    assert.equal(new Set(ids).size, ids.length, `duplicate revision ids: ${ids}`);
}

// --- 7. inherited part: section 2 has no footer of its own -----------------------------------------------------
{
    // Documented behavior (resolvePartSelector): `section` is matched against appliesToSections, so an inheriting
    // section resolves to the part it inherits; a section that defines its own part resolves to that one.
    const parts = inspectParts(multi);
    assert.deepEqual(parts['word/footer2.xml'].sections, [0]);
    assert.ok(parts['word/footer2.xml'].appliesToSections.includes(1), 'inspect reports the inheriting section');
    const viaSection = await run(multi, [replaceIn({ kind: 'footer', section: 1 }, 'Shared footer', 'Shared footer (rev)')]);
    assert.equal(viaSection.status, 'ok', viaSection.error?.message);
    assert.deepEqual(viaSection.artifactsChanged, ['word/footer2.xml']);
    const viaPath = await run(multi, [replaceIn('word/footer2.xml', 'Shared footer', 'Shared footer (rev)')]);
    assert.equal(viaPath.status, 'ok');
    assert.deepEqual(viaPath.artifactsChanged, viaSection.artifactsChanged);
    const own = await run(multi, [replaceIn({ kind: 'header', section: 1 }, 'Section two header', 'Section two header (rev)')]);
    assert.deepEqual(own.artifactsChanged, ['word/header4.xml']);
}

console.log('PASS: header and footer edge cases');
