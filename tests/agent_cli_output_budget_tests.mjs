import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { buildZip } from '../scripts/lib/minimal-zip.mjs';
import { executeCli, runCli } from '../node/cli.js';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const contentTypes = `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`;
const rels = `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`;
const paragraph = (index, text) => `<w:p w:paraId="${index.toString(16).padStart(8, '0')}"><w:r><w:t>${text}</w:t></w:r></w:p>`;
const commonText = index => `Common email clause ${index}: ${'contract language '.repeat(105)}`;
const manyXml = `<w:document xmlns:w="${W}"><w:body>${Array.from({ length: 70 }, (_, index) => paragraph(index + 1, commonText(index + 1))).join('')}<w:sectPr/></w:body></w:document>`;
const hugeText = `Oversize email clause: ${'x'.repeat(60000)}`;
const hugeXml = `<w:document xmlns:w="${W}"><w:body>${paragraph(1, hugeText)}<w:sectPr/></w:body></w:document>`;
const packageFor = documentXml => buildZip([
    { name: '[Content_Types].xml', data: contentTypes },
    { name: 'word/document.xml', data: documentXml },
    { name: 'word/_rels/document.xml.rels', data: rels }
]);
const directory = await mkdtemp(path.join(tmpdir(), 'docx-cli-budget-'));

try {
    const input = path.join(directory, 'many.docx');
    const hugeInput = path.join(directory, 'huge.docx');
    await writeFile(input, packageFor(manyXml));
    await writeFile(hugeInput, packageFor(hugeXml));

    const broadExtract = await executeCli(['extract', input, '--search', 'email']);
    const broadBytes = Buffer.byteLength(JSON.stringify(broadExtract, null, 2)) + 1;
    assert(broadBytes <= 48 * 1024, `broad extract used ${broadBytes} bytes`);
    assert.equal(broadExtract.selection.totalMatches, 70);
    assert(broadExtract.selection.returnedMatches <= 20);
    assert.equal(broadExtract.selection.truncated, true);
    assert.equal(typeof broadExtract.selection.nextAfter, 'number');
    assert.equal(broadExtract.paragraphs.every(item => item.exactText.includes('Common email clause')), true);

    const detailed = await executeCli(['inspect', input, '--non-empty']);
    const detailedBytes = Buffer.byteLength(JSON.stringify(detailed, null, 2)) + 1;
    assert(detailedBytes <= 48 * 1024, `detailed inspection used ${detailedBytes} bytes`);
    assert.equal(detailed.selection.totalMatches, 70);
    assert.equal(detailed.selection.truncated, true);
    assert.match(detailed.notes[0], /Prefer extract --search/);

    const firstPage = await executeCli(['extract', input, '--search', 'common', '--limit', '7']);
    assert.deepEqual(firstPage.paragraphs.map(item => item.index), [1, 2, 3, 4, 5, 6, 7]);
    assert.equal(firstPage.selection.nextAfter, 7);
    const secondPage = await executeCli(['extract', input, '--search', 'common', '--limit', '7', '--after', '7']);
    assert.deepEqual(secondPage.paragraphs.map(item => item.index), [8, 9, 10, 11, 12, 13, 14]);
    assert.equal(new Set([...firstPage.paragraphs, ...secondPage.paragraphs].map(item => item.index)).size, 14);

    const ranged = await executeCli(['extract', input, '--range', '1:70']);
    const rangedBytes = Buffer.byteLength(JSON.stringify(ranged, null, 2)) + 1;
    assert(rangedBytes <= 48 * 1024, `ranged extract used ${rangedBytes} bytes`);
    assert.equal(ranged.selection.totalMatches, 70);
    assert.equal(ranged.selection.truncated, true);

    const unbounded = await executeCli(['extract', input, '--all']);
    assert.equal(unbounded.paragraphs.length, 70);
    assert(Buffer.byteLength(JSON.stringify(unbounded, null, 2)) > 48 * 1024);

    const huge = await executeCli(['extract', hugeInput, '--search', 'oversize']);
    assert.equal(huge.paragraphs.length, 1);
    assert.equal(huge.paragraphs[0].exactText, hugeText);
    assert.equal(huge.selection.oversizeItem, true);
    assert(Buffer.byteLength(JSON.stringify(huge, null, 2)) > 48 * 1024);

    for (const args of [
        ['--limit', '0'],
        ['--limit', '201'],
        ['--after', '0'],
        ['--around', '21', '--search', 'email'],
        ['--around', '1.5', '--search', 'email'],
        ['--all', '--limit', '2']
    ]) {
        const invalid = await executeCli(['extract', input, ...args]);
        assert.equal(invalid.error.code, 'INVALID_FILTER', args.join(' '));
    }

    let stdout = '';
    const exitCode = await runCli(['extract', input, '--search', 'email'], {
        stdout: { write: value => { stdout += value; } }
    });
    assert.equal(exitCode, 0);
    assert(Buffer.byteLength(stdout) <= 48 * 1024);
    assert.doesNotThrow(() => JSON.parse(stdout));
} finally {
    await rm(directory, { recursive: true, force: true });
}

console.log('agent CLI output budget tests passed');
