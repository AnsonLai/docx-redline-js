import assert from 'node:assert/strict';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { buildZip } from '../scripts/lib/minimal-zip.mjs';
import { executeCli, runCli } from '../node/cli.js';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const contentTypes = `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`;
const rels = `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`;
const documentXml = `<w:document xmlns:w="${W}"><w:body><w:p w:paraId="A1"><w:r><w:t xml:space="preserve">  Exact\ttext  </w:t></w:r></w:p><w:sectPr/></w:body></w:document>`;
const fixture = buildZip([{name:'[Content_Types].xml',data:contentTypes},{name:'word/document.xml',data:documentXml},{name:'word/_rels/document.xml.rels',data:rels},{name:'custom.bin',data:Buffer.from([9,8,7])}]);

const directory = await mkdtemp(path.join(tmpdir(), 'docx-redline-cli-'));
try {
    const input = path.join(directory, 'input.docx'); const operationsFile = path.join(directory, 'operations.json'); const authorlessFile = path.join(directory, 'authorless.json');
    await writeFile(input, fixture);
    await writeFile(operationsFile, JSON.stringify([{ type:'replace', target:{ exactText:'  Exact\ttext  ', paragraphId:'A1' }, modified:'Updated', author:'CLI Editor' }]));
    await writeFile(authorlessFile, JSON.stringify([{ type:'replace', target:{ exactText:'  Exact\ttext  ' }, modified:'Updated' }]));

    const inspected = await executeCli(['inspect', input, '--non-empty']);
    assert.equal(inspected.status, 'ok'); assert.equal(inspected.paragraphs.length, 1);
    const extracted = await executeCli(['extract', input]);
    assert.equal(extracted.paragraphs[0].exactText, '  Exact\ttext  ');
    const preflight = await executeCli(['preflight', input, '--operations', operationsFile, '--author', 'CLI Editor']);
    assert.equal(preflight.valid, true);
    const missingAuthor = await executeCli(['apply', input, '--operations', authorlessFile]);
    assert.equal(missingAuthor.status, 'ok');
    assert.equal(missingAuthor.authorsUsed[0], 'AI Redliner');

    const output = path.join(directory, 'output.docx');
    const applied = await executeCli(['apply', input, '--operations', operationsFile, '--output', output]);
    assert.equal(applied.status, 'ok'); assert.equal(applied.written, true); assert.equal(applied.outputPath, output);
    assert.deepEqual(await readFile(input), fixture);
    assert.equal((await executeCli(['validate', output])).valid, true);
    const accepted = await executeCli(['accept', output, '--author', 'CLI Editor']);
    assert.equal(accepted.written, true); assert.equal((await executeCli(['validate', accepted.outputPath])).valid, true);
    const rejected = await executeCli(['reject', output, '--author', 'CLI Editor']);
    assert.equal(rejected.written, true); assert.equal((await executeCli(['validate', rejected.outputPath])).valid, true);

    const commentOps = path.join(directory, 'comments.json'); const commented = path.join(directory, 'commented.docx');
    await writeFile(commentOps, JSON.stringify([{ type:'comment', target:{ exactText:'  Exact\ttext  ' }, commentContent:'Review', author:'Reviewer' }]));
    const commentResult = await executeCli(['apply', input, '--operations', commentOps, '--output', commented]);
    assert.equal(commentResult.written, true);
    const removed = await executeCli(['delete-comments', commented, '--author', 'Reviewer']);
    assert.equal(removed.written, true); assert.equal(removed.commentsRemoved, 1);
    const refused = await executeCli(['apply', input, '--operations', operationsFile, '--output', output, '--no-overwrite']);
    assert.equal(refused.error.code, 'OUTPUT_EXISTS');

    // Inline one-liner test
    const inlineOutput = path.join(directory, 'inline_output.docx');
    const inlineApplied = await executeCli(['apply', input, '--target', '  Exact\ttext  ', '--modified', 'Inline Replaced', '--author', 'Inline Author', '--output', inlineOutput]);
    assert.equal(inlineApplied.status, 'ok');
    assert.equal(inlineApplied.written, true);
    const inlineExtracted = await executeCli(['extract', inlineOutput]);
    assert.ok(inlineExtracted.paragraphs[0].exactText.includes('Inline Replaced'));

    let stdout = '';
    const exitCode = await runCli(['extract', input, '--range', '1:1'], { stdout: { write: value => { stdout += value; } } });
    assert.equal(exitCode, 0); assert.equal(JSON.parse(stdout).paragraphs[0].exactText, '  Exact\ttext  ');
    const schemaText = await readFile(new URL('../docs/schemas/document-operations.schema.json', import.meta.url), 'utf8');
    assert.doesNotThrow(() => JSON.parse(schemaText));
} finally {
    await rm(directory, { recursive: true, force: true });
}
console.log('agent CLI tests passed');
