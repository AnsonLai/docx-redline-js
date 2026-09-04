import assert from 'node:assert/strict';
import { access, mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { spawnSync } from 'node:child_process';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { buildZip } from '../scripts/lib/minimal-zip.mjs';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const repositoryRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const applyWrapper = path.join(repositoryRoot, 'scripts', 'apply_changes.mjs');
const extractWrapper = path.join(repositoryRoot, 'scripts', 'extract_text.mjs');
const contentTypes = `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`;
const rels = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>';
const documentXml = `<w:document xmlns:w="${W}"><w:body>`
    + '<w:p><w:r><w:t>First paragraph.</w:t></w:r></w:p>'
    + '<w:p><w:r><w:t>during the\u00a0Subscription Term\u00a0only</w:t></w:r></w:p>'
    + '<w:p><w:r><w:t>Third paragraph.</w:t></w:r></w:p>'
    + '<w:sectPr/></w:body></w:document>';
const fixture = buildZip([
    { name: '[Content_Types].xml', data: contentTypes },
    { name: 'word/document.xml', data: documentXml },
    { name: 'word/_rels/document.xml.rels', data: rels }
]);

function invoke(script, args) {
    const result = spawnSync(process.execPath, [script, ...args], {
        cwd: repositoryRoot,
        encoding: 'utf8'
    });
    const stdout = result.stdout.trim();
    assert.ok(stdout, `wrapper emitted no JSON; stderr=${result.stderr}`);
    return { ...result, json: JSON.parse(stdout) };
}

const directory = await mkdtemp(path.join(tmpdir(), 'docx-redline-plugin-wrapper-'));
try {
    const input = path.join(directory, 'input.docx');
    await writeFile(input, fixture);

    const indexed = invoke(extractWrapper, [input, '--index', '2']);
    assert.equal(indexed.status, 0);
    assert.equal(indexed.json.indexBase, 1);
    assert.deepEqual(indexed.json.paragraphs.map(item => [item.index, item.ref]), [[2, 'P2']]);
    assert.equal(indexed.json.paragraphs[0].exactText, 'during the\u00a0Subscription Term\u00a0only');

    const ranged = invoke(extractWrapper, [input, '--range', '1-2']);
    assert.equal(ranged.status, 0);
    assert.deepEqual(ranged.json.paragraphs.map(item => item.index), [1, 2]);

    const malformed = invoke(extractWrapper, [input, '--range', 'bad']);
    assert.notEqual(malformed.status, 0);
    assert.equal(malformed.json.error.code, 'INVALID_FILTER');

    const successfulOperations = path.join(directory, 'successful-changes.json');
    const successfulOutput = path.join(directory, 'successful.docx');
    await writeFile(successfulOperations, JSON.stringify({ changes: [{
        type: 'comment',
        target: 'during the Subscription Term only',
        commentContent: 'Review the whole paragraph.'
    }] }));
    const successful = invoke(applyWrapper, [input, successfulOperations, successfulOutput]);
    assert.equal(successful.status, 0, successful.stderr);
    assert.equal(successful.json.status, 'ok');
    assert.equal(successful.json.written, true);
    assert.equal(successful.json.outputPath, successfulOutput);
    assert.deepEqual(successful.json.authorsUsed, ['Agent']);
    await access(successfulOutput);

    const failingOperations = path.join(directory, 'failing-changes.json');
    const failedOutput = path.join(directory, 'must-not-exist.docx');
    await writeFile(failingOperations, JSON.stringify({ changes: [
        {
            type: 'replace', target: 'First paragraph.', modified: 'Changed.',
            generateRedlines: false
        },
        {
            type: 'comment', target: 'during the Subscription Term only',
            textToComment: 'missing anchor', commentContent: 'Must fail.'
        }
    ] }));
    const failed = invoke(applyWrapper, [input, failingOperations, failedOutput]);
    assert.notEqual(failed.status, 0);
    assert.equal(failed.json.status, 'error');
    assert.equal(failed.json.written, false);
    assert.equal(failed.json.rolledBack, true);
    assert.equal(failed.json.outputPath, null);
    assert.equal(failed.json.results[1].error.code, 'ANCHOR_NOT_FOUND');
    await assert.rejects(access(failedOutput));
    assert.deepEqual(await readFile(input), fixture);

    const existingOutput = path.join(directory, 'existing.docx');
    const sentinel = Buffer.from('leave this file untouched');
    await writeFile(existingOutput, sentinel);
    const failedExisting = invoke(applyWrapper, [input, failingOperations, existingOutput, '--force']);
    assert.notEqual(failedExisting.status, 0);
    assert.deepEqual(await readFile(existingOutput), sentinel);
} finally {
    await rm(directory, { recursive: true, force: true });
}

console.log('plugin wrapper compatibility tests passed');
