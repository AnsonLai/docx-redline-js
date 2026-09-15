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
const nbsp = '\u00a0';

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

    const revisionsInput = path.join(directory, 'revisions.docx');
    const revisionsXml = `<w:document xmlns:w="${W}"><w:body>
<w:p w:paraId="D1"><w:del w:id="1" w:author="Old Reviewer"><w:r><w:delText>Removed clause.</w:delText></w:r></w:del></w:p>
<w:p w:paraId="I1"><w:ins w:id="2" w:author="New Reviewer"><w:r><w:t>Added clause.</w:t></w:r></w:ins></w:p>
<w:sectPr/></w:body></w:document>`;
    await writeFile(revisionsInput, buildZip([
        {name:'[Content_Types].xml',data:contentTypes},
        {name:'word/document.xml',data:revisionsXml},
        {name:'word/_rels/document.xml.rels',data:rels}
    ]));
    const revisedExtract = await executeCli(['extract', revisionsInput, '--revised']);
    assert.deepEqual(revisedExtract.paragraphs.map(item => ({
        exactText: item.exactText,
        hasRevisions: item.hasRevisions,
        revisionAuthors: item.revisionAuthors,
        deleted: item.deleted
    })), [
        { exactText: '', hasRevisions: true, revisionAuthors: ['Old Reviewer'], deleted: true },
        { exactText: 'Added clause.', hasRevisions: true, revisionAuthors: ['New Reviewer'], deleted: undefined }
    ]);
    const rejectedRevisedExtract = await executeCli(['extract', revisionsInput, '--revised', '--view', 'rejected']);
    assert.equal(rejectedRevisedExtract.paragraphs[1].exactText, '');
    assert.equal(rejectedRevisedExtract.paragraphs[1].inserted, true);
    const preflight = await executeCli(['preflight', input, '--operations', operationsFile, '--author', 'CLI Editor']);
    assert.equal(preflight.valid, true);
    const missingAuthor = await executeCli(['apply', input, '--operations', authorlessFile]);
    assert.equal(missingAuthor.status, 'ok');
    assert.equal(missingAuthor.authorsUsed[0], 'AI Redliner');

    const output = path.join(directory, 'output.docx');
    const applied = await executeCli(['apply', input, '--operations', operationsFile, '--output', output]);
    assert.equal(applied.status, 'ok'); assert.equal(applied.written, true); assert.equal(applied.outputPath, output);
    assert.equal(applied.completion, true);
    assert.equal('documentXml' in applied, false);
    assert.equal(applied.validation.originalIssues.total, 0);
    assert.equal(applied.validation.generatedIssues.total, 0);
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

    // WP-02: one exact localized replacement may use a strong inline target.
    const localizedOutput = path.join(directory, 'localized_output.docx');
    const localizedApplied = await executeCli([
        'apply', input,
        '--target-id', 'A1',
        '--find', 'Exact',
        '--replace', 'Precise',
        '--author', 'Inline Author',
        '--output', localizedOutput
    ]);
    assert.equal(localizedApplied.status, 'ok', JSON.stringify(localizedApplied));
    assert.equal(localizedApplied.written, true);
    assert.equal(localizedApplied.completion, true);
    assert.equal((await executeCli(['extract', localizedOutput])).paragraphs[0].exactText, '  Precise\ttext  ');

    const localizedDeleteOutput = path.join(directory, 'localized_delete_output.docx');
    const localizedDeleted = await executeCli([
        'apply', input,
        '--target-ref', '1',
        '--find', 'Exact',
        '--replace', '',
        '--output', localizedDeleteOutput
    ]);
    assert.equal(localizedDeleted.status, 'ok', JSON.stringify(localizedDeleted));
    assert.equal((await executeCli(['extract', localizedDeleteOutput])).paragraphs[0].exactText, '  \ttext  ');

    const speculativeInlinePatch = await executeCli([
        'apply', input, '--find', 'Exact', '--replace', 'Precise'
    ]);
    assert.equal(speculativeInlinePatch.status, 'ok', JSON.stringify(speculativeInlinePatch));
    assert.equal(speculativeInlinePatch.completion, true);
    assert.equal(speculativeInlinePatch.results[0].change.verification.acceptedViewMatchesCompiledText, true);
    const mixedInlineForms = await executeCli([
        'apply', input, '--target-id', 'A1', '--modified', 'Complete',
        '--find', 'Exact', '--replace', 'Precise'
    ]);
    assert.equal(mixedInlineForms.error.code, 'INVALID_OPERATION');

    // WP09a: target matching may treat ASCII spaces and NBSPs as equivalent,
    // but mutation coordinates and Reject must use the exact resolved source.
    const whitespaceInput = path.join(directory, 'whitespace-source.docx');
    const whitespaceOutput = path.join(directory, 'whitespace-output.docx');
    const whitespaceOps = path.join(directory, 'whitespace-operations.json');
    const sourceText = `The operator handbook is available at${nbsp}docs.example/handbook/${nbsp}(the “Handbook”).`;
    const callerTarget = 'The operator handbook is available at docs.example/handbook/ (the “Handbook”).';
    const requestedText = 'The operator handbook, current at activation, is available at docs.example/handbook/ (the “Handbook”).';
    const whitespaceXml = `<w:document xmlns:w="${W}"><w:body><w:p w:paraId="W9"><w:r><w:t>${sourceText}</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`;
    await writeFile(whitespaceInput, buildZip([
        {name:'[Content_Types].xml',data:contentTypes},
        {name:'word/document.xml',data:whitespaceXml},
        {name:'word/_rels/document.xml.rels',data:rels}
    ]));
    await writeFile(whitespaceOps, JSON.stringify([{
        type: 'replace',
        target: { exactText: callerTarget, paragraphId: 'W9' },
        modified: requestedText,
        author: 'Reviewer B',
        existingRevisions: 'slice-cross-author'
    }]));
    const whitespaceApplied = await executeCli(['apply', whitespaceInput, '--operations', whitespaceOps, '--output', whitespaceOutput]);
    assert.equal(whitespaceApplied.status, 'ok', JSON.stringify(whitespaceApplied));
    assert.equal(whitespaceApplied.written, true);
    assert.equal(whitespaceApplied.completion, true);
    assert.equal(whitespaceApplied.results[0].resolvedTarget.targetTextMatch.mode, 'space_equivalent');
    assert.equal(whitespaceApplied.results[0].resolvedTarget.targetTextMatch.differenceCount, 2);
    assert.equal(whitespaceApplied.results[0].resolvedTarget.text, undefined, 'CLI must not echo resolved clause text');
    assert.equal(whitespaceApplied.results[0].resolvedTarget.targetTextMatch.differences, undefined);
    assert.equal(whitespaceApplied.results[0].receipt, undefined);
    assert.equal(whitespaceApplied.receipts.length, 1);
    assert.equal((await executeCli(['extract', whitespaceOutput])).paragraphs[0].exactText, requestedText);
    const whitespaceRejected = await executeCli(['reject', whitespaceOutput, '--author', 'Reviewer B']);
    assert.equal(whitespaceRejected.written, true);
    assert.equal((await executeCli(['extract', whitespaceRejected.outputPath])).paragraphs[0].exactText, sourceText);

    let stdout = '';
    const exitCode = await runCli(['extract', input, '--range', '1:1'], { stdout: { write: value => { stdout += value; } } });
    assert.equal(exitCode, 0); assert.equal(JSON.parse(stdout).paragraphs[0].exactText, '  Exact\ttext  ');
    const schemaText = await readFile(new URL('../docs/schemas/document-operations.schema.json', import.meta.url), 'utf8');
    const schema = JSON.parse(schemaText);
    assert(schema.$defs.base.properties.existingRevisions.enum.includes('slice-cross-author'));
    assert.equal(schema.$defs.localizedReplacement.required.includes('find'), true);
    assert.equal(schema.$defs.localizedReplacement.required.includes('replace'), true);
} finally {
    await rm(directory, { recursive: true, force: true });
}
console.log('agent CLI tests passed');
