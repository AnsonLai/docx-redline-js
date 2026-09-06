import assert from 'node:assert/strict';
import { access, mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { buildZip } from '../scripts/lib/minimal-zip.mjs';
import { executeCli, runCli } from '../node/cli.js';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const types = `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`;
const rels = `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`;
const xml = `<w:document xmlns:w="${W}"><w:body><w:p><w:r><w:t>Body</w:t></w:r></w:p><w:tbl><w:tr><w:tc><w:p><w:r><w:t>Table Needle</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p><w:ins w:id="4" w:author="Prior" w:date="2026-01-01T00:00:00Z"><w:r><w:t>Revised</w:t></w:r></w:ins></w:p><w:p/><w:sectPr/></w:body></w:document>`;
const fixture = buildZip([{name:'[Content_Types].xml',data:types},{name:'word/document.xml',data:xml},{name:'word/_rels/document.xml.rels',data:rels}]);
const directory = await mkdtemp(path.join(tmpdir(), 'docx-redline-cli-edge-'));

try {
    const input = path.join(directory, 'sample.docx'); await writeFile(input, fixture);
    assert.equal((await executeCli([])).error.code, 'COMMAND_REQUIRED');
    assert.equal((await executeCli(['wat', input])).error.code, 'UNKNOWN_COMMAND');
    assert.equal((await executeCli(['inspect'])).error.code, 'INPUT_REQUIRED');
    assert.equal((await executeCli(['inspect', path.join(directory, 'absent.docx')])).error.code, 'INPUT_READ_FAILED');
    assert.equal((await executeCli(['apply', input])).error.code, 'OPERATIONS_REQUIRED');

    const malformed = path.join(directory, 'malformed.json'); await writeFile(malformed, '{nope');
    assert.equal((await executeCli(['preflight', input, '--operations', malformed])).error.code, 'INVALID_OPERATIONS_FILE');
    const wrongShape = path.join(directory, 'wrong.json'); await writeFile(wrongShape, '{"items":[]}');
    assert.equal((await executeCli(['preflight', input, '--operations', wrongShape])).error.code, 'INVALID_OPERATIONS_FILE');

    assert.deepEqual((await executeCli(['inspect', input, '--table'])).paragraphs.map(item => item.text), ['Table Needle']);
    assert.deepEqual((await executeCli(['inspect', input, '--body', '--non-empty'])).paragraphs.map(item => item.text), ['Body', 'Revised']);
    assert.deepEqual((await executeCli(['extract', input, '--search', 'needle'])).paragraphs.map(item => item.exactText), ['Table Needle']);
    assert.deepEqual((await executeCli(['inspect', input, '--revised'])).paragraphs.map(item => item.text), ['Revised']);
    const singleIndex = await executeCli(['extract', input, '--index', '3']);
    assert.equal(singleIndex.indexBase, 1);
    assert.deepEqual(singleIndex.paragraphs.map(item => [item.index, item.ref]), [[3, 'P3']]);
    assert.deepEqual((await executeCli(['extract', input, '--indexes=1,3'])).paragraphs.map(item => item.index), [1,3]);
    assert.deepEqual((await executeCli(['extract', input, '--range=2:3'])).paragraphs.map(item => item.index), [2,3]);
    assert.deepEqual((await executeCli(['extract', input, '--range=2-3'])).paragraphs.map(item => item.index), [2,3]);
    assert.deepEqual((await executeCli(['extract', input, '--range=2,3'])).paragraphs.map(item => item.index), [2,3]);
    assert.deepEqual((await executeCli(['extract', input, '--range=3:3'])).paragraphs.map(item => item.index), [3]);
    assert.deepEqual((await executeCli(['extract', input, '--range=20:21'])).paragraphs, []);

    for (const args of [
        ['--index', '0'],
        ['--index', '-1'],
        ['--index'],
        ['--indexes', '1,nope'],
        ['--range', '0:2'],
        ['--range', '3:2'],
        ['--range', '2:'],
        ['--range', '1.5:2'],
        ['--index', '1', '--range', '1:2'],
        ['--view', 'future']
    ]) {
        const invalidFilter = await executeCli(['extract', input, ...args]);
        assert.equal(invalidFilter.status, 'error', `expected ${args.join(' ')} to fail`);
        assert.equal(invalidFilter.error.code, 'INVALID_FILTER');
    }
    const unknownOption = await executeCli(['extract', input, '--indxe', '2']);
    assert.equal(unknownOption.error.code, 'UNKNOWN_OPTION');
    const wrongCommandOption = await executeCli(['extract', input, '--output', path.join(directory, 'wrong.docx')]);
    assert.equal(wrongCommandOption.error.code, 'UNKNOWN_OPTION');
    const extraArgument = await executeCli(['extract', input, 'unexpected']);
    assert.equal(extraArgument.error.code, 'UNEXPECTED_ARGUMENT');

    const invalidTarget = path.join(directory, 'invalid-target.json');
    await writeFile(invalidTarget, JSON.stringify({ operations:[{ type:'replace', target:{ exactText:'Absent' }, modified:'No', author:'Editor' }] }));
    const failedOutput = path.join(directory, 'must-not-exist.docx');
    const failed = await executeCli(['apply', input, '--operations', invalidTarget, '--output', failedOutput]);
    assert.equal(failed.status, 'error'); assert.equal(failed.written, false);
    await assert.rejects(access(failedOutput)); assert.deepEqual(await readFile(input), fixture);

    const invalidAnchor = path.join(directory, 'invalid-anchor.json');
    await writeFile(invalidAnchor, JSON.stringify({ operations:[
        { type:'replace', target:{ exactText:'Body' }, modified:'Changed', author:'Editor', generateRedlines:false },
        { type:'comment', target:{ exactText:'Body' }, textToComment:'missing anchor', commentContent:'Nope', author:'Reviewer' }
    ] }));
    const failedAnchorOutput = path.join(directory, 'anchor-must-not-exist.docx');
    const failedAnchor = await executeCli(['apply', input, '--operations', invalidAnchor, '--output', failedAnchorOutput]);
    assert.equal(failedAnchor.status, 'error');
    assert.equal(failedAnchor.written, false);
    assert.equal(failedAnchor.results[1].error.code, 'ANCHOR_NOT_FOUND');
    await assert.rejects(access(failedAnchorOutput));
    assert.deepEqual(await readFile(input), fixture);

    const validOps = path.join(directory, 'valid.json');
    await writeFile(validOps, JSON.stringify({ operations:[{ type:'replace', target:{ exactText:'Body' }, modified:'Changed', author:'Editor', generateRedlines:false }] }));
    const defaultResult = await executeCli(['apply', input, '--operations', validOps]);
    assert.equal(defaultResult.outputPath, path.join(directory, 'sample.redlined.docx'));
    const originalDefault = await readFile(defaultResult.outputPath);
    const forced = await executeCli(['apply', input, '--operations', validOps, '--output', defaultResult.outputPath, '--force']);
    assert.equal(forced.written, true); assert.deepEqual(await readFile(defaultResult.outputPath), originalDefault);

    const aliasOutput = path.join(directory, 'aliases.docx');
    const aliased = await executeCli(['apply', input, '--operations-file', validOps, '-a', 'Editor', '-o', aliasOutput]);
    assert.equal(aliased.written, true);
    assert.equal(aliased.outputPath, aliasOutput);
    const baselineValidation = await executeCli(['validate', aliasOutput, '--baseline', input]);
    assert.equal(baselineValidation.status, 'ok');
    assert.equal(baselineValidation.valid, true);
    assert.deepEqual(baselineValidation.introducedIssues, []);
    assert.equal(baselineValidation.baseline, input);
    assert.equal((await executeCli(['validate', aliasOutput, '--baseline', path.join(directory, 'absent.docx')])).error.code, 'BASELINE_READ_FAILED');

    const invalidOnceXml = xml.replace('>Body<', '> Body<');
    const invalidTwiceXml = invalidOnceXml.replace('>Table Needle<', '> Table Needle<');
    const invalidOnce = path.join(directory, 'invalid-once.docx');
    const invalidTwice = path.join(directory, 'invalid-twice.docx');
    await writeFile(invalidOnce, buildZip([{name:'[Content_Types].xml',data:types},{name:'word/document.xml',data:invalidOnceXml},{name:'word/_rels/document.xml.rels',data:rels}]));
    await writeFile(invalidTwice, buildZip([{name:'[Content_Types].xml',data:types},{name:'word/document.xml',data:invalidTwiceXml},{name:'word/_rels/document.xml.rels',data:rels}]));
    const introducedValidation = await executeCli(['validate', invalidTwice, '--baseline', invalidOnce]);
    assert.equal(introducedValidation.status, 'error');
    assert.equal(introducedValidation.valid, false);
    assert.equal(introducedValidation.introducedIssues.filter(issue => issue.code === 'MISSING_SPACE_PRESERVE').length, 1);

    const inPlace = path.join(directory, 'in-place.docx'); await writeFile(inPlace, fixture);
    const inPlaceResult = await executeCli(['apply', inPlace, '--operations', validOps, '--in-place']);
    assert.equal(inPlaceResult.outputPath, inPlace); assert.notDeepEqual(await readFile(inPlace), fixture);

    let errorJson = '';
    const usageExit = await runCli(['unknown'], { stdout:{ write:value => { errorJson += value; } } });
    assert.equal(usageExit, 2); assert.equal(JSON.parse(errorJson).error.code, 'UNKNOWN_COMMAND');
    let validationJson = '';
    const badZip = path.join(directory, 'bad.docx'); await writeFile(badZip, 'bad');
    const validationExit = await runCli(['validate', badZip], { stdout:{ write:value => { validationJson += value; } } });
    assert.equal(validationExit, 2); assert.doesNotThrow(() => JSON.parse(validationJson));
    let filterJson = '';
    const filterExit = await runCli(['extract', input, '--range', 'bad'], { stdout:{ write:value => { filterJson += value; } } });
    assert.notEqual(filterExit, 0);
    assert.equal(JSON.parse(filterJson).error.code, 'INVALID_FILTER');
} finally {
    await rm(directory, { recursive:true, force:true });
}
console.log('agent CLI edge tests passed');
