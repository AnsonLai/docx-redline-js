import assert from 'node:assert/strict';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { buildZip } from '../scripts/lib/minimal-zip.mjs';
import { executeCli, runCli } from '../node/cli.js';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const contentTypes = '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>';
const rels = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>';

function escapeXml(value) {
    return String(value).replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;');
}

function fixture(paragraphs) {
    const body = paragraphs.map(({ id, text, heading = false }) => (
        `<w:p w:paraId="${id}">${heading ? '<w:pPr><w:pStyle w:val="Heading1"/></w:pPr>' : ''}<w:r><w:t>${escapeXml(text)}</w:t></w:r></w:p>`
    )).join('');
    const documentXml = `<w:document xmlns:w="${W}"><w:body>${body}<w:sectPr/></w:body></w:document>`;
    return buildZip([
        { name: '[Content_Types].xml', data: contentTypes },
        { name: 'word/document.xml', data: documentXml },
        { name: 'word/_rels/document.xml.rels', data: rels }
    ]);
}

async function writeFixture(directory, name, paragraphs) {
    const file = path.join(directory, name);
    await writeFile(file, fixture(paragraphs));
    return file;
}

const directory = await mkdtemp(path.join(tmpdir(), 'docx-redline-speculative-'));
try {
    const provision = [
        { id: 'A1', text: 'Background.' },
        { id: 'A2', text: 'Section 4.1 — Termination', heading: true },
        { id: 'A3', text: 'This provision applies during the Term.' },
        { id: 'A4', text: 'Either party may terminate upon thirty (30) days written notice.' },
        { id: 'A5', text: 'See Section 4.1 for the applicable procedure.' }
    ];

    const globalInput = await writeFixture(directory, 'global.docx', provision);
    const globalOutput = path.join(directory, 'global-output.docx');
    const global = await executeCli([
        'apply', globalInput,
        '--find', 'thirty (30) days',
        '--replace', 'sixty (60) days',
        '--profile', 'agent',
        '--output', globalOutput
    ]);
    assert.equal(global.status, 'ok', JSON.stringify(global));
    assert.equal(global.written, true);
    assert.equal(global.completion, true);
    assert.equal(global.results[0].change.kind, 'localized_replacement');
    assert.equal(global.results[0].change.committed, true);
    assert.equal(global.results[0].change.finalDisposition, 'applied');
    assert.equal(global.results[0].change.target.paragraphId, 'A4');
    assert.match(global.results[0].change.replacements[0].beforeExcerpt, /thirty \(30\) days/);
    assert.match(global.results[0].change.replacements[0].afterExcerpt, /sixty \(60\) days/);
    assert.equal(global.results[0].change.verification.acceptedViewMatchesCompiledText, true);
    assert.equal((await executeCli(['extract', globalOutput, '--search', 'sixty (60) days'])).paragraphs.length, 1);

    const nbspInput = await writeFixture(directory, 'nbsp.docx', [
        { id: 'NB1', text: '3.2 Reference Materials', heading: true },
        {
            id: 'NB2',
            text: 'The toolkit includes an assembly chart stored at docs.example/assembly. In addition, operators should consult the Calibration Manual located at\u00a0docs.example/calibration/\u00a0(the “Manual”).'
        },
        { id: 'NB3', text: 'The assembly chart remains available at docs.example/assembly.' }
    ]);
    const nbspOutput = path.join(directory, 'nbsp-output.docx');
    const nbspResult = await executeCli([
        'apply', nbspInput,
        '--search', 'In addition, operators',
        '--context-range', '0:0',
        '--find', 'Calibration Manual located at docs.example/calibration/ (the “Manual”).',
        '--replace', 'Calibration Manual located at docs.example/calibration/, as published when the device was activated (the “Manual”).',
        '--profile', 'agent',
        '--output', nbspOutput
    ]);
    assert.equal(nbspResult.status, 'ok', JSON.stringify(nbspResult));
    assert.equal(nbspResult.completion, true);
    assert.equal(nbspResult.results[0].change.target.paragraphId, 'NB2');
    assert.equal(nbspResult.results[0].change.context.anchorMatchCount, 1);
    assert.equal(nbspResult.results[0].change.replacements[0].matchMode, 'space_equivalent');
    const nbspExtract = await executeCli(['extract', nbspOutput, '--index', '2']);
    assert.equal(
        nbspExtract.paragraphs[0].exactText,
        'The toolkit includes an assembly chart stored at docs.example/assembly. In addition, operators should consult the Calibration Manual located at\u00a0docs.example/calibration/, as published when the device was activated\u00a0(the “Manual”).'
    );
    assert.match(nbspExtract.paragraphs[0].exactText, /assembly chart stored at docs\.example\/assembly/);
    assert.equal((await executeCli(['extract', nbspOutput, '--search', 'calendar date'])).paragraphs.length, 0);

    const exactPriorityInput = await writeFixture(directory, 'exact-priority.docx', [
        { id: 'EP1', text: 'The handbook is at docs.example/handbook/ today.' },
        { id: 'EP2', text: 'The handbook is at\u00a0docs.example/handbook/\u00a0today.' }
    ]);
    const exactPriority = await executeCli([
        'apply', exactPriorityInput,
        '--find', 'handbook is at docs.example/handbook/ today',
        '--replace', 'handbook is at docs.example/handbook/ at launch'
    ]);
    assert.equal(exactPriority.status, 'ok', JSON.stringify(exactPriority));
    assert.equal(exactPriority.results[0].change.target.paragraphId, 'EP1');
    assert.equal(exactPriority.results[0].change.replacements[0].matchMode, undefined);

    const ambiguousNbspInput = await writeFixture(directory, 'ambiguous-nbsp.docx', [
        { id: 'AN1', text: 'The handbook is at\u00a0docs.example/handbook/.' },
        { id: 'AN2', text: 'A second handbook is at\u00a0docs.example/handbook/.' }
    ]);
    const ambiguousNbspOutput = path.join(directory, 'ambiguous-nbsp-output.docx');
    const ambiguousNbsp = await executeCli([
        'apply', ambiguousNbspInput,
        '--find', 'handbook is at docs.example/handbook/.',
        '--replace', 'handbook is at docs.example/handbook/ at launch.',
        '--output', ambiguousNbspOutput
    ]);
    assert.equal(ambiguousNbsp.status, 'error');
    assert.equal(ambiguousNbsp.error.code, 'AMBIGUOUS_TARGET');
    assert.equal(ambiguousNbsp.error.candidates.length, 2);
    await assert.rejects(readFile(ambiguousNbspOutput), error => error.code === 'ENOENT');

    // Repeated visible section references are valid anchors when their union still
    // leaves one eligible patch paragraph. The directional range excludes text
    // before the heading and the heading itself.
    const directionalOutput = path.join(directory, 'directional-output.docx');
    const directional = await executeCli([
        'apply', globalInput,
        '--search', 'Section 4.1',
        '--context-range', '1:3',
        '--find', 'thirty (30) days',
        '--replace', 'forty-five (45) days',
        '--output', directionalOutput
    ]);
    assert.equal(directional.status, 'ok', JSON.stringify(directional));
    assert.equal(directional.results[0].change.context.search, 'Section 4.1');
    assert.equal(directional.results[0].change.context.range, '1:3');
    assert.equal(directional.results[0].change.context.anchorMatchCount, 2);
    assert.match(directional.results[0].change.target.humanReference, /Termination/);

    const directionalCases = [
        { range: '0:0', targetIndex: 4, find: 'zero target' },
        { range: '0:3', targetIndex: 7, find: 'plus three target' },
        { range: '1:3', targetIndex: 5, find: 'plus one target' },
        { range: '-3:-1', targetIndex: 1, find: 'minus three target' },
        { range: '-3:3', targetIndex: 2, find: 'symmetric target' }
    ];
    for (const [caseIndex, testCase] of directionalCases.entries()) {
        const paragraphs = Array.from({ length: 7 }, (_, index) => ({
            id: `E${index + 1}`,
            text: index + 1 === 4 ? 'Anchor provision.' : `Filler ${index + 1}.`
        }));
        paragraphs[testCase.targetIndex - 1].text = testCase.targetIndex === 4
            ? `Anchor provision with ${testCase.find}.`
            : `The ${testCase.find} applies.`;
        const rangeInput = await writeFixture(directory, `range-${caseIndex}.docx`, paragraphs);
        const rangeOutput = path.join(directory, `range-${caseIndex}-output.docx`);
        const ranged = await executeCli([
            'apply', rangeInput,
            '--search', 'Anchor provision',
            '--context-range', testCase.range,
            '--find', testCase.find,
            '--replace', `replacement ${caseIndex}`,
            '--output', rangeOutput
        ]);
        assert.equal(ranged.status, 'ok', `${testCase.range}: ${JSON.stringify(ranged)}`);
        assert.equal(ranged.results[0].change.context.range, testCase.range);
        assert.equal(ranged.results[0].change.target.paragraphId, `E${testCase.targetIndex}`);
    }

    const boundaryInput = await writeFixture(directory, 'boundary.docx', [
        { id: 'F1', text: 'Boundary anchor and old boundary value.' },
        { id: 'F2', text: 'Following paragraph.' }
    ]);
    const boundary = await executeCli([
        'apply', boundaryInput,
        '--search', 'Boundary anchor',
        '--context-range', '-3:0',
        '--find', 'old boundary value',
        '--replace', 'new boundary value'
    ]);
    assert.equal(boundary.status, 'ok', JSON.stringify(boundary));

    const beforeOnly = await writeFixture(directory, 'before-only.docx', [
        { id: 'B1', text: 'The fee is thirty dollars.' },
        { id: 'B2', text: 'Section 4.1 — Fees', heading: true },
        { id: 'B3', text: 'No fee appears here.' }
    ]);
    const excludedBefore = await executeCli([
        'apply', beforeOnly,
        '--search', 'Section 4.1',
        '--context-range', '1:3',
        '--find', 'thirty dollars',
        '--replace', 'sixty dollars'
    ]);
    assert.equal(excludedBefore.status, 'error');
    assert.equal(excludedBefore.error.code, 'PATCH_SOURCE_NOT_FOUND');
    assert.equal(excludedBefore.error.recovery.requiresReinspection, true);

    const repeatedWithinOne = await writeFixture(directory, 'repeated-within-one.docx', [
        { id: 'B4', text: 'The old value applies, and the old value remains.' }
    ]);
    const repeatedOutput = path.join(directory, 'repeated-within-one-output.docx');
    const repeatedSpan = await executeCli([
        'apply', repeatedWithinOne,
        '--find', 'old value',
        '--replace', 'new value',
        '--output', repeatedOutput
    ]);
    assert.equal(repeatedSpan.status, 'error');
    assert.equal(repeatedSpan.results[0].error.code, 'AMBIGUOUS_PATCH_SOURCE');
    assert.equal(repeatedSpan.results[0].error.candidates.length, 2);
    await assert.rejects(readFile(repeatedOutput), error => error.code === 'ENOENT');

    const selectedOccurrence = await executeCli([
        'apply', repeatedWithinOne,
        '--find', 'old value',
        '--replace', 'new value',
        '--occurrence', '2'
    ]);
    assert.equal(selectedOccurrence.status, 'ok', JSON.stringify(selectedOccurrence));
    assert.equal(selectedOccurrence.results[0].change.replacements[0].occurrence, 2);

    const staleIdOutput = path.join(directory, 'stale-id-output.docx');
    const staleId = await executeCli([
        'apply', globalInput,
        '--target-id', 'DOES-NOT-EXIST',
        '--find', 'thirty (30) days',
        '--replace', 'sixty (60) days',
        '--output', staleIdOutput
    ]);
    assert.equal(staleId.status, 'error');
    assert.equal(staleId.results[0].error.code, 'TARGET_NOT_FOUND');
    await assert.rejects(readFile(staleIdOutput), error => error.code === 'ENOENT');

    const overlapInput = await writeFixture(directory, 'overlap.docx', [
        { id: 'C1', text: 'Anchor text.' },
        { id: 'C2', text: 'Change old value here.' },
        { id: 'C3', text: 'Another anchor text.' }
    ]);
    const overlap = await executeCli([
        'apply', overlapInput,
        '--search', 'anchor text',
        '--around', '1',
        '--find', 'old value',
        '--replace', 'new value'
    ]);
    assert.equal(overlap.status, 'ok', JSON.stringify(overlap));
    assert.equal(overlap.results[0].change.context.anchorMatchCount, 2);
    assert.equal(overlap.results[0].change.context.range, '-1:1');

    const longText = `  ${'leading context '.repeat(12)}old literal${' trailing context'.repeat(12)}  `;
    const longInput = await writeFixture(directory, 'long-summary.docx', [
        { id: 'C4', text: longText }
    ]);
    const longResult = await executeCli([
        'apply', longInput,
        '--find', 'old literal',
        '--replace', 'new literal'
    ]);
    assert.equal(longResult.status, 'ok', JSON.stringify(longResult));
    const longChange = longResult.results[0].change;
    assert(longChange.replacements[0].beforeExcerpt.length <= 120);
    assert(longChange.replacements[0].afterExcerpt.length <= 120);
    assert.match(longChange.replacements[0].beforeExcerpt, /old literal/);
    assert.match(longChange.replacements[0].afterExcerpt, /new literal/);
    assert.equal(JSON.stringify(longChange).includes(longText), false);

    const ambiguousInput = await writeFixture(directory, 'ambiguous.docx', [
        { id: 'D1', text: 'Payment is due in thirty days.' },
        { id: 'D2', text: 'Notice is due in thirty days.' }
    ]);
    const protectedOutput = path.join(directory, 'protected.docx');
    const sentinel = Buffer.from('do-not-overwrite');
    await writeFile(protectedOutput, sentinel);
    const ambiguous = await executeCli([
        'apply', ambiguousInput,
        '--find', 'thirty days',
        '--replace', 'sixty days',
        '--occurrence', '1',
        '--output', protectedOutput,
        '--force'
    ]);
    assert.equal(ambiguous.status, 'error');
    assert.equal(ambiguous.error.code, 'AMBIGUOUS_TARGET');
    assert.equal(ambiguous.error.candidates.length, 2);
    assert.equal(ambiguous.error.recovery.action, 'choose_candidate');
    assert.deepEqual(await readFile(protectedOutput), sentinel);
    let stdout = '';
    const ambiguousExitCode = await runCli([
        'apply', ambiguousInput,
        '--find', 'thirty days',
        '--replace', 'sixty days',
        '--profile', 'agent'
    ], { stdout: { write: value => { stdout += value; } } });
    assert.equal(ambiguousExitCode, 2);
    assert.equal(JSON.parse(stdout).error.code, 'AMBIGUOUS_TARGET');

    const missingAnchor = await executeCli([
        'apply', globalInput,
        '--search', 'Section 99.9',
        '--context-range', '1:3',
        '--find', 'thirty (30) days',
        '--replace', 'sixty (60) days'
    ]);
    assert.equal(missingAnchor.error.code, 'TARGET_NOT_FOUND');
    assert.equal(missingAnchor.error.context.search, 'Section 99.9');

    const invalidRanges = [
        ['--context-range', '3:1'],
        ['--context-range', '-21:1'],
        ['--context-range', 'word'],
        ['--context-range', '0:1', '--around', '1']
    ];
    for (const rangeArgs of invalidRanges) {
        const invalid = await executeCli([
            'apply', globalInput,
            '--search', 'Section 4.1',
            ...rangeArgs,
            '--find', 'thirty (30) days',
            '--replace', 'sixty (60) days'
        ]);
        assert.equal(invalid.error.code, 'INVALID_FILTER', JSON.stringify(invalid));
    }
} finally {
    await rm(directory, { recursive: true, force: true });
}

console.log('CLI speculative apply tests passed');
