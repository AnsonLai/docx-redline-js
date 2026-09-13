import assert from 'node:assert/strict';
import { access, mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { Readable } from 'node:stream';
import { executeCli, runCli } from '../node/cli.js';
import { openDocx } from '../node/index.js';

const fixtureUrl = new URL('./fixtures/sample_doc_test.docx', import.meta.url);
const fixture = await readFile(fixtureUrl);
const directory = await mkdtemp(path.join(tmpdir(), 'docx-agent-cli-protocol-'));

async function runWithStdin(args, payload) {
    let stdout = '';
    const exitCode = await runCli(args, {
        stdin: Readable.from([payload]),
        stdout: { write: value => { stdout += value; } }
    });
    return { exitCode, result: JSON.parse(stdout) };
}

try {
    const input = path.join(directory, 'input.docx');
    await writeFile(input, fixture);
    const version = await executeCli(['version']);
    assert.equal(version.contractVersion, 5);
    assert(version.capabilities.includes('operations-stdin'));
    assert(version.capabilities.includes('agent-profile-v1'));

    const inspected = await executeCli(['extract', input, '--index', '53']);
    const target = inspected.paragraphs[0];
    const desiredText = `${target.exactText} — confirmed under § 2.`;
    const operation = {
        operationId: 'stdin-unicode-edit',
        type: 'redline',
        target: {
            exactText: target.exactText,
            paragraphId: target.paragraphId,
            fingerprint: target.fingerprint
        },
        modified: desiredText
    };
    const output = path.join(directory, 'stdin-output.docx');
    const packageRevision = openDocx(fixture).getRevisionToken();
    const applied = await runWithStdin([
        'apply', input,
        '--operations', '-',
        '--profile', 'agent',
        '--author', 'Stdin Agent',
        '--output', output
    ], JSON.stringify({ operations: [operation], expectedRevision: packageRevision }));
    assert.equal(applied.exitCode, 0, JSON.stringify(applied.result));
    assert.equal(applied.result.status, 'ok');
    assert.equal(applied.result.completion, true);
    assert.equal(applied.result.executionProfile, 'agent');
    assert.deepEqual(applied.result.effectiveOptions, {
        author: 'Stdin Agent',
        atomic: true,
        strictTargets: true,
        validate: true,
        generateRedlines: true,
        existingRevisions: 'merge-same-author',
        requireComplete: true
    });
    assert.deepEqual(await readFile(input), fixture, 'stdin apply must not overwrite the source');
    assert.equal((await executeCli(['extract', output, '--index', '53'])).paragraphs[0].exactText, desiredText);
    assert.equal((await executeCli(['validate', output, '--baseline', input])).valid, true);
    const rejected = await executeCli(['reject', output, '--author', 'Stdin Agent']);
    assert.equal(rejected.status, 'ok');
    assert.equal((await executeCli(['extract', rejected.outputPath, '--index', '53'])).paragraphs[0].exactText, target.exactText);

    const partialOutput = path.join(directory, 'partial-output.docx');
    const partial = await runWithStdin([
        'apply', input,
        '--operations', '-',
        '--profile', 'agent',
        '--atomic=false',
        '--output', partialOutput
    ], JSON.stringify([
        operation,
        { type: 'redline', target: 'Missing stdin target.', modified: 'Still missing.' }
    ]));
    assert.equal(partial.exitCode, 3);
    assert.equal(partial.result.status, 'partial');
    assert.equal(partial.result.written, true);
    assert.equal(partial.result.completion, false);
    assert.equal(partial.result.effectiveOptions.atomic, false);
    assert.equal(partial.result.effectiveOptions.requireComplete, true);
    assert.equal(partial.result.retryPlan.base, 'output');
    await access(partialOutput);

    const failedOutput = path.join(directory, 'atomic-failure.docx');
    const failed = await runWithStdin([
        'apply', input,
        '--operations', '-',
        '--profile', 'agent',
        '--output', failedOutput
    ], JSON.stringify([
        operation,
        { type: 'redline', target: 'Missing atomic target.', modified: 'Still missing.' }
    ]));
    assert.equal(failed.exitCode, 2);
    assert.equal(failed.result.status, 'error');
    assert.equal(failed.result.written, false);
    assert.equal(failed.result.retryPlan.base, 'original');
    await assert.rejects(access(failedOutput));

    const malformed = await executeCli([
        'apply', input, '--operations', '-'
    ], { stdin: Readable.from(['{not json']) });
    assert.equal(malformed.status, 'error');
    assert.equal(malformed.error.code, 'INVALID_OPERATIONS_FILE');
    assert.equal(malformed.error.category, 'request_fixable');

    const unknownProfile = await executeCli([
        'apply', input, '--operations', '-', '--profile', 'fast-and-loose'
    ], { stdin: Readable.from(['[]']) });
    assert.equal(unknownProfile.error.code, 'INVALID_PROFILE');
} finally {
    await rm(directory, { recursive: true, force: true });
}

console.log('agent CLI protocol tests passed');
