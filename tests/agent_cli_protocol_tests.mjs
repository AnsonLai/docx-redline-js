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
    assert.equal(version.contractVersion, 8);
    assert(version.capabilities.includes('operations-stdin'));
    assert(version.capabilities.includes('agent-safety-profile-v2'));
    assert(version.capabilities.includes('deduplicated-cli-receipts'));
    assert(version.capabilities.includes('compact-cli-json-v1'));
    assert.equal(version.capabilities.includes('agent-profile-v1'), false);
    assert(version.capabilities.includes('command-help-v1'));
    assert(version.capabilities.includes('inspection-context-v1'));
    assert(version.capabilities.includes('bounded-inspection-v1'));
    assert(version.capabilities.includes('human-document-references-v1'));
    assert(version.capabilities.includes('localized-replacements-v1'));
    assert(version.capabilities.includes('speculative-search-apply-v1'));
    assert(version.capabilities.includes('localized-change-summary-v1'));

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
    const nativeResult = await openDocx(fixture).applyOperations([operation], {
        author: 'Native Agent',
        atomic: true
    });
    assert(nativeResult.results[0].receipt, 'Node facade retains per-operation receipts');
    assert.equal(nativeResult.results[0].resolvedTarget.targetTextMatch.mode, 'exact');
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
        atomic: false,
        strictTargets: true,
        validate: true,
        generateRedlines: true,
        existingRevisions: 'merge-same-author',
        requireComplete: true
    });
    assert.equal(applied.result.results[0].receipt, undefined);
    assert.equal(applied.result.receipts.length, 1);
    assert.equal(applied.result.receipts[0].operationIndex, 1);
    assert.equal(applied.result.receipts[0].authorUsed, 'Stdin Agent');
    assert.equal(applied.result.results[0].resolvedTarget.targetTextMatch, undefined);
    assert.deepEqual(await readFile(input), fixture, 'stdin apply must not overwrite the source');
    assert.equal((await executeCli(['extract', output, '--index', '53'])).paragraphs[0].exactText, desiredText);
    assert.equal((await executeCli(['validate', output, '--baseline', input])).valid, true);
    const rejected = await executeCli(['reject', output, '--author', 'Stdin Agent']);
    assert.equal(rejected.status, 'ok');
    assert.equal((await executeCli(['extract', rejected.outputPath, '--index', '53'])).paragraphs[0].exactText, target.exactText);

    const operationsFile = path.join(directory, 'unicode-operations.json');
    const fileOutput = path.join(directory, 'file-output.docx');
    await writeFile(operationsFile, JSON.stringify({ operations: [operation], expectedRevision: packageRevision }));
    const fileApplied = await executeCli([
        'apply', input,
        '--operations', operationsFile,
        '--profile', 'agent',
        '--author', 'Stdin Agent',
        '--output', fileOutput
    ]);
    assert.equal(fileApplied.completion, true, JSON.stringify(fileApplied));
    assert.equal((await executeCli(['extract', fileOutput, '--index', '53'])).paragraphs[0].exactText, desiredText);
    assert.deepEqual(fileApplied.effectiveOptions, applied.result.effectiveOptions);
    assert.deepEqual(fileApplied.authorsUsed, applied.result.authorsUsed);

    const partialOutput = path.join(directory, 'partial-output.docx');
    const partial = await runWithStdin([
        'apply', input,
        '--operations', '-',
        '--profile', 'agent',
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
    assert.equal(partial.result.effectiveOptions.author, 'AI Redliner');
    assert.equal(partial.result.receipts[0].authorUsed, 'AI Redliner');
    assert.equal(partial.result.retryPlan.base, 'output');
    await access(partialOutput);

    const failedOutput = path.join(directory, 'atomic-failure.docx');
    const failed = await runWithStdin([
        'apply', input,
        '--operations', '-',
        '--profile', 'agent',
        '--atomic',
        '--output', failedOutput
    ], JSON.stringify([
        operation,
        { type: 'redline', target: 'Missing atomic target.', modified: 'Still missing.' }
    ]));
    assert.equal(failed.exitCode, 2);
    assert.equal(failed.result.status, 'error');
    assert.equal(failed.result.written, false);
    assert.equal(failed.result.retryPlan.base, 'original');
    assert.equal(failed.result.effectiveOptions.atomic, true);
    assert.equal(failed.result.results.every(item => item.receipt === undefined), true);
    assert.equal(failed.result.receipts.length, 2);
    assert.deepEqual(failed.result.receipts.map(item => item.operationIndex), [1, 2]);
    assert.equal(failed.result.receipts.some(item => item.finalDisposition === 'rolled_back'), true);
    await assert.rejects(access(failedOutput));

    const policyOutput = path.join(directory, 'policy-output.docx');
    const policy = await runWithStdin([
        'apply', input,
        '--operations', '-',
        '--profile', 'agent',
        '--atomic=false',
        '--existing-revisions', 'slice-cross-author',
        '--output', policyOutput
    ], JSON.stringify([operation]));
    assert.equal(policy.exitCode, 0, JSON.stringify(policy.result));
    assert.equal(policy.result.effectiveOptions.atomic, false);
    assert.equal(policy.result.effectiveOptions.existingRevisions, 'slice-cross-author');

    const compactOutput = path.join(directory, 'compact-output.docx');
    let compactStdout = '';
    const compactExit = await runCli([
        'apply', input,
        '--operations', '-',
        '--profile', 'agent',
        '--compact',
        '--output', compactOutput
    ], {
        stdin: Readable.from([JSON.stringify([operation])]),
        stdout: { write: value => { compactStdout += value; } }
    });
    assert.equal(compactExit, 0);
    assert.equal(compactStdout.includes('\n  "'), false, 'compact mode emits one-line JSON');
    assert.equal(JSON.parse(compactStdout).completion, true);

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
