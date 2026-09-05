import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';

import {
    discoverTestFiles,
    formatTestRun,
    resolveTestConcurrency,
    runTestFile,
    runTestFiles
} from '../scripts/run-tests.mjs';

const directory = await mkdtemp(join(tmpdir(), 'docx-redline-test-runner-'));
try {
    await writeFile(join(directory, 'setup-xml-provider.mjs'), 'throw new Error("must be excluded");');
    await writeFile(join(directory, 'z-pass.mjs'), 'console.log(`pass:${process.pid}`);');
    await writeFile(join(directory, 'a-pass.mjs'), 'console.log(`pass:${process.pid}`);');
    await writeFile(join(directory, 'marker.mjs'), 'console.log("TEST FAILED despite exit zero");');
    await writeFile(join(directory, 'nonzero.mjs'), 'console.log("captured stdout"); console.error("captured stderr"); process.exit(3);');
    await writeFile(join(directory, 'timeout.mjs'), 'setTimeout(() => {}, 1000);');

    assert.deepEqual(discoverTestFiles(directory), [
        'a-pass.mjs',
        'marker.mjs',
        'nonzero.mjs',
        'timeout.mjs',
        'z-pass.mjs'
    ]);
    assert.equal(resolveTestConcurrency('1'), 1);
    assert.equal(resolveTestConcurrency('4'), 4);
    assert.ok(resolveTestConcurrency() >= 1 && resolveTestConcurrency() <= 4);
    assert.throws(() => resolveTestConcurrency('0'), /positive integer/);
    assert.throws(() => resolveTestConcurrency('fast'), /positive integer/);

    let active = 0;
    let maximumActive = 0;
    const scheduled = await runTestFiles(['g', 'f', 'e', 'd', 'c', 'b', 'a'], {
        concurrency: 3,
        async runFile(file) {
            active++;
            maximumActive = Math.max(maximumActive, active);
            await new Promise(resolve => setTimeout(resolve, 5));
            active--;
            return { file, passed: true, stdout: '', stderr: '', error: null };
        }
    });
    assert.equal(maximumActive, 3);
    assert.deepEqual(scheduled.map(result => result.file), ['a', 'b', 'c', 'd', 'e', 'f', 'g']);

    const passingFiles = ['z-pass.mjs', 'a-pass.mjs'];
    const serial = await runTestFiles(passingFiles, { testDir: directory, concurrency: 1 });
    const parallel = await runTestFiles(passingFiles, { testDir: directory, concurrency: 2 });
    assert.deepEqual(serial.map(result => [result.file, result.passed]), [
        ['a-pass.mjs', true],
        ['z-pass.mjs', true]
    ]);
    assert.deepEqual(
        parallel.map(result => [result.file, result.passed]),
        serial.map(result => [result.file, result.passed])
    );
    assert.equal(new Set(parallel.map(result => result.stdout.match(/pass:(\d+)/)?.[1])).size, 2,
        'each parallel test must execute in its own process');

    const marker = await runTestFile('marker.mjs', { testDir: directory });
    assert.equal(marker.passed, false);
    assert.match(marker.error.message, /failure marker/i);

    const nonzero = await runTestFile('nonzero.mjs', { testDir: directory });
    assert.equal(nonzero.passed, false);
    const nonzeroSummary = formatTestRun([nonzero]);
    assert.match(nonzeroSummary.errorOutput, /captured stdout/);
    assert.match(nonzeroSummary.errorOutput, /captured stderr/);
    assert.equal(nonzeroSummary.failed, 1);

    const timeout = await runTestFile('timeout.mjs', { testDir: directory, timeout: 30 });
    assert.equal(timeout.passed, false);
    assert.match(timeout.error.message, /timed out|timeout/i);

    const mixedFiles = ['nonzero.mjs', ...passingFiles];
    const mixedSerial = await runTestFiles(mixedFiles, {
        testDir: directory,
        concurrency: 1
    });
    const mixed = await runTestFiles(mixedFiles, {
        testDir: directory,
        concurrency: 3
    });
    assert.deepEqual(mixed.map(result => result.file), ['a-pass.mjs', 'nonzero.mjs', 'z-pass.mjs']);
    assert.deepEqual(mixed.map(result => result.passed), [true, false, true]);
    assert.deepEqual(
        mixed.map(result => [result.file, result.passed]),
        mixedSerial.map(result => [result.file, result.passed])
    );
} finally {
    await rm(directory, { recursive: true, force: true });
}

console.log('PASS: Phase 6 bounded parallel test runner preserves isolation and failures');
