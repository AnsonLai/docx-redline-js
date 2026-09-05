import { readdirSync } from 'node:fs';
import * as os from 'node:os';
import { join, dirname, resolve } from 'node:path';
import { execFile } from 'node:child_process';
import { fileURLToPath } from 'node:url';

const scriptPath = fileURLToPath(import.meta.url);
const scriptDir = dirname(scriptPath);
const defaultTestDir = join(scriptDir, '..', 'tests');
const failOutputPattern = /(?:❌\s*(?:FAIL|FAILED|FAILURE|TEST FAILED)|\bTEST FAILED\b)/i;

export function discoverTestFiles(testDir = defaultTestDir) {
  return readdirSync(testDir)
    .filter(file => file.endsWith('.mjs') && file !== 'setup-xml-provider.mjs')
    .sort();
}

export function resolveTestConcurrency(value = process.env.DOCX_TEST_CONCURRENCY) {
  if (value == null || value === '') {
    const available = os.availableParallelism?.() ?? os.cpus().length;
    return Math.max(1, Math.min(4, available));
  }
  if (!/^\d+$/.test(String(value)) || Number(value) < 1) {
    throw new Error('DOCX_TEST_CONCURRENCY must be a positive integer.');
  }
  return Number(value);
}

export function runTestFile(file, options = {}) {
  const testDir = options.testDir || defaultTestDir;
  const timeout = options.timeout ?? 30000;
  const filePath = join(testDir, file);

  return new Promise(resolveResult => {
    execFile(process.execPath, [filePath], {
      cwd: options.cwd || join(scriptDir, '..'),
      env: options.env || process.env,
      timeout,
      maxBuffer: options.maxBuffer ?? 1024 * 1024,
      encoding: 'utf8',
      windowsHide: true
    }, (error, stdout = '', stderr = '') => {
      let executionError = error;
      if (!executionError && failOutputPattern.test(stdout)) {
        executionError = new Error(`Test printed a failure marker while exiting successfully:\n${stdout}`);
      }
      resolveResult({
        file,
        passed: !executionError,
        stdout: String(stdout),
        stderr: String(stderr),
        error: executionError || null
      });
    });
  });
}

export async function runTestFiles(files, options = {}) {
  const orderedFiles = [...files].sort();
  const concurrency = Math.min(resolveTestConcurrency(options.concurrency), Math.max(orderedFiles.length, 1));
  const execute = options.runFile || runTestFile;
  const results = new Array(orderedFiles.length);
  let nextIndex = 0;

  async function worker() {
    while (true) {
      const index = nextIndex++;
      if (index >= orderedFiles.length) return;
      results[index] = await execute(orderedFiles[index], options);
    }
  }

  await Promise.all(Array.from({ length: concurrency }, () => worker()));
  return results;
}

export function formatTestRun(results) {
  const lines = results.map(result => `  ${result.file} ... ${result.passed ? 'PASS' : 'FAIL'}`);
  const failures = results.filter(result => !result.passed);
  const passed = results.length - failures.length;
  lines.push('', `${passed} passed, ${failures.length} failed out of ${results.length} tests`);
  const failureLines = [];
  for (const failure of failures) {
    const detail = [failure.stdout, failure.stderr, failure.error?.message].filter(Boolean).join('\n');
    failureLines.push('', `--- ${failure.file} ---`, detail);
  }
  return {
    output: `${lines.join('\n')}\n`,
    errorOutput: failureLines.length > 0 ? `${failureLines.join('\n')}\n` : '',
    passed,
    failed: failures.length
  };
}

export async function main(options = {}) {
  const testDir = options.testDir || defaultTestDir;
  const files = options.files || discoverTestFiles(testDir);
  const results = await runTestFiles(files, { ...options, testDir });
  const summary = formatTestRun(results);
  const write = options.write || (text => process.stdout.write(text));
  const writeError = options.writeError || (text => process.stderr.write(text));
  write(summary.output);
  if (summary.errorOutput) writeError(summary.errorOutput);
  return summary;
}

if (process.argv[1] && resolve(process.argv[1]) === resolve(scriptPath)) {
  try {
    const summary = await main();
    if (summary.failed > 0) process.exitCode = 1;
  } catch (error) {
    console.error(error?.stack || error?.message || String(error));
    process.exitCode = 1;
  }
}
