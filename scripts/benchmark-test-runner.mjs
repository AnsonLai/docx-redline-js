import { mkdirSync, writeFileSync } from 'node:fs';
import * as os from 'node:os';
import { dirname, join } from 'node:path';
import { performance } from 'node:perf_hooks';
import { spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';

const repoRoot = join(dirname(fileURLToPath(import.meta.url)), '..');
const measuredIterations = Number.parseInt(process.env.DOCX_TEST_BENCH_ITERATIONS || '3', 10);
const warmups = 1;
const parallelConcurrency = Math.max(1, Math.min(4, os.availableParallelism?.() ?? os.cpus().length));

function run(concurrency) {
    const started = performance.now();
    const result = spawnSync(process.execPath, ['scripts/run-tests.mjs'], {
        cwd: repoRoot,
        env: { ...process.env, DOCX_TEST_CONCURRENCY: String(concurrency) },
        encoding: 'utf8',
        windowsHide: true,
        maxBuffer: 10 * 1024 * 1024
    });
    if (result.status !== 0) {
        process.stderr.write(result.stdout || '');
        process.stderr.write(result.stderr || '');
        throw new Error(`Test runner failed in concurrency=${concurrency} benchmark mode.`);
    }
    return performance.now() - started;
}

function percentile(values, ratio) {
    const sorted = values.slice().sort((a, b) => a - b);
    return sorted[Math.min(sorted.length - 1, Math.floor(sorted.length * ratio))];
}

for (let index = 0; index < warmups; index++) {
    run(1);
    run(parallelConcurrency);
}

const serial = [];
const parallel = [];
for (let index = 0; index < measuredIterations; index++) {
    serial.push(run(1));
    parallel.push(run(parallelConcurrency));
}

const result = {
    runtime: { node: process.version, platform: process.platform, arch: process.arch },
    fixture: { warmups, measuredIterations, parallelConcurrency },
    serialMs: { median: percentile(serial, 0.5), p95: percentile(serial, 0.95), samples: serial },
    parallelMs: { median: percentile(parallel, 0.5), p95: percentile(parallel, 0.95), samples: parallel }
};
result.medianSpeedup = result.serialMs.median / result.parallelMs.median;
result.medianReductionPercent = (1 - (result.parallelMs.median / result.serialMs.median)) * 100;

const outputDir = join(repoRoot, 'tmp', 'benchmarks');
mkdirSync(outputDir, { recursive: true });
writeFileSync(join(outputDir, 'test-runner-latest.json'), `${JSON.stringify(result, null, 2)}\n`);
process.stdout.write(`${JSON.stringify(result, null, 2)}\n`);
