import assert from 'node:assert/strict';
import { mkdir, mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { performance } from 'node:perf_hooks';
import { Readable } from 'node:stream';
import { configureLogger } from '../adapters/logger.js';
import { executeCli } from '../node/cli.js';

configureLogger({ log() {}, warn() {}, error() {} });

const fixturePath = path.resolve('tests/fixtures/sample_doc_test.docx');
const reportPath = path.resolve('tmp/benchmarks/localized-turn-reduction-latest.json');
const iterations = Math.max(5, Number.parseInt(process.env.DOCX_LOCALIZED_BENCH_ITERATIONS || '11', 10));
const warmups = Math.max(1, Number.parseInt(process.env.DOCX_LOCALIZED_BENCH_WARMUPS || '3', 10));
const author = 'Localized Performance Benchmark';

const cases = Object.freeze([
    {
        id: 'global-short-punctuation',
        index: 37,
        find: 'Materials',
        replace: 'Materials.'
    },
    {
        id: 'global-long-provision',
        index: 53,
        find: 'addresses listed above',
        replace: 'addresses and email contacts listed above'
    },
    {
        id: 'directional-heading-scope',
        index: 32,
        search: '4. TERM',
        contextRange: '1:1',
        find: 'five (5) years',
        replace: 'three (3) years'
    }
]);

function percentile(values, ratio) {
    const sorted = [...values].sort((left, right) => left - right);
    return sorted[Math.min(sorted.length - 1, Math.ceil(sorted.length * ratio) - 1)];
}

function timingSummary(samples) {
    const values = samples.map(sample => sample.elapsedMs);
    return {
        medianMs: Number(percentile(values, 0.5).toFixed(2)),
        p95Ms: Number(percentile(values, 0.95).toFixed(2))
    };
}

function byteLength(value) {
    return Buffer.byteLength(typeof value === 'string' ? value : JSON.stringify(value), 'utf8');
}

function tokenProxy(value) {
    return Math.ceil(byteLength(value) / 4);
}

function percentReduction(before, after) {
    return Number(((1 - after / before) * 100).toFixed(2));
}

function canonicalTarget(paragraph) {
    return {
        exactText: paragraph.exactText,
        ...(paragraph.paragraphId ? { paragraphId: paragraph.paragraphId } : {}),
        ...(paragraph.fingerprint ? { fingerprint: paragraph.fingerprint } : {}),
        inTable: paragraph.inTable,
        revisionView: paragraph.revisionView
    };
}

function speculativeRequest(testCase) {
    return {
        command: 'apply',
        ...(testCase.search ? {
            search: testCase.search,
            contextRange: testCase.contextRange
        } : {}),
        find: testCase.find,
        replace: testCase.replace,
        profile: 'agent'
    };
}

const fixture = await readFile(fixturePath);
const directory = await mkdtemp(path.join(tmpdir(), 'docx-localized-benchmark-'));
const inputPath = path.join(directory, 'input.docx');
await writeFile(inputPath, fixture);

async function baselineSample(testCase, sampleIndex) {
    const started = performance.now();
    const extraction = await executeCli(['extract', inputPath, '--index', String(testCase.index)]);
    assert.equal(extraction.status, 'ok');
    const paragraph = extraction.paragraphs[0];
    const desiredText = paragraph.exactText.replace(testCase.find, testCase.replace);
    assert.notEqual(desiredText, paragraph.exactText);
    const operation = {
        type: 'redline',
        target: canonicalTarget(paragraph),
        modified: desiredText
    };
    const payload = JSON.stringify({ operations: [operation] });
    const outputPath = path.join(directory, `${testCase.id}-baseline-${sampleIndex}.docx`);
    const applied = await executeCli([
        'apply', inputPath,
        '--operations', '-',
        '--profile', 'agent',
        '--author', author,
        '--output', outputPath
    ], { stdin: Readable.from([payload]) });
    const elapsedMs = performance.now() - started;
    assert.equal(applied.status, 'ok', JSON.stringify(applied));
    assert.equal(applied.completion, true);
    const accepted = await executeCli(['extract', outputPath, '--index', String(testCase.index)]);
    assert.equal(accepted.paragraphs[0].exactText, desiredText);
    return {
        elapsedMs,
        generatedRequestBytes: byteLength(payload),
        generatedTokenProxy: tokenProxy(payload),
        providerVisibleResponseBytes: byteLength(extraction) + byteLength(applied),
        toolTurns: 2,
        desiredText
    };
}

async function speculativeSample(testCase, sampleIndex, expectedText) {
    const request = speculativeRequest(testCase);
    const outputPath = path.join(directory, `${testCase.id}-speculative-${sampleIndex}.docx`);
    const args = [
        'apply', inputPath,
        ...(testCase.search ? [
            '--search', testCase.search,
            '--context-range', testCase.contextRange
        ] : []),
        '--find', testCase.find,
        '--replace', testCase.replace,
        '--profile', 'agent',
        '--author', author,
        '--output', outputPath
    ];
    const started = performance.now();
    const applied = await executeCli(args);
    const elapsedMs = performance.now() - started;
    assert.equal(applied.status, 'ok', JSON.stringify(applied));
    assert.equal(applied.completion, true);
    assert.equal(applied.results[0].change.committed, true);
    assert.equal(applied.results[0].change.verification.acceptedViewMatchesCompiledText, true);
    const accepted = await executeCli(['extract', outputPath, '--index', String(testCase.index)]);
    assert.equal(accepted.paragraphs[0].exactText, expectedText);
    return {
        elapsedMs,
        generatedRequestBytes: byteLength(request),
        generatedTokenProxy: tokenProxy(request),
        providerVisibleResponseBytes: byteLength(applied),
        toolTurns: 1
    };
}

async function measureCase(testCase) {
    for (let index = 0; index < warmups; index += 1) {
        const baseline = await baselineSample(testCase, `warmup-${index}`);
        await speculativeSample(testCase, `warmup-${index}`, baseline.desiredText);
    }
    const baseline = [];
    const speculative = [];
    for (let index = 0; index < iterations; index += 1) {
        const baselineResult = await baselineSample(testCase, index);
        baseline.push(baselineResult);
        speculative.push(await speculativeSample(testCase, index, baselineResult.desiredText));
    }
    const baselineShape = baseline[0];
    const speculativeShape = speculative[0];
    return {
        id: testCase.id,
        selector: testCase.search
            ? { search: testCase.search, contextRange: testCase.contextRange }
            : { scope: 'global' },
        baseline06Style: {
            localEngineAndIo: timingSummary(baseline),
            toolTurns: baselineShape.toolTurns,
            generatedRequestBytes: baselineShape.generatedRequestBytes,
            generatedTokenProxy: baselineShape.generatedTokenProxy,
            providerVisibleResponseBytes: baselineShape.providerVisibleResponseBytes
        },
        contract8Speculative: {
            localEngineAndIo: timingSummary(speculative),
            toolTurns: speculativeShape.toolTurns,
            generatedRequestBytes: speculativeShape.generatedRequestBytes,
            generatedTokenProxy: speculativeShape.generatedTokenProxy,
            providerVisibleResponseBytes: speculativeShape.providerVisibleResponseBytes
        },
        reduction: {
            toolTurnsPercent: percentReduction(baselineShape.toolTurns, speculativeShape.toolTurns),
            generatedRequestBytesPercent: percentReduction(
                baselineShape.generatedRequestBytes,
                speculativeShape.generatedRequestBytes
            ),
            generatedTokenProxyPercent: percentReduction(
                baselineShape.generatedTokenProxy,
                speculativeShape.generatedTokenProxy
            ),
            medianLocalEngineAndIoPercent: percentReduction(
                timingSummary(baseline).medianMs,
                timingSummary(speculative).medianMs
            )
        }
    };
}

let report;
try {
    const results = [];
    for (const testCase of cases) results.push(await measureCase(testCase));
    assert((await readFile(inputPath)).equals(fixture));
    report = {
        generatedAt: new Date().toISOString(),
        environment: { node: process.version, platform: process.platform, arch: process.arch },
        configuration: { iterations, warmups, fixture: path.relative(process.cwd(), fixturePath) },
        comparison: {
            baseline: '0.6.2-style focused extract followed by a full-paragraph canonical operation',
            candidate: 'contract-8 targetless localized apply',
            controlledRuntime: 'Both paths run on the current source tree so timing isolates workflow shape, not release/runtime drift.'
        },
        measurements: {
            localEngineAndIo: 'Measured in-process with performance.now(); includes DOCX reads, inspection, mutation, validation, and output write.',
            generatedTokenProxy: 'Deterministic ceil(UTF-8 request bytes / 4); not a provider tokenizer or billed-token count.',
            providerVisibleResponseBytes: 'Serialized JSON bytes returned across all workflow tool turns.',
            modelAndNetwork: {
                measured: false,
                reason: 'Provider reasoning, streaming, transport, and end-to-end harness latency require external transcript instrumentation.'
            }
        },
        results,
        invariants: {
            sourceByteIdentical: true,
            acceptedOutputVerified: true,
            localizedChangeEvidenceVerified: true
        }
    };
} finally {
    await rm(directory, { recursive: true, force: true });
}

await mkdir(path.dirname(reportPath), { recursive: true });
await writeFile(reportPath, `${JSON.stringify(report, null, 2)}\n`, 'utf8');
console.log(JSON.stringify(report, null, 2));
