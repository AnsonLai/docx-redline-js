import assert from 'node:assert/strict';
import { mkdir, mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { performance } from 'node:perf_hooks';
import { Readable } from 'node:stream';
import { configureLogger } from '../adapters/logger.js';
import { executeCli } from '../node/cli.js';
import { buildZip } from './lib/minimal-zip.mjs';

configureLogger({ log() {}, warn() {}, error() {} });

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const contentTypes = '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>';
const rels = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>';
const reportPath = path.resolve('tmp/benchmarks/agent-surface-simplification-latest.json');
const iterations = Math.max(3, Number.parseInt(process.env.DOCX_AGENT_SURFACE_ITERATIONS || '7', 10));
const warmups = Math.max(0, Number.parseInt(process.env.DOCX_AGENT_SURFACE_WARMUPS || '1', 10));
const author = 'Synthetic Workflow Benchmark';

function escapeXml(value) {
    return String(value).replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;');
}

function paragraph(id, text) {
    return `<w:p w:paraId="${id}"><w:r><w:t>${escapeXml(text)}</w:t></w:r></w:p>`;
}

function deletedParagraph(id, text) {
    return `<w:p w:paraId="${id}"><w:pPr><w:rPr><w:del w:id="41" w:author="Earlier Reviewer" w:date="2026-01-01T00:00:00Z"/></w:rPr></w:pPr><w:del w:id="42" w:author="Earlier Reviewer" w:date="2026-01-01T00:00:00Z"><w:r><w:delText>${escapeXml(text)}</w:delText></w:r></w:del></w:p>`;
}

function fixture(body) {
    const documentXml = `<w:document xmlns:w="${W}"><w:body>${body}<w:sectPr/></w:body></w:document>`;
    return buildZip([
        { name: '[Content_Types].xml', data: contentTypes },
        { name: 'word/document.xml', data: documentXml },
        { name: 'word/_rels/document.xml.rels', data: rels }
    ]);
}

const scenarios = Object.freeze([
    {
        id: 'literal-long-paragraph',
        kind: 'literal',
        input: fixture(paragraph('L1', `${'Background material remains unchanged. '.repeat(12)}The inspection cycle is thirty days. ${'Additional guidance remains unchanged. '.repeat(12)}`)),
        indexes: [1],
        find: 'thirty days',
        replace: 'forty-five days'
    },
    {
        id: 'semantic-contextual-rewrite',
        kind: 'semantic',
        input: fixture(paragraph('S1', 'The studio may cancel a reservation after written notice. The operator must return all issued equipment.')),
        indexes: [1],
        desired: ['Either the studio or the operator may cancel a reservation after written notice. The operator must return all issued equipment.']
    },
    {
        id: 'repeated-visible-literal',
        kind: 'repeated',
        input: fixture(
            paragraph('R1', 'Place the blue marker beside the north sensor.')
            + paragraph('R2', 'Place the blue marker beside the south sensor.')
        ),
        indexes: [1],
        find: 'blue marker',
        replace: 'green marker'
    },
    {
        id: 'rejected-view-restore',
        kind: 'restore',
        input: fixture(
            deletedParagraph('D1', 'The rover calibration log must be retained for two operating cycles.')
            + paragraph('D2', 'Routine maintenance continues under the current field guide.')
        ),
        indexes: [1],
        restoreText: 'The rover calibration log must be retained for two operating cycles.'
    },
    {
        id: 'nbsp-localized-target',
        kind: 'nbsp',
        input: fixture(paragraph('N1', 'The Calibration Manual is stored at\u00a0docs.example/calibration/\u00a0(the “Manual”).')),
        indexes: [1],
        find: 'Calibration Manual is stored at docs.example/calibration/ (the “Manual”).',
        replace: 'Calibration Manual is stored at docs.example/calibration/, current at activation (the “Manual”).',
        desired: ['The Calibration Manual is stored at\u00a0docs.example/calibration/, current at activation\u00a0(the “Manual”).']
    },
    {
        id: 'independent-two-target-batch',
        kind: 'batch',
        input: fixture(
            paragraph('B1', 'The north sensor is checked every ten days.')
            + paragraph('B2', 'The south sensor is checked every twenty days.')
        ),
        indexes: [1, 2],
        patches: [
            { find: 'ten days', replace: 'twelve days' },
            { find: 'twenty days', replace: 'twenty-four days' }
        ]
    }
]);

const policies = Object.freeze([
    { id: 'v0.6.2-extract-full', label: 'Extraction first with complete modified paragraphs' },
    { id: 'v0.7.1-speculation-led', label: 'Speculative literals with extract/apply recovery' },
    { id: 'simplified-extract-localized', label: 'Extraction first with localized inspected-target payloads' }
]);

function percentile(values, ratio) {
    const sorted = [...values].sort((left, right) => left - right);
    return sorted[Math.min(sorted.length - 1, Math.ceil(sorted.length * ratio) - 1)];
}

function summaries(samples, field) {
    const values = samples.map(sample => sample[field]);
    return {
        median: Number(percentile(values, 0.5).toFixed(2)),
        p95: Number(percentile(values, 0.95).toFixed(2))
    };
}

function aggregateSummary(policy, samples) {
    const total = field => samples.reduce((sum, sample) => sum + sample[field], 0);
    return {
        policy: policy.id,
        successfulRuns: samples.length,
        totalToolCalls: total('toolCalls'),
        averageToolCalls: Number((total('toolCalls') / samples.length).toFixed(2)),
        totalFailedCalls: total('failedCalls'),
        totalGeneratedEditRequestBytes: total('generatedEditRequestBytes'),
        toolCalls: summaries(samples, 'toolCalls'),
        generatedEditRequestBytes: summaries(samples, 'generatedEditRequestBytes'),
        localEngineAndIoMs: summaries(samples, 'elapsedMs')
    };
}

function bytes(value) {
    return Buffer.byteLength(typeof value === 'string' ? value : JSON.stringify(value), 'utf8');
}

function target(paragraphRecord, includeExactText = false) {
    return {
        ...(includeExactText ? { exactText: paragraphRecord.exactText } : {}),
        paragraphId: paragraphRecord.paragraphId,
        fingerprint: paragraphRecord.fingerprint,
        revisionView: paragraphRecord.revisionView,
        inTable: paragraphRecord.inTable
    };
}

function tracker(inputPath, outputPrefix) {
    const metrics = {
        toolCalls: 0,
        failedCalls: 0,
        requestBytes: 0,
        responseBytes: 0,
        elapsedMs: 0,
        generatedEditRequestBytes: 0
    };
    let outputCounter = 0;
    return {
        metrics,
        outputPath() {
            outputCounter += 1;
            return `${outputPrefix}-${outputCounter}.docx`;
        },
        async call(args, payload = null, editRequest = null) {
            metrics.toolCalls += 1;
            metrics.requestBytes += bytes(args) + (payload == null ? 0 : bytes(payload));
            if (editRequest != null) metrics.generatedEditRequestBytes += bytes(editRequest);
            const started = performance.now();
            const result = await executeCli(args, payload == null ? {} : { stdin: Readable.from([payload]) });
            metrics.elapsedMs += performance.now() - started;
            metrics.responseBytes += bytes(result);
            if (result.status === 'error' || result.status === 'partial') metrics.failedCalls += 1;
            return result;
        },
        inputPath
    };
}

async function extractRecords(run, scenario, view = 'accepted') {
    const result = await run.call([
        'extract', run.inputPath,
        '--indexes', scenario.indexes.join(','),
        '--view', view
    ]);
    assert.equal(result.status, 'ok', JSON.stringify(result));
    assert.equal(result.paragraphs.length, scenario.indexes.length);
    return result.paragraphs;
}

function desiredTexts(scenario, records) {
    if (scenario.desired) return scenario.desired;
    if (scenario.kind === 'batch') {
        return records.map((record, index) => record.exactText.replace(
            scenario.patches[index].find,
            scenario.patches[index].replace
        ));
    }
    return [records[0].exactText.replace(scenario.find, scenario.replace)];
}

async function applyOperations(run, scenario, records, localized) {
    const desired = desiredTexts(scenario, records);
    const operations = records.map((record, index) => ({
        type: scenario.kind === 'restore' ? 'restore' : 'redline',
        target: target(record, !localized || scenario.kind === 'semantic'),
        ...(scenario.kind === 'restore'
            ? { modified: scenario.restoreText }
            : localized && scenario.kind !== 'semantic'
                ? { replacements: [scenario.kind === 'batch' ? scenario.patches[index] : { find: scenario.find, replace: scenario.replace }] }
                : { modified: desired[index] })
    }));
    const payload = JSON.stringify({ operations });
    const outputPath = run.outputPath();
    const result = await run.call([
        'apply', run.inputPath,
        '--operations', '-',
        '--profile', 'agent',
        '--author', author,
        '--output', outputPath
    ], payload, operations);
    assert.equal(result.completion, true, JSON.stringify(result));
    return { outputPath, desired };
}

async function inlineRestore(run, record) {
    const outputPath = run.outputPath();
    const editRequest = { restore: true, targetId: record.paragraphId };
    const result = await run.call([
        'apply', run.inputPath,
        '--restore',
        '--target-id', record.paragraphId,
        '--profile', 'agent',
        '--author', author,
        '--output', outputPath
    ], null, editRequest);
    assert.equal(result.completion, true, JSON.stringify(result));
    return { outputPath, desired: [record.exactText] };
}

async function speculative(run, scenario) {
    const outputPath = run.outputPath();
    const editRequest = { find: scenario.find, replace: scenario.replace };
    const result = await run.call([
        'apply', run.inputPath,
        '--find', scenario.find,
        '--replace', scenario.replace,
        '--profile', 'agent',
        '--author', author,
        '--output', outputPath
    ], null, editRequest);
    return { result, outputPath };
}

async function verifyOutput(outputPath, scenario, desired) {
    if (scenario.kind === 'restore') {
        const result = await executeCli(['extract', outputPath, '--search', 'rover calibration log']);
        assert.equal(result.paragraphs.length, 1);
        assert.equal(result.paragraphs[0].exactText, desired[0]);
        return;
    }
    const result = await executeCli(['extract', outputPath, '--indexes', scenario.indexes.join(',')]);
    assert.deepEqual(result.paragraphs.map(item => item.exactText), desired);
}

async function runPolicy(run, scenario, policy) {
    let applied;
    if (policy.id === 'v0.6.2-extract-full') {
        const records = await extractRecords(run, scenario, scenario.kind === 'restore' ? 'rejected' : 'accepted');
        applied = await applyOperations(run, scenario, records, false);
    } else if (policy.id === 'v0.7.1-speculation-led' && ['literal', 'nbsp', 'repeated'].includes(scenario.kind)) {
        const attempt = await speculative(run, scenario);
        if (attempt.result.completion === true) {
            const source = await executeCli(['extract', run.inputPath, '--indexes', scenario.indexes.join(',')]);
            applied = { outputPath: attempt.outputPath, desired: desiredTexts(scenario, source.paragraphs) };
        } else {
            assert.equal(scenario.kind, 'repeated');
            assert.equal(attempt.result.error.code, 'AMBIGUOUS_TARGET');
            const records = await extractRecords(run, scenario);
            applied = await applyOperations(run, scenario, records, true);
        }
    } else if (scenario.kind === 'restore') {
        const records = await extractRecords(run, scenario, 'rejected');
        applied = policy.id === 'v0.6.2-extract-full'
            ? await applyOperations(run, scenario, records, false)
            : await inlineRestore(run, records[0]);
    } else {
        const records = await extractRecords(run, scenario);
        applied = await applyOperations(run, scenario, records, policy.id !== 'v0.6.2-extract-full');
    }
    await verifyOutput(applied.outputPath, scenario, applied.desired);
    return run.metrics;
}

async function measure(directory, scenario, policy, count, phase) {
    const samples = [];
    for (let index = 0; index < count; index += 1) {
        const inputPath = path.join(directory, `${scenario.id}-input.docx`);
        const run = tracker(inputPath, path.join(directory, `${scenario.id}-${policy.id}-${phase}-${index}`));
        samples.push(await runPolicy(run, scenario, policy));
    }
    return samples;
}

const directory = await mkdtemp(path.join(tmpdir(), 'docx-agent-surface-'));
let report;
try {
    for (const scenario of scenarios) {
        await writeFile(path.join(directory, `${scenario.id}-input.docx`), scenario.input);
    }
    const results = [];
    const aggregateSamples = new Map(policies.map(policy => [policy.id, []]));
    for (const scenario of scenarios) {
        for (const policy of policies) {
            if (warmups > 0) await measure(directory, scenario, policy, warmups, 'warmup');
            const samples = await measure(directory, scenario, policy, iterations, 'sample');
            aggregateSamples.get(policy.id).push(...samples);
            results.push({
                scenario: scenario.id,
                policy: policy.id,
                successfulRuns: samples.length,
                toolCalls: summaries(samples, 'toolCalls'),
                failedCalls: summaries(samples, 'failedCalls'),
                generatedEditRequestBytes: summaries(samples, 'generatedEditRequestBytes'),
                providerVisibleResponseBytes: summaries(samples, 'responseBytes'),
                localEngineAndIoMs: summaries(samples, 'elapsedMs')
            });
        }
    }
    report = {
        generatedAt: new Date().toISOString(),
        environment: { node: process.version, platform: process.platform, arch: process.arch },
        configuration: { iterations, warmups, currentEngine: true },
        policies,
        scenarios: scenarios.map(({ id, kind }) => ({ id, kind })),
        measurements: {
            correctness: 'Every output is re-extracted and compared with the predetermined accepted-view text.',
            toolCalls: 'Workflow calls only; benchmark-only output verification is excluded.',
            generatedEditRequestBytes: 'Serialized operation objects or speculative edit arguments; not provider tokens.',
            localEngineAndIoMs: 'In-process CLI reads, extraction, mutation, validation, and writes measured with performance.now().',
            modelAndNetwork: {
                measured: false,
                reason: 'A deterministic replay cannot measure model reasoning, provider token use, transport, or stochastic instruction following.'
            }
        },
        summaryByPolicy: policies.map(policy => aggregateSummary(policy, aggregateSamples.get(policy.id))),
        results,
        limitations: [
            'The three policies are deterministic workflow replays, not independent model runs.',
            'Use external repeated agent transcripts to validate reasoning-token and end-to-end latency effects.'
        ]
    };
} finally {
    await rm(directory, { recursive: true, force: true });
}

await mkdir(path.dirname(reportPath), { recursive: true });
await writeFile(reportPath, `${JSON.stringify(report, null, 2)}\n`, 'utf8');
console.log(JSON.stringify(report, null, 2));
