import assert from 'node:assert/strict';
import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { performance } from 'node:perf_hooks';
import { configureLogger } from '../adapters/logger.js';
import { createExampleAgentSession } from '../examples/agent-session-wrapper.mjs';
import { openDocx } from '../node/index.js';
import { AGENT_PERFORMANCE_CASES } from './lib/agent-performance-cases.mjs';

configureLogger({ info() {}, warn() {}, error() {} });

const fixtureUrl = new URL('../tests/fixtures/sample_doc_test.docx', import.meta.url);
const crossAuthorFixtureUrl = new URL('../tests/fixtures/cross-author-slicing/insert-interior-pending.docx', import.meta.url);
const fixture = await readFile(fixtureUrl);
const iterations = Math.max(3, Number.parseInt(process.env.DOCX_AGENT_BENCH_ITERATIONS || '7', 10));
const warmups = Math.max(1, Number.parseInt(process.env.DOCX_AGENT_BENCH_WARMUPS || '2', 10));
const author = 'Agent Performance Benchmark';

function percentile(values, ratio) {
    const sorted = [...values].sort((left, right) => left - right);
    return sorted[Math.min(sorted.length - 1, Math.ceil(sorted.length * ratio) - 1)];
}

function summarize(samples) {
    const timings = samples.map(sample => sample.elapsedMs);
    return {
        medianMs: Number(percentile(timings, 0.5).toFixed(2)),
        p95Ms: Number(percentile(timings, 0.95).toFixed(2)),
        medianHeapDeltaBytes: percentile(samples.map(sample => sample.heapDeltaBytes), 0.5),
        applyRequestBytes: samples[0].applyRequestBytes,
        protocolCalls: samples[0].protocolCalls,
        outputBytes: samples[0].outputBytes
    };
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

function canonicalOperation(edit) {
    const common = {
        operationId: edit.operationId,
        target: canonicalTarget(edit.paragraph)
    };
    if (typeof edit.desiredText === 'string') {
        return { ...common, type: 'redline', modified: edit.desiredText };
    }
    return { ...common, type: 'comment', commentContent: edit.commentContent };
}

function handleEdit(edit, handle) {
    return {
        operationId: edit.operationId,
        target: handle,
        ...(typeof edit.desiredText === 'string'
            ? { desiredText: edit.desiredText }
            : { commentContent: edit.commentContent })
    };
}

function resultBuffer(result) {
    assert.equal(result.status, 'ok');
    assert.equal((result.results || []).some(item => item.status === 'error'), false);
    return Buffer.from(result.toBuffer());
}

function verifyLifecycle(output, edits) {
    const document = openDocx(output);
    const acceptedTexts = new Set(document.inspect().paragraphs.map(paragraph => paragraph.exactText));
    const rejectedTexts = new Set(document.inspect({ revisionView: 'rejected' }).paragraphs.map(paragraph => paragraph.exactText));
    for (const edit of edits) {
        if (typeof edit.desiredText !== 'string') continue;
        assert(acceptedTexts.has(edit.desiredText), `${edit.operationId} accepted text did not match.`);
        assert(rejectedTexts.has(edit.paragraph.exactText), `${edit.operationId} rejected text did not restore source.`);
    }
}

async function runCanonical(task) {
    const heapBefore = process.memoryUsage().heapUsed;
    const started = performance.now();
    const inspectionDocument = openDocx(fixture);
    const inspection = inspectionDocument.inspect({ indexes: task.indexes });
    const edits = task.buildEdits(inspection.paragraphs);
    const operations = edits.map(canonicalOperation);
    const expectedRevision = inspectionDocument.getRevisionToken();
    const mutationDocument = openDocx(fixture);
    const result = await mutationDocument.applyOperations(operations, {
        author,
        atomic: true,
        continueOnError: true,
        strictTargets: true,
        validate: true,
        generateRedlines: true,
        existingRevisions: 'merge-same-author',
        expectedRevision
    });
    const elapsedMs = performance.now() - started;
    const output = resultBuffer(result);
    verifyLifecycle(output, edits);
    return {
        elapsedMs,
        heapDeltaBytes: process.memoryUsage().heapUsed - heapBefore,
        applyRequestBytes: Buffer.byteLength(JSON.stringify({ operations, expectedRevision })),
        protocolCalls: 2,
        outputBytes: output.length
    };
}

async function runExampleSession(task) {
    const heapBefore = process.memoryUsage().heapUsed;
    const started = performance.now();
    const session = createExampleAgentSession(fixture, {
        profile: { author }
    });
    const inspection = session.inspect({ indexes: task.indexes });
    assert.equal(inspection.ok, true);
    const edits = task.buildEdits(inspection.targets.map(target => ({ exactText: target.exactText })));
    const handlesByText = new Map(inspection.targets.map(target => [target.exactText, target.handle]));
    const requestEdits = edits.map(edit => handleEdit(edit, handlesByText.get(edit.paragraph.exactText)));
    const result = await session.applyEdits(requestEdits);
    const elapsedMs = performance.now() - started;
    assert.equal(result.ok, true);
    const output = Buffer.from(result.outputBytes);
    verifyLifecycle(output, edits);
    return {
        elapsedMs,
        heapDeltaBytes: process.memoryUsage().heapUsed - heapBefore,
        applyRequestBytes: Buffer.byteLength(JSON.stringify({ edits: requestEdits })),
        protocolCalls: 2,
        outputBytes: output.length
    };
}

async function measureTask(task) {
    for (let index = 0; index < warmups; index += 1) {
        await runCanonical(task);
        await runExampleSession(task);
    }
    const canonicalSamples = [];
    const sessionSamples = [];
    for (let index = 0; index < iterations; index += 1) {
        canonicalSamples.push(await runCanonical(task));
        sessionSamples.push(await runExampleSession(task));
    }
    const canonical = summarize(canonicalSamples);
    const exampleSession = summarize(sessionSamples);
    return {
        id: task.id,
        description: task.description,
        canonical,
        exampleSession,
        applyRequestByteReductionPercent: Number(
            ((1 - exampleSession.applyRequestBytes / canonical.applyRequestBytes) * 100).toFixed(2)
        )
    };
}

async function orderDependencyDiagnostic() {
    const inspected = openDocx(fixture).inspect();
    const splitTarget = inspected.paragraphs.find(paragraph => paragraph.index === 2);
    const laterTargets = [11, 12, 13, 14, 18, 19, 20, 21, 22]
        .map(index => inspected.paragraphs.find(paragraph => paragraph.index === index));
    const splitPoint = splitTarget.exactText.indexOf(' by and between:');
    assert(splitPoint > 0);
    const split = {
        type: 'redline',
        target: canonicalTarget(splitTarget),
        modified: `${splitTarget.exactText.slice(0, splitPoint)}.\nThe parties are identified below:`,
        generateRedlines: false
    };
    const later = laterTargets.map(paragraph => ({
        type: 'redline',
        target: canonicalTarget(paragraph),
        modified: `${paragraph.exactText} Reviewed.`,
        generateRedlines: false
    }));
    const apply = operations => openDocx(fixture).applyOperations(operations, {
        author,
        atomic: true,
        continueOnError: true,
        strictTargets: true,
        validate: true,
        generateRedlines: false
    });
    const splitFirst = await apply([split, ...later]);
    const splitLast = await apply([...later, split]);
    return {
        case: 'ten-independent-edits-after-early-split',
        splitFirst: {
            status: splitFirst.status,
            written: splitFirst.written,
            rolledBack: splitFirst.rolledBack === true,
            errorCodes: (splitFirst.results || []).map(item => item.error?.code).filter(Boolean)
        },
        splitLast: {
            status: splitLast.status,
            written: splitLast.written,
            rolledBack: splitLast.rolledBack === true,
            errorCodes: (splitLast.results || []).map(item => item.error?.code).filter(Boolean)
        },
        expectedFutureBehavior: 'Both permutations succeed after WP-04 source-target binding.'
    };
}

async function staleHandleDiagnostic() {
    const session = createExampleAgentSession(fixture, { profile: { author } });
    const target = session.inspect({ indexes: [53] }).targets[0];
    const applied = await session.applyEdits([{ target: target.handle, desiredText: `${target.exactText} Additional copy.` }]);
    assert.equal(applied.ok, true);
    const stale = await session.applyEdits([{ target: target.handle, desiredText: target.exactText }]);
    return {
        case: 'stale-session-handle',
        status: stale.status,
        code: stale.error?.code,
        recoveryAction: stale.error?.recovery?.action
    };
}

async function ambiguousTargetDiagnostic() {
    const result = await openDocx(fixture).applyOperations([
        { type: 'redline', target: 'By: [Name]', modified: 'By: [Authorized Signatory]' }
    ], { author, atomic: true, strictTargets: true, validate: true });
    return {
        case: 'ambiguous-target',
        status: result.status,
        code: result.results?.[0]?.error?.code,
        candidateCount: result.results?.[0]?.error?.candidates?.length || 0
    };
}

async function foreignRevisionDiagnostic() {
    const input = await readFile(crossAuthorFixtureUrl);
    const document = openDocx(input);
    const paragraph = document.inspect({ revisedOnly: true }).paragraphs[0];
    const result = await document.applyOperations([{
        type: 'redline',
        target: canonicalTarget(paragraph),
        modified: `${paragraph.exactText} Confirmed.`
    }], {
        author,
        atomic: true,
        strictTargets: true,
        validate: true,
        existingRevisions: 'merge-same-author'
    });
    return {
        case: 'foreign-revision-policy',
        status: result.status,
        code: result.results?.[0]?.error?.code,
        written: result.written
    };
}

const taskResults = [];
for (const task of AGENT_PERFORMANCE_CASES) taskResults.push(await measureTask(task));

const report = {
    generatedAt: new Date().toISOString(),
    environment: {
        node: process.version,
        platform: process.platform,
        arch: process.arch
    },
    scope: {
        fixture: 'tests/fixtures/sample_doc_test.docx',
        iterations,
        warmups,
        measured: [
            'native inspect-and-apply wall time',
            'heap delta',
            'serialized apply request bytes',
            'accepted/rejected text fidelity'
        ],
        notMeasured: [
            'LLM reasoning time',
            'provider tokenization',
            'tool transport latency',
            'Claude or OpenCode end-to-end wall time'
        ]
    },
    taskResults,
    diagnostics: [
        await orderDependencyDiagnostic(),
        await staleHandleDiagnostic(),
        await ambiguousTargetDiagnostic(),
        await foreignRevisionDiagnostic()
    ],
    note: 'This benchmark is observational. Timing and heap values are not correctness gates; automated fidelity tests remain authoritative.'
};

await mkdir(new URL('../tmp/benchmarks/', import.meta.url), { recursive: true });
await writeFile(
    new URL('../tmp/benchmarks/agent-workflow-latest.json', import.meta.url),
    `${JSON.stringify(report, null, 2)}\n`,
    'utf8'
);
console.log(JSON.stringify(report, null, 2));
