import { mkdir, writeFile } from 'node:fs/promises';
import { performance } from 'node:perf_hooks';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { configureXmlProvider, parseOoxmlSafe } from '../adapters/xml-adapter.js';
import { configureLogger } from '../adapters/logger.js';
import { buildTargetReferenceSnapshot } from '../core/paragraph-targeting.js';
import {
    applyOperationToDocumentXml,
    applyOperationsToDocumentXml
} from '../services/standalone-operation-runner.js';

configureXmlProvider({ DOMParser, XMLSerializer });
configureLogger({ info() {}, warn() {}, error() {} });

const paragraphCount = Math.max(100, Number.parseInt(process.env.DOCX_BENCH_PARAGRAPHS || '1000', 10));
const operationCount = Math.max(1, Number.parseInt(process.env.DOCX_BENCH_OPERATIONS || '10', 10));
const iterations = Math.max(3, Number.parseInt(process.env.DOCX_BENCH_ITERATIONS || '7', 10));
const warmups = Math.max(1, Number.parseInt(process.env.DOCX_BENCH_WARMUPS || '2', 10));
const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const NS_W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';

const paragraphs = Array.from({ length: paragraphCount }, (_, index) => {
    const id = (index + 1).toString(16).toUpperCase().padStart(8, '0');
    return `<w:p w14:paraId="${id}"><w:r><w:t>Benchmark paragraph ${index + 1} with stable unique content.</w:t></w:r></w:p>`;
}).join('');
const source = `<w:document xmlns:w="${NS_W}" xmlns:w14="${NS_W14}"><w:body>${paragraphs}<w:sectPr/></w:body></w:document>`;
const step = Math.max(1, Math.floor(paragraphCount / (operationCount + 1)));
const operations = Array.from({ length: operationCount }, (_, index) => {
    const paragraphNumber = Math.min(paragraphCount, step * (index + 1));
    return {
        type: 'replace',
        target: {
            paragraphId: paragraphNumber.toString(16).toUpperCase().padStart(8, '0'),
            exactText: `Benchmark paragraph ${paragraphNumber} with stable unique content.`
        },
        modified: `Benchmark paragraph ${paragraphNumber} with verified updated content.`,
        author: 'Session Benchmark'
    };
});

function percentile(values, ratio) {
    const sorted = [...values].sort((a, b) => a - b);
    return sorted[Math.min(sorted.length - 1, Math.ceil(sorted.length * ratio) - 1)];
}

async function runBatch() {
    const instrumentation = { parses: 0, serializations: 0 };
    const heapBefore = process.memoryUsage().heapUsed;
    const started = performance.now();
    const result = await applyOperationsToDocumentXml(source, operations, 'Session Benchmark', null, {
        generateRedlines: false,
        _sessionInstrumentation: {
            onDocumentParse: () => { instrumentation.parses += 1; },
            onDocumentSerialize: () => { instrumentation.serializations += 1; }
        }
    });
    const elapsedMs = performance.now() - started;
    if (!result.hasChanges || result.results.some(item => item.status !== 'applied')) {
        throw new Error('Live-session benchmark batch did not apply every operation.');
    }
    return {
        elapsedMs,
        heapDeltaBytes: process.memoryUsage().heapUsed - heapBefore,
        outputBytes: Buffer.byteLength(result.documentXml),
        ...instrumentation
    };
}

async function runSequential() {
    const parsed = parseOoxmlSafe(source, 'application/xml');
    const context = { targetRefSnapshot: buildTargetReferenceSnapshot(parsed.doc) };
    const instrumentation = { parses: 0, serializations: 0 };
    const heapBefore = process.memoryUsage().heapUsed;
    const started = performance.now();
    let documentXml = source;
    for (const operation of operations) {
        const result = await applyOperationToDocumentXml(documentXml, operation, 'Session Benchmark', context, {
            generateRedlines: false,
            _sessionInstrumentation: {
                onDocumentParse: () => { instrumentation.parses += 1; },
                onDocumentSerialize: () => { instrumentation.serializations += 1; }
            }
        });
        if (!result.hasChanges) throw new Error('Sequential benchmark operation did not apply.');
        documentXml = result.documentXml;
    }
    return {
        elapsedMs: performance.now() - started,
        heapDeltaBytes: process.memoryUsage().heapUsed - heapBefore,
        outputBytes: Buffer.byteLength(documentXml),
        ...instrumentation
    };
}

for (let index = 0; index < warmups; index += 1) {
    await runBatch();
    await runSequential();
}

const batchSamples = [];
const sequentialSamples = [];
for (let index = 0; index < iterations; index += 1) {
    batchSamples.push(await runBatch());
    sequentialSamples.push(await runSequential());
}

function summarize(samples) {
    const timings = samples.map(sample => sample.elapsedMs);
    return {
        medianMs: Number(percentile(timings, 0.5).toFixed(2)),
        p95Ms: Number(percentile(timings, 0.95).toFixed(2)),
        medianHeapDeltaBytes: percentile(samples.map(sample => sample.heapDeltaBytes), 0.5),
        parseCount: samples[0].parses,
        serializeCount: samples[0].serializations,
        outputBytes: samples[0].outputBytes
    };
}

const batch = summarize(batchSamples);
const sequential = summarize(sequentialSamples);
const report = {
    generatedAt: new Date().toISOString(),
    environment: { node: process.version, platform: process.platform, arch: process.arch },
    fixture: { paragraphCount, operationCount, iterations, warmups },
    batch,
    sequential,
    medianSpeedup: Number((sequential.medianMs / batch.medianMs).toFixed(2)),
    note: 'Timing and heap figures are observational. Semantic correctness tests remain the release gate.'
};

await mkdir(new URL('../tmp/benchmarks/', import.meta.url), { recursive: true });
await writeFile(
    new URL('../tmp/benchmarks/operation-session-latest.json', import.meta.url),
    `${JSON.stringify(report, null, 2)}\n`,
    'utf8'
);
console.log(JSON.stringify(report, null, 2));
