import { performance } from 'node:perf_hooks';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { configureXmlProvider } from '../adapters/xml-adapter.js';
import {
    buildParagraphMetadataIndex,
    buildTargetReferenceSnapshot,
    resolveTargetParagraph
} from '../core/paragraph-targeting.js';

configureXmlProvider({ DOMParser, XMLSerializer });

const paragraphCount = Number.parseInt(process.env.DOCX_TARGET_BENCH_PARAGRAPHS || '10000', 10);
const operationCount = Number.parseInt(process.env.DOCX_TARGET_BENCH_OPERATIONS || '100', 10);
const measuredIterations = Number.parseInt(process.env.DOCX_TARGET_BENCH_ITERATIONS || '7', 10);
const warmups = 2;
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';

const body = Array.from({ length: paragraphCount }, (_, index) =>
    `<w:p w14:paraId="${String(index + 1).padStart(8, '0')}"><w:r><w:t>Benchmark paragraph ${index + 1}</w:t></w:r></w:p>`
).join('');
const xml = `<w:document xmlns:w="${W}" xmlns:w14="${W14}"><w:body>${body}<w:sectPr/></w:body></w:document>`;
const targets = Array.from({ length: operationCount }, (_, index) => {
    const paragraphIndex = 1 + Math.floor(index * (paragraphCount - 1) / Math.max(1, operationCount - 1));
    return { text: `Benchmark paragraph ${paragraphIndex}`, index: paragraphIndex };
});

function percentile(values, ratio) {
    const sorted = values.slice().sort((a, b) => a - b);
    return sorted[Math.min(sorted.length - 1, Math.floor(sorted.length * ratio))];
}

function measure(useCache) {
    const doc = new DOMParser().parseFromString(xml, 'application/xml');
    const started = performance.now();
    const paragraphMetadataIndex = useCache ? buildParagraphMetadataIndex(doc) : null;
    buildTargetReferenceSnapshot(doc, paragraphMetadataIndex);
    for (const targetDescriptor of targets) {
        resolveTargetParagraph(doc, {
            targetDescriptor,
            strictAmbiguity: true,
            paragraphMetadataIndex
        });
    }
    return performance.now() - started;
}

for (let index = 0; index < warmups; index++) {
    measure(false);
    measure(true);
}

const uncached = [];
const cached = [];
for (let index = 0; index < measuredIterations; index++) {
    uncached.push(measure(false));
    cached.push(measure(true));
}

const result = {
    runtime: { node: process.version, platform: process.platform, arch: process.arch },
    fixture: { paragraphCount, operationCount, warmups, measuredIterations },
    uncachedMs: { median: percentile(uncached, 0.5), p95: percentile(uncached, 0.95) },
    sessionCachedMs: { median: percentile(cached, 0.5), p95: percentile(cached, 0.95) }
};
result.medianSpeedup = result.uncachedMs.median / result.sessionCachedMs.median;
process.stdout.write(`${JSON.stringify(result, null, 2)}\n`);
