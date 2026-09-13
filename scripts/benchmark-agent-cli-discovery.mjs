import assert from 'node:assert/strict';
import { mkdir, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { performance } from 'node:perf_hooks';
import { executeCli } from '../node/cli.js';

const fixture = path.resolve('tests/fixtures/sample_doc_test.docx');
const output = path.resolve('tmp/benchmarks/agent-cli-discovery-latest.json');
const before = Object.freeze({
    source: 'Observed immediately before WP-01/WP-02 on 2026-09-12',
    globalHelpBytes: 171,
    applyHelpBytes: 171,
    helpPayloadsIdentical: true,
    unscopedInspectBytes: 72387,
    unscopedInspectParagraphs: 63,
    broadExtractBytes: 13968,
    broadExtractParagraphs: 31,
    receiptPayloadDuplicated: true,
    legacyContextCommands: 2
});

const bytes = value => Buffer.byteLength(JSON.stringify(value, null, 2), 'utf8') + 1;

async function measure(args) {
    const started = performance.now();
    const result = await executeCli(args);
    return {
        elapsedMs: Number((performance.now() - started).toFixed(2)),
        bytes: bytes(result),
        paragraphs: result.paragraphs?.length ?? null,
        selection: result.selection || null,
        result
    };
}

const globalHelp = await measure(['--help']);
const applyHelp = await measure(['apply', '--help']);
const unscopedInspect = await measure(['inspect', fixture, '--non-empty']);
const broadExtract = await measure(['extract', fixture, '--search', 'the']);
const contextualExtract = await measure(['extract', fixture, '--search', 'agreement', '--around', '2', '--limit', '1']);
const upperCase = await executeCli(['extract', fixture, '--search', 'AGREEMENT', '--limit', '5']);
const lowerCase = await executeCli(['extract', fixture, '--search', 'agreement', '--limit', '5']);

assert.notDeepEqual(globalHelp.result, applyHelp.result);
assert(applyHelp.result.examples.some(item => item.operation?.type === 'restore'));
assert(unscopedInspect.bytes <= 48 * 1024);
assert(broadExtract.bytes <= 48 * 1024);
assert.equal(unscopedInspect.selection.truncated, true);
assert.equal(contextualExtract.result.paragraphs.some(item => item.selectionRole === 'context'), true);
assert.deepEqual(upperCase.paragraphs.map(item => item.index), lowerCase.paragraphs.map(item => item.index));

const report = {
    date: '2026-09-12',
    fixture,
    before,
    after: {
        globalHelp: { bytes: globalHelp.bytes, elapsedMs: globalHelp.elapsedMs },
        applyHelp: { bytes: applyHelp.bytes, elapsedMs: applyHelp.elapsedMs },
        unscopedInspect: {
            bytes: unscopedInspect.bytes,
            elapsedMs: unscopedInspect.elapsedMs,
            paragraphs: unscopedInspect.paragraphs,
            selection: unscopedInspect.selection
        },
        broadExtract: {
            bytes: broadExtract.bytes,
            elapsedMs: broadExtract.elapsedMs,
            paragraphs: broadExtract.paragraphs,
            selection: broadExtract.selection
        },
        contextualExtract: {
            bytes: contextualExtract.bytes,
            elapsedMs: contextualExtract.elapsedMs,
            paragraphs: contextualExtract.paragraphs,
            selection: contextualExtract.selection,
            commands: 1
        },
        searchCaseInsensitive: true
    },
    boundaries: {
        softOutputBytes: 48 * 1024,
        exactParagraphsAreNeverTruncated: true,
        providerTokensEstimated: false,
        providerWallTimeEstimated: false
    }
};

await mkdir(path.dirname(output), { recursive: true });
await writeFile(output, `${JSON.stringify(report, null, 2)}\n`);
console.log(JSON.stringify(report, null, 2));
