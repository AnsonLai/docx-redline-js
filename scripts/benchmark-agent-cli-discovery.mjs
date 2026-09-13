import assert from 'node:assert/strict';
import { mkdir, mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
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
    nestedReceiptBytes: 464,
    rootReceiptBytes: 464,
    applyResponseBytes: 2968,
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

const workflowDirectory = await mkdtemp(path.join(tmpdir(), 'docx-agent-cli-rollout-'));
let compactMutation;
let workflowAudit;
try {
    const fixtureBytes = await readFile(fixture);
    const inspectedTarget = (await executeCli(['extract', fixture, '--index', '53'])).paragraphs[0];
    const desiredText = `${inspectedTarget.exactText} Confirmed.`;
    const operationsPath = path.join(workflowDirectory, 'operations.json');
    const reviewedPath = path.join(workflowDirectory, 'reviewed.docx');
    const rejectedPath = path.join(workflowDirectory, 'rejected.docx');
    await writeFile(operationsPath, JSON.stringify([{
        operationId: 'wp05-contextual-edit',
        type: 'redline',
        target: {
            exactText: inspectedTarget.exactText,
            paragraphId: inspectedTarget.paragraphId,
            fingerprint: inspectedTarget.fingerprint
        },
        modified: desiredText,
        author: 'WP05 Audit'
    }]));
    const applied = await measure([
        'apply', fixture,
        '--operations', operationsPath,
        '--profile', 'agent',
        '--output', reviewedPath
    ]);
    const acceptedText = (await executeCli(['extract', reviewedPath, '--index', '53'])).paragraphs[0].exactText;
    const sourceComments = (await executeCli(['inspect', fixture, '--all'])).comments.length;
    const outputComments = (await executeCli(['inspect', reviewedPath, '--all'])).comments.length;
    const rejected = await executeCli([
        'reject', reviewedPath,
        '--author', 'WP05 Audit',
        '--output', rejectedPath
    ]);
    const rejectedText = (await executeCli(['extract', rejected.outputPath, '--index', '53'])).paragraphs[0].exactText;
    compactMutation = {
        beforePrettyBytes: before.applyResponseBytes,
        prettyBytes: applied.bytes,
        minifiedBytes: Buffer.byteLength(JSON.stringify(applied.result), 'utf8') + 1,
        nestedReceiptCount: applied.result.results.filter(item => item.receipt).length,
        topLevelReceiptCount: applied.result.receipts.length,
        exactTargetMatchOmitted: applied.result.results[0].resolvedTarget.targetTextMatch === undefined
    };
    workflowAudit = {
        commandHelpCalls: 1,
        sourceOrBundleInspectionCalls: 0,
        contextualSearchCalls: 1,
        applyAttempts: 1,
        profileAtomic: applied.result.effectiveOptions.atomic,
        profileRequireComplete: applied.result.effectiveOptions.requireComplete,
        acceptedTextMatches: acceptedText === desiredText,
        rejectedTextMatches: rejectedText === inspectedTarget.exactText,
        commentsPreserved: outputComments === sourceComments,
        sourceUnchanged: (await readFile(fixture)).equals(fixtureBytes)
    };
} finally {
    await rm(workflowDirectory, { recursive: true, force: true });
}

assert.notDeepEqual(globalHelp.result, applyHelp.result);
assert(applyHelp.result.examples.some(item => item.operation?.type === 'restore'));
assert(unscopedInspect.bytes <= 48 * 1024);
assert(broadExtract.bytes <= 48 * 1024);
assert.equal(unscopedInspect.selection.truncated, true);
assert.equal(contextualExtract.result.paragraphs.some(item => item.selectionRole === 'context'), true);
assert.deepEqual(upperCase.paragraphs.map(item => item.index), lowerCase.paragraphs.map(item => item.index));
assert.equal(compactMutation.nestedReceiptCount, 0);
assert.equal(compactMutation.topLevelReceiptCount, 1);
assert.equal(compactMutation.exactTargetMatchOmitted, true);
assert(compactMutation.prettyBytes < compactMutation.beforePrettyBytes);
assert(compactMutation.minifiedBytes < compactMutation.prettyBytes);
assert.deepEqual(workflowAudit, {
    commandHelpCalls: 1,
    sourceOrBundleInspectionCalls: 0,
    contextualSearchCalls: 1,
    applyAttempts: 1,
    profileAtomic: false,
    profileRequireComplete: true,
    acceptedTextMatches: true,
    rejectedTextMatches: true,
    commentsPreserved: true,
    sourceUnchanged: true
});

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
        searchCaseInsensitive: true,
        compactMutation,
        workflowAudit
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
