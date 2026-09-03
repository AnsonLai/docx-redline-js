import assert from 'assert/strict';
import { execFileSync } from 'child_process';
import { readFileSync } from 'fs';
import { fileURLToPath } from 'url';

import { WORD_TASK_CASES } from './fixtures/word-task-cases.mjs';
import {
    buildVisualReviewManifest,
    selectVisualReviewCases,
    VISUAL_STRUCTURES
} from '../scripts/prepare-word-visual-review.mjs';
import {
    buildCorpusVisualReviewManifest,
    selectCorpusVisualReviewCases
} from '../scripts/prepare-corpus-word-visual-review.mjs';

const selected = selectVisualReviewCases();
assert.ok(selected.length >= 15, 'visual lane should cover a broad layout-sensitive sample');
assert.deepEqual(selected.map(item => item.name), [...selected.map(item => item.name)].sort());

for (const structure of VISUAL_STRUCTURES) {
    const catalogueHasStructure = WORD_TASK_CASES.some(testCase =>
        testCase.coverageMetadata.structures.includes(structure)
    );
    if (!catalogueHasStructure) continue;
    assert.ok(
        selected.some(testCase => testCase.coverageMetadata.structures.includes(structure)),
        `visual lane has no fixture for ${structure}`
    );
}

const historicallyReviewed = WORD_TASK_CASES.filter(testCase =>
    testCase.coverageMetadata.oracles.includes('ai-word-visual-preflight')
);
assert.ok(historicallyReviewed.length >= 22);
assert.ok(historicallyReviewed.every(testCase => selected.includes(testCase)));

const manifest = buildVisualReviewManifest();
assert.equal(manifest.cases.length, selected.length);
assert.equal(manifest.certification.status, 'pending');
assert.equal(manifest.word.version, null);
for (const entry of manifest.cases) {
    assert.equal(entry.renderStatus, 'pending');
    assert.deepEqual(Object.keys(entry.views), ['allMarkup', 'acceptAll', 'rejectAll']);
    for (const view of Object.values(entry.views)) {
        assert.equal(view.status, 'pending');
        assert.match(view.pdf, /^[a-z0-9-]+-(?:allMarkup|acceptAll|rejectAll)\.pdf$/);
        assert.equal(view.pages, null);
        assert.equal(view.bytes, null);
    }
}

assert.throws(() => selectVisualReviewCases(WORD_TASK_CASES, ['not-a-real-case']), /Unknown visual-review case/);
assert.deepEqual(
    selectVisualReviewCases(WORD_TASK_CASES, ['legal-locked-field-adjacent-replacement']).map(item => item.name),
    ['legal-locked-field-adjacent-replacement']
);

const repoRoot = new URL('../', import.meta.url);
const outputPath = new URL('../tmp/word-visual-review/test-manifest.json', import.meta.url);
execFileSync(process.execPath, [
    'scripts/prepare-word-visual-review.mjs',
    `--output=${fileURLToPath(outputPath)}`,
    '--case=administrative-tab-aligned-status'
], { cwd: repoRoot, encoding: 'utf8' });
const written = JSON.parse(readFileSync(outputPath, 'utf8'));
assert.equal(written.cases.length, 1);
assert.equal(written.cases[0].name, 'administrative-tab-aligned-status');
assert.equal(written.certification.status, 'pending');

const corpusSelected = selectCorpusVisualReviewCases();
assert.equal(corpusSelected.length, 40);
assert.equal(corpusSelected.filter(item => item.shape === 'list').length, 22);
assert.equal(corpusSelected.filter(item => item.shape === 'table-form').length, 16);
assert.equal(corpusSelected.filter(item => item.shape === 'page-header').length, 2);
assert.equal(corpusSelected.filter(item => item.operationCount === 1).length, 31);
assert.equal(corpusSelected.filter(item => item.operationCount === 3).length, 4);
assert.equal(corpusSelected.filter(item => item.operationCount === 8).length, 5);
const corpusManifest = buildCorpusVisualReviewManifest();
assert.equal(corpusManifest.cases.length, 40);
assert.ok(corpusManifest.cases.every(item => item.identity.startsWith('superdoc:')));
assert.ok(corpusManifest.cases.every(item => item.renderStatus === 'pending'));
assert.throws(() => selectCorpusVisualReviewCases(['not-a-real-case']), /Unknown corpus visual-review case/);
const corpusVisualRunner = readFileSync(
    new URL('../scripts/word-com-corpus-visual-suite.ps1', import.meta.url),
    'utf8'
);
assert.match(corpusVisualRunner, /RPC_E_DISCONNECTED/);
assert.match(corpusVisualRunner, /restarting and retrying once/);
assert.match(corpusVisualRunner, /fixtures-\$PID/);
assert.match(corpusVisualRunner, /build-test-dashboard\.mjs.*--corpus-fixtures-dir\s+\$fixturesPath/s);

console.log(`PASS: Word visual-review lanes (${selected.length} synthetic and ${corpusSelected.length} real cases)`);
