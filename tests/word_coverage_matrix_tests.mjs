import assert from 'assert/strict';
import { execFileSync } from 'child_process';
import { readFileSync } from 'fs';

import {
    COVERAGE_STRUCTURES,
    COVERAGE_TASKS,
    loadCoverageCatalogue,
    validateCoveragePriorities,
    validateScenarioIdentities
} from '../scripts/lib/word-coverage-catalogue.mjs';
import { validateCoverageMetadata } from '../scripts/lib/word-coverage-metadata.mjs';

const { cases, priorities } = loadCoverageCatalogue();
assert.equal(cases.length, 107);
assert.equal(cases.filter(item => item.lane === 'synthetic').length, 47);
assert.equal(cases.filter(item => item.lane === 'superdoc').length, 60);

const matrix = validateCoveragePriorities(cases, priorities);
assert.equal(matrix.size, COVERAGE_TASKS.length * COVERAGE_STRUCTURES.length);
assert.ok(matrix.get('replace/table').length > 0);
assert.ok(matrix.get('replace/comment').length > 0);
assert.ok(matrix.get('accept-reject/prior-revisions').length > 0);

assert.throws(
    () => validateCoverageMetadata({
        task: 'unknown-task',
        structures: ['plain-paragraph'],
        oracles: ['synthetic-word'],
        manualReview: { status: 'missing' }
    }, 'bad-case'),
    /unknown coverage task/
);
assert.throws(
    () => validateCoverageMetadata({
        task: 'replace',
        structures: ['mystery-structure'],
        oracles: ['synthetic-word'],
        manualReview: { status: 'missing' }
    }, 'bad-case'),
    /unknown structures label/
);
assert.throws(
    () => validateCoverageMetadata({
        task: 'replace',
        structures: ['plain-paragraph'],
        oracles: [],
        manualReview: { status: 'missing' }
    }, 'bad-case'),
    /oracles must be a non-empty array/
);
assert.throws(
    () => validateScenarioIdentities([{ identity: 'same' }, { identity: 'same' }]),
    /Duplicate or missing scenario identity/
);
assert.throws(
    () => validateCoveragePriorities(cases, {
        highPriorityCells: [{ task: 'comment', structure: 'plain-paragraph' }],
        emptyCellDispositions: []
    }),
    /lacks a disposition/
);

const repoRoot = new URL('../', import.meta.url);
const firstReport = execFileSync(process.execPath, ['scripts/report-word-coverage.mjs', '--json'], {
    cwd: repoRoot,
    encoding: 'utf8'
});
const secondReport = execFileSync(process.execPath, ['scripts/report-word-coverage.mjs', '--json'], {
    cwd: repoRoot,
    encoding: 'utf8'
});
assert.equal(secondReport, firstReport, 'coverage report must be deterministic');
const report = JSON.parse(firstReport);
assert.equal(report.totals.cases, cases.length);
assert.equal(report.totals.manualMissing, cases.length);
assert.equal(report.totals.aiPreflight, 22);
assert.equal(report.totals.syntheticWord, 47);
assert.equal(report.totals.realDocumentWord, 60);

execFileSync(process.execPath, ['scripts/prepare-word-review.mjs', '--cycle=2'], {
    cwd: repoRoot,
    encoding: 'utf8'
});
const review = JSON.parse(readFileSync(new URL('../tmp/word-manual-review/review-manifest-cycle-2.json', import.meta.url)));
assert.equal(review.humanSignOff.status, 'pending');
assert.ok(review.cases.length >= Math.ceil(47 * 0.2) + 2);
assert.ok(review.cases.every(item => Object.values(item.views).every(status => status === 'pending')));
assert.ok(review.cases.some(item => item.lane === 'superdoc' && item.category === 'legal'));
assert.ok(review.cases.some(item => item.lane === 'superdoc' && item.category === 'administrative'));

console.log(`PASS: explicit Word coverage matrix (${cases.length} normalized cases)`);
