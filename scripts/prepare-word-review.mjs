import { execFileSync } from 'child_process';
import { mkdirSync, writeFileSync } from 'fs';
import { dirname, join } from 'path';
import { fileURLToPath } from 'url';

import { loadCoverageCatalogue } from './lib/word-coverage-catalogue.mjs';

const repoRoot = dirname(dirname(fileURLToPath(import.meta.url)));
const cycleArg = process.argv.find(argument => argument.startsWith('--cycle='));
const cycle = Number(cycleArg?.split('=')[1] ?? 0);
if (!Number.isInteger(cycle) || cycle < 0) throw new Error('--cycle must be a non-negative integer');

function changedFiles() {
    try {
        return execFileSync('git', [
            'status', '--short', '--untracked-files=all', '--',
            'tests/fixtures/word-task-cases.mjs',
            'tests/fixtures/word-task-coverage.mjs',
            'tests/corpus/superdoc-word-scenarios.json',
            'tests/corpus/superdoc-word-coverage.json'
        ], { cwd: repoRoot, encoding: 'utf8' });
    } catch {
        return '';
    }
}

function rotate(items, count, offset) {
    if (items.length === 0) return [];
    return Array.from({ length: Math.min(count, items.length) }, (_, index) => items[(offset + index) % items.length]);
}

const { cases } = loadCoverageCatalogue();
const synthetic = cases.filter(item => item.lane === 'synthetic');
const corpus = cases.filter(item => item.lane === 'superdoc');
const changes = changedFiles();
const selected = new Map();
const add = (item, reason) => {
    const existing = selected.get(item.identity);
    selected.set(item.identity, { ...item, reasons: [...new Set([...(existing?.reasons || []), reason])] });
};

if (/word-task-(?:cases|coverage)\.mjs/.test(changes)) {
    for (const item of synthetic) add(item, 'Synthetic catalogue metadata or cases changed');
}
if (/superdoc-word-(?:scenarios|coverage)\.json/.test(changes)) {
    for (const item of corpus) add(item, 'SuperDoc scenario catalogue changed');
}

const sampleSize = Math.ceil(synthetic.length * 0.2);
for (const item of rotate(synthetic, sampleSize, (cycle * sampleSize) % synthetic.length)) {
    add(item, `Rotating 20% synthetic sample, cycle ${cycle}`);
}
for (const category of ['legal', 'administrative']) {
    const categoryCases = corpus.filter(item => item.category === category);
    add(categoryCases[cycle % categoryCases.length], `Required ${category} SuperDoc sample`);
}

const manifest = {
    schemaVersion: 1,
    cycle,
    humanSignOff: { status: 'pending', reviewer: null, wordBuild: null },
    instructions: 'Inspect All Markup, Accept All, and Reject All. This manifest does not record a pass.',
    cases: [...selected.values()].sort((a, b) => a.identity.localeCompare(b.identity)).map(item => ({
        identity: item.identity,
        lane: item.lane,
        category: item.category,
        task: item.metadata.task,
        structures: item.metadata.structures,
        reasons: item.reasons,
        views: { allMarkup: 'pending', acceptAll: 'pending', rejectAll: 'pending' }
    }))
};
const outputDir = join(repoRoot, 'tmp', 'word-manual-review');
mkdirSync(outputDir, { recursive: true });
const outputPath = join(outputDir, `review-manifest-cycle-${cycle}.json`);
writeFileSync(outputPath, `${JSON.stringify(manifest, null, 2)}\n`);
console.log(`Prepared ${manifest.cases.length} cases for human review: ${outputPath}`);
