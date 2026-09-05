import { execFileSync } from 'child_process';
import { existsSync, readFileSync } from 'fs';
import { join, resolve } from 'path';

const repoRoot = process.cwd();
const syntheticDir = join(repoRoot, 'tmp', 'dashboard-docx');
const corpusSourceDir = join(repoRoot, 'tmp', 'superdoc-corpus');
const corpusFixturesArgIndex = process.argv.indexOf('--corpus-fixtures-dir');
if (corpusFixturesArgIndex >= 0 && !process.argv[corpusFixturesArgIndex + 1]) {
    throw new Error('--corpus-fixtures-dir requires a path');
}
const suppliedCorpusFixturesDir = corpusFixturesArgIndex >= 0;
const corpusFixturesDir = suppliedCorpusFixturesDir
    ? resolve(repoRoot, process.argv[corpusFixturesArgIndex + 1])
    : join(repoRoot, 'tmp', 'superdoc-word-fixtures');
const manifest = JSON.parse(readFileSync(
    join(repoRoot, 'tests', 'corpus', 'superdoc-english-legal-administrative.json'),
    'utf8'
));
const run = (script, args = []) => execFileSync(process.execPath, [script, ...args], {
    cwd: repoRoot,
    stdio: 'inherit'
});

run('scripts/export-validation-fixtures.mjs', ['--output-dir', syntheticDir]);

const corpusReady = manifest.documents.every(item =>
    existsSync(join(corpusSourceDir, `${item.id}.docx`))
);
if (corpusReady && suppliedCorpusFixturesDir) {
    if (!existsSync(join(corpusFixturesDir, 'suite.json'))) {
        throw new Error(`Supplied corpus fixture directory has no suite.json: ${corpusFixturesDir}`);
    }
} else if (corpusReady) {
    run('scripts/prepare-superdoc-word-corpus.mjs');
} else {
    console.warn('Real-document corpus is not downloaded; embedding synthetic DOCX previews only.');
    console.warn('Run npm run test:corpus:word once to fetch and validate the pinned corpus.');
}

const lane1FixturesArgIndex = process.argv.indexOf('--lane1-fixtures-dir');
if (lane1FixturesArgIndex >= 0 && !process.argv[lane1FixturesArgIndex + 1]) {
    throw new Error('--lane1-fixtures-dir requires a path');
}
const suppliedLane1FixturesDir = lane1FixturesArgIndex >= 0;
const lane1FixturesDir = suppliedLane1FixturesDir
    ? resolve(repoRoot, process.argv[lane1FixturesArgIndex + 1])
    : join(repoRoot, 'tmp', 'lane1-docx');

if (!existsSync(join(lane1FixturesDir, 'manifest.json'))) {
    run('scripts/export-lane1-fixtures.mjs', ['--output-dir', lane1FixturesDir]);
}

const args = ['scripts/generate-test-dashboard.mjs', '--fixtures-dir', syntheticDir];
if (corpusReady) args.push('--corpus-fixtures-dir', corpusFixturesDir);
if (existsSync(join(lane1FixturesDir, 'manifest.json'))) {
    args.push('--lane1-fixtures-dir', lane1FixturesDir);
}
run(args[0], args.slice(1));
