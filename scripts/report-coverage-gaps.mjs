import { existsSync, readFileSync } from 'fs';
import { relative, resolve } from 'path';
import { fileURLToPath } from 'url';

const repoRoot = resolve(fileURLToPath(new URL('..', import.meta.url)));
const coveragePath = resolve(repoRoot, 'coverage', 'coverage-final.json');
const baselinePath = resolve(repoRoot, 'tests', 'coverage-data', 'phase3-baseline.json');
const reviewedGapsPath = resolve(repoRoot, 'tests', 'coverage-data', 'phase3-reviewed-gaps.json');
const productionRoots = ['index.js', 'adapters/', 'core/', 'engine/', 'pipeline/', 'services/', 'orchestration/'];
const priorities = new Map([
    ['services/numbering-helpers.js', 'P0'],
    ['orchestration/route-plan.js', 'P0'],
    ['orchestration/list-markdown.js', 'P0'],
    ['pipeline/patching.js', 'P0'],
    ['engine/format-span-application.js', 'P0'],
    ['orchestration/list-structural-fallback.js', 'P1'],
    ['engine/table-mode.js', 'P1'],
    ['core/table-targeting.js', 'P1'],
    ['services/standalone-operation-runner.js', 'P1'],
    ['pipeline/pipeline.js', 'P1']
]);

if (!existsSync(coveragePath)) {
    throw new Error('coverage/coverage-final.json is missing; run npm run test:coverage first');
}

const raw = JSON.parse(readFileSync(coveragePath, 'utf8'));
const files = [];
for (const [absolutePath, coverage] of Object.entries(raw)) {
    const file = relative(repoRoot, absolutePath).replaceAll('\\', '/');
    if (!productionRoots.some(root => root.endsWith('/') ? file.startsWith(root) : file === root)) continue;
    const uncoveredFunctions = Object.entries(coverage.fnMap || {})
        .filter(([id]) => (coverage.f?.[id] || 0) === 0)
        .map(([, fn]) => ({ name: fn.name || '(anonymous)', line: fn.decl?.start?.line || fn.loc?.start?.line }));
    const functionHits = Object.values(coverage.f || {});
    const branchHits = Object.values(coverage.b || {}).flat();
    files.push({
        file,
        priority: priorities.get(file) || 'P2',
        functions: { covered: functionHits.filter(hits => hits > 0).length, total: functionHits.length },
        branches: { covered: branchHits.filter(hits => hits > 0).length, total: branchHits.length },
        uncoveredFunctions
    });
}
files.sort((a, b) => a.priority.localeCompare(b.priority) || a.file.localeCompare(b.file));

const baseline = existsSync(baselinePath) ? JSON.parse(readFileSync(baselinePath, 'utf8')) : null;
const reviewedGaps = existsSync(reviewedGapsPath) ? JSON.parse(readFileSync(reviewedGapsPath, 'utf8')) : { items: [] };
const regressions = [];
if (baseline) {
    const currentByFile = new Map(files.map(item => [item.file, item]));
    for (const expected of baseline.targetFiles) {
        const current = currentByFile.get(expected.file);
        if (!current || current.functions.covered < expected.functionsCovered
            || current.branches.covered < expected.branchesCovered) {
            regressions.push({ expected, current: current || null });
        }
    }
}

const actualPriorityGaps = new Set(files
    .filter(item => item.priority === 'P0' || item.priority === 'P1')
    .flatMap(item => item.uncoveredFunctions.map(fn => `${item.file}:${fn.line}`)));
const classifiedPriorityGaps = new Set((reviewedGaps.items || [])
    .flatMap(item => (item.lines || []).map(line => `${item.file}:${line}`)));
const classificationErrors = [
    ...[...actualPriorityGaps]
        .filter(key => !classifiedPriorityGaps.has(key))
        .map(key => `Unclassified priority gap: ${key}`),
    ...[...classifiedPriorityGaps]
        .filter(key => !actualPriorityGaps.has(key))
        .map(key => `Stale priority-gap classification: ${key}`)
];

const report = {
    schemaVersion: 1,
    production: {
        functionsCovered: files.reduce((sum, item) => sum + item.functions.covered, 0),
        functionsTotal: files.reduce((sum, item) => sum + item.functions.total, 0),
        branchesCovered: files.reduce((sum, item) => sum + item.branches.covered, 0),
        branchesTotal: files.reduce((sum, item) => sum + item.branches.total, 0)
    },
    files,
    regressions,
    classificationErrors
};

if (process.argv.includes('--json')) {
    console.log(JSON.stringify(report, null, 2));
} else {
    console.log(`Production functions: ${report.production.functionsCovered}/${report.production.functionsTotal}`);
    console.log(`Production branches: ${report.production.branchesCovered}/${report.production.branchesTotal}`);
    for (const item of files.filter(file => file.uncoveredFunctions.length > 0)) {
        console.log(`\n${item.priority} ${item.file} (${item.functions.covered}/${item.functions.total} functions, ${item.branches.covered}/${item.branches.total} branches)`);
        for (const fn of item.uncoveredFunctions) console.log(`  ${fn.line}: ${fn.name}`);
    }
}

if (regressions.length > 0 || classificationErrors.length > 0) {
    console.error(`Coverage regression in ${regressions.length} targeted file(s)`);
    for (const issue of classificationErrors) console.error(issue);
    process.exitCode = 1;
}
