import {
    COVERAGE_STRUCTURES,
    COVERAGE_TASKS,
    loadCoverageCatalogue,
    validateCoveragePriorities
} from './lib/word-coverage-catalogue.mjs';

const { cases, priorities } = loadCoverageCatalogue();
const cells = validateCoveragePriorities(cases, priorities);
const dispositionByCell = new Map(
    priorities.emptyCellDispositions.map(item => [`${item.task}/${item.structure}`, item])
);
const highPriority = new Set(
    priorities.highPriorityCells.map(item => `${item.task}/${item.structure}`)
);

const report = {
    schemaVersion: 1,
    totals: {
        cases: cases.length,
        synthetic: cases.filter(item => item.lane === 'synthetic').length,
        superdoc: cases.filter(item => item.lane === 'superdoc').length,
        manualMissing: cases.filter(item => item.metadata.manualReview.status === 'missing').length,
        manualStale: cases.filter(item => item.metadata.manualReview.status === 'stale').length,
        aiPreflight: cases.filter(item => item.metadata.oracles.includes('ai-word-visual-preflight')).length,
        syntheticWord: cases.filter(item => item.metadata.oracles.includes('synthetic-word')).length,
        realDocumentWord: cases.filter(item => item.metadata.oracles.includes('real-document-word')).length
    },
    tasks: COVERAGE_TASKS.map(task => ({
        task,
        structures: COVERAGE_STRUCTURES.map(structure => {
            const key = `${task}/${structure}`;
            return {
                structure,
                count: cells.get(key).length,
                highPriority: highPriority.has(key),
                disposition: dispositionByCell.get(key) || null,
                cases: cells.get(key)
            };
        })
    })),
    missingVisualReview: cases
        .filter(item => item.metadata.manualReview.status !== 'current')
        .map(item => ({
            identity: item.identity,
            status: item.metadata.manualReview.status,
            aiPreflight: item.metadata.oracles.includes('ai-word-visual-preflight')
        }))
};

if (process.argv.includes('--json')) {
    console.log(JSON.stringify(report, null, 2));
} else {
    console.log('# Word task/structure coverage matrix\n');
    console.log(`Cases: ${report.totals.cases} (${report.totals.synthetic} synthetic, ${report.totals.superdoc} SuperDoc)`);
    console.log(`Human review: ${report.totals.manualMissing} missing, ${report.totals.manualStale} stale\n`);
    console.log(`Automated Word: ${report.totals.syntheticWord} synthetic, ${report.totals.realDocumentWord} real-document; AI visual preflight: ${report.totals.aiPreflight}\n`);
    console.log(`| Task | ${COVERAGE_STRUCTURES.join(' | ')} |`);
    console.log(`|---|${COVERAGE_STRUCTURES.map(() => '---:').join('|')}|`);
    for (const row of report.tasks) {
        console.log(`| ${row.task} | ${row.structures.map(cell => cell.count || '—').join(' | ')} |`);
    }
    console.log('\n## High-priority empty cells\n');
    const empty = report.tasks.flatMap(row => row.structures
        .filter(cell => cell.highPriority && cell.count === 0)
        .map(cell => ({ task: row.task, ...cell })));
    for (const cell of empty) {
        console.log(`- ${cell.task} / ${cell.structure}: ${cell.disposition.status} — ${cell.disposition.reason} Dependency: ${cell.disposition.dependency}.`);
    }
    console.log('\nUse `--json` for stable machine-readable case identities and review status.');
}
