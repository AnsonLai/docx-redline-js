import { mkdirSync, readFileSync, writeFileSync } from 'fs';
import { dirname, join, resolve } from 'path';
import { fileURLToPath } from 'url';

const manifest = JSON.parse(readFileSync(
    new URL('../tests/corpus/superdoc-english-legal-administrative.json', import.meta.url),
    'utf8'
));
const catalogue = JSON.parse(readFileSync(
    new URL('../tests/corpus/superdoc-word-scenarios.json', import.meta.url),
    'utf8'
));
const sourceById = new Map(manifest.documents.map(item => [item.id, item]));
const VIEW_NAMES = Object.freeze(['allMarkup', 'acceptAll', 'rejectAll']);

export function selectCorpusVisualReviewCases(requestedKeys = []) {
    const requested = new Set(requestedKeys);
    const cases = catalogue.scenarios.map((scenario, index) => {
        const scenarioKey = scenario.key || scenario.id;
        const sourceId = scenario.sourceId || scenario.id;
        const source = sourceById.get(sourceId);
        const caseToken = scenario.key || sourceId.slice(0, 12);
        return {
            scenarioKey,
            name: `${String(index + 1).padStart(2, '0')}-${source.type}-${caseToken}`,
            category: source.type,
            shape: scenario.shape,
            operationCount: (scenario.operations || [scenario.operation]).length
        };
    });
    if (requested.size > 0) {
        const known = new Set(cases.map(item => item.scenarioKey));
        const unknown = [...requested].filter(key => !known.has(key));
        if (unknown.length) throw new Error(`Unknown corpus visual-review case(s): ${unknown.join(', ')}`);
        return cases.filter(item => requested.has(item.scenarioKey));
    }
    return cases.filter(item => !/^[a-f0-9]{64}$/.test(item.scenarioKey));
}

export function buildCorpusVisualReviewManifest(requestedKeys = []) {
    const selected = selectCorpusVisualReviewCases(requestedKeys);
    return {
        schemaVersion: 1,
        reviewType: 'render-only-corpus-visual-preflight',
        certification: {
            status: 'pending',
            reviewer: null,
            reviewedAt: null,
            note: 'Rendering proves that Word produced inspectable evidence; it does not certify visual correctness.'
        },
        word: { version: null, build: null },
        cases: selected.map(item => ({
            identity: `superdoc:${item.scenarioKey}`,
            ...item,
            renderStatus: 'pending',
            views: Object.fromEntries(VIEW_NAMES.map(view => [view, {
                status: 'pending',
                pdf: `${item.name}-${view}.pdf`,
                pages: null,
                bytes: null
            }]))
        }))
    };
}

function parseArgs(argv) {
    const requestedKeys = argv.filter(value => value.startsWith('--case=')).map(value => value.slice(7));
    const outputArgument = argv.find(value => value.startsWith('--output='));
    return {
        requestedKeys,
        outputPath: outputArgument
            ? resolve(process.cwd(), outputArgument.slice(9))
            : join(process.cwd(), 'tmp', 'superdoc-word-visual-review', 'manifest.json')
    };
}

const isCli = process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url);
if (isCli) {
    const { requestedKeys, outputPath } = parseArgs(process.argv.slice(2));
    const review = buildCorpusVisualReviewManifest(requestedKeys);
    mkdirSync(dirname(outputPath), { recursive: true });
    writeFileSync(outputPath, `${JSON.stringify(review, null, 2)}\n`, 'utf8');
    console.log(`Prepared ${review.cases.length} new SuperDoc scenarios for Word visual rendering: ${outputPath}`);
}
