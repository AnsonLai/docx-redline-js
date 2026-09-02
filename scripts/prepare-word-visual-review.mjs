import { mkdirSync, writeFileSync } from 'fs';
import { dirname, join, resolve } from 'path';
import { fileURLToPath } from 'url';

import { WORD_TASK_CASES } from '../tests/fixtures/word-task-cases.mjs';

export const VISUAL_STRUCTURES = Object.freeze([
    'formatted-runs',
    'list',
    'table',
    'bookmark',
    'hyperlink',
    'content-control',
    'tab-break',
    'field',
    'comment',
    'note',
    'header-footer',
    'section-boundary'
]);

const VISUAL_STRUCTURE_SET = new Set(VISUAL_STRUCTURES);
const VIEW_NAMES = Object.freeze(['allMarkup', 'acceptAll', 'rejectAll']);

export function selectVisualReviewCases(cases = WORD_TASK_CASES, requestedNames = []) {
    const requested = new Set(requestedNames);
    if (requested.size > 0) {
        const knownNames = new Set(cases.map(testCase => testCase.name));
        const unknown = [...requested].filter(name => !knownNames.has(name));
        if (unknown.length > 0) throw new Error(`Unknown visual-review case(s): ${unknown.join(', ')}`);
    }

    return cases
        .filter(testCase => requested.size > 0
            ? requested.has(testCase.name)
            : testCase.coverageMetadata.structures.some(structure => VISUAL_STRUCTURE_SET.has(structure)))
        .sort((a, b) => a.name.localeCompare(b.name));
}

export function buildVisualReviewManifest(cases = WORD_TASK_CASES, requestedNames = []) {
    const selected = selectVisualReviewCases(cases, requestedNames);
    return {
        schemaVersion: 1,
        reviewType: 'render-only-visual-preflight',
        certification: {
            status: 'pending',
            reviewer: null,
            reviewedAt: null,
            note: 'Rendering proves that Word produced inspectable evidence; it does not certify visual correctness.'
        },
        word: { version: null, build: null },
        cases: selected.map(testCase => ({
            identity: `synthetic:${testCase.name}`,
            name: testCase.name,
            category: testCase.category,
            task: testCase.coverageMetadata.task,
            structures: [...testCase.coverageMetadata.structures],
            renderStatus: 'pending',
            views: Object.fromEntries(VIEW_NAMES.map(view => [view, {
                status: 'pending',
                pdf: `${testCase.name}-${view}.pdf`,
                pages: null,
                bytes: null
            }]))
        }))
    };
}

function parseArgs(argv) {
    const requestedNames = argv
        .filter(argument => argument.startsWith('--case='))
        .map(argument => argument.slice('--case='.length))
        .filter(Boolean);
    const outputArgument = argv.find(argument => argument.startsWith('--output='));
    return {
        requestedNames,
        outputPath: outputArgument
            ? resolve(process.cwd(), outputArgument.slice('--output='.length))
            : join(process.cwd(), 'tmp', 'word-visual-review', 'manifest.json')
    };
}

const isCli = process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url);
if (isCli) {
    const { requestedNames, outputPath } = parseArgs(process.argv.slice(2));
    const manifest = buildVisualReviewManifest(WORD_TASK_CASES, requestedNames);
    mkdirSync(dirname(outputPath), { recursive: true });
    writeFileSync(outputPath, `${JSON.stringify(manifest, null, 2)}\n`, 'utf8');
    console.log(`Prepared ${manifest.cases.length} layout-sensitive cases for Word visual rendering: ${outputPath}`);
}
