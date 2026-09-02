import { readFileSync } from 'fs';

import { WORD_TASK_CASES } from '../../tests/fixtures/word-task-cases.mjs';
import {
    COVERAGE_ORACLES,
    COVERAGE_STRUCTURES,
    COVERAGE_TASKS,
    sortCoverageMetadata,
    validateCoverageMetadata
} from './word-coverage-metadata.mjs';

const corpusCatalogue = JSON.parse(readFileSync(
    new URL('../../tests/corpus/superdoc-word-scenarios.json', import.meta.url),
    'utf8'
));
const corpusManifest = JSON.parse(readFileSync(
    new URL('../../tests/corpus/superdoc-english-legal-administrative.json', import.meta.url),
    'utf8'
));
const corpusCoverage = JSON.parse(readFileSync(
    new URL('../../tests/corpus/superdoc-word-coverage.json', import.meta.url),
    'utf8'
));
const priorities = JSON.parse(readFileSync(
    new URL('../../tests/fixtures/coverage-matrix-priorities.json', import.meta.url),
    'utf8'
));

const manifestById = new Map(corpusManifest.documents.map(document => [document.id, document]));

export function validateScenarioIdentities(cases) {
    const identities = new Set();
    for (const item of cases) {
        if (!item.identity || identities.has(item.identity)) {
            throw new Error(`Duplicate or missing scenario identity: ${item.identity || '<missing>'}`);
        }
        identities.add(item.identity);
    }
}

function elementCount(xml, localName) {
    return (xml.match(new RegExp(`<w:${localName}(?:\\s|/|>)`, 'g')) || []).length;
}

function assertSyntheticClaim(testCase, structure) {
    const xml = testCase.sourceDocumentXml || '';
    const rules = {
        'plain-paragraph': () => !testCase.sourceText?.includes('\n') && !testCase.original.includes('\n'),
        'multi-paragraph': () => elementCount(xml, 'p') > 1 || testCase.sourceText?.includes('\n')
            || testCase.original.includes('\n') || testCase.modified.includes('\n'),
        'formatted-runs': () => /<(?:w:)?(?:b|i|u)(?:\s|\/|>)/.test(xml) || /(?:\*\*|\*|\+\+)/.test(testCase.modified),
        list: () => elementCount(xml, 'numPr') > 0,
        table: () => elementCount(xml, 'tbl') > 0,
        bookmark: () => elementCount(xml, 'bookmarkStart') > 0 && elementCount(xml, 'bookmarkEnd') > 0,
        hyperlink: () => elementCount(xml, 'hyperlink') > 0,
        'content-control': () => elementCount(xml, 'sdt') > 0,
        'tab-break': () => elementCount(xml, 'tab') > 0 || elementCount(xml, 'br') > 0,
        field: () => elementCount(xml, 'fldChar') > 0,
        comment: () => elementCount(xml, 'commentRangeStart') > 0 && Boolean(testCase.packageParts?.commentsXml),
        note: () => (elementCount(xml, 'footnoteReference') > 0 && Boolean(testCase.packageParts?.footnotesXml))
            || (elementCount(xml, 'endnoteReference') > 0 && Boolean(testCase.packageParts?.endnotesXml)),
        'header-footer': () => (elementCount(xml, 'headerReference') > 0 && Boolean(testCase.packageParts?.headers))
            || (elementCount(xml, 'footerReference') > 0 && Boolean(testCase.packageParts?.footers)),
        'section-boundary': () => elementCount(xml, 'sectPr') > 0,
        'prior-revisions': () => elementCount(xml, 'ins') > 0 || elementCount(xml, 'del') > 0
    };
    if (!rules[structure]()) {
        throw new Error(`synthetic:${testCase.name}: ${structure} is not supported by the fixture`);
    }
}

function syntheticTask(testCase) {
    if (testCase.expectAtomicRollback) return 'mixed-batch';
    if (testCase.expectNoOp) return 'accept-reject';
    if (testCase.task.startsWith('apply-')) return 'format';
    if (testCase.task.includes('deletion') || testCase.task.startsWith('delete-')) return 'delete';
    if (testCase.task.includes('insertion') || testCase.task.startsWith('insert-')
        || testCase.task === 'preserve-dollar-delimiters') return 'insert';
    return 'replace';
}

function corpusStructures(scenario) {
    const labels = new Set([scenario.shape, ...scenario.coverage]);
    const structures = new Set();
    if (scenario.shape === 'list') structures.add('list');
    if (scenario.shape === 'table-form') structures.add('table');
    if (scenario.shape === 'body-paragraph' || scenario.shape === 'administrative-layout' || scenario.shape === 'legal-apparatus') {
        structures.add('plain-paragraph');
    }
    const mappings = [
        [['table', 'tables', 'table-cell'], 'table'],
        [['numbering', 'numbered-list'], 'list'],
        [['bookmark', 'bookmarks', 'bookmark-adjacency'], 'bookmark'],
        [['header', 'headers', 'footer', 'footers'], 'header-footer'],
        [['footnotes-part'], 'note'],
        [['tabs'], 'tab-break'],
        [['field-adjacency'], 'field'],
        [['multi-section', 'section-properties'], 'section-boundary']
    ];
    for (const [sourceLabels, structure] of mappings) {
        if (sourceLabels.some(label => labels.has(label))) structures.add(structure);
    }
    return [...structures];
}

function corpusScenarioKey(scenario) {
    return scenario.key || scenario.id;
}

function corpusSourceId(scenario) {
    return scenario.sourceId || scenario.id;
}

function corpusOperations(scenario) {
    return scenario.operations || [scenario.operation];
}

export function loadCoverageCatalogue() {
    const synthetic = WORD_TASK_CASES.map(testCase => {
        const identity = `synthetic:${testCase.name}`;
        const metadata = sortCoverageMetadata(validateCoverageMetadata(testCase.coverageMetadata, identity));
        if (metadata.task !== syntheticTask(testCase)) {
            throw new Error(`${identity}: task ${metadata.task} is unsupported by fixture task ${testCase.task}`);
        }
        for (const structure of metadata.structures) assertSyntheticClaim(testCase, structure);
        return { identity, lane: 'synthetic', category: testCase.category, detail: testCase.task, metadata };
    });

    const corpus = corpusCatalogue.scenarios.map(scenario => {
        const scenarioKey = corpusScenarioKey(scenario);
        const sourceId = corpusSourceId(scenario);
        const identity = `superdoc:${scenarioKey}`;
        const source = manifestById.get(sourceId);
        if (!source) throw new Error(`${identity}: source is absent from the pinned corpus manifest`);
        const operations = corpusOperations(scenario);
        if (operations.some(operation => operation?.type !== 'replace')) {
            throw new Error(`${identity}: corpus scenarios currently support replace operations only`);
        }
        const declaredStructures = corpusCoverage.structuresByScenario[scenarioKey];
        const supportedStructures = corpusStructures(scenario).sort();
        if (!declaredStructures || JSON.stringify([...declaredStructures].sort()) !== JSON.stringify(supportedStructures)) {
            throw new Error(`${identity}: declared structures do not match reviewed fixture labels`);
        }
        const metadata = sortCoverageMetadata(validateCoverageMetadata({
            ...corpusCoverage.defaults,
            task: corpusCoverage.tasksByScenario?.[scenarioKey]
                || (operations.length > 1 ? 'mixed-batch' : 'replace'),
            structures: declaredStructures
        }, identity));
        const expectedTask = operations.length > 1 ? 'mixed-batch' : 'replace';
        if (metadata.task !== expectedTask) {
            throw new Error(`${identity}: task ${metadata.task} does not match ${operations.length} operation(s)`);
        }
        return { identity, lane: 'superdoc', category: source.type, detail: scenario.shape, metadata };
    });
    const cases = [...synthetic, ...corpus];
    if (Object.keys(corpusCoverage.structuresByScenario).length !== corpus.length) {
        throw new Error('SuperDoc coverage metadata contains a missing or unknown scenario identity');
    }
    validateScenarioIdentities(cases);
    return { cases, priorities };
}

export function buildCoverageMatrix(cases) {
    const cells = new Map();
    for (const task of COVERAGE_TASKS) {
        for (const structure of COVERAGE_STRUCTURES) cells.set(`${task}/${structure}`, []);
    }
    for (const testCase of cases) {
        for (const structure of testCase.metadata.structures) {
            cells.get(`${testCase.metadata.task}/${structure}`).push(testCase.identity);
        }
    }
    return cells;
}

export function validateCoveragePriorities(cases, priorityConfig = priorities) {
    const cells = buildCoverageMatrix(cases);
    const dispositions = new Map(priorityConfig.emptyCellDispositions.map(item => [`${item.task}/${item.structure}`, item]));
    for (const item of [...priorityConfig.highPriorityCells, ...priorityConfig.emptyCellDispositions]) {
        if (!COVERAGE_TASKS.includes(item.task) || !COVERAGE_STRUCTURES.includes(item.structure)) {
            throw new Error(`Unknown priority cell: ${item.task}/${item.structure}`);
        }
    }
    for (const cell of priorityConfig.highPriorityCells) {
        const key = `${cell.task}/${cell.structure}`;
        if (cells.get(key).length === 0 && !dispositions.has(key)) {
            throw new Error(`High-priority empty cell lacks a disposition: ${key}`);
        }
    }
    for (const [key, disposition] of dispositions) {
        if (!['planned', 'excluded'].includes(disposition.status) || !disposition.reason || !disposition.dependency) {
            throw new Error(`${key}: disposition requires status, reason, and dependency`);
        }
        if (cells.get(key).length > 0) throw new Error(`${key}: obsolete empty-cell disposition is still present`);
    }
    return cells;
}

export { COVERAGE_ORACLES, COVERAGE_STRUCTURES, COVERAGE_TASKS };
