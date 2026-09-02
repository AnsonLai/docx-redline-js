import '../tests/setup-xml-provider.mjs';

import assert from 'assert';
import { mkdirSync, readFileSync, writeFileSync } from 'fs';
import { dirname, join, resolve } from 'path';
import { fileURLToPath } from 'url';

import { applyOperationsToDocumentXml } from '../services/standalone-operation-runner.js';
import { applyRedlineToOxml } from '../index.js';
import {
    acceptTrackedChangesInOoxml,
    rejectTrackedChangesInOoxml
} from '../services/revision-comment-management.js';
import { buildZip } from './lib/minimal-zip.mjs';
import { unzipEntries } from './lib/zip-reader.mjs';
import { loadCoverageCatalogue } from './lib/word-coverage-catalogue.mjs';

const scriptDir = dirname(fileURLToPath(import.meta.url));
const repoRoot = dirname(scriptDir);
const manifest = JSON.parse(readFileSync(join(repoRoot, 'tests', 'corpus', 'superdoc-english-legal-administrative.json'), 'utf8'));
const catalogue = JSON.parse(readFileSync(join(repoRoot, 'tests', 'corpus', 'superdoc-word-scenarios.json'), 'utf8'));
function argumentPath(name, fallback) {
    const index = process.argv.indexOf(name);
    if (index < 0) return fallback;
    if (!process.argv[index + 1]) throw new Error(`${name} requires a path`);
    return resolve(process.cwd(), process.argv[index + 1]);
}

const inputDir = argumentPath('--input-dir', join(repoRoot, 'tmp', 'superdoc-corpus'));
const outputDir = argumentPath('--output-dir', join(repoRoot, 'tmp', 'superdoc-word-fixtures'));
const manifestById = new Map(manifest.documents.map(document => [document.id, document]));
const normalizedCoverageById = new Map(loadCoverageCatalogue().cases
    .filter(item => item.lane === 'superdoc')
    .map(item => [item.identity.slice('superdoc:'.length), item.metadata]));

mkdirSync(outputDir, { recursive: true });

function packageWithPartXml(sourceEntries, partName, partXml) {
    return buildZip(Array.from(sourceEntries, ([name, data]) => ({
        name,
        data: name === partName ? partXml : data
    })));
}

function assertInsertionFormatting(scenario, trackedXml) {
    const forbidden = scenario.formatExpectations?.insertionsForbidRunProperties || [];
    const requiredValues = scenario.formatExpectations?.insertionsRequireRunPropertyValues || {};
    if (forbidden.length === 0 && Object.keys(requiredValues).length === 0) return;

    const xmlDoc = new DOMParser().parseFromString(trackedXml, 'application/xml');
    const insertions = Array.from(xmlDoc.getElementsByTagName('*'))
        .filter(node => node.localName === 'ins');
    assert.ok(insertions.length > 0, `${scenario.key}: expected tracked insertions for format validation`);

    for (const insertion of insertions) {
        const runs = Array.from(insertion.getElementsByTagName('*'))
            .filter(node => node.localName === 'r' && Array.from(node.childNodes)
                .some(child => child.localName === 't'));
        const runProperties = runs.flatMap(run => Array.from(run.childNodes)
            .filter(node => node.localName === 'rPr'));
        for (const propertyName of forbidden) {
            const inherited = runProperties.some(rPr => Array.from(rPr.getElementsByTagName('*'))
                .some(node => node.localName === propertyName));
            assert.equal(inherited, false,
                `${scenario.key}: inserted text unexpectedly inherited w:${propertyName}`);
        }
        for (const [propertyName, expectedValue] of Object.entries(requiredValues)) {
            assert.ok(runs.length > 0, `${scenario.key}: expected inserted text runs for format validation`);
            for (const run of runs) {
                const rPr = Array.from(run.childNodes).find(node => node.localName === 'rPr');
                const property = rPr && Array.from(rPr.childNodes)
                    .find(node => node.localName === propertyName);
                const actualValue = property?.getAttribute('w:val') || property?.getAttribute('val');
                assert.equal(actualValue, String(expectedValue),
                    `${scenario.key}: inserted text must retain w:${propertyName}=${expectedValue}`);
            }
        }
    }
}

function countMatches(text, pattern) {
    return Array.from(text.matchAll(pattern)).length;
}

function assertSourceComplexity(scenario, source, sourceEntries, documentXml) {
    const expected = scenario.structuralExpectations;
    if (!expected) return;

    const observed = {
        words: source.wordCount,
        tables: countMatches(documentXml, /<w:tbl(?:\s|>)/g),
        listParagraphs: countMatches(documentXml, /<w:numPr(?:\s|>)/g),
        sections: countMatches(documentXml, /<w:sectPr(?:\s|>)/g),
        headers: Array.from(sourceEntries.keys()).filter(name => /^word\/header\d+\.xml$/.test(name)).length,
        footers: Array.from(sourceEntries.keys()).filter(name => /^word\/footer\d+\.xml$/.test(name)).length
    };
    for (const [label, minimum] of Object.entries({
        words: expected.minWords,
        tables: expected.minTables,
        listParagraphs: expected.minListParagraphs,
        sections: expected.minSections,
        headers: expected.minHeaders,
        footers: expected.minFooters
    })) {
        if (minimum === undefined) continue;
        assert.ok(observed[label] >= minimum,
            `${scenario.key}: expected at least ${minimum} ${label}, observed ${observed[label]}`);
    }
}

const suite = {
    generatedAt: new Date().toISOString(),
    source: manifest.source,
    attribution: manifest.attribution,
    datasetLicense: manifest.datasetLicense,
    cases: []
};

for (const [index, scenario] of catalogue.scenarios.entries()) {
    const scenarioKey = scenario.key || scenario.id;
    const sourceId = scenario.sourceId || scenario.id;
    const operations = scenario.operations || [scenario.operation];
    const revisionPart = scenario.part || 'word/document.xml';
    const source = manifestById.get(sourceId);
    assert.ok(source, `Scenario ${scenarioKey} is absent from the pinned manifest`);

    const sourcePath = join(inputDir, `${sourceId}.docx`);
    const sourceBytes = readFileSync(sourcePath);
    const sourceEntries = unzipEntries(sourceBytes);
    const sourcePartXml = sourceEntries.get(revisionPart)?.toString('utf8');
    assert.ok(sourcePartXml, `${scenarioKey} has no ${revisionPart}`);
    const documentXml = sourceEntries.get('word/document.xml')?.toString('utf8');
    assert.ok(documentXml, `${scenarioKey} has no word/document.xml`);
    assertSourceComplexity(scenario, source, sourceEntries, documentXml);

    let trackedXml;
    if (revisionPart === 'word/document.xml') {
        const result = await applyOperationsToDocumentXml(
            sourcePartXml,
            operations,
            catalogue.author,
            null,
            { generateRedlines: true, existingRevisions: 'reject-input' }
        );
        assert.notEqual(result.status, 'error', `${scenarioKey}: ${result.error?.message || result.status}`);
        assert.equal(result.hasChanges, true, `${scenarioKey}: operation was a no-op`);
        trackedXml = result.documentXml;
    } else {
        assert.equal(operations.length, 1, `${scenarioKey}: related-part scenarios support one operation`);
        const operation = operations[0];
        const result = await applyRedlineToOxml(sourcePartXml, operation.target, operation.modified, {
            generateRedlines: true,
            author: catalogue.author,
            existingRevisions: 'reject-input'
        });
        assert.equal(result.status, 'ok', `${scenarioKey}: ${result.error?.message || result.status}`);
        assert.equal(result.hasChanges, true, `${scenarioKey}: operation was a no-op`);
        trackedXml = result.oxml;
    }

    assert.ok(
        trackedXml.includes('<w:ins') || trackedXml.includes('<w:del'),
        `${scenarioKey}: tracked revision markup missing`
    );
    assertInsertionFormatting(scenario, trackedXml);

    const caseToken = scenario.key || sourceId.slice(0, 12);
    const caseName = `${String(index + 1).padStart(2, '0')}-${source.type}-${caseToken}`;
    const accepted = acceptTrackedChangesInOoxml(trackedXml, { allAuthors: true });
    const rejected = rejectTrackedChangesInOoxml(trackedXml, { allAuthors: true });
    assert.notEqual(accepted.status, 'error', `${scenarioKey}: accepting revisions failed`);
    assert.notEqual(rejected.status, 'error', `${scenarioKey}: rejecting revisions failed`);

    // Emit complete comparison packages for the dashboard. The Windows suite
    // may independently repackage the tracked state before its hash checks.
    writeFileSync(join(outputDir, `${caseName}.source.docx`), sourceBytes);
    writeFileSync(join(outputDir, `${caseName}.docx`), packageWithPartXml(sourceEntries, revisionPart, trackedXml));
    writeFileSync(join(outputDir, `${caseName}.accepted.docx`), packageWithPartXml(sourceEntries, revisionPart, accepted.oxml));
    writeFileSync(join(outputDir, `${caseName}.rejected.docx`), packageWithPartXml(sourceEntries, revisionPart, rejected.oxml));
    writeFileSync(join(outputDir, `${caseName}.document.xml`), trackedXml);
    writeFileSync(join(outputDir, `${caseName}.expected.json`), `${JSON.stringify({
        name: caseName,
        scenarioKey,
        sourceId,
        revisionPart,
        assertionScope: revisionPart.startsWith('word/header') ? 'headers' : 'document',
        assertionMode: 'word-source-exact',
        replacements: operations.map(operation => ({
            originalTarget: operation.target,
            modifiedTarget: operation.modified
        })),
        originalTarget: operations[0].target,
        modifiedTarget: operations[0].modified,
        sourceText: operations.map(operation => operation.target).join('\n'),
        expectedAcceptedText: operations.map(operation => operation.modified).join('\n'),
        expectedRejectedText: operations.map(operation => operation.target).join('\n'),
        shape: scenario.shape,
        coverage: scenario.coverage,
        coverageMetadata: normalizedCoverageById.get(scenarioKey)
    }, null, 2)}\n`);
    suite.cases.push({
        name: caseName,
        scenarioKey,
        sourceId,
        revisionPart,
        type: source.type,
        shape: scenario.shape,
        coverageMetadata: normalizedCoverageById.get(scenarioKey)
    });
    console.log(`Prepared ${caseName}`);
}

writeFileSync(join(outputDir, 'suite.json'), `${JSON.stringify(suite, null, 2)}\n`);
console.log(`Prepared ${suite.cases.length} reviewed SuperDoc Word fixtures in ${outputDir}`);
