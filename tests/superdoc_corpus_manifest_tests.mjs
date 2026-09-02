import assert from 'assert/strict';
import { readFileSync } from 'fs';
import { buildZip } from '../scripts/lib/minimal-zip.mjs';
import { unzipEntries } from '../scripts/lib/zip-reader.mjs';

const manifest = JSON.parse(readFileSync(
    new URL('./corpus/superdoc-english-legal-administrative.json', import.meta.url),
    'utf8'
));
const catalogue = JSON.parse(readFileSync(
    new URL('./corpus/superdoc-word-scenarios.json', import.meta.url),
    'utf8'
));

assert.equal(manifest.datasetLicense, 'ODC-By 1.0');
assert.match(manifest.attribution, /SuperDoc/);
assert.equal(manifest.documents.length, 23);
assert.equal(manifest.documents.filter(document => document.type === 'legal').length, 11);
assert.equal(manifest.documents.filter(document => document.type === 'administrative').length, 12);

const ids = new Set();
for (const document of manifest.documents) {
    assert.match(document.id, /^[a-f0-9]{64}$/);
    assert.match(document.downloadSha256, /^[a-f0-9]{64}$/);
    assert.ok(!ids.has(document.id), `Duplicate corpus id: ${document.id}`);
    ids.add(document.id);
    assert.equal(document.language, 'en');
    assert.ok(['legal', 'administrative'].includes(document.type));
    assert.ok(document.confidence >= manifest.selectionPolicy.minimumClassificationConfidence);
    assert.equal(document.downloadUrl, `https://docxcorp.us/documents/${document.id}.docx`);
}

assert.ok(catalogue.scenarios.length > manifest.documents.length);
assert.deepEqual(
    new Set(catalogue.scenarios.map(scenario => scenario.sourceId || scenario.id)),
    ids,
    'Every pinned source must have at least one reviewed deterministic scenario'
);
assert.deepEqual(
    new Set(catalogue.scenarios.map(scenario => scenario.shape)),
    new Set(['body-paragraph', 'legal-apparatus', 'list', 'table-form', 'administrative-layout', 'page-header'])
);
const scenarioKeys = new Set();
for (const scenario of catalogue.scenarios) {
    const key = scenario.key || scenario.id;
    assert.ok(!scenarioKeys.has(key), `Duplicate scenario key: ${key}`);
    scenarioKeys.add(key);
    assert.ok(scenario.review.length > 20, `${key}: review record is missing`);
    const operations = scenario.operations || [scenario.operation];
    assert.ok(operations.length > 0, `${key}: operations are missing`);
    for (const operation of operations) {
        assert.equal(operation.type, 'replace');
        assert.ok(operation.target.length > 0);
        assert.ok(operation.modified.length > 0);
        assert.notEqual(operation.target, operation.modified);
    }
    if (scenario.part) assert.match(scenario.part, /^word\/header[1-9][0-9]*\.xml$/);
}

const multiChangeScenarios = catalogue.scenarios.filter(scenario => scenario.operations?.length > 1);
assert.equal(multiChangeScenarios.length, 9);
assert.equal(multiChangeScenarios.filter(scenario => scenario.shape === 'list').length, 3);
assert.equal(multiChangeScenarios.filter(scenario => scenario.shape === 'table-form').length, 6);
assert.equal(multiChangeScenarios.reduce((total, scenario) => total + scenario.operations.length, 0), 52);
assert.ok(multiChangeScenarios.every(scenario => scenario.operations.length >= 3));
const complexityGuardedScenarios = catalogue.scenarios.filter(scenario => scenario.structuralExpectations);
assert.equal(complexityGuardedScenarios.length, 5);
assert.ok(complexityGuardedScenarios.every(scenario => scenario.operations.length === 8));
assert.ok(complexityGuardedScenarios.every(scenario =>
    scenario.structuralExpectations.minTables > 0 || scenario.structuralExpectations.minListParagraphs > 0
));
const headerScenarios = catalogue.scenarios.filter(scenario => scenario.shape === 'page-header');
assert.equal(headerScenarios.length, 2);
assert.ok(headerScenarios.every(scenario => scenario.part === 'word/header1.xml'));

const zipRoundTrip = unzipEntries(buildZip([
    { name: 'word/document.xml', data: '<document>compressible compressible compressible</document>' },
    { name: 'word/media/raw.bin', data: Buffer.from([0, 1, 2, 3]) }
]));
assert.equal(zipRoundTrip.get('word/document.xml').toString(), '<document>compressible compressible compressible</document>');
assert.deepEqual(zipRoundTrip.get('word/media/raw.bin'), Buffer.from([0, 1, 2, 3]));

console.log(`PASS: reviewed SuperDoc corpus (${manifest.documents.length} pinned documents, ${catalogue.scenarios.length} scenarios)`);
