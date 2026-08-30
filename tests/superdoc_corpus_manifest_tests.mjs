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
assert.equal(manifest.documents.length, 20);
assert.equal(manifest.documents.filter(document => document.type === 'legal').length, 10);
assert.equal(manifest.documents.filter(document => document.type === 'administrative').length, 10);

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

assert.equal(catalogue.scenarios.length, manifest.documents.length);
assert.deepEqual(
    new Set(catalogue.scenarios.map(scenario => scenario.id)),
    ids,
    'Every pinned source must have exactly one reviewed deterministic scenario'
);
assert.deepEqual(
    new Set(catalogue.scenarios.map(scenario => scenario.shape)),
    new Set(['body-paragraph', 'legal-apparatus', 'list', 'table-form', 'administrative-layout'])
);
for (const scenario of catalogue.scenarios) {
    assert.ok(scenario.review.length > 20, `${scenario.id}: review record is missing`);
    assert.equal(scenario.operation.type, 'replace');
    assert.ok(scenario.operation.target.length > 0);
    assert.ok(scenario.operation.modified.length > 0);
    assert.notEqual(scenario.operation.target, scenario.operation.modified);
}

const zipRoundTrip = unzipEntries(buildZip([
    { name: 'word/document.xml', data: '<document>compressible compressible compressible</document>' },
    { name: 'word/media/raw.bin', data: Buffer.from([0, 1, 2, 3]) }
]));
assert.equal(zipRoundTrip.get('word/document.xml').toString(), '<document>compressible compressible compressible</document>');
assert.deepEqual(zipRoundTrip.get('word/media/raw.bin'), Buffer.from([0, 1, 2, 3]));

console.log(`PASS: reviewed SuperDoc corpus (${manifest.documents.length} pinned scenarios)`);
