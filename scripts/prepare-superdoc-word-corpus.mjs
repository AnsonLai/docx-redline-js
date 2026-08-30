import '../tests/setup-xml-provider.mjs';

import assert from 'assert';
import { mkdirSync, readFileSync, writeFileSync } from 'fs';
import { dirname, join } from 'path';
import { fileURLToPath } from 'url';

import { applyOperationToDocumentXml } from '../services/standalone-operation-runner.js';
import { unzipEntries } from './lib/zip-reader.mjs';

const scriptDir = dirname(fileURLToPath(import.meta.url));
const repoRoot = dirname(scriptDir);
const manifest = JSON.parse(readFileSync(join(repoRoot, 'tests', 'corpus', 'superdoc-english-legal-administrative.json'), 'utf8'));
const catalogue = JSON.parse(readFileSync(join(repoRoot, 'tests', 'corpus', 'superdoc-word-scenarios.json'), 'utf8'));
const inputDir = join(repoRoot, 'tmp', 'superdoc-corpus');
const outputDir = join(repoRoot, 'tmp', 'superdoc-word-fixtures');
const manifestById = new Map(manifest.documents.map(document => [document.id, document]));

mkdirSync(outputDir, { recursive: true });

const suite = {
    generatedAt: new Date().toISOString(),
    source: manifest.source,
    attribution: manifest.attribution,
    datasetLicense: manifest.datasetLicense,
    cases: []
};

for (const [index, scenario] of catalogue.scenarios.entries()) {
    const source = manifestById.get(scenario.id);
    assert.ok(source, `Scenario ${scenario.id} is absent from the pinned manifest`);

    const sourcePath = join(inputDir, `${scenario.id}.docx`);
    const sourceBytes = readFileSync(sourcePath);
    const sourceEntries = unzipEntries(sourceBytes);
    const documentXml = sourceEntries.get('word/document.xml')?.toString('utf8');
    assert.ok(documentXml, `${scenario.id} has no word/document.xml`);

    const result = await applyOperationToDocumentXml(
        documentXml,
        scenario.operation,
        catalogue.author,
        null,
        { generateRedlines: true, existingRevisions: 'reject-input' }
    );

    assert.equal(result.status, 'ok', `${scenario.id}: ${result.error?.message || result.status}`);
    assert.equal(result.hasChanges, true, `${scenario.id}: operation was a no-op`);
    assert.ok(
        result.documentXml.includes('<w:ins') || result.documentXml.includes('<w:del'),
        `${scenario.id}: tracked revision markup missing`
    );

    const caseName = `${String(index + 1).padStart(2, '0')}-${source.type}-${scenario.id.slice(0, 12)}`;
    // Start from the original package. The Windows-only suite replaces just
    // word/document.xml with System.IO.Compression, then independently hashes
    // every untouched part before asking Word to open the result.
    writeFileSync(join(outputDir, `${caseName}.docx`), sourceBytes);
    writeFileSync(join(outputDir, `${caseName}.document.xml`), result.documentXml);
    writeFileSync(join(outputDir, `${caseName}.expected.json`), `${JSON.stringify({
        name: caseName,
        sourceId: scenario.id,
        assertionMode: 'word-source-exact',
        originalTarget: scenario.operation.target,
        modifiedTarget: scenario.operation.modified,
        shape: scenario.shape,
        coverage: scenario.coverage
    }, null, 2)}\n`);
    suite.cases.push({ name: caseName, sourceId: scenario.id, type: source.type, shape: scenario.shape });
    console.log(`Prepared ${caseName}`);
}

writeFileSync(join(outputDir, 'suite.json'), `${JSON.stringify(suite, null, 2)}\n`);
console.log(`Prepared ${suite.cases.length} reviewed SuperDoc Word fixtures in ${outputDir}`);
