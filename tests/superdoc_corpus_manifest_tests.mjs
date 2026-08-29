import assert from 'assert/strict';
import { readFileSync } from 'fs';

const manifest = JSON.parse(readFileSync(
    new URL('./corpus/superdoc-english-legal-administrative.json', import.meta.url),
    'utf8'
));

assert.equal(manifest.datasetLicense, 'ODC-By 1.0');
assert.match(manifest.attribution, /SuperDoc/);
assert.ok(manifest.documents.length >= 6);

const ids = new Set();
for (const document of manifest.documents) {
    assert.match(document.id, /^[a-f0-9]{64}$/);
    assert.ok(!ids.has(document.id), `Duplicate corpus id: ${document.id}`);
    ids.add(document.id);
    assert.equal(document.language, 'en');
    assert.ok(['legal', 'administrative'].includes(document.type));
    assert.ok(document.confidence >= manifest.selectionPolicy.minimumClassificationConfidence);
    assert.equal(document.downloadUrl, `https://docxcorp.us/documents/${document.id}.docx`);
}

console.log(`PASS: pinned SuperDoc corpus manifest (${manifest.documents.length} references)`);
