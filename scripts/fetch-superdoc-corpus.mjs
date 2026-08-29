import { createHash } from 'crypto';
import { mkdirSync, readFileSync, writeFileSync } from 'fs';
import { dirname, join } from 'path';
import { fileURLToPath } from 'url';

const scriptDir = dirname(fileURLToPath(import.meta.url));
const repoRoot = dirname(scriptDir);
const manifestPath = join(repoRoot, 'tests', 'corpus', 'superdoc-english-legal-administrative.json');
const manifest = JSON.parse(readFileSync(manifestPath, 'utf8'));
const requestedIds = [];

for (let index = 2; index < process.argv.length; index++) {
    if (process.argv[index] === '--id' && process.argv[index + 1]) {
        requestedIds.push(process.argv[++index]);
    } else {
        throw new Error(`Unknown argument: ${process.argv[index]}. Use --id <pinned-sha256>.`);
    }
}

if (requestedIds.length === 0) {
    console.error('No document selected. Fetch one or more pinned references with:');
    console.error('  npm run corpus:fetch:superdoc -- --id <sha256> [--id <sha256>]');
    console.error('Pinned ids:');
    for (const document of manifest.documents) console.error(`  ${document.id}  ${document.type}  ${document.filename}`);
    process.exit(2);
}

const byId = new Map(manifest.documents.map(document => [document.id, document]));
const outputDir = join(repoRoot, 'tmp', 'superdoc-corpus');
mkdirSync(outputDir, { recursive: true });

for (const id of requestedIds) {
    const document = byId.get(id);
    if (!document) throw new Error(`Document ${id} is not in the pinned SuperDoc reference manifest.`);

    const response = await fetch(document.downloadUrl);
    if (!response.ok) throw new Error(`Failed to fetch ${id}: HTTP ${response.status}`);
    const bytes = Buffer.from(await response.arrayBuffer());
    const digest = createHash('sha256').update(bytes).digest('hex');
    if (digest !== id) throw new Error(`Hash mismatch for ${id}: received ${digest}`);

    writeFileSync(join(outputDir, `${id}.docx`), bytes);
    writeFileSync(join(outputDir, `${id}.source.json`), `${JSON.stringify({
        ...document,
        dataset: manifest.name,
        attribution: manifest.attribution,
        datasetLicense: manifest.datasetLicense,
        licenseUrl: manifest.licenseUrl
    }, null, 2)}\n`, 'utf8');
    console.log(`Fetched ${id}.docx (${bytes.length} bytes)`);
}
