import assert from 'assert/strict';
import { execFileSync } from 'child_process';
import { existsSync, readFileSync } from 'fs';
import { join } from 'path';
import { fileURLToPath } from 'url';

import { DOMParser as XmldomParser, XMLSerializer } from '@xmldom/xmldom';

import { WORD_TASK_CASES } from './fixtures/word-task-cases.mjs';

// docx-preview expects the browser's firstElementChild convenience property.
// Add only that DOM compatibility surface so parseAsync can validate the exact
// embedded packages without pretending to test browser layout in Node.
const probe = new XmldomParser().parseFromString('<root><child/></root>', 'application/xml');
for (const prototype of [Object.getPrototypeOf(probe), Object.getPrototypeOf(probe.documentElement)]) {
    if (!Object.getOwnPropertyDescriptor(prototype, 'firstElementChild')) {
        Object.defineProperty(prototype, 'firstElementChild', {
            configurable: true,
            get() {
                for (let child = this.firstChild; child; child = child.nextSibling) {
                    if (child.nodeType === 1) return child;
                }
                return null;
            }
        });
    }
}
globalThis.DOMParser = XmldomParser;
globalThis.XMLSerializer = XMLSerializer;
globalThis.Node = { ELEMENT_NODE: 1 };

const { parseAsync } = await import('docx-preview');
const repoRoot = fileURLToPath(new URL('../', import.meta.url));
const fixturesDir = fileURLToPath(new URL('../tmp/docxjs-dashboard-fixtures', import.meta.url));
execFileSync(process.execPath, [
    'scripts/export-validation-fixtures.mjs',
    '--output-dir',
    fixturesDir
], { cwd: repoRoot, encoding: 'utf8', maxBuffer: 10 * 1024 * 1024 });

let parsed = 0;
for (const testCase of WORD_TASK_CASES) {
    for (const state of ['source', 'tracked', 'accepted', 'rejected']) {
        const suffix = state === 'tracked' ? '' : `.${state}`;
        const bytes = readFileSync(join(fixturesDir, `${testCase.name}${suffix}.docx`));
        const document = await parseAsync(bytes, {
            experimental: true,
            renderChanges: state === 'tracked'
        });
        assert.ok(document.documentPart?.body, `${testCase.name} ${state} has no document body`);
        parsed++;
    }
}

assert.equal(parsed, WORD_TASK_CASES.length * 4);

const corpusFixturesDir = join(repoRoot, 'tmp', 'superdoc-word-fixtures');
const corpusSuitePath = join(corpusFixturesDir, 'suite.json');
if (existsSync(corpusSuitePath)) {
    const suite = JSON.parse(readFileSync(corpusSuitePath, 'utf8'));
    for (const testCase of suite.cases) {
        for (const state of ['source', 'tracked', 'accepted', 'rejected']) {
            const suffix = state === 'tracked' ? '' : `.${state}`;
            const path = join(corpusFixturesDir, `${testCase.name}${suffix}.docx`);
            assert.ok(existsSync(path), `${testCase.name} ${state} package is missing`);
            const document = await parseAsync(readFileSync(path), {
                experimental: true,
                renderChanges: state === 'tracked'
            });
            assert.ok(document.documentPart?.body, `${testCase.name} ${state} has no document body`);
            parsed++;
        }
    }
}

console.log(`PASS: docxjs parsed ${parsed} embedded comparison documents`);
