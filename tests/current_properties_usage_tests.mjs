/**
 * Lint-style guard: descendant lookups cross w:pPrChange / w:rPrChange snapshots, so reading a property by
 * descendant search makes historical values look current. Source files must use core/current-properties.js
 * (direct-child reads) unless the site is listed here with a reason.
 */
import assert from 'assert';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const SCAN_DIRS = ['core', 'services', 'engine', 'pipeline', 'orchestration', 'document'];

const PROPERTY_NAMES = ['numPr', 'numId', 'ilvl', 'pStyle', 'outlineLvl', 'rFonts', 'sz', 'b', 'i', 'u'];
const DESCENDANT_LOOKUPS = [
    'getElementsByTagNameNS', 'getElementsByTagName', 'getWordElementsByLocalName', 'getFirstElementByTagNS',
    'getFirstElementByTag', 'getElementsByTagNSOrTag', 'getFirstElementByTagNSOrTag', 'getElementsByTagNS',
    'getWordDescendants', 'querySelector', 'querySelectorAll'
];

// file -> reason. A listed file may use descendant lookups for these names.
const ALLOWLIST = new Map([
    ['core/current-properties.js', 'the accessor itself: getAllReferencedNumIdElements is the explicit historical-inclusive helper'],
    ['services/standalone-docx-plumbing.js', 'package-level "does any numPr exist anywhere" check decides whether a numbering part is required; snapshots count'],
]);

const QUOTE = "['\"`]";
const callPattern = new RegExp(
    String.raw`\b(?:${DESCENDANT_LOOKUPS.join("|")})\s*\((?:[^()]|\([^()]*\))*?${QUOTE}(?:w:)?(?:${PROPERTY_NAMES.join("|")})${QUOTE}`,
    "g"
);

 
function walk(dir) {
    const out = [];
    for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
        const full = path.join(dir, entry.name);
        if (entry.isDirectory()) out.push(...walk(full));
        else if (entry.name.endsWith('.js')) out.push(full);
    }
    return out;
}

export function findViolations(source) {
    const hits = [];
    for (const match of source.matchAll(callPattern)) {
        hits.push(source.slice(0, match.index).split('\n').length);
    }
    return hits;
}

// Self-check: the pattern catches the known bad shapes and ignores direct-child reads.
assert.ok(findViolations(`x.getElementsByTagNameNS(NS_W, 'numPr')`).length === 1);
assert.ok(findViolations(`getWordDescendants(rPr, 'b').length`).length === 1);
assert.ok(findViolations(`getFirstElementByTagNS(pPr, NS_W, 'w:pStyle')`).length === 1);
assert.ok(findViolations(`getElementsByTagNameNS(NS_W, 'p')`).length === 0);
assert.ok(findViolations(`getCurrentChild(rPr, 'b')`).length === 0);

const violations = [];
for (const dir of SCAN_DIRS) {
    for (const file of walk(path.join(root, dir))) {
        const rel = path.relative(root, file).split(path.sep).join('/');
        if (ALLOWLIST.has(rel)) continue;
        for (const line of findViolations(fs.readFileSync(file, 'utf8'))) {
            violations.push(`${rel}:${line}`);
        }
    }
}

assert.deepStrictEqual(
    violations,
    [],
    `Descendant property lookups cross pPrChange/rPrChange snapshots; use core/current-properties.js or allowlist with a reason:\n${violations.join('\n')}`
);
for (const rel of ALLOWLIST.keys()) {
    assert.ok(fs.existsSync(path.join(root, rel)), `stale allowlist entry: ${rel}`);
}
console.log('PASS: current properties usage tests');
