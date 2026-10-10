/**
 * Architecture guard (WP-6): a tracked paragraph mark (`w:ins` / `w:del` under `w:pPr/w:rPr`) is constructed in exactly
 * one module, `engine/paragraph-mark-revisions.js`. Every write path (reconstruction, list generation, table
 * generation, document-level shortcuts) must call its helpers so no path can forget a rule the others have.
 *
 * Same technique as tests/no_word_api_index_check.mjs: scan the source (comments stripped) with patterns.
 *
 * Detected constructions outside the helper:
 *   1. string templates that open an rPr with an ins/del/move marker: `<w:rPr><w:ins ...`
 *   2. DOM construction of a w:ins / w:del / w:moveFrom / w:moveTo element (createWordElement, createElementNS,
 *      createElement, createTrackChange) that is attached to an rPr within a few lines (the paragraph mark is the only rPr child ins/del)
 *   3. the retired helper names (markParagraphMarkInserted / markParagraphMarkDeleted)
 *
 * ALLOWLIST: file -> reason. Empty by design; add an entry only with a written reason.
 */
import assert from 'assert';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const SCAN_DIRS = ['adapters', 'core', 'document', 'engine', 'node', 'orchestration', 'pipeline', 'services'];
const HELPER = 'engine/paragraph-mark-revisions.js';
const ALLOWLIST = new Map([
    // ['path/from/root.js', 'reason'],
]);

const REVISION_ELEMENT = String.raw`(?:w:)?(?:ins|del|moveFrom|moveTo)`;
const QUOTE = "['\"`]";
const TEMPLATE_RPR_MARKER = new RegExp(String.raw`<w:rPr[^>]*>\s*(?:\$\{[^}]*\}\s*)?<w:(?:ins|del|moveFrom|moveTo)\b`, 'g');
const DOM_CONSTRUCTION = new RegExp(
    String.raw`\b(?:createWordElement|createElementNS|createElement)\s*\((?:[^()]|\([^()]*\))*?${QUOTE}${REVISION_ELEMENT}${QUOTE}`,
    'g'
);
const TRACK_CHANGE_FACTORY = /\bcreateTrackChange\s*\(/g;
const RETIRED_NAMES = /\bmarkParagraphMark(?:Inserted|Deleted)\b/g;
const ATTACHED_TO_RPR = /\b\w*rPr\w*\s*\.\s*(?:appendChild|insertBefore)\s*\(/i;
const WINDOW = 6;

function stripComments(source) {
    return source
        .replace(/\/\*[\s\S]*?\*\//g, match => match.replace(/[^\n]/g, ' '))
        .replace(/(^|[^:'"`])\/\/.*$/gm, '$1');
}

function walk(dir) {
    const out = [];
    for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
        const full = path.join(dir, entry.name);
        if (entry.isDirectory()) out.push(...walk(full));
        else if (entry.name.endsWith('.js')) out.push(full);
    }
    return out;
}

const lineOf = (source, index) => source.slice(0, index).split('\n').length;

export function findViolations(rawSource) {
    const source = stripComments(rawSource);
    const lines = source.split('\n');
    const hits = [];
    for (const match of source.matchAll(TEMPLATE_RPR_MARKER)) {
        hits.push({ line: lineOf(source, match.index), rule: 'rPr string template with a revision marker' });
    }
    for (const regex of [DOM_CONSTRUCTION, TRACK_CHANGE_FACTORY]) {
        for (const match of source.matchAll(regex)) {
            if (regex === TRACK_CHANGE_FACTORY) {
                // createTrackChange builds run wrappers; it only counts when it is used on paragraph properties.
                const call = source.slice(match.index, match.index + 200);
                if (!/['"`](?:ins|del)['"`]/.test(call)) continue;
            }
            const line = lineOf(source, match.index);
            const window = lines.slice(line - 1, line + WINDOW).join('\n');
            if (ATTACHED_TO_RPR.test(window)) {
                hits.push({ line, rule: 'revision element created and attached to an rPr' });
            }
        }
    }
    for (const match of source.matchAll(RETIRED_NAMES)) {
        hits.push({ line: lineOf(source, match.index), rule: 'retired markParagraphMark* helper' });
    }
    return hits;
}

// Self-checks: the patterns must catch each construction style.
const SAMPLE_TEMPLATE = '`<w:pPr><w:rPr><w:ins w:id="${id}"/></w:rPr></w:pPr>`';
const SAMPLE_DOM = "const pPr = x; const rPr = y; const m = createWordElement(doc, 'w:del'); rPr.appendChild(m);";
const SAMPLE_RUN_WRAPPER = "const ins = createWordElement(xmlDoc, 'w:ins');\nins.appendChild(run);";
assert.equal(findViolations(SAMPLE_TEMPLATE).length, 1, 'template construction is detected');
assert.equal(findViolations(SAMPLE_DOM).length, 1, 'DOM construction attached to an rPr is detected');
assert.equal(findViolations(SAMPLE_RUN_WRAPPER).length, 0, 'a run-level wrapper is not a paragraph mark');
assert.equal(findViolations('// <w:rPr><w:ins/> in a comment\nconst a = 1;').length, 0, 'comments are ignored');

const violations = [];
for (const dir of SCAN_DIRS) {
    const absolute = path.join(root, dir);
    if (!fs.existsSync(absolute)) continue;
    for (const file of walk(absolute)) {
        const relative = path.relative(root, file).replace(/\\/g, '/');
        if (relative === HELPER || ALLOWLIST.has(relative)) continue;
        for (const hit of findViolations(fs.readFileSync(file, 'utf8'))) {
            violations.push(`${relative}:${hit.line} ${hit.rule}`);
        }
    }
}

assert.deepEqual(violations, [], [
    `Paragraph-mark revisions must be constructed only in ${HELPER}.`,
    'Call markParagraphInserted / markParagraphDeleted / setParagraphPropertiesMarkRevision / paragraphMarkPPrXml instead:',
    ...violations
].join('\n'));
assert.ok(fs.existsSync(path.join(root, HELPER)), 'the helper module exists');

console.log('PASS: paragraph-mark revisions are constructed in a single module');
