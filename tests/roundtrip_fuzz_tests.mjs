import './setup-xml-provider.mjs';

import { assertRoundTrip } from './helpers/roundtrip.mjs';

/**
 * Seeded fuzz harness for the accept/reject round-trip invariant.
 *
 * Generates random paragraph structures (run counts, run properties,
 * hyperlinks, bookmarks, unicode, boundary whitespace) crossed with random
 * edits (replace/delete/insert at arbitrary and run-boundary offsets), then
 * asserts the Phase 1 round-trip invariant on each case.
 *
 * Deterministic by default so CI cannot flake. To explore new inputs or
 * reproduce a failure:
 *
 *   FUZZ_SEED=<caseSeed> FUZZ_ITERATIONS=1 node tests/roundtrip_fuzz_tests.mjs
 */

const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

const BASE_SEED = (Number(process.env.FUZZ_SEED) >>> 0) || 20260704;
const ITERATIONS = Math.max(1, Number(process.env.FUZZ_ITERATIONS) || 100);
const MAX_REPORTED_FAILURES = 5;

function mulberry32(seed) {
    let a = seed >>> 0;
    return function next() {
        a |= 0;
        a = (a + 0x6D2B79F5) | 0;
        let t = Math.imul(a ^ (a >>> 15), 1 | a);
        t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
        return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
    };
}

function makeRng(seed) {
    const next = mulberry32(seed);
    return {
        next,
        int: max => Math.floor(next() * max),
        pick: arr => arr[Math.floor(next() * arr.length)],
        chance: p => next() < p
    };
}

const WORDS = [
    'alpha', 'beta', 'gamma', 'delta', 'omega', 'contract', 'clause', 'party',
    'agreement', 'notice', 'term', 'section', 'payment', 'delivery', 'liability',
    'provision', 'warranty', 'remedy', 'schedule', 'annex'
];
const UNICODE_WORDS = ['条款', '合同', '契約', 'データ', 'ελληνικά', 'русский', '😀', '🚀', 'naïve', 'café'];
const RPR_POOL = [
    '',
    '<w:rPr><w:b/></w:rPr>',
    '<w:rPr><w:i/></w:rPr>',
    '<w:rPr><w:u w:val="single"/></w:rPr>',
    '<w:rPr><w:color w:val="FF0000"/></w:rPr>',
    '<w:rPr><w:b/><w:i/></w:rPr>'
];

function escapeXmlText(text) {
    return text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function randomWord(rng) {
    return rng.chance(0.15) ? rng.pick(UNICODE_WORDS) : rng.pick(WORDS);
}

function randomText(rng, minWords, maxWords) {
    const count = minWords + rng.int(maxWords - minWords + 1);
    return Array.from({ length: count }, () => randomWord(rng)).join(' ');
}

function textRunXml(text, rPr) {
    if (text === '') return '';
    const space = /^\s|\s$/.test(text) ? ' xml:space="preserve"' : '';
    return `<w:r>${rPr}<w:t${space}>${escapeXmlText(text)}</w:t></w:r>`;
}

/**
 * Splits `text` into 1..6 runs at random code-point boundaries and decorates
 * the result with random run properties, an optional hyperlink, and optional
 * bookmark markers. Returns { paragraphXml, text, cuts } where `cuts` are the
 * code-point offsets of run boundaries (used to aim edits at boundaries).
 */
function generateParagraph(rng) {
    const text = randomText(rng, 5, 16);
    const points = Array.from(text);

    const runCount = 1 + rng.int(6);
    const cuts = new Set();
    while (cuts.size < runCount - 1) {
        const cut = 1 + rng.int(points.length - 1);
        cuts.add(cut);
    }
    const boundaries = [0, ...[...cuts].sort((a, b) => a - b), points.length];

    const runs = [];
    for (let i = 0; i < boundaries.length - 1; i++) {
        runs.push({
            text: points.slice(boundaries[i], boundaries[i + 1]).join(''),
            rPr: rng.pick(RPR_POOL)
        });
    }

    const pieces = runs.map(run => textRunXml(run.text, run.rPr));

    if (runs.length >= 3 && rng.chance(0.15)) {
        const target = 1 + rng.int(runs.length - 2);
        pieces[target] = `<w:hyperlink r:id="rId9">${pieces[target]}</w:hyperlink>`;
    }

    if (rng.chance(0.15)) {
        const target = rng.int(pieces.length);
        pieces[target] = `<w:bookmarkStart w:id="3" w:name="fuzzMark"/>${pieces[target]}<w:bookmarkEnd w:id="3"/>`;
    }

    return {
        paragraphXml: `<w:p>${pieces.join('')}</w:p>`,
        text,
        cuts: boundaries.slice(1, -1)
    };
}

/**
 * Applies one random edit to `text` (code-point safe). Edits sometimes snap
 * to run boundaries, the positions where surgical splitting is most fragile.
 */
function applyRandomEdit(rng, text, cuts) {
    const points = Array.from(text);

    const pickOffset = maxExclusive => {
        if (cuts.length > 0 && rng.chance(0.3)) {
            const cut = rng.pick(cuts);
            if (cut < maxExclusive) return cut;
        }
        return rng.int(maxExclusive);
    };

    const op = rng.pick(['replace', 'replace', 'delete', 'insert', 'whitespace']);

    if (op === 'insert') {
        const at = pickOffset(points.length + 1);
        points.splice(at, 0, ...Array.from(` ${randomWord(rng)} `));
        return points.join('');
    }

    const start = pickOffset(Math.max(1, points.length - 1));
    const span = 1 + rng.int(Math.min(20, points.length - start));

    if (op === 'delete') {
        if (span >= points.length) return `${randomWord(rng)}`;
        points.splice(start, span);
        return points.join('');
    }

    const replacement = op === 'whitespace'
        ? ` ${randomWord(rng)}  ${randomWord(rng)} `
        : randomText(rng, 1, 3);
    points.splice(start, span, ...Array.from(replacement));
    return points.join('');
}

function generateCase(caseSeed) {
    const rng = makeRng(caseSeed);
    const { paragraphXml, text, cuts } = generateParagraph(rng);

    let modified = applyRandomEdit(rng, text, cuts);
    if (rng.chance(0.25)) {
        modified = applyRandomEdit(rng, modified, []);
    }
    if (modified.replace(/\s+/g, ' ').trim() === text.replace(/\s+/g, ' ').trim()) {
        modified = `${modified} ${randomWord(rng)}`;
    }

    const oxml = `<w:document xmlns:w="${NS_W}" xmlns:r="${NS_R}"><w:body>${paragraphXml}<w:sectPr/></w:body></w:document>`;
    return { oxml, original: text, modified };
}

let failures = 0;

for (let i = 0; i < ITERATIONS; i++) {
    const caseSeed = (BASE_SEED + i) >>> 0;
    const testCase = generateCase(caseSeed);

    try {
        await assertRoundTrip(testCase.oxml, testCase.original, testCase.modified);
    } catch (error) {
        failures++;
        console.error(`\nFUZZ CASE FAILED (seed ${caseSeed})`);
        console.error(`  reproduce: FUZZ_SEED=${caseSeed} FUZZ_ITERATIONS=1 node tests/roundtrip_fuzz_tests.mjs`);
        console.error(`  original: ${JSON.stringify(testCase.original)}`);
        console.error(`  modified: ${JSON.stringify(testCase.modified)}`);
        console.error(`  oxml: ${testCase.oxml}`);
        console.error(`  ${error?.message || error}`);
        if (failures >= MAX_REPORTED_FAILURES) {
            console.error(`\nStopping after ${MAX_REPORTED_FAILURES} failures.`);
            break;
        }
    }
}

if (failures > 0) {
    console.error(`\n${failures} fuzz case(s) failed out of ${ITERATIONS} (base seed ${BASE_SEED}).`);
    process.exit(1);
}

console.log(`PASS: ${ITERATIONS} fuzz round-trip cases (base seed ${BASE_SEED})`);
