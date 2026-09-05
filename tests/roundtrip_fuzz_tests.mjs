import './setup-xml-provider.mjs';

import assert from 'assert/strict';

import { computeWordDiffs } from '../pipeline/diff-engine.js';
import { elementsByLocalName, parseXmlFragment } from './helpers/ooxml-assertions.mjs';
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

function ensureChanged(rng, original, modified) {
    if (modified.replace(/\s+/g, ' ').trim() !== original.replace(/\s+/g, ' ').trim()) {
        return modified;
    }
    return `${modified} ${randomWord(rng)}`;
}

function wrapBody(inner, { withSectPr = true } = {}) {
    const sectPr = withSectPr ? '<w:sectPr/>' : '';
    return `<w:document xmlns:w="${NS_W}" xmlns:r="${NS_R}"><w:body>${inner}${sectPr}</w:body></w:document>`;
}

/* --- shape: single paragraph (the original generator) --------------------- */

function shapeParagraph(rng) {
    const { paragraphXml, text, cuts } = generateParagraph(rng);

    let modified = applyRandomEdit(rng, text, cuts);
    if (rng.chance(0.25)) modified = applyRandomEdit(rng, modified, []);

    return {
        shape: 'paragraph',
        oxml: wrapBody(paragraphXml),
        original: text,
        modified: ensureChanged(rng, text, modified)
    };
}

/* --- shape: multi-paragraph body ------------------------------------------ */

function shapeMultiParagraph(rng) {
    const count = 2 + rng.int(4);
    const paragraphs = Array.from({ length: count }, () => generateParagraph(rng));
    const texts = paragraphs.map(entry => entry.text);

    const targetIndex = rng.int(count);
    const operation = rng.pick(['edit', 'edit', 'delete', 'insert']);
    const modifiedTexts = texts.slice();

    if (operation === 'delete' && count > 1) {
        modifiedTexts.splice(targetIndex, 1);
    } else if (operation === 'insert') {
        modifiedTexts.splice(targetIndex, 0, randomText(rng, 3, 8));
    } else {
        modifiedTexts[targetIndex] = applyRandomEdit(rng, texts[targetIndex], paragraphs[targetIndex].cuts);
    }

    const original = texts.join('\n');
    const modified = modifiedTexts.join('\n');

    return {
        shape: 'multiParagraph',
        oxml: wrapBody(paragraphs.map(entry => entry.paragraphXml).join('')),
        original,
        modified: ensureChanged(rng, original, modified)
    };
}

/* --- shape: table cell ----------------------------------------------------- */

function shapeTableCell(rng) {
    const rows = 1 + rng.int(3);
    const cols = 1 + rng.int(2);
    const cells = [];

    for (let r = 0; r < rows; r++) {
        const rowCells = [];
        for (let c = 0; c < cols; c++) rowCells.push(generateParagraph(rng));
        cells.push(rowCells);
    }

    const targetRow = rng.int(rows);
    const targetCol = rng.int(cols);
    const target = cells[targetRow][targetCol];
    const modified = ensureChanged(rng, target.text, applyRandomEdit(rng, target.text, target.cuts));

    const tableXml = `<w:tbl xmlns:w="${NS_W}" xmlns:r="${NS_R}">${cells
        .map(row => `<w:tr>${row.map(cell => `<w:tc>${cell.paragraphXml}</w:tc>`).join('')}</w:tr>`)
        .join('')}</w:tbl>`;

    // The engine isolates the matched cell paragraph, so the resolved text is
    // that paragraph alone rather than the whole table.
    return { shape: 'tableCell', oxml: tableXml, original: target.text, modified };
}

/* --- shape: whitespace-hostile -------------------------------------------- */

function shapeWhitespace(rng) {
    const left = randomText(rng, 2, 5);
    const right = randomText(rng, 2, 5);
    const separator = rng.pick(['  ', '   ', ' ', ' ']);
    const useTab = rng.chance(0.4);

    const text = `${left}${separator}${right}`;
    const tabPosition = useTab ? rng.pick(['leading', 'middle', 'trailing']) : null;
    let runs = textRunXml(text, '');
    let original = text;
    if (tabPosition === 'leading') {
        runs = `<w:r><w:tab/></w:r>${textRunXml(`${left} ${right}`, '')}`;
        original = `\t${left} ${right}`;
    } else if (tabPosition === 'middle') {
        runs = `${textRunXml(left, '')}<w:r><w:tab/></w:r>${textRunXml(right, '')}`;
        original = `${left}\t${right}`;
    } else if (tabPosition === 'trailing') {
        runs = `${textRunXml(`${left} ${right}`, '')}<w:r><w:tab/></w:r>`;
        original = `${left} ${right}\t`;
    }

    const replacement = randomWord(rng);
    const modified = tabPosition === 'leading'
        ? `\t${left} ${right} ${replacement}`
        : tabPosition === 'middle'
            ? `${left}\t${right} ${replacement}`
            : tabPosition === 'trailing'
                ? `${left} ${right} ${replacement}\t`
                : `${left}${separator}${right} ${replacement}`;

    return {
        shape: 'whitespace',
        oxml: wrapBody(`<w:p>${runs}</w:p>`),
        original,
        modified: ensureChanged(rng, original, modified)
    };
}

/* --- shape: manually numbered heading expanded into a list --------------- */

function shapeSuppressedHeadingList(rng) {
    const label = `${rng.pick(['A', 'B', 'C', 'D'])}.`;
    const heading = randomWord(rng).toUpperCase();
    const font = rng.pick(['Times New Roman', 'Arial', 'Calibri']);
    const size = rng.pick(['20', '22', '24']);
    const itemCount = 2 + rng.int(3);
    const marker = rng.pick(['*', '-']);
    const original = `${label}\t${heading}`;
    const items = Array.from({ length: itemCount }, () => randomText(rng, 3, 7));
    const modified = items.map(item => `${marker} ${item}`).join('\n');
    const typography = `<w:rFonts w:ascii="${font}" w:hAnsi="${font}"/><w:sz w:val="${size}"/>`;
    const paragraphXml = `<w:p>`
        + `<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="0"/></w:numPr>`
        + `<w:rPr>${typography}<w:b/><w:u w:val="single"/></w:rPr></w:pPr>`
        + `<w:r><w:rPr>${typography}<w:b/></w:rPr><w:t>${label}</w:t></w:r>`
        + '<w:r><w:tab/></w:r>'
        + `<w:r><w:rPr>${typography}<w:b/><w:u w:val="single"/></w:rPr><w:t>${escapeXmlText(heading)}</w:t></w:r>`
        + '</w:p>';

    return {
        shape: 'suppressedHeadingList',
        oxml: wrapBody(paragraphXml),
        original,
        modified,
        // Markdown marker removal and the package fragment's terminal sentinel
        // make normalized fidelity appropriate here; the fixed regression pins
        // the exact physical Accept/Reject paragraph sequence separately.
        options: { fidelity: 'normalized', expectedAcceptedText: items.join('\n') },
        expectedListItems: itemCount,
        expectedFont: font,
        expectedSize: size
    };
}

/* --- shape: complex field with editable surrounding text ----------------- */

function shapeComplexField(rng) {
    const prefix = randomText(rng, 2, 4);
    const suffix = randomText(rng, 2, 4);
    const result = `${1 + rng.int(20)}.${1 + rng.int(9)}`;
    const fieldId = 1000 + rng.int(9000);
    const instruction = ` REF FuzzBookmark${fieldId} \\h `;
    const paragraphXml = `<w:p>`
        + textRunXml(`${prefix} `, '')
        + '<w:r><w:fldChar w:fldCharType="begin" w:fldLock="true"/></w:r>'
        + `<w:r><w:instrText xml:space="preserve">${escapeXmlText(instruction)}</w:instrText></w:r>`
        + '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
        + textRunXml(result, '')
        + '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
        + textRunXml(` ${suffix}`, '')
        + '</w:p>';
    const original = `${prefix} ${result} ${suffix}`;
    const modified = `${prefix} ${result} ${suffix} ${randomWord(rng)}`;

    return {
        shape: 'complexField',
        oxml: wrapBody(paragraphXml),
        original,
        modified,
        requiredFieldInstruction: instruction
    };
}

/* --- shape: pre-existing revisions from another author -------------------- */

function shapeExistingRevisions(rng) {
    const keep = randomText(rng, 2, 5);
    const inserted = randomText(rng, 1, 3);
    const deleted = randomText(rng, 1, 3);

    const paragraphXml = `<w:p>`
        + textRunXml(`${keep} `, '')
        + `<w:ins w:id="9001" w:author="Prior" w:date="2020-01-01T00:00:00Z">${textRunXml(inserted, '')}</w:ins>`
        + `<w:del w:id="9002" w:author="Prior" w:date="2020-01-01T00:00:00Z"><w:r><w:delText>${escapeXmlText(deleted)}</w:delText></w:r></w:del>`
        + `</w:p>`;

    // After accept-all-first the visible text is keep + inserted; the w:del is gone.
    const original = `${keep} ${inserted}`;
    const modified = `${keep} ${inserted} ${randomWord(rng)}`;

    return {
        shape: 'existingRevisions',
        oxml: wrapBody(paragraphXml),
        original,
        modified,
        options: { existingRevisions: 'accept-all-first' }
    };
}

function shapeExistingRevisionsNoOp(rng) {
    const keep = randomText(rng, 2, 5);
    const inserted = randomText(rng, 1, 3);
    const paragraphXml = `<w:p>`
        + textRunXml(`${keep} `, '')
        + `<w:ins w:id="9101" w:author="Prior" w:date="2020-01-01T00:00:00Z">${textRunXml(inserted, '')}</w:ins>`
        + `</w:p>`;
    const visible = `${keep} ${inserted}`;
    return {
        shape: 'existingRevisionsNoOp',
        oxml: wrapBody(paragraphXml),
        original: visible,
        modified: visible,
        options: { existingRevisions: 'accept-all-first' }
    };
}

const SHAPES = [
    shapeParagraph,
    shapeParagraph,
    shapeMultiParagraph,
    shapeTableCell,
    shapeWhitespace,
    shapeSuppressedHeadingList,
    shapeComplexField,
    shapeExistingRevisions,
    shapeExistingRevisionsNoOp
];

function generateCase(caseSeed) {
    const rng = makeRng(caseSeed);
    const shape = SHAPES[caseSeed % SHAPES.length];
    return shape(rng);
}

/*
 * Known gaps: failures that Phase 1 discovered and a later phase owns. They are
 * counted and reported every run so they cannot be forgotten, but they do not
 * fail the build. Delete an entry the moment its phase lands -- if the matching
 * failures stop appearing, the run reports that the entry is stale.
 */
const KNOWN_GAPS = [];

function classifyFailure(error) {
    return KNOWN_GAPS.find(gap => gap.matches(error)) || null;
}

let failures = 0;
const shapeCounts = new Map();
const knownGapCounts = new Map();

for (let i = 0; i < ITERATIONS; i++) {
    const caseSeed = (BASE_SEED + i) >>> 0;
    const testCase = generateCase(caseSeed);
    shapeCounts.set(testCase.shape, (shapeCounts.get(testCase.shape) || 0) + 1);

    try {
        const roundTrip = await assertRoundTrip(testCase.oxml, testCase.original, testCase.modified, testCase.options || {});
        if (testCase.shape === 'complexField') {
            const parsed = parseXmlFragment(roundTrip.redlined.oxml);
            const fieldChars = elementsByLocalName(parsed, 'fldChar');
            const instructions = elementsByLocalName(parsed, 'instrText');
            assert.equal(fieldChars.length, 3, 'complex field scaffolding must survive fuzz edit');
            assert.deepEqual(instructions.map(node => node.textContent), [testCase.requiredFieldInstruction]);
            for (const node of [...fieldChars, ...instructions]) {
                assert.equal(node.parentNode?.localName, 'r', `${node.localName} must remain inside w:r`);
            }
        }
        if (testCase.shape === 'existingRevisionsNoOp') {
            assert.equal(roundTrip.redlined.hasChanges, false);
            assert.equal(roundTrip.redlined.oxml, testCase.oxml);
        }
        if (testCase.shape === 'suppressedHeadingList') {
            const parsed = parseXmlFragment(roundTrip.redlined.oxml);
            const insertedParagraphs = elementsByLocalName(parsed, 'p').filter(paragraph => (
                elementsByLocalName(paragraph, 'ins').some(ins => ins.parentNode === paragraph)
            ));
            assert.equal(insertedParagraphs.length, testCase.expectedListItems);
            for (const paragraph of insertedParagraphs) {
                const numId = elementsByLocalName(paragraph, 'numId')[0];
                assert.ok(Number.parseInt(numId?.getAttribute('w:val'), 10) > 0);
                const textInsertion = elementsByLocalName(paragraph, 'ins').find(ins => ins.parentNode === paragraph);
                const insertedRPr = elementsByLocalName(textInsertion, 'rPr')[0];
                assert.equal(elementsByLocalName(insertedRPr, 'rFonts')[0]?.getAttribute('w:ascii'), testCase.expectedFont);
                assert.equal(elementsByLocalName(insertedRPr, 'sz')[0]?.getAttribute('w:val'), testCase.expectedSize);
                assert.equal(elementsByLocalName(insertedRPr, 'b').length, 0);
                assert.equal(elementsByLocalName(insertedRPr, 'u').length, 0);
            }
        }
    } catch (error) {
        const knownGap = classifyFailure(error);
        if (knownGap) {
            knownGapCounts.set(knownGap.id, (knownGapCounts.get(knownGap.id) || 0) + 1);
            continue;
        }

        failures++;
        console.error(`\nFUZZ CASE FAILED (seed ${caseSeed}, shape ${testCase.shape})`);
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

// One deterministic corpus member crosses the old 65,536-token boundary on
// every run without making the ordinary randomized shapes prohibitively large.
try {
    const uniqueTokens = Array.from({ length: 70000 }, (_, index) => `u${BASE_SEED}_${index}`);
    const original = uniqueTokens.join(' ');
    uniqueTokens[35000] = `changed_${BASE_SEED}`;
    const modified = uniqueTokens.join(' ');
    const diffs = computeWordDiffs(original, modified);
    assert.equal(diffs.filter(([op]) => op !== 1).map(([, text]) => text).join(''), original);
    assert.equal(diffs.filter(([op]) => op !== -1).map(([, text]) => text).join(''), modified);
    shapeCounts.set('highUniqueToken', 1);
} catch (error) {
    failures++;
    console.error(`\nFUZZ CASE FAILED (seed ${BASE_SEED}, shape highUniqueToken)`);
    console.error(`  ${error?.message || error}`);
}

if (failures > 0) {
    console.error(`\n${failures} fuzz case(s) failed out of ${ITERATIONS} (base seed ${BASE_SEED}).`);
    process.exit(1);
}

const shapeSummary = [...shapeCounts.entries()].map(([name, count]) => `${name}=${count}`).join(' ');
console.log(`PASS: ${ITERATIONS} fuzz round-trip cases (base seed ${BASE_SEED}) [${shapeSummary}]`);

for (const gap of KNOWN_GAPS) {
    const count = knownGapCounts.get(gap.id) || 0;
    if (count > 0) {
        console.warn(`  KNOWN-GAP ${gap.id} (${gap.phase}): ${count} case(s) -- ${gap.note}`);
    } else {
        console.warn(`  KNOWN-GAP ${gap.id} (${gap.phase}): 0 cases -- possibly fixed; re-check and remove this entry`);
    }
}
