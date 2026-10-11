/**
 * Pure, seeded generators of valid document operations for the reject-roundtrip matrix (WP-3).
 *
 * Nothing here touches the library under test: input is the list of inspected paragraphs
 * (`openDocx(bytes).inspect().paragraphs`) and a seeded RNG; output is canonical operation JSON plus
 * the text each operation intends to leave in the accepted view.
 */

export function mulberry32(seed) {
    let a = seed >>> 0;
    return function next() {
        a |= 0;
        a = (a + 0x6D2B79F5) | 0;
        let t = Math.imul(a ^ (a >>> 15), 1 | a);
        t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
        return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
    };
}

export function makeRng(seed) {
    const next = mulberry32(seed);
    return {
        next,
        int: max => Math.floor(next() * max),
        pick: arr => arr[Math.floor(next() * arr.length)],
        chance: p => next() < p
    };
}

/** Same convention as tests/roundtrip_fuzz_tests.mjs: FUZZ_SEED env, else the fixed default. */
export function resolveBaseSeed(defaultSeed, env = process.env) {
    return (Number(env.FUZZ_SEED) >>> 0) || defaultSeed;
}

const NEW_WORDS = ['amended', 'revised', 'superseding', 'updated', 'binding', 'mutual', 'prompt'];
const NEW_SENTENCES = [
    'The parties agree to the revised terms set out here.',
    'Payment is due within thirty days of the invoice date.',
    'Either party may terminate on written notice.'
];
const NEW_LINES = ['Replacement line alpha', 'Replacement line beta', 'Replacement line gamma', 'Replacement line delta'];
const FORMAT_PROPERTIES = [{ bold: true }, { italic: true }, { underline: true }, { bold: true, italic: true }];
const PARAGRAPH_FORMAT_PROPERTIES = [{ alignment: 'center' }, { alignment: 'right' }, { alignment: 'justify' }];

export const OPERATION_KINDS = [
    'replace-word',
    'replace-sentence',
    'replace-paragraph',
    'range-replace',
    'range-replace-empty',
    'insert-after',
    'delete',
    'replacements',
    'format',
    'paragraph-format',
    'paragraph-format-numbering',
    'plain-to-list',
    'list-format-convert-plain',
    'last-paragraph-replace'
];

// Typed header markers, read from inspected text with a regex of our own (not the library's grammar).
const TYPED_MARKER = /^(?:(\d{1,3})|([A-Za-z]))([.)])[ \t]+(\S[\s\S]*)$|^([-•])[ \t]+(\S[\s\S]*)$/;
const ROMAN_LETTERS = new Set(['I', 'V', 'X', 'L', 'C', 'D', 'M', 'i', 'v', 'x', 'l', 'c', 'd', 'm']);
function typedMarker(text) {
    const match = TYPED_MARKER.exec(String(text || ''));
    if (!match) return null;
    if (match[5]) return { family: 'bullet', rest: match[6] };
    if (match[1]) return { family: `decimal${match[3]}`, rest: match[4] };
    const letter = match[2];
    const upper = letter === letter.toUpperCase();
    return {
        family: `${upper ? 'upper' : 'lower'}Letter${match[3]}`,
        rest: match[4],
        // A lone roman-looking letter is ambiguous by design; state the format so the case exercises a conversion.
        format: ROMAN_LETTERS.has(letter) ? (upper ? 'upperLetter' : 'lowerLetter') : null
    };
}

const targetOf = paragraph => ({
    paragraphId: paragraph.paragraphId,
    fingerprint: paragraph.fingerprint,
    exactText: paragraph.exactText
});

const WORD_PATTERN = /[A-Za-z]{4,}/g;
const wordMatches = text => Array.from(String(text).matchAll(WORD_PATTERN));
const occurrenceOf = (matches, match) => matches.filter(m => m[0] === match[0] && m.index <= match.index).length;
const countOf = (matches, match) => matches.filter(m => m[0] === match[0]).length;

/** Replaces one rng-chosen word of `text`; returns the new text plus the replacement word. */
function swapOneWord(rng, text) {
    const matches = wordMatches(text);
    if (matches.length === 0) return null;
    const match = rng.pick(matches);
    const replacement = rng.pick(NEW_WORDS);
    if (replacement === match[0]) return null;
    return {
        replacement,
        modified: text.slice(0, match.index) + replacement + text.slice(match.index + match[0].length)
    };
}

function sentenceSplit(text) {
    const match = /^(.*?[.!?])(\s+)(\S[\s\S]*)$/.exec(text);
    return match ? { first: match[1], rest: match[3], gap: match[2] } : null;
}

/** Picks an inclusive [start, end] paragraph range (same container, at most `maxSpan` apart), or null. */
function pickRange(rng, paragraphs, maxSpan) {
    const pairs = [];
    for (let i = 0; i < paragraphs.length; i++) {
        for (let j = i + 1; j <= Math.min(paragraphs.length - 1, i + maxSpan); j++) {
            const a = paragraphs[i];
            const b = paragraphs[j];
            if (Boolean(a.inTable) !== Boolean(b.inTable)) continue;
            if (a.inTable && JSON.stringify(a.table) !== JSON.stringify(b.table)) continue;
            pairs.push([i, j]);
        }
    }
    if (pairs.length === 0) return null;
    const [startIndex, endIndex] = rng.pick(pairs);
    return { start: paragraphs[startIndex], end: paragraphs[endIndex], startIndex, endIndex };
}

const linesFor = (rng, count) => Array.from({ length: count }, (_, i) => `${rng.pick(NEW_LINES)} ${i + 1}`);
const listItemLines = (rng, count) => Array.from({ length: count }, (_, i) => `${rng.pick(['alpha', 'beta', 'gamma', 'delta'])} item ${i + 1}`);

/**
 * Builds one operation of `kind`, or null when the fixture offers no valid site for it.
 *
 * @returns {null | { kind: string, operation: object, expectAccepted: string[], paragraphIndex: number }}
 */
export function generateOperation(rng, paragraphs, kind, serial = 0) {
    const n = paragraphs.length;
    if (n === 0) return null;
    const indexed = paragraphs.map((q, i) => ({ q, i }));
    const withText = indexed.filter(({ q }) => q.exactText && q.exactText.trim());
    const pickWithText = () => (withText.length ? rng.pick(withText) : null);
    const operationId = `matrix-${kind}-${serial}`;
    const base = { operationId, type: 'redline' };

    switch (kind) {
        case 'replace-word': {
            const site = pickWithText();
            const swap = site && swapOneWord(rng, site.q.exactText);
            if (!swap) return null;
            return { kind, paragraphIndex: site.i, expectAccepted: [swap.replacement],
                operation: { ...base, target: targetOf(site.q), modified: swap.modified } };
        }
        case 'replace-sentence': {
            const site = pickWithText();
            if (!site) return null;
            const split = sentenceSplit(site.q.exactText);
            const sentence = rng.pick(NEW_SENTENCES);
            const modified = split ? `${sentence}${split.gap}${split.rest}` : sentence;
            return { kind, paragraphIndex: site.i, expectAccepted: [sentence],
                operation: { ...base, target: targetOf(site.q), modified } };
        }
        case 'replace-paragraph': {
            const site = pickWithText();
            if (!site) return null;
            const modified = `${rng.pick(NEW_SENTENCES)} (${serial})`;
            return { kind, paragraphIndex: site.i, expectAccepted: [modified],
                operation: { ...base, target: targetOf(site.q), modified } };
        }
        case 'range-replace': {
            const range = pickRange(rng, paragraphs, 3);
            if (!range) return null;
            const lines = linesFor(rng, 1 + rng.int(3));
            return { kind, paragraphIndex: range.startIndex, expectAccepted: lines,
                operation: { ...base, target: targetOf(range.start), targetEnd: targetOf(range.end), modified: lines.join('\n') } };
        }
        case 'range-replace-empty': {
            // Ranges made only of empty paragraphs (empty list items, spacers).
            const emptyRuns = [];
            for (let i = 0; i < n - 1; i++) {
                if (!paragraphs[i].exactText && !paragraphs[i + 1].exactText) {
                    let j = i + 1;
                    while (j + 1 < n && !paragraphs[j + 1].exactText) j++;
                    emptyRuns.push([i, j]);
                    i = j;
                }
            }
            if (emptyRuns.length === 0) return null;
            const [startIndex, endIndex] = rng.pick(emptyRuns);
            const lines = linesFor(rng, 1 + rng.int(2));
            return { kind, paragraphIndex: startIndex, expectAccepted: lines,
                operation: { ...base, target: targetOf(paragraphs[startIndex]), targetEnd: targetOf(paragraphs[endIndex]), modified: lines.join('\n') } };
        }
        case 'insert-after': {
            const site = pickWithText();
            if (!site) return null;
            const added = `${rng.pick(NEW_SENTENCES)} [${serial}]`;
            return { kind, paragraphIndex: site.i, expectAccepted: [added],
                operation: { ...base, target: targetOf(site.q), modified: `${site.q.exactText}\n${added}` } };
        }
        case 'delete': {
            const site = pickWithText();
            if (!site) return null;
            return { kind, paragraphIndex: site.i, expectAccepted: [],
                operation: { operationId, type: 'delete', target: targetOf(site.q), modified: '' } };
        }
        case 'replacements': {
            const site = pickWithText();
            const matches = site ? wordMatches(site.q.exactText) : [];
            if (matches.length === 0) return null;
            const match = rng.pick(matches);
            const replace = rng.pick(NEW_WORDS);
            if (replace === match[0]) return null;
            const occurrence = occurrenceOf(matches, match);
            return { kind, paragraphIndex: site.i, expectAccepted: [replace],
                operation: { ...base, target: targetOf(site.q),
                    replacements: [{ find: match[0], replace, ...(countOf(matches, match) > 1 ? { occurrence } : {}) }] } };
        }
        case 'format': {
            const site = pickWithText();
            const matches = site ? wordMatches(site.q.exactText) : [];
            if (matches.length === 0) return null;
            const match = rng.pick(matches);
            return { kind, paragraphIndex: site.i, expectAccepted: [],
                operation: { operationId, type: 'format', target: targetOf(site.q), textToFormat: match[0],
                    ...(countOf(matches, match) > 1 ? { textOccurrence: occurrenceOf(matches, match) } : {}),
                    properties: rng.pick(FORMAT_PROPERTIES) } };
        }
        case 'paragraph-format': {
            const site = rng.pick(indexed);
            return { kind, paragraphIndex: site.i, expectAccepted: [],
                operation: { operationId, type: 'paragraph-format', target: targetOf(site.q), properties: rng.pick(PARAGRAPH_FORMAT_PROPERTIES) } };
        }
        case 'paragraph-format-numbering': {
            // Binds a paragraph to a list the document already defines (numId/level read from inspect), rebinds a list
            // item to another existing list/level, or removes a list item's direct binding.
            const listed = indexed.filter(({ q }) => q.list && q.list.numId != null);
            if (listed.length === 0) return null;
            const source = rng.pick(listed).q.list;
            const level = Number.isInteger(Number(source.level)) ? Number(source.level) : 0;
            const site = rng.pick(indexed);
            const remove = site.q.list && rng.chance(0.4);
            return { kind, paragraphIndex: site.i, expectAccepted: [],
                operation: { operationId, type: 'paragraph-format', target: targetOf(site.q),
                    properties: { numbering: remove ? null : { numId: Number(source.numId), level } } } };
        }
        case 'plain-to-list': {
            const plain = withText.filter(({ q }) => !q.list && !q.inTable);
            if (plain.length === 0) return null;
            const site = rng.pick(plain);
            const items = listItemLines(rng, 2 + rng.int(2));
            const ordered = rng.chance(0.5);
            return { kind, paragraphIndex: site.i, expectAccepted: items,
                operation: { ...base, target: targetOf(site.q),
                    modified: items.map((item, i) => (ordered ? `${i + 1}. ${item}` : `- ${item}`)).join('\n') } };
        }
        case 'list-format-convert-plain': {
            // Plain paragraphs whose text starts with a typed marker; sometimes a run of same-family headers.
            const typed = indexed
                .filter(({ q }) => !q.list && q.exactText)
                .map(entry => ({ ...entry, marker: typedMarker(entry.q.exactText) }))
                .filter(entry => entry.marker);
            if (typed.length === 0) return null;
            const start = rng.pick(typed);
            let end = start;
            if (rng.chance(0.4)) {
                const later = typed.filter(entry => entry.i > start.i && entry.i - start.i <= 4
                    && entry.marker.family === start.marker.family
                    && Boolean(entry.q.inTable) === Boolean(start.q.inTable)
                    && JSON.stringify(entry.q.table || null) === JSON.stringify(start.q.table || null));
                if (later.length > 0) end = rng.pick(later);
            }
            const span = typed.filter(entry => entry.i >= start.i && entry.i <= end.i);
            const list = { mode: 'convert-plain', ...(start.marker.format ? { format: start.marker.format } : {}) };
            return { kind, paragraphIndex: start.i, expectAccepted: span.map(entry => entry.marker.rest),
                operation: { operationId, type: 'list-format', target: targetOf(start.q),
                    ...(end !== start ? { targetEnd: targetOf(end.q) } : {}), list } };
        }
        case 'last-paragraph-replace': {
            let index = n - 1;
            while (index > 0 && !(paragraphs[index].exactText || '').trim()) index--;
            if (!(paragraphs[index].exactText || '').trim()) return null;
            const modified = `${rng.pick(NEW_SENTENCES)} (last ${serial})`;
            return { kind, paragraphIndex: index, expectAccepted: [modified],
                operation: { ...base, target: targetOf(paragraphs[index]), modified } };
        }
        default:
            throw new Error(`Unknown operation kind '${kind}'.`);
    }
}

/**
 * Enumerates `count` operations for a fixture, cycling through every kind so each is exercised, with
 * rng-drawn sites and parameters. Kinds without a valid site on this fixture are reported, not hidden.
 *
 * @returns {{ cases: object[], unavailable: string[] }}
 */
export function generateOperationsForFixture(rng, paragraphs, count) {
    const cases = [];
    const unavailable = new Set();
    // Kinds with no valid site on this fixture are skipped (and reported); the cycle continues so the
    // requested number of cases is still produced.
    for (let i = 0; cases.length < count && i < count * 3; i++) {
        const kind = OPERATION_KINDS[i % OPERATION_KINDS.length];
        let generated = null;
        for (let attempt = 0; attempt < 8 && !generated; attempt++) {
            generated = generateOperation(rng, paragraphs, kind, i);
        }
        if (generated) cases.push(generated);
        else unavailable.add(kind);
    }
    return { cases, unavailable: [...unavailable] };
}
