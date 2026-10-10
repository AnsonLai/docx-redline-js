import './setup-xml-provider.mjs';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { parseOoxmlSafe } from '../adapters/xml-adapter.js';
import { projectWordRevisionView } from '../core/word-revision-semantics.js';
import { rejectTrackedChangesInOoxml, acceptTrackedChangesInOoxml } from '../index.js';

/**
 * Pins the revision simulator (core/word-revision-semantics.js) and the library's
 * reject/accept against what desktop Word actually did on Reject All / Accept All.
 * Expectations are recorded by scripts/word-revision-semantics-matrix.ps1 (needs Word);
 * this suite needs no Word.
 *
 * KNOWN_DIVERGENCES documents where we currently differ from Word. The suite fails if a
 * listed divergence disappears (delete the entry) or an unlisted one appears.
 */

const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const dir = join(dirname(fileURLToPath(import.meta.url)), 'fixtures', 'word-revision-semantics');
const expectations = JSON.parse(readFileSync(join(dir, 'expectations.json'), 'utf8'));

const KNOWN_DIVERGENCES = {
    'ins-mark-last-prev-deleted': { simulator: 'Word removes the last body paragraph when its inserted mark is rejected (previous mark becomes final); simulator keeps an empty paragraph' },
    'ins-mark-last-prev-plain': { simulator: 'Word removes the last body paragraph when its inserted mark is rejected (previous mark becomes final); simulator keeps an empty paragraph' },
    'ins-mark-last-plain-style': { simulator: 'Word removes the last body paragraph when its inserted mark is rejected (previous mark becomes final); simulator keeps an empty paragraph' },
    'del-mark-last': { simulator: 'Word removes the last body paragraph when its deleted mark is accepted; simulator keeps an empty paragraph' },
    'del-mark-last-list-prev-plain': { simulator: 'Word removes the last body paragraph when its deleted mark is accepted; simulator keeps an empty paragraph' },
    'ins-mark-cell-last-of-two': { library: 'Word keeps an empty cell paragraph when the last of two cell paragraphs has its inserted mark rejected; library drops it' },
    'del-mark-cell-sole': {
        simulator: 'Word removes the whole table when its only cell paragraph is fully deleted on Accept; simulator keeps an empty cell paragraph',
        library: 'Word removes the whole table when its only cell paragraph is fully deleted on Accept; library keeps an empty cell paragraph'
    },
    'del-mark-cell-last-of-two': { simulator: 'Word collapses the cell (no table left) when the last cell paragraph mark is deleted on Accept; simulator keeps an empty cell paragraph' }
};

const wordParagraphs = list => (Array.isArray(list) ? list : [list]).map(p => ({
    text: String(p.text).replace(/\f/g, ''),
    list: Number(p.listType) !== 0
}));

const simulatorParagraphs = (doc, view) => projectWordRevisionView(doc, view).map(r => ({
    text: r.text,
    list: Boolean(r.numPr)
}));

function libraryParagraphs(oxml) {
    const parsed = parseOoxmlSafe(oxml);
    const xmlDoc = parsed.doc;
    return Array.from(xmlDoc.getElementsByTagNameNS(NS_W, 'p')).map(p => {
        let text = '';
        const walk = node => {
            for (const child of Array.from(node.childNodes || [])) {
                if (child.nodeType !== 1) continue;
                const name = child.localName;
                if (name === 'pPr') continue;
                if (name === 't') text += child.textContent;
                else if (name === 'tab') text += '\t';
                else if (name === 'br' || name === 'cr') text += '\n';
                else walk(child);
            }
        };
        walk(p);
        const pPr = Array.from(p.childNodes).find(c => c.nodeType === 1 && c.localName === 'pPr');
        const list = Boolean(pPr && Array.from(pPr.childNodes).some(c => c.nodeType === 1 && c.localName === 'numPr'));
        return { text, list };
    });
}

const show = list => list.map(p => `${JSON.stringify(p.text)}${p.list ? '[list]' : ''}`).join(' | ');
const equal = (a, b) => JSON.stringify(a) === JSON.stringify(b);

const failures = [];
const seenDivergences = new Set();
let checks = 0;

for (const testCase of expectations.cases) {
    const docXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="${NS_W}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>${testCase.bodyXml}</w:body></w:document>`;
    const parsed = parseOoxmlSafe(docXml);
    const doc = parsed.doc;
    const known = KNOWN_DIVERGENCES[testCase.name] || {};
    const diffs = { simulator: [], library: [] };

    for (const [view, wordList] of [['rejected', testCase.reject], ['accepted', testCase.accept]]) {
        const word = wordParagraphs(wordList);
        const sim = simulatorParagraphs(doc, view);
        const lib = libraryParagraphs((view === 'rejected'
            ? rejectTrackedChangesInOoxml(docXml, { allAuthors: true })
            : acceptTrackedChangesInOoxml(docXml, { allAuthors: true })).oxml);
        checks += 2;
        if (!equal(word, sim)) diffs.simulator.push(`${view}: word=${show(word)} ours=${show(sim)}`);
        if (!equal(word, lib)) diffs.library.push(`${view}: word=${show(word)} ours=${show(lib)}`);
    }

    for (const subject of ['simulator', 'library']) {
        const diverges = diffs[subject].length > 0;
        const listed = Boolean(known[subject]);
        if (diverges) seenDivergences.add(`${testCase.name}/${subject}`);
        if (diverges && !listed) failures.push(`${testCase.name}: ${subject} diverges from Word (unlisted)\n    ${diffs[subject].join('\n    ')}`);
        if (!diverges && listed) failures.push(`${testCase.name}: ${subject} now matches Word; remove it from KNOWN_DIVERGENCES`);
    }
}

for (const name of Object.keys(KNOWN_DIVERGENCES)) {
    if (!expectations.cases.some(c => c.name === name)) failures.push(`KNOWN_DIVERGENCES names unknown case '${name}'`);
}

if (process.env.WORD_SEMANTICS_REPORT) console.log([...seenDivergences].join('\n'));

if (failures.length) {
    console.error(`❌ FAIL word revision semantics (${failures.length}):\n${failures.join('\n')}`);
    process.exit(1);
}
console.log(`✅ word revision semantics: ${expectations.cases.length} cases, ${checks} comparisons against Word ${expectations.wordBuild}`);
