import assert from 'node:assert/strict';

import './setup-xml-provider.mjs';
import { acceptTrackedChangesInOoxml, rejectTrackedChangesInOoxml } from '../index.js';
import { parseOoxmlSafe } from '../adapters/xml-adapter.js';
import { applyOperationsToDocumentXml } from '../services/standalone-operation-runner.js';

// Regression for appending a Markdown table after a paragraph (library-only form of the consumer
// "replace final paragraph with retained text + table" model):
//   WP-5a  a same-author tracked underline on the paragraph must survive a later table append
//   WP-5b  a Markdown underline sent together with the table must be applied
//   WP-6   every paragraph mark the append creates is a tracked-inserted mark, so Word's
//          Reject All restores exactly the source paragraphs (no extra empty paragraph)

const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const AUTHOR = 'Table Append Fidelity';
const OPTIONS = {
    atomic: true,
    structuredContent: true,
    pairReplacements: true,
    generateRedlines: true,
    existingRevisions: 'merge-same-author'
};
const SOURCE = [
    'Opening paragraph 1.', 'Opening paragraph 2.', 'Opening paragraph 3.',
    'Opening paragraph 4.', 'Opening paragraph 5.', 'Opening paragraph 6.',
    'Final paragraph stays unchanged.'
];
const FINAL = SOURCE[6];
const TABLE = '| Mountain | River | Forest |\n| --- | --- | --- |\n| Ocean | Valley | Canyon |\n| Meadow | Desert | Island |';
const CELLS = ['Mountain', 'River', 'Forest', 'Ocean', 'Valley', 'Canyon', 'Meadow', 'Desert', 'Island'];
const SOURCE_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="${NS_W}"><w:body>`
    + SOURCE.map(text => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`).join('')
    + '<w:sectPr/></w:body></w:document>';

const elements = (node, name) => Array.from(node.getElementsByTagNameNS('*', name));
const bodyOf = xml => {
    const parsed = parseOoxmlSafe(xml);
    assert.ok(parsed.doc, parsed.error?.message || 'parseable OOXML');
    return elements(parsed.doc, 'body')[0];
};
const blocks = xml => Array.from(bodyOf(xml).childNodes).filter(node => node.nodeType === 1);
const paragraphsOf = xml => blocks(xml).filter(node => node.localName === 'p');
const inside = (node, name, root) => {
    for (let cur = node.parentNode; cur && cur !== root; cur = cur.parentNode) if (cur.localName === name) return true;
    return false;
};
const markType = paragraph => {
    for (const type of ['ins', 'del']) {
        if (elements(paragraph, type).some(n => n.parentNode?.localName === 'rPr' && n.parentNode.parentNode?.localName === 'pPr'
            && n.parentNode.parentNode.parentNode === paragraph)) return type;
    }
    return null;
};
const rejectedFragment = p => elements(p, 't').concat(elements(p, 'delText'))
    .filter(n => !inside(n, 'ins', p)).map(n => n.textContent).join('');
const acceptedFragment = p => elements(p, 't').filter(n => !inside(n, 'del', p)).map(n => n.textContent).join('');

// Word semantics for body-level paragraphs: a rejected inserted mark (accepted deleted mark) merges the
// paragraph into the following one; text inside w:ins is dropped on Reject. A trailing carry with no
// following paragraph is the removed final paragraph.
function wordParagraphs(xml, action) {
    const out = [];
    let carry = '';
    for (const block of blocks(xml)) {
        if (block.localName !== 'p') continue;
        const mark = markType(block);
        const text = carry + (action === 'reject' ? rejectedFragment(block) : acceptedFragment(block));
        const markGoes = action === 'reject' ? mark === 'ins' : mark === 'del';
        if (markGoes) { carry = text; continue; }
        carry = '';
        out.push(text);
    }
    if (carry) out.push(carry);
    return out;
}
const libraryParagraphs = (xml, fn) => paragraphsOf(fn(xml, { allAuthors: true }).oxml)
    .map(p => elements(p, 't').map(n => n.textContent).join(''));
const resolve = (xml, fn) => fn(xml, { allAuthors: true }).oxml;
const hasUnderline = xml => /<w:u\b/.test(xml.replace(/<w:rPrChange\b[\s\S]*?<\/w:rPrChange>/g, ''));
const cellsOf = xml => elements(bodyOf(xml), 'tc').map(tc => elements(tc, 't').map(n => n.textContent).join(''));
const underlinedText = xml => paragraphsOf(xml)
    .filter(p => elements(p, 'r').some(r => elements(r, 'u').some(u => u.parentNode?.localName === 'rPr' && u.parentNode.parentNode === r)))
    .map(p => elements(p, 't').map(n => n.textContent).join(''));

async function run(documentXml, operation) {
    const result = await applyOperationsToDocumentXml(documentXml, [{
        type: 'redline', structuredContent: true, ...operation
    }], AUTHOR, {}, OPTIONS);
    assert.ok(result.documentXml && result.hasChanges, JSON.stringify(result.results || result.error || result));
    return result.documentXml;
}

// The structural invariant: every body paragraph mark that is not one of the source's marks is a
// tracked-inserted mark, so treating inserted marks as Word does leaves exactly the source marks.
function assertParagraphMarkAccounting(xml, label) {
    assertWordNativeInsertedTable(xml, label);
    const paragraphs = paragraphsOf(xml);
    const untracked = paragraphs.filter(p => markType(p) !== 'ins');
    assert.equal(untracked.length, SOURCE.length, `${label}: untracked paragraph marks must equal the source paragraph count`);
    assert.deepEqual(untracked.map(p => elements(p, 't').filter(n => !inside(n, 'ins', p)).map(n => n.textContent).join('')),
        SOURCE, `${label}: untracked marks carry the source paragraphs`);
    // A table may not be the last body block: Word would add an untracked paragraph after it.
    const last = blocks(xml).filter(n => n.localName !== 'sectPr').at(-1);
    assert.equal(last.localName, 'p', `${label}: a paragraph must follow a final table`);
    assert.equal(markType(last), 'ins', `${label}: the paragraph after a final table must be a tracked-inserted mark`);
}

// Word only treats a table as inserted when each row carries trPr/ins (w:ins around w:tbl is ignored).
function assertWordNativeInsertedTable(xml, label) {
    const body = bodyOf(xml);
    for (const table of elements(body, 'tbl')) {
        assert.notEqual(table.parentNode.localName, 'ins', `${label}: w:ins must not wrap w:tbl`);
        const rows = Array.from(table.childNodes).filter(n => n.localName === 'tr');
        assert.equal(rows.length, 3, `${label}: row count`);
        for (const row of rows) {
            const trPr = Array.from(row.childNodes).find(n => n.localName === 'trPr');
            assert.ok(trPr && Array.from(trPr.childNodes).some(n => n.localName === 'ins'), `${label}: row carries trPr/ins`);
            for (const p of elements(row, 'p')) assert.equal(markType(p), 'ins', `${label}: cell paragraph mark is tracked-inserted`);
        }
    }
}

function assertRejectRestoresSource(xml, label) {
    assert.deepEqual(wordParagraphs(xml, 'reject'), SOURCE, `${label}: Word-semantics Reject All`);
    assert.deepEqual(libraryParagraphs(xml, rejectTrackedChangesInOoxml), SOURCE, `${label}: library Reject All`);
    const rejected = resolve(xml, rejectTrackedChangesInOoxml);
    assert.equal(/<w:tbl\b/.test(rejected), false, `${label}: Reject All removes the table`);
    assert.equal(hasUnderline(rejected), false, `${label}: Reject All removes the underline`);
}

function assertAcceptKeepsIntent(xml, label) {
    const accepted = resolve(xml, acceptTrackedChangesInOoxml);
    assert.deepEqual(cellsOf(accepted), CELLS, `${label}: table cells`);
    assert.equal(blocks(accepted).filter(n => n.localName === 'tbl').length, 1, `${label}: one table`);
    assert.deepEqual(underlinedText(accepted), [FINAL], `${label}: P7 underlined after Accept All`);
    const wordAccepted = wordParagraphs(xml, 'accept');
    assert.deepEqual(wordAccepted.slice(0, 7), SOURCE, `${label}: accepted text`);
}

// WP-5a: turn 1 underlines P7; turn 2 appends the table on the tracked turn-1 output.
{
    const turn1 = await run(SOURCE_XML, {
        target: { index: 7, exactText: FINAL }, targetRef: 'P7',
        replacements: [{ find: FINAL, replace: `++${FINAL}++` }]
    });
    assert.equal(hasUnderline(resolve(turn1, acceptTrackedChangesInOoxml)), true, 'turn 1 underline works');
    const turn2 = await run(turn1, {
        target: { index: 7, exactText: FINAL }, targetRef: 'P7', modified: `${FINAL}\n${TABLE}`
    });
    assertAcceptKeepsIntent(turn2, 'WP-5a');
    assertRejectRestoresSource(turn2, 'WP-5a');
    assertParagraphMarkAccounting(turn2, 'WP-5a');
}

// WP-5b: one replacement string carrying a new Markdown underline and the table.
for (const separator of ['\n', '\n\n']) {
    const label = `WP-5b (${JSON.stringify(separator)})`;
    const tracked = await run(SOURCE_XML, {
        target: { index: 7, exactText: FINAL }, targetRef: 'P7', modified: `++${FINAL}++${separator}${TABLE}`
    });
    assertAcceptKeepsIntent(tracked, label);
    assertRejectRestoresSource(tracked, label);
    assertParagraphMarkAccounting(tracked, label);
}

// WP-6: plain append (no formatting) at the end of the document.
{
    const tracked = await run(SOURCE_XML, {
        target: { index: 7, exactText: FINAL }, targetRef: 'P7', modified: `${FINAL}\n${TABLE}`
    });
    const accepted = resolve(tracked, acceptTrackedChangesInOoxml);
    assert.deepEqual(cellsOf(accepted), CELLS);
    assert.equal(hasUnderline(accepted), false);
    assertRejectRestoresSource(tracked, 'WP-6 plain');
    assertParagraphMarkAccounting(tracked, 'WP-6 plain');
    // The source paragraph is retained, not rewritten as a deletion plus a copy.
    assert.equal(elements(bodyOf(tracked), 'del').length, 0, 'retained paragraph is not deleted and re-inserted');
}

// Mid-document append: the following paragraph already ends the table, so no extra paragraph is added.
{
    const tracked = await run(SOURCE_XML, {
        target: { index: 3, exactText: SOURCE[2] }, targetRef: 'P3', modified: `${SOURCE[2]}\n${TABLE}`
    });
    assertWordNativeInsertedTable(tracked, 'mid-document');
    assert.equal(paragraphsOf(tracked).length, SOURCE.length, 'no extra paragraph for a mid-document table');
    assert.equal(paragraphsOf(tracked).filter(p => markType(p) === 'ins').length, 0);
    assert.deepEqual(wordParagraphs(tracked, 'reject'), SOURCE);
    assert.deepEqual(libraryParagraphs(tracked, rejectTrackedChangesInOoxml), SOURCE);
    assert.deepEqual(cellsOf(resolve(tracked, acceptTrackedChangesInOoxml)), CELLS);
}

console.log('PASS: table append fidelity regression tests');
