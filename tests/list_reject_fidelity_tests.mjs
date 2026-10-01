import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';

import './setup-xml-provider.mjs';
import { acceptTrackedChangesInOoxml, rejectTrackedChangesInOoxml } from '../index.js';
import { parseOoxmlSafe } from '../adapters/xml-adapter.js';
import { unzipDocx } from '../document/zip-archive.js';
import { applyOperationsToDocumentXml } from '../services/standalone-operation-runner.js';

// Regression: Word-authored nested bullet list. Reject All in real Word must restore the exact
// source paragraph boundaries. Word semantics modelled here:
//   - a paragraph mark in w:pPr/w:rPr/w:ins is removed on Reject (content merges into the FOLLOWING paragraph)
//   - a paragraph mark in w:pPr/w:rPr/w:del is restored on Reject
//   - w:ins content is removed, w:del content is restored

const FIXTURE = new URL('./fixtures/agentic-lists/nested-lists-source.docx', import.meta.url);
const files = unzipDocx(readFileSync(FIXTURE));
const decoder = new TextDecoder();
const part = name => decoder.decode(files.get(name));
const documentXml = part('word/document.xml');
const parts = { numberingXml: part('word/numbering.xml'), stylesXml: part('word/styles.xml') };
const AUTHOR = 'List Reject Fidelity';
const OPTIONS = {
    atomic: true,
    strictTargets: true,
    structuredContent: true,
    generateRedlines: true,
    existingRevisions: 'merge-same-author',
    date: '2026-09-30T12:00:00Z'
};

const elements = (node, name) => Array.from(node.getElementsByTagNameNS('*', name));
const paragraphsOf = xml => {
    const parsed = parseOoxmlSafe(xml);
    assert.ok(parsed.doc, parsed.error?.message || 'parseable OOXML');
    return elements(parsed.doc, 'p');
};
const markType = paragraph => {
    for (const type of ['ins', 'del']) {
        if (elements(paragraph, type).some(n => n.parentNode?.localName === 'rPr' && n.parentNode.parentNode?.localName === 'pPr'
            && n.parentNode.parentNode.parentNode === paragraph)) return type;
    }
    return null;
};
const inside = (node, name, root) => {
    for (let cur = node.parentNode; cur && cur !== root; cur = cur.parentNode) if (cur.localName === name) return true;
    return false;
};
// text a paragraph carries after Reject (ins removed, del restored)
const rejectedFragment = p => elements(p, 't').concat(elements(p, 'delText'))
    .filter(n => !inside(n, 'ins', p)).map(n => n.textContent).join('');
const acceptedFragment = p => elements(p, 't').filter(n => !inside(n, 'del', p)).map(n => n.textContent).join('');

function wordParagraphs(xml, action) {
    const out = [];
    let carry = '';
    for (const p of paragraphsOf(xml)) {
        const mark = markType(p);
        const text = carry + (action === 'reject' ? rejectedFragment(p) : acceptedFragment(p));
        const markGoes = action === 'reject' ? mark === 'ins' : mark === 'del';
        if (markGoes) { carry = text; continue; }
        carry = '';
        out.push(text);
    }
    if (carry) out.push(carry);
    return out;
}
const libraryParagraphs = (xml, fn) => paragraphsOf(fn(xml, { allAuthors: true }).oxml).map(p => elements(p, 't').map(n => n.textContent).join(''));

async function run(operation) {
    const result = await applyOperationsToDocumentXml(documentXml, [operation], AUTHOR, parts, OPTIONS);
    assert.ok(result.documentXml && result.hasChanges, JSON.stringify(result.results || result));
    return result.documentXml;
}

const sourceTexts = paragraphsOf(documentXml).map(p => elements(p, 't').map(n => n.textContent).join(''));
// Target `index` values are the consumer's (one past the zero-based paragraph position).
const PLAIN = 1, ROOT = 2, ANCHOR = 3;
assert.equal(sourceTexts[PLAIN], 'Plain paragraph before bullet list.');
assert.equal(sourceTexts[ROOT], 'Bullet Root A');
assert.equal(sourceTexts[ANCHOR], 'Bullet Insertion Anchor');

// 1. Plain-anchor insertion before a list: inserted paragraph must carry a tracked-inserted mark.
{
    const tracked = await run({
        type: 'redline',
        target: { index: 2, exactText: sourceTexts[PLAIN] },
        modified: `${sourceTexts[PLAIN]}\nPlanner plain insertion after paragraph`
    });
    const added = paragraphsOf(tracked).find(p => acceptedFragment(p) === 'Planner plain insertion after paragraph');
    assert.ok(added, 'inserted paragraph exists');
    assert.equal(markType(added), 'ins', 'inserted paragraph mark must be tracked as inserted');
    assert.deepEqual(wordParagraphs(tracked, 'reject'), sourceTexts, 'Word Reject All restores source paragraphs exactly');
    assert.equal(paragraphsOf(tracked).filter(p => markType(p) !== 'ins').length, sourceTexts.length,
        'untracked paragraph marks equal source paragraph count');
    assert.deepEqual(libraryParagraphs(tracked, rejectTrackedChangesInOoxml), sourceTexts, 'library reject restores source');
    assert.deepEqual(wordParagraphs(tracked, 'accept'),
        sourceTexts.toSpliced(PLAIN + 1, 0, 'Planner plain insertion after paragraph'));
    assert.deepEqual(libraryParagraphs(tracked, acceptTrackedChangesInOoxml),
        sourceTexts.toSpliced(PLAIN + 1, 0, 'Planner plain insertion after paragraph'));
}

// 2. Range replacement of two list paragraphs: both original paragraph marks must survive Reject.
{
    const tracked = await run({
        type: 'redline',
        target: { index: 3, exactText: sourceTexts[ROOT] },
        targetEnd: { index: 4, exactText: sourceTexts[ANCHOR] },
        modified: '- Planner replacement parent\n    - Planner replacement child'
    });
    assert.deepEqual(wordParagraphs(tracked, 'reject'), sourceTexts, 'Word Reject All keeps the two source list paragraphs separate');
    assert.deepEqual(libraryParagraphs(tracked, rejectTrackedChangesInOoxml), sourceTexts, 'library reject keeps source paragraphs');
    const expectedAccepted = sourceTexts.toSpliced(ROOT, 2, 'Planner replacement parent', 'Planner replacement child');
    assert.deepEqual(wordParagraphs(tracked, 'accept'), expectedAccepted);
    assert.deepEqual(libraryParagraphs(tracked, acceptTrackedChangesInOoxml), expectedAccepted);
    const deletedParagraphs = paragraphsOf(tracked).filter(p => elements(p, 'delText').length > 0);
    assert.equal(deletedParagraphs.length, 2, 'each source paragraph keeps its own deleted paragraph');
    assert.ok(deletedParagraphs.every(p => markType(p) === 'del'), 'each deleted source paragraph has a tracked-deleted mark');
    const bullets = deletedParagraphs.map(p => elements(p, 'ilvl')[0]?.getAttribute('w:val'));
    assert.deepEqual(bullets, ['0', '1'], 'source list levels are preserved on the deleted paragraphs');
}

// 3. Range replacement where a middle source paragraph is empty: its mark must survive Reject too.
{
    const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
    const item = text => `<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>${text ? `<w:r><w:t>${text}</w:t></w:r>` : ''}</w:p>`;
    const emptyDoc = `<w:document xmlns:w="${W}"><w:body><w:p><w:r><w:t>Intro</w:t></w:r></w:p>${item('Alpha')}${item('')}${item('Gamma')}<w:sectPr/></w:body></w:document>`;
    const result = await applyOperationsToDocumentXml(emptyDoc, [{
        type: 'redline',
        target: { index: 2, exactText: 'Alpha' },
        targetEnd: { index: 4, exactText: 'Gamma' },
        modified: '- One\n- Two'
    }], AUTHOR, {}, { atomic: true, strictTargets: true, structuredContent: true, generateRedlines: true });
    assert.equal(result.status, 'ok', JSON.stringify(result.error || result.results));
    const tracked = result.documentXml;
    assert.deepEqual(wordParagraphs(tracked, 'reject'), ['Intro', 'Alpha', '', 'Gamma']);
    assert.deepEqual(libraryParagraphs(tracked, rejectTrackedChangesInOoxml), ['Intro', 'Alpha', '', 'Gamma']);
    assert.deepEqual(wordParagraphs(tracked, 'accept'), ['Intro', 'One', 'Two']);
    assert.deepEqual(libraryParagraphs(tracked, acceptTrackedChangesInOoxml), ['Intro', 'One', 'Two']);
    assert.equal(paragraphsOf(tracked).filter(p => markType(p) === 'del').length, 3, 'all three source marks tracked as deleted');
}

// 4. All-empty range and a single empty paragraph: source marks survive Reject without any source text.
{
    const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
    const empty = '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr></w:p>';
    const build = count => `<w:document xmlns:w="${W}"><w:body><w:p><w:r><w:t>Intro</w:t></w:r></w:p>${empty.repeat(count)}<w:sectPr/></w:body></w:document>`;
    const apply = (xml, operation) => applyOperationsToDocumentXml(xml, [operation], AUTHOR, {}, { atomic: true, strictTargets: true, structuredContent: true, generateRedlines: true });
    const cases = [
        { name: 'all-empty range', count: 2, operation: { type: 'redline', target: { index: 2, exactText: '' }, targetEnd: { index: 3, exactText: '' }, modified: '- One\n- Two' }, accepted: ['Intro', 'One', 'Two'] },
        { name: 'single empty paragraph (explicit one-paragraph range)', count: 1, operation: { type: 'redline', target: { index: 2, exactText: '' }, targetEnd: { index: 2, exactText: '' }, modified: '- One\n- Two' }, accepted: ['Intro', 'One', 'Two'] }
    ];
    // Without targetEnd an empty exactText cannot identify a paragraph: refused before any mutation.
    const refused = await apply(build(1), { type: 'redline', target: { index: 2, exactText: '' }, modified: '- One\n- Two' });
    assert.equal(refused.status, 'error');
    assert.equal(refused.documentXml, build(1), 'refusal leaves the document untouched');
    for (const { name, count, operation, accepted } of cases) {
        const result = await apply(build(count), operation);
        assert.equal(result.status, 'ok', `${name}: ${JSON.stringify(result.error || result.results)}`);
        const expectedRejected = ['Intro', ...Array(count).fill('')];
        assert.deepEqual(wordParagraphs(result.documentXml, 'reject'), expectedRejected, `${name}: Word reject`);
        assert.deepEqual(libraryParagraphs(result.documentXml, rejectTrackedChangesInOoxml), expectedRejected, `${name}: library reject`);
        assert.deepEqual(wordParagraphs(result.documentXml, 'accept'), accepted, `${name}: Word accept`);
        assert.deepEqual(libraryParagraphs(result.documentXml, acceptTrackedChangesInOoxml), accepted, `${name}: library accept`);
        assert.equal(paragraphsOf(result.documentXml).filter(p => markType(p) === 'del').length, count, `${name}: source marks tracked as deleted`);
    }
}

console.log('PASS: list reject fidelity regression tests');
