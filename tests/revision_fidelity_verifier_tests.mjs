import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';

import './setup-xml-provider.mjs';
import { parseOoxmlSafe } from '../adapters/xml-adapter.js';
import { unzipDocx } from '../document/zip-archive.js';
import { applyOperationsToDocumentXml } from '../services/standalone-operation-runner.js';
import { projectWordRevisionView, wordRevisionParagraphTexts } from '../core/word-revision-semantics.js';
import { verifyRevisionFidelity } from '../services/revision-fidelity-verifier.js';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const META = 'w:id="900" w:author="Defect" w:date="2026-09-30T00:00:00Z"';
const doc = body => {
    const parsed = parseOoxmlSafe(`<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>${body}<w:sectPr/></w:body></w:document>`);
    assert.ok(parsed.doc, parsed.error?.message);
    return parsed.doc;
};
const p = (inner, pPr = '') => `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}${inner}</w:p>`;
const run = text => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const delRun = text => `<w:del ${META}><w:r><w:delText xml:space="preserve">${text}</w:delText></w:r></w:del>`;
const insRun = text => `<w:ins ${META}><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:ins>`;
const insMark = `<w:rPr><w:ins ${META}/></w:rPr>`;
const delMark = `<w:rPr><w:del ${META}/></w:rPr>`;
const numPr = (numId, ilvl = 0) => `<w:numPr><w:ilvl w:val="${ilvl}"/><w:numId w:val="${numId}"/></w:numPr>`;
const invariantOf = (before, after) => {
    const result = verifyRevisionFidelity({ before: doc(before), after: doc(after) });
    return result.ok ? null : result.violations[0].invariant;
};

// 1. Simulator: Word paragraph-mark rules.
{
    const tracked = doc(p(run('Keep')) + p(insRun('New'), insMark) + p(run('Next')));
    assert.deepEqual(wordRevisionParagraphTexts(tracked, 'rejected'), ['Keep', 'Next']);
    assert.deepEqual(wordRevisionParagraphTexts(tracked, 'accepted'), ['Keep', 'New', 'Next']);
    // Untracked mark with inserted content: Word keeps an empty paragraph on Reject.
    const untrackedMark = doc(p(run('Keep')) + p(insRun('New')) + p(run('Next')));
    assert.deepEqual(wordRevisionParagraphTexts(untrackedMark, 'rejected'), ['Keep', '', 'Next']);
    // Deleted mark merges content into the following paragraph on Accept.
    const deleted = doc(p(run('A') + delRun('gone'), delMark) + p(run('B')));
    assert.deepEqual(wordRevisionParagraphTexts(deleted, 'accepted'), ['AB']);
    assert.deepEqual(wordRevisionParagraphTexts(deleted, 'rejected'), ['Agone', 'B']);
    // The last mark of a container survives even when tracked.
    const lastInCell = doc(`<w:tbl><w:tr><w:tc>${p(run('cell'), insMark)}</w:tc></w:tr></w:tbl>`);
    assert.deepEqual(wordRevisionParagraphTexts(lastInCell, 'rejected'), ['cell']);
    // Inserted rows vanish on Reject; deleted rows on Accept.
    const rows = doc(`<w:tbl><w:tr><w:trPr><w:ins ${META}/></w:trPr><w:tc>${p(insRun('row'))}</w:tc></w:tr><w:tr><w:tc>${p(run('kept'))}</w:tc></w:tr></w:tbl>`);
    assert.deepEqual(wordRevisionParagraphTexts(rows, 'rejected'), ['kept']);
    // pPrChange restores the historical properties on Reject.
    const changed = doc(p(run('Item'), `<w:pStyle w:val="Heading1"/><w:pPrChange ${META}><w:pPr>${numPr(4)}</w:pPr></w:pPrChange>`));
    const [rejected] = projectWordRevisionView(changed, 'rejected');
    const [accepted] = projectWordRevisionView(changed, 'accepted');
    assert.deepEqual(rejected.numPr, { numId: '4', ilvl: '0' });
    assert.equal(accepted.numPr, null);
    assert.match(accepted.props, /pStyle/);
    // Breaks and tabs are part of the canonical text; hyperlink spans are recorded.
    const linked = doc(p(`<w:hyperlink r:id="rId5">${run('site')}</w:hyperlink>${run('.')}<w:r><w:br/><w:tab/></w:r>`));
    const [record] = projectWordRevisionView(linked, 'rejected');
    assert.equal(record.text, 'site.\n\t');
    assert.deepEqual(record.wrappers, ['hyperlink:rId5@0-4']);
}

// 2. Regression corpus: one defective output per row of the plan's "Why now" table.
{
    // Plain-anchor insertion without a tracked inserted mark (Reject left an empty paragraph).
    assert.equal(invariantOf(
        p(run('Plain')) + p(run('Bullet'), numPr(1)),
        p(run('Plain')) + p(insRun('Inserted')) + p(run('Bullet'), numPr(1))
    ), 'REJECT_PARAGRAPH_COUNT');
    // List-range replacement collapsing two source paragraphs into one deleted paragraph
    // (Reject restores "AlphaBeta" as one paragraph).
    assert.equal(invariantOf(
        p(run('Alpha'), numPr(1)) + p(run('Beta'), numPr(1)),
        p(delRun('Alpha') + delRun('Beta'), numPr(1) + delMark) + p(insRun('One'), numPr(1))
    ), 'REJECT_PARAGRAPH_TEXT');
    assert.equal(invariantOf(
        p(run('Alpha'), numPr(1)) + p(run('Beta'), numPr(1)) + p(run('Tail')),
        p(delRun('Alpha') + delRun('Beta'), numPr(1) + delMark) + p(insRun('One'), numPr(1) + insMark) + p(run('Tail'))
    ), 'REJECT_PARAGRAPH_COUNT');
    // Localized replacement that pulled the adjacent plain "." into the hyperlink.
    assert.equal(invariantOf(
        p(`<w:hyperlink r:id="rId5">${run('site')}</w:hyperlink>${run('.')}`),
        p(`<w:hyperlink r:id="rId5">${delRun('site')}${insRun('page')}${run('.')}</w:hyperlink>`)
    ), 'REJECT_WRAPPERS');
    // Reject All reordering text across a manual line break.
    assert.equal(invariantOf(
        p(`${run('first')}<w:r><w:br/></w:r>${run('second')}`),
        p(`${run('second')}${insRun('X')}<w:r><w:br/></w:r>${run('first')}`)
    ), 'REJECT_PARAGRAPH_TEXT');
    // Facade path remapping deleted source paragraphs to a new generic numId.
    assert.equal(invariantOf(
        p(run('Item'), numPr(3)) + p(run('Tail')),
        p(delRun('Item'), numPr(7) + delMark) + p(insRun('New'), numPr(7) + insMark) + p(run('Tail'))
    ), 'REJECT_NUMBERING_REFS');
    // Historical properties reported/written as current: pPrChange lost on a tracked edit.
    assert.equal(invariantOf(
        p(run('Heading'), '<w:pStyle w:val="Heading1"/>'),
        p(run('Heading') + insRun('!'), '<w:pStyle w:val="Normal"/>')
    ), 'REJECT_PARAGRAPH_PROPS');
    // Correct tracked equivalents of the same edits pass.
    assert.equal(invariantOf(
        p(run('Plain')) + p(run('Bullet'), numPr(1)),
        p(run('Plain')) + p(insRun('Inserted'), insMark) + p(run('Bullet'), numPr(1))
    ), null);
    assert.equal(invariantOf(
        p(run('Heading'), '<w:pStyle w:val="Heading1"/>'),
        p(run('Heading'), `<w:pStyle w:val="Normal"/><w:pPrChange ${META}><w:pPr><w:pStyle w:val="Heading1"/></w:pPr></w:pPrChange>`)
    ), null);
}

// 3. Mutation test of the verifier against real library output.
{
    const files = unzipDocx(readFileSync(new URL('./fixtures/agentic-lists/nested-lists-source.docx', import.meta.url)));
    const decoder = new TextDecoder();
    const part = name => decoder.decode(files.get(name));
    const documentXml = part('word/document.xml');
    const parts = { numberingXml: part('word/numbering.xml'), stylesXml: part('word/styles.xml') };
    const source = parseOoxmlSafe(documentXml).doc;
    const texts = wordRevisionParagraphTexts(source, 'rejected');
    const options = { atomic: true, strictTargets: true, structuredContent: true, generateRedlines: true, verifyRevisionFidelity: 'enforce' };
    const apply = async operation => {
        const result = await applyOperationsToDocumentXml(documentXml, [operation], 'Verifier', parts, options);
        assert.equal(result.status, 'ok', JSON.stringify(result.error || result.results));
        return result.documentXml;
    };
    const check = xml => verifyRevisionFidelity({ before: source, after: parseOoxmlSafe(xml).doc });

    const inserted = await apply({ type: 'redline', target: { index: 2, exactText: texts[1] }, modified: `${texts[1]}\nAdded paragraph` });
    assert.equal(check(inserted).ok, true);
    const droppedMark = inserted.replace(/(<w:pPr>(?:(?!<\/w:pPr>).)*?<w:rPr>(?:(?!<\/w:rPr>).)*?)<w:ins\b[^>]*\/>/s, '$1');
    assert.notEqual(droppedMark, inserted);
    assert.equal(check(droppedMark).violations[0].invariant, 'REJECT_PARAGRAPH_COUNT');

    const listRange = await apply({
        type: 'redline',
        target: { index: 3, exactText: texts[2] },
        targetEnd: { index: 4, exactText: texts[3] },
        modified: '- Replacement parent\n    - Replacement child'
    });
    assert.equal(check(listRange).ok, true);
    const firstDeleted = listRange.indexOf('<w:delText');
    const numIdAt = listRange.lastIndexOf('<w:numId w:val="', firstDeleted);
    const swapped = `${listRange.slice(0, numIdAt)}<w:numId w:val="999"/>${listRange.slice(listRange.indexOf('/>', numIdAt) + 2)}`;
    assert.equal(check(swapped).violations[0].invariant, 'REJECT_NUMBERING_REFS');

    const edited = await apply({ type: 'redline', target: { index: 2, exactText: texts[1] }, modified: 'Plain sentence before bullet list.' });
    assert.equal(check(edited).ok, true);
    const lostDeletion = edited.replace(/<w:del\b[^>]*[^/]>(?:(?!<\/w:del>).)*<\/w:del>/s, '');
    assert.notEqual(lostDeletion, edited);
    assert.equal(check(lostDeletion).violations[0].invariant, 'REJECT_PARAGRAPH_TEXT');
}

// 4. Batches reuse the previous operation's projection as the baseline; an untracked commit
//    in between must invalidate it (otherwise the next tracked op would be compared stale).
{
    const body = ['Alpha clause.', 'Beta clause.', 'Gamma clause.', 'Delta clause.'].map(t => p(run(t))).join('');
    const xml = `<w:document xmlns:w="${W}"><w:body>${body}<w:sectPr/></w:body></w:document>`;
    const result = await applyOperationsToDocumentXml(xml, [
        { type: 'redline', target: 'Alpha clause.', modified: 'Alpha term.' },
        { type: 'redline', target: 'Beta clause.', modified: 'Beta term.', generateRedlines: false },
        { type: 'redline', target: 'Gamma clause.', modified: 'Gamma term.' }
    ], 'Verifier', {}, { atomic: true, verifyRevisionFidelity: 'enforce' });
    assert.equal(result.status, 'ok', JSON.stringify(result.error || result.results));
    assert.deepEqual(wordRevisionParagraphTexts(parseOoxmlSafe(result.documentXml).doc, 'rejected'),
        ['Alpha clause.', 'Beta term.', 'Gamma clause.', 'Delta clause.']);
}

console.log('PASS: revision fidelity verifier tests');
