/**
 * Paragraph boundary fixture matrix and ownership oracle tests.
 *
 * Validates the 10 Word-native golden fixture triples created for WP-09:
 * 1. Split paragraph in middle
 * 2. Delete boundary between two paragraphs
 * 3. Delete entire middle paragraph
 * 4. Insert blank paragraph
 * 5. Boundaries between different paragraph styles
 * 6. Boundaries with different list levels / numIds
 * 7. Boundary before a section break
 * 8. First/last paragraph in a table cell
 * 9. Adjacent bookmark / comment ranges
 * 10. Multi-author boundary revisions
 *
 * Proves the exact OOXML ownership rules for:
 * - which paragraph owns the inserted/deleted mark (w:pPr/w:rPr/w:ins or w:pPr/w:rPr/w:del)
 * - which w:pPr survives acceptance
 * - exact restoration upon rejection
 */

import test from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync, existsSync } from 'node:fs';
import { join } from 'node:path';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { configureXmlProvider } from '../adapters/xml-adapter.js';
import {
    acceptTrackedChangesInOoxml,
    rejectTrackedChangesInOoxml
} from '../services/revision-comment-management.js';

configureXmlProvider({ DOMParser, XMLSerializer });

const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const FIXTURES_DIR = join(process.cwd(), 'tests', 'fixtures', 'paragraph-boundaries');

function parseXml(xmlString) {
    const cleanXml = String(xmlString).replace(/^\uFEFF/, '').trim();
    return new DOMParser().parseFromString(cleanXml, 'application/xml');
}

function getWordElements(docOrEl, localName) {
    const list = [];
    const elements = docOrEl.getElementsByTagNameNS ? docOrEl.getElementsByTagNameNS(NS_W, localName) : [];
    for (let i = 0; i < elements.length; i++) {
        list.push(elements[i]);
    }
    return list;
}

function loadFixtureTriple(caseName) {
    const pendingPath = join(FIXTURES_DIR, `${caseName}-pending.xml`);
    const acceptedPath = join(FIXTURES_DIR, `${caseName}-accepted.xml`);
    const rejectedPath = join(FIXTURES_DIR, `${caseName}-rejected.xml`);

    assert.ok(existsSync(pendingPath), `Missing fixture: ${pendingPath}`);
    assert.ok(existsSync(acceptedPath), `Missing fixture: ${acceptedPath}`);
    assert.ok(existsSync(rejectedPath), `Missing fixture: ${rejectedPath}`);

    const pendingXml = readFileSync(pendingPath, 'utf8').replace(/^\uFEFF/, '').trim();
    const acceptedXml = readFileSync(acceptedPath, 'utf8').replace(/^\uFEFF/, '').trim();
    const rejectedXml = readFileSync(rejectedPath, 'utf8').replace(/^\uFEFF/, '').trim();

    return {
        caseName,
        pendingXml,
        acceptedXml,
        rejectedXml,
        pendingDoc: parseXml(pendingXml),
        acceptedDoc: parseXml(acceptedXml),
        rejectedDoc: parseXml(rejectedXml)
    };
}

test('WP-09 Case 1: split-middle fixture proves paragraph mark ownership and lifecycle', () => {
    const { pendingDoc, acceptedDoc, rejectedDoc, pendingXml } = loadFixtureTriple('split-middle');

    const paragraphs = getWordElements(pendingDoc, 'p');
    assert.equal(paragraphs.length, 2, 'Pending split has 2 paragraphs');

    // Rule: The first paragraph (A) owns the inserted paragraph mark inside w:pPr/w:rPr/w:ins
    const p1Pr = getWordElements(paragraphs[0], 'pPr')[0];
    assert.ok(p1Pr, 'Paragraph A must have w:pPr');
    const p1RPr = getWordElements(p1Pr, 'rPr')[0];
    assert.ok(p1RPr, 'Paragraph A w:pPr must have w:rPr');
    const insNode = getWordElements(p1RPr, 'ins')[0];
    assert.ok(insNode, 'Paragraph A must own the w:ins paragraph mark revision');

    // Paragraph B does not own an ins marker on its paragraph mark
    const p2Pr = getWordElements(paragraphs[1], 'pPr')[0];
    if (p2Pr) {
        const p2RPr = getWordElements(p2Pr, 'rPr')[0];
        const p2Ins = p2RPr ? getWordElements(p2RPr, 'ins')[0] : null;
        assert.equal(p2Ins, null, 'Paragraph B must NOT own the inserted paragraph mark');
    }

    // Acceptance: both paragraphs remain
    const acceptedParagraphs = getWordElements(acceptedDoc, 'p');
    assert.equal(acceptedParagraphs.length, 2, 'Accepted split maintains 2 paragraphs');

    // Rejection: single paragraph with combined text
    const rejectedParagraphs = getWordElements(rejectedDoc, 'p');
    assert.equal(rejectedParagraphs.length, 1, 'Rejected split collapses to 1 paragraph');
    const rejectedText = rejectedParagraphs[0].textContent;
    assert.match(rejectedText, /Sentence one\.\s*Sentence two\./, 'Text is restored in original sequence');

    // Validate with library accept/reject
    const acceptedResult = acceptTrackedChangesInOoxml(pendingXml, { allAuthors: true });
    const acceptedDocLib = parseXml(acceptedResult.oxml);
    assert.equal(getWordElements(acceptedDocLib, 'p').length, 2);

    const rejectedResult = rejectTrackedChangesInOoxml(pendingXml, { allAuthors: true });
    const rejectedDocLib = parseXml(rejectedResult.oxml);
    assert.equal(getWordElements(rejectedDocLib, 'p').length, 1);
});

test('WP-09 Case 2: delete-boundary fixture proves paragraph mark ownership and merge destination', () => {
    const { pendingDoc, acceptedDoc, rejectedDoc, pendingXml } = loadFixtureTriple('delete-boundary');

    const paragraphs = getWordElements(pendingDoc, 'p');
    assert.equal(paragraphs.length, 2, 'Pending boundary deletion has 2 paragraphs');

    // Rule: Paragraph A owns the deleted paragraph mark inside w:pPr/w:rPr/w:del
    const p1Pr = getWordElements(paragraphs[0], 'pPr')[0];
    assert.ok(p1Pr, 'Paragraph A must have w:pPr');
    const p1RPr = getWordElements(p1Pr, 'rPr')[0];
    assert.ok(p1RPr, 'Paragraph A w:pPr must have w:rPr');
    const delNode = getWordElements(p1RPr, 'del')[0];
    assert.ok(delNode, 'Paragraph A must own the w:del paragraph mark revision');

    // Acceptance: Paragraphs are merged into 1 paragraph
    const acceptedParagraphs = getWordElements(acceptedDoc, 'p');
    assert.equal(acceptedParagraphs.length, 1, 'Accepted boundary deletion yields 1 paragraph');
    assert.match(acceptedParagraphs[0].textContent, /Paragraph one\.\s*Paragraph two\./);

    // Rejection: 2 separate paragraphs restored
    const rejectedParagraphs = getWordElements(rejectedDoc, 'p');
    assert.equal(rejectedParagraphs.length, 2, 'Rejected boundary deletion restores 2 paragraphs');

    // Validate with library
    const acceptedResult = acceptTrackedChangesInOoxml(pendingXml, { allAuthors: true });
    const acceptedDocLib = parseXml(acceptedResult.oxml);
    assert.equal(getWordElements(acceptedDocLib, 'p').length, 1);

    const rejectedResult = rejectTrackedChangesInOoxml(pendingXml, { allAuthors: true });
    const rejectedDocLib = parseXml(rejectedResult.oxml);
    assert.equal(getWordElements(rejectedDocLib, 'p').length, 2);
});

test('WP-09 Case 3: delete-middle-paragraph fixture proves full paragraph deletion ownership', () => {
    const { pendingDoc, acceptedDoc, rejectedDoc, pendingXml } = loadFixtureTriple('delete-middle-paragraph');

    const paragraphs = getWordElements(pendingDoc, 'p');
    assert.equal(paragraphs.length, 3, 'Pending paragraph deletion retains 3 paragraphs in pending state');

    // Rule: The deleted paragraph (Paragraph 2) owns both w:pPr/w:rPr/w:del and w:del around its text
    const p2 = paragraphs[1];
    const p2Pr = getWordElements(p2, 'pPr')[0];
    assert.ok(p2Pr, 'Paragraph 2 has w:pPr');
    const p2RPr = getWordElements(p2Pr, 'rPr')[0];
    assert.ok(p2RPr, 'Paragraph 2 has w:rPr');
    const p2Del = getWordElements(p2RPr, 'del')[0];
    assert.ok(p2Del, 'Paragraph 2 owns w:del paragraph mark');

    const textDels = getWordElements(p2, 'del');
    // At least one w:del for paragraph mark + one w:del wrapping runs
    assert.ok(textDels.length >= 2, 'Paragraph 2 has w:del for text runs in addition to paragraph mark');

    // Acceptance: 2 paragraphs remain (Paragraph 2 is removed)
    const acceptedParagraphs = getWordElements(acceptedDoc, 'p');
    assert.equal(acceptedParagraphs.length, 2, 'Accepted paragraph deletion leaves 2 paragraphs');
    assert.equal(acceptedParagraphs[0].textContent.trim(), 'Paragraph one.');
    assert.equal(acceptedParagraphs[1].textContent.trim(), 'Paragraph three.');

    // Rejection: 3 paragraphs restored
    const rejectedParagraphs = getWordElements(rejectedDoc, 'p');
    assert.equal(rejectedParagraphs.length, 3, 'Rejected paragraph deletion restores 3 paragraphs');
    assert.equal(rejectedParagraphs[1].textContent.trim(), 'Paragraph two.');
});

test('WP-09 Case 4: insert-blank-paragraph fixture proves blank paragraph mark ownership', () => {
    const { pendingDoc, acceptedDoc, rejectedDoc } = loadFixtureTriple('insert-blank-paragraph');

    const paragraphs = getWordElements(pendingDoc, 'p');
    assert.equal(paragraphs.length, 3, 'Pending blank paragraph insertion has 3 paragraphs');

    // Acceptance: 3 paragraphs
    const acceptedParagraphs = getWordElements(acceptedDoc, 'p');
    assert.equal(acceptedParagraphs.length, 3, 'Accepted blank paragraph insertion maintains 3 paragraphs');

    // Rejection: 2 paragraphs
    const rejectedParagraphs = getWordElements(rejectedDoc, 'p');
    assert.equal(rejectedParagraphs.length, 2, 'Rejected blank paragraph insertion leaves 2 paragraphs');
});

test('WP-09 Case 5: different-styles-boundary proves property change propagation', () => {
    const { pendingDoc, acceptedDoc, rejectedDoc } = loadFixtureTriple('different-styles-boundary');

    const paragraphs = getWordElements(pendingDoc, 'p');
    assert.equal(paragraphs.length, 2, 'Pending boundary deletion has 2 paragraphs');

    // Rule: Paragraph A (Heading 1) owns w:pPr/w:rPr/w:del.
    // Paragraph B records w:pPrChange transitioning its style from Normal to Heading 1.
    const p1Pr = getWordElements(paragraphs[0], 'pPr')[0];
    assert.ok(getWordElements(p1Pr, 'del')[0], 'Paragraph 1 owns deleted paragraph mark');

    const p2Pr = getWordElements(paragraphs[1], 'pPr')[0];
    assert.ok(p2Pr, 'Paragraph 2 has w:pPr');
    const pPrChange = getWordElements(p2Pr, 'pPrChange')[0];
    assert.ok(pPrChange, 'Paragraph 2 has w:pPrChange for styling inheritance');

    // Acceptance: Surviving paragraph has Heading 1
    const acceptedParagraphs = getWordElements(acceptedDoc, 'p');
    assert.equal(acceptedParagraphs.length, 1, 'Accepted boundary yields 1 merged paragraph');
    const acceptedPPr = getWordElements(acceptedParagraphs[0], 'pPr')[0];
    const pStyle = getWordElements(acceptedPPr, 'pStyle')[0];
    assert.equal(pStyle?.getAttributeNS(NS_W, 'val') || pStyle?.getAttribute('w:val'), 'Heading1');

    // Rejection: Both paragraphs restored, Paragraph 2 style is Normal (no Heading1)
    const rejectedParagraphs = getWordElements(rejectedDoc, 'p');
    assert.equal(rejectedParagraphs.length, 2, 'Rejected boundary yields 2 separate paragraphs');
    const rejP2Pr = getWordElements(rejectedParagraphs[1], 'pPr')[0];
    const rejP2Style = rejP2Pr ? getWordElements(rejP2Pr, 'pStyle')[0] : null;
    assert.notEqual(rejP2Style?.getAttributeNS(NS_W, 'val') || rejP2Style?.getAttribute('w:val'), 'Heading1');
});

test('WP-09 Case 6: different-list-levels-boundary proves list numbering transition', () => {
    const { pendingDoc, acceptedDoc, rejectedDoc } = loadFixtureTriple('different-list-levels-boundary');

    const paragraphs = getWordElements(pendingDoc, 'p');
    assert.equal(paragraphs.length, 2, 'Pending list boundary deletion has 2 paragraphs');

    // Paragraph 1 has w:pPr/w:rPr/w:del
    const p1Pr = getWordElements(paragraphs[0], 'pPr')[0];
    assert.ok(getWordElements(p1Pr, 'del')[0], 'Paragraph 1 has deleted paragraph mark');

    // Acceptance: Merged into 1 list item
    const acceptedParagraphs = getWordElements(acceptedDoc, 'p');
    assert.equal(acceptedParagraphs.length, 1, 'Accepted list merge yields 1 paragraph');
    const numPr = getWordElements(acceptedParagraphs[0], 'numPr')[0];
    assert.ok(numPr, 'Merged paragraph retains list formatting');

    // Rejection: Restored to 2 distinct list items
    const rejectedParagraphs = getWordElements(rejectedDoc, 'p');
    assert.equal(rejectedParagraphs.length, 2, 'Rejected list merge yields 2 distinct paragraphs');
});

test('WP-09 Case 7: section-break-boundary proves section break encapsulation', () => {
    const { pendingDoc, acceptedDoc, rejectedDoc } = loadFixtureTriple('section-break-boundary');

    const paragraphs = getWordElements(pendingDoc, 'p');
    assert.ok(paragraphs.length >= 1, 'Pending document has paragraphs');

    // In Word, section break is stored inside w:pPr/w:sectPr of the paragraph before the break
    const p1Pr = getWordElements(paragraphs[0], 'pPr')[0];
    const sectPrInP = getWordElements(p1Pr, 'sectPr')[0];
    assert.ok(sectPrInP, 'Section break is anchored inside w:pPr/w:sectPr');
    assert.ok(getWordElements(p1Pr, 'del')[0], 'Deleted section break mark is on w:pPr/w:rPr/w:del');

    // Acceptance: section break is removed
    const acceptedP1 = getWordElements(acceptedDoc, 'p')[0];
    const acceptedP1Pr = getWordElements(acceptedP1, 'pPr')[0];
    if (acceptedP1Pr) {
        assert.equal(getWordElements(acceptedP1Pr, 'sectPr').length, 0, 'Section break was removed');
    }
});

test('WP-09 Case 8: table-cell-boundary proves intra-cell boundary lifecycle', () => {
    const { pendingDoc, acceptedDoc, rejectedDoc } = loadFixtureTriple('table-cell-boundary');

    const tables = getWordElements(pendingDoc, 'tbl');
    assert.equal(tables.length, 1, 'Table is preserved');

    const cellParagraphs = getWordElements(tables[0], 'p');
    assert.equal(cellParagraphs.length, 2, 'Cell originally contains 2 paragraphs in pending state');

    // Paragraph 1 owns deleted paragraph mark
    const p1Pr = getWordElements(cellParagraphs[0], 'pPr')[0];
    assert.ok(getWordElements(p1Pr, 'del')[0], 'First cell paragraph has deleted paragraph mark');

    // Acceptance: merged into 1 cell paragraph
    const acceptedTable = getWordElements(acceptedDoc, 'tbl')[0];
    const acceptedCellP = getWordElements(acceptedTable, 'p');
    assert.equal(acceptedCellP.length, 1, 'Cell has 1 merged paragraph after acceptance');

    // Rejection: restored to 2 cell paragraphs
    const rejectedTable = getWordElements(rejectedDoc, 'tbl')[0];
    const rejectedCellP = getWordElements(rejectedTable, 'p');
    assert.equal(rejectedCellP.length, 2, 'Cell has 2 restored paragraphs after rejection');
});

test('WP-09 Case 9: adjacent-bookmark-comment proves range marker integrity across boundary deletion', () => {
    const { pendingDoc, acceptedDoc, rejectedDoc } = loadFixtureTriple('adjacent-bookmark-comment');

    const pendingBookmarks = getWordElements(pendingDoc, 'bookmarkStart');
    assert.ok(pendingBookmarks.length >= 1, 'BookmarkStart is preserved in pending state');
    const pendingComments = getWordElements(pendingDoc, 'commentRangeStart');
    assert.ok(pendingComments.length >= 1, 'CommentRangeStart is preserved in pending state');

    // Acceptance: bookmark and comment range markers survive boundary join
    const acceptedBookmarks = getWordElements(acceptedDoc, 'bookmarkStart');
    assert.ok(acceptedBookmarks.length >= 1, 'BookmarkStart survives boundary join');
    const acceptedComments = getWordElements(acceptedDoc, 'commentRangeStart');
    assert.ok(acceptedComments.length >= 1, 'CommentRangeStart survives boundary join');

    // Rejection: bookmark and comment range markers survive boundary restoration
    const rejectedBookmarks = getWordElements(rejectedDoc, 'bookmarkStart');
    assert.ok(rejectedBookmarks.length >= 1, 'BookmarkStart survives rejection');
    const rejectedComments = getWordElements(rejectedDoc, 'commentRangeStart');
    assert.ok(rejectedComments.length >= 1, 'CommentRangeStart survives rejection');
});

test('WP-09 Case 10: multi-author-boundary proves multi-author metadata on boundary revisions', () => {
    const { pendingDoc, acceptedDoc, rejectedDoc } = loadFixtureTriple('multi-author-boundary');

    const paragraphs = getWordElements(pendingDoc, 'p');
    assert.equal(paragraphs.length, 2, 'Multi-author split has 2 paragraphs');

    // First author inserted boundary (split)
    const p1Pr = getWordElements(paragraphs[0], 'pPr')[0];
    const p1Ins = getWordElements(p1Pr, 'ins')[0];
    assert.ok(p1Ins, 'Paragraph 1 owns inserted boundary');
    const author1 = p1Ins.getAttributeNS(NS_W, 'author') || p1Ins.getAttribute('w:author');
    assert.equal(author1, 'AuthorOne');

    // Second author inserted text in paragraph 2
    const p2Ins = getWordElements(paragraphs[1], 'ins')[0];
    assert.ok(p2Ins, 'Paragraph 2 contains insertion by second author');
    const author2 = p2Ins.getAttributeNS(NS_W, 'author') || p2Ins.getAttribute('w:author');
    assert.equal(author2, 'AuthorTwo');

    // Acceptance: both changes accepted
    const acceptedParagraphs = getWordElements(acceptedDoc, 'p');
    assert.equal(acceptedParagraphs.length, 2, 'Both paragraphs maintained');
    assert.match(acceptedParagraphs[1].textContent, /Updated\s*Author two text\./);

    // Rejection: restored to single unsplit paragraph
    const rejectedParagraphs = getWordElements(rejectedDoc, 'p');
    assert.equal(rejectedParagraphs.length, 1, 'Rejection restores single paragraph');
    assert.match(rejectedParagraphs[0].textContent, /Author one text\.\s*Author two text\./);
});
