import './setup-xml-provider.mjs';

import assert from 'assert/strict';

import { ingestWordOoxmlToPlainText } from '../index.js';
import { extractExactVisibleText, normalizeParagraphBreaks } from './helpers/ooxml-assertions.mjs';

/*
 * Self-tests for the verification oracle itself.
 *
 * assertRoundTrip is only as strong as the extractor behind it. These cases pin
 * down that extractExactVisibleText really is lossless where the production
 * reader is lossy -- if this file ever starts agreeing with
 * ingestWordOoxmlToPlainText on whitespace, the round-trip suite has quietly
 * gone blind again.
 */

const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

function doc(body) {
    return `<w:document xmlns:w="${NS_W}"><w:body>${body}</w:body></w:document>`;
}

function para(inner) {
    return `<w:p>${inner}</w:p>`;
}

function t(text) {
    const space = /^\s|\s$/.test(text) ? ' xml:space="preserve"' : '';
    return `<w:r><w:t${space}>${text}</w:t></w:r>`;
}

const failures = [];

function check(name, fn) {
    try {
        fn();
    } catch (error) {
        failures.push(`${name}: ${error.message}`);
    }
}

// --- the whole point: whitespace the production reader throws away ---------

check('preserves a double space that the production reader collapses', () => {
    const xml = doc(para(t('Section  1 applies.')));
    assert.equal(extractExactVisibleText(xml), 'Section  1 applies.');
    assert.equal(ingestWordOoxmlToPlainText(xml), 'Section 1 applies.');
});

check('preserves w:tab as a tab where the production reader emits a space', () => {
    const xml = doc(para(`${t('A')}<w:r><w:tab/></w:r>${t('B')}`));
    assert.equal(extractExactVisibleText(xml), 'A\tB');
    assert.equal(ingestWordOoxmlToPlainText(xml), 'A B');
});

check('preserves trailing whitespace that the production reader trims', () => {
    const xml = doc(para(t('alpha beta ')));
    assert.equal(extractExactVisibleText(xml), 'alpha beta ');
    assert.equal(ingestWordOoxmlToPlainText(xml), 'alpha beta');
});

check('preserves leading whitespace', () => {
    const xml = doc(para(t('  indented')));
    assert.equal(extractExactVisibleText(xml), '  indented');
});

check('maps w:br to a newline and w:noBreakHyphen to U+2011', () => {
    const xml = doc(para(`${t('one')}<w:r><w:br/></w:r>${t('two')}<w:r><w:noBreakHyphen/></w:r>${t('three')}`));
    assert.equal(extractExactVisibleText(xml), 'one\ntwo‑three');
});

// --- accepted-view semantics ----------------------------------------------

check('hides w:del content and shows w:ins content', () => {
    const xml = doc(para([
        t('keep '),
        `<w:del w:id="1" w:author="A" w:date="2026-01-01T00:00:00Z"><w:r><w:delText>gone </w:delText></w:r></w:del>`,
        `<w:ins w:id="2" w:author="A" w:date="2026-01-01T00:00:00Z">${t('added')}</w:ins>`
    ].join('')));
    assert.equal(extractExactVisibleText(xml), 'keep added');
});

check('hides w:moveFrom content and shows w:moveTo content', () => {
    const xml = doc([
        para(`<w:moveFrom w:id="1" w:author="A" w:date="2026-01-01T00:00:00Z"><w:r><w:delText>moved</w:delText></w:r></w:moveFrom>`),
        para(`<w:moveTo w:id="2" w:author="A" w:date="2026-01-01T00:00:00Z">${t('moved')}</w:moveTo>`)
    ].join(''));
    assert.equal(extractExactVisibleText(xml), '\nmoved');
});

check('ignores field instruction text and paragraph properties', () => {
    const xml = doc(para([
        '<w:pPr><w:jc w:val="center"/></w:pPr>',
        '<w:r><w:instrText> DATE </w:instrText></w:r>',
        t('visible')
    ].join('')));
    assert.equal(extractExactVisibleText(xml), 'visible');
});

check('separates paragraphs with a newline', () => {
    const xml = doc(`${para(t('one'))}${para(t('two'))}`);
    assert.equal(extractExactVisibleText(xml), 'one\ntwo');
});

check('merges a paragraph whose mark is deleted into the next one', () => {
    const xml = doc([
        para(`<w:pPr><w:rPr><w:del w:id="9" w:author="A" w:date="2026-01-01T00:00:00Z"/></w:rPr></w:pPr>${t('one')}`),
        para(t('two'))
    ].join(''));
    assert.equal(extractExactVisibleText(xml), 'onetwo');
});

check('reads table cell paragraphs', () => {
    const xml = doc(`<w:tbl><w:tr><w:tc>${para(t('cell one'))}</w:tc><w:tc>${para(t('cell two'))}</w:tc></w:tr></w:tbl>`);
    assert.equal(extractExactVisibleText(xml), 'cell one\ncell two');
});

// --- the oracle must be able to FAIL --------------------------------------

check('detects a w:tab downgraded to a literal space', () => {
    // A w:tab rewritten as a space is the regression the normalized oracle
    // cannot see: both sides read as "A B" once whitespace is collapsed.
    const withTab = doc(para(`${t('A')}<w:r><w:tab/></w:r>${t('B')}`));
    const tabAsSpace = doc(para(`${t('A')}${t(' ')}${t('B')}`));

    assert.notEqual(
        extractExactVisibleText(withTab),
        extractExactVisibleText(tabAsSpace),
        'exact extractor must distinguish a tab from a space'
    );
    assert.equal(
        ingestWordOoxmlToPlainText(withTab).replace(/\s+/g, ' '),
        ingestWordOoxmlToPlainText(tabAsSpace).replace(/\s+/g, ' '),
        'the old normalized oracle genuinely could not distinguish these'
    );
});

check('detects a collapsed double space', () => {
    assert.notEqual(
        extractExactVisibleText(doc(para(t('Section  1')))),
        extractExactVisibleText(doc(para(t('Section 1'))))
    );
});

check('detects lost trailing whitespace', () => {
    assert.notEqual(
        extractExactVisibleText(doc(para(t('alpha ')))),
        extractExactVisibleText(doc(para(t('alpha'))))
    );
});

// --- paragraph-break normalization is the ONLY normalization --------------

check('normalizeParagraphBreaks folds blank lines but not spaces or tabs', () => {
    assert.equal(normalizeParagraphBreaks('a\n\n\nb'), 'a\nb');
    assert.equal(normalizeParagraphBreaks('a\r\nb'), 'a\nb');
    assert.equal(normalizeParagraphBreaks('a  b'), 'a  b', 'double space must survive');
    assert.equal(normalizeParagraphBreaks('a\tb'), 'a\tb', 'tab must survive');
    assert.equal(normalizeParagraphBreaks('a '), 'a ', 'trailing space must survive');
});

if (failures.length > 0) {
    console.error('FAIL: round-trip oracle tests');
    for (const failure of failures) console.error(`  - ${failure}`);
    process.exit(1);
}

console.log('PASS: round-trip oracle tests');
