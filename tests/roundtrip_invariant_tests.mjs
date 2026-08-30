import './setup-xml-provider.mjs';

import assert from 'assert/strict';

import { acceptTrackedChangesInOoxml, applyRedlineToOxml } from '../index.js';
import { assertRoundTrip, assertRoundTripStructure } from './helpers/roundtrip.mjs';
import { elementsByLocalName, parseXmlFragment } from './helpers/ooxml-assertions.mjs';

const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

function paragraph(inner, attrs = '') {
    return `<w:p xmlns:w="${NS_W}" xmlns:r="${NS_R}"${attrs}>${inner}</w:p>`;
}

function textRun(text, rPr = '') {
    const needsPreserve = /^\s|\s$/.test(text);
    const space = needsPreserve ? ' xml:space="preserve"' : '';
    return `<w:r>${rPr}<w:t${space}>${text}</w:t></w:r>`;
}

function document(body) {
    return `<w:document xmlns:w="${NS_W}" xmlns:r="${NS_R}"><w:body>${body}<w:sectPr/></w:body></w:document>`;
}

const singleRun = paragraph(textRun('The quick brown fox.'));

const multiRun = paragraph([
    textRun('Alpha', '<w:rPr><w:b/></w:rPr>'),
    textRun(' beta ', '<w:rPr><w:i/></w:rPr>'),
    textRun('gamma', '<w:rPr><w:u w:val="single"/></w:rPr>')
].join(''));

const whitespace = paragraph(textRun('foo '));

const sentenceDeletion = paragraph(textRun('Keep this. Delete this sentence.'));

const hyperlink = paragraph([
    textRun('Before '),
    '<w:hyperlink r:id="rId5">',
    textRun('link text'),
    '</w:hyperlink>',
    textRun(' after')
].join(''));

const tableCell = `<w:tbl xmlns:w="${NS_W}" xmlns:r="${NS_R}">
  <w:tr>
    <w:tc>
      ${paragraph(textRun('Cell old text'))}
    </w:tc>
  </w:tr>
</w:tbl>`;

const markdownBold = paragraph(textRun('Make word bold'));

const proofErrAndField = paragraph([
    '<w:proofErr w:type="spellStart"/>',
    textRun('Shown '),
    '<w:proofErr w:type="spellEnd"/>',
    '<w:r><w:fldChar w:fldCharType="begin"/></w:r>',
    '<w:r><w:instrText> DATE </w:instrText></w:r>',
    '<w:r><w:fldChar w:fldCharType="separate"/></w:r>',
    textRun('value'),
    '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
].join(''));

const unicode = paragraph(textRun('Hello 世界 😀'));

// Whitespace-hostile cases. These only became meaningful once assertRoundTrip
// started comparing at 'exact' fidelity -- under the old normalized comparison
// every one of them passed vacuously.
const doubleSpace = paragraph(textRun('Section  1 applies here.'));
const tabbedRuns = paragraph([textRun('Name'), '<w:r><w:tab/></w:r>', textRun('old value')].join(''));
const trailingSpace = paragraph(textRun('alpha beta '));
const leadingSpace = paragraph(textRun('  indented text'));
const lineBreak = paragraph([textRun('first line'), '<w:r><w:br/></w:r>', textRun('second line')].join(''));
const twoParagraphsBody = `<w:document xmlns:w="${NS_W}" xmlns:r="${NS_R}"><w:body>${paragraph(textRun('para one'))}${paragraph(textRun('para two'))}</w:body></w:document>`;
const twoParagraphsWithSectPr = document(`${paragraph(textRun('para one'))}${paragraph(textRun('para two'))}`);

const corpus = [
    {
        name: 'single-run paragraph, one word replaced mid-sentence',
        oxml: singleRun,
        original: 'The quick brown fox.',
        modified: 'The quick red fox.'
    },
    {
        name: 'multi-run paragraph, edit spanning a run boundary',
        oxml: multiRun,
        original: 'Alpha beta gamma',
        modified: 'Alpha better gamma'
    },
    {
        name: 'leading/trailing whitespace significant',
        oxml: whitespace,
        original: 'foo ',
        modified: 'bar  baz '
    },
    {
        name: 'pure insertion at start of paragraph',
        oxml: paragraph(textRun('world')),
        original: 'world',
        modified: 'Hello world'
    },
    {
        name: 'pure insertion at end of paragraph',
        oxml: paragraph(textRun('Hello')),
        original: 'Hello',
        modified: 'Hello world'
    },
    {
        name: 'pure deletion of an entire sentence',
        oxml: sentenceDeletion,
        original: 'Keep this. Delete this sentence.',
        modified: 'Keep this.'
    },
    {
        name: 'edit outside hyperlink preserves link paragraph',
        oxml: hyperlink,
        original: 'Before link text after',
        modified: 'Ahead link text after'
    },
    {
        name: 'edit inside a table cell paragraph',
        oxml: tableCell,
        original: 'Cell old text',
        modified: 'Cell new text'
    },
    {
        name: 'markdown formatting added around an existing word',
        oxml: markdownBold,
        original: 'Make word bold',
        modified: 'Make **word** bold'
    },
    {
        name: 'proofErr markers and simple field scaffolding',
        oxml: proofErrAndField,
        original: 'Shown value',
        modified: 'Shown updated value'
    },
    {
        name: 'unicode emoji and CJK replacement',
        oxml: unicode,
        original: 'Hello 世界 😀',
        modified: 'Hello 世界朋友 😀'
    },
    {
        name: 'double space survives an edit elsewhere in the paragraph',
        oxml: doubleSpace,
        original: 'Section  1 applies here.',
        modified: 'Section  1 governs here.'
    },
    {
        name: 'double space introduced by the edit is preserved',
        oxml: paragraph(textRun('alpha beta gamma')),
        original: 'alpha beta gamma',
        modified: 'alpha  beta  gamma delta'
    },
    {
        name: 'w:tab survives an edit in an adjacent run',
        oxml: tabbedRuns,
        original: 'Name\told value',
        modified: 'Name\tnew value'
    },
    {
        name: 'trailing space is preserved',
        oxml: trailingSpace,
        original: 'alpha beta ',
        modified: 'alpha gamma '
    },
    {
        name: 'w:br survives an edit in an adjacent run',
        oxml: lineBreak,
        original: 'first line\nsecond line',
        modified: 'first line\nsecond row'
    },
    {
        name: 'multi-paragraph body, edit in the second paragraph',
        oxml: twoParagraphsBody,
        original: 'para one\npara two',
        modified: 'para one\npara three'
    },
    {
        name: 'w:sectPr stays last in the body after a reconstruction edit',
        oxml: twoParagraphsWithSectPr,
        original: 'para one\npara two',
        modified: 'para one\npara three'
    },
    {
        name: 'leading whitespace is preserved',
        oxml: leadingSpace,
        original: '  indented text',
        modified: '  indented copy'
    }
];

async function runCorpus() {
    const skipped = [];
    for (const testCase of corpus) {
        if (testCase.knownGap) {
            skipped.push(`${testCase.name} -- KNOWN-GAP: ${testCase.knownGap}`);
            continue;
        }
        await assertRoundTrip(testCase.oxml, testCase.original, testCase.modified);
    }
    for (const entry of skipped) {
        console.warn(`  SKIP: ${entry}`);
    }
}

async function runConsecutiveEdits() {
    const first = await assertRoundTrip(
        singleRun,
        'The quick brown fox.',
        'The quick red fox.'
    );
    const accepted = acceptTrackedChangesInOoxml(first.redlined.oxml, { allAuthors: true });

    assert.ok(accepted.oxml.includes('red'), 'accepted first edit should contain the first modified text');

    await assertRoundTrip(
        accepted.oxml,
        'The quick red fox.',
        'The quick silver fox.'
    );
}

async function runPartialTargetPreservation() {
    const untouchedSecond = '<w:p w:rsidR="22222222"><w:r><w:t>untouched two</w:t></w:r></w:p>';
    const untouchedThird = '<w:p w:rsidR="33333333"><w:r><w:t>untouched three</w:t></w:r></w:p>';
    const source = document(`${paragraph(textRun('alpha beta gamma'))}${untouchedSecond}${untouchedThird}`);

    const result = await applyRedlineToOxml(source, 'alpha beta gamma', 'alpha beta delta', {
        generateRedlines: true,
        author: 'RoundTrip'
    });

    assert.equal(result.status, 'ok');
    assert.ok(result.oxml.includes(untouchedSecond), 'second untargeted paragraph should survive byte-identical');
    assert.ok(result.oxml.includes(untouchedThird), 'third untargeted paragraph should survive byte-identical');
    assertRoundTripStructure(result.oxml);
}

async function runSoftBreakPreservation() {
    const result = await assertRoundTrip(
        lineBreak,
        'first line\nsecond line',
        'first line\nsecond row'
    );
    const doc = parseXmlFragment(result.redlined.oxml);
    assert.equal(elementsByLocalName(doc, 'br').length, 1, 'adjacent edit should preserve the original w:br');
    assert.equal(elementsByLocalName(doc, 'p').length, 1, 'soft break must not create another paragraph');
}

async function run() {
    await runCorpus();
    await runConsecutiveEdits();
    await runPartialTargetPreservation();
    await runSoftBreakPreservation();
    console.log('PASS: round-trip invariant tests');
}

run().catch(err => {
    console.error('FAIL:', err.message);
    process.exit(1);
});
