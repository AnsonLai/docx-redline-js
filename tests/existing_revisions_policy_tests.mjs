import './setup-xml-provider.mjs';

import assert from 'assert/strict';

import {
    applyRedlineToOxml,
    containsTrackedChanges,
    ingestWordOoxmlToPlainText
} from '../index.js';
import { assertRoundTrip } from './helpers/roundtrip.mjs';

const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

function parse(xml) {
    const doc = new DOMParser().parseFromString(xml, 'application/xml');
    const parseError = doc.getElementsByTagName('parsererror')[0];
    assert(!parseError, parseError?.textContent || 'XML parse error');
    return doc;
}

function paragraph(inner) {
    return `<w:p xmlns:w="${NS_W}">${inner}</w:p>`;
}

function revisionParagraph() {
    return paragraph([
        '<w:r><w:t xml:space="preserve">A </w:t></w:r>',
        '<w:del w:id="1" w:author="Human" w:date="2026-01-01T00:00:00Z"><w:r><w:delText>old</w:delText></w:r></w:del>',
        '<w:ins w:id="2" w:author="Human" w:date="2026-01-01T00:00:00Z"><w:r><w:t>new</w:t></w:r></w:ins>',
        '<w:r><w:t xml:space="preserve"> end</w:t></w:r>'
    ].join(''));
}

async function testDefaultNormalizesBeforeRedlining() {
    const source = revisionParagraph();
    const result = await assertRoundTrip(source, 'A new end', 'A newer end');

    assert.equal(result.redlined.hasChanges, true);
    assert.ok(!result.redlined.oxml.includes('w:author="Human"'), 'input revisions should be normalized before new redlines are generated');
    assert.ok(result.redlined.oxml.includes('w:author="RoundTrip"'), 'new revisions should use the caller author');
}

async function testExplicitRejectInputRejectsExistingRevisions() {
    const source = revisionParagraph();
    const result = await applyRedlineToOxml(source, 'A new end', 'A newer end', {
        generateRedlines: true,
        author: 'Agent',
        existingRevisions: 'reject-input'
    });

    assert.equal(result.status, 'error');
    assert.equal(result.error?.code, 'EXISTING_REVISIONS');
    assert.equal(result.hasChanges, false);
    assert.equal(result.oxml, source);
}

function testPlainTextTreatsExistingRevisionsLikeAcceptedView() {
    const source = revisionParagraph();
    assert.equal(ingestWordOoxmlToPlainText(source), 'A new end');
}

function testContainsTrackedChangesPositiveMarkers() {
    const cases = [
        ['ins', paragraph('<w:ins w:id="1" w:author="A" w:date="2026-01-01T00:00:00Z"><w:r><w:t>x</w:t></w:r></w:ins>')],
        ['del', paragraph('<w:del w:id="1" w:author="A" w:date="2026-01-01T00:00:00Z"><w:r><w:delText>x</w:delText></w:r></w:del>')],
        ['moveFrom', paragraph('<w:moveFrom w:id="1" w:author="A" w:date="2026-01-01T00:00:00Z"><w:r><w:delText>x</w:delText></w:r></w:moveFrom>')],
        ['moveTo', paragraph('<w:moveTo w:id="1" w:author="A" w:date="2026-01-01T00:00:00Z"><w:r><w:t>x</w:t></w:r></w:moveTo>')],
        ['rPrChange', paragraph('<w:r><w:rPr><w:rPrChange w:id="1" w:author="A" w:date="2026-01-01T00:00:00Z"><w:rPr/></w:rPrChange></w:rPr><w:t>x</w:t></w:r>')],
        ['pPrChange', paragraph('<w:pPr><w:pPrChange w:id="1" w:author="A" w:date="2026-01-01T00:00:00Z"><w:pPr/></w:pPrChange></w:pPr><w:r><w:t>x</w:t></w:r>')],
        ['cellIns', `<w:tbl xmlns:w="${NS_W}"><w:tr><w:tc><w:tcPr><w:cellIns w:id="1" w:author="A" w:date="2026-01-01T00:00:00Z"/></w:tcPr><w:p><w:r><w:t>x</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`],
        ['cellDel', `<w:tbl xmlns:w="${NS_W}"><w:tr><w:tc><w:tcPr><w:cellDel w:id="1" w:author="A" w:date="2026-01-01T00:00:00Z"/></w:tcPr><w:p><w:r><w:t>x</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`],
        ['paragraph mark ins', paragraph('<w:pPr><w:rPr><w:ins w:id="1" w:author="A" w:date="2026-01-01T00:00:00Z"/></w:rPr></w:pPr><w:r><w:t>x</w:t></w:r>')],
        ['paragraph mark del', paragraph('<w:pPr><w:rPr><w:del w:id="1" w:author="A" w:date="2026-01-01T00:00:00Z"/></w:rPr></w:pPr><w:r><w:t>x</w:t></w:r>')]
    ];

    for (const [label, xml] of cases) {
        assert.equal(containsTrackedChanges(parse(xml)), true, `Expected ${label} to be detected`);
    }
}

function testContainsTrackedChangesNegativeMarkers() {
    const clean = paragraph('<w:r><w:t>clean</w:t></w:r>');
    const commentsAndBookmarks = paragraph([
        '<w:bookmarkStart w:id="1" w:name="bm"/>',
        '<w:commentRangeStart w:id="2"/>',
        '<w:r><w:t>commented</w:t></w:r>',
        '<w:commentRangeEnd w:id="2"/>',
        '<w:r><w:commentReference w:id="2"/></w:r>',
        '<w:bookmarkEnd w:id="1"/>'
    ].join(''));

    assert.equal(containsTrackedChanges(parse(clean)), false);
    assert.equal(containsTrackedChanges(parse(commentsAndBookmarks)), false);
}

async function run() {
    await testDefaultNormalizesBeforeRedlining();
    await testExplicitRejectInputRejectsExistingRevisions();
    testPlainTextTreatsExistingRevisionsLikeAcceptedView();
    testContainsTrackedChangesPositiveMarkers();
    testContainsTrackedChangesNegativeMarkers();
    console.log('PASS: existing revisions policy tests');
}

run().catch(err => {
    console.error('FAIL:', err.message);
    process.exit(1);
});
