import './setup-xml-provider.mjs';

import assert from 'assert/strict';
import { applyRedlineToOxml } from '../engine/oxml-engine.js';
import { applySurgicalMode } from '../engine/surgical-mode.js';
import {
    directChildByLocalName,
    elementsByLocalName,
    formatIsEnabled,
    parseXml,
    serializeXml,
    textIndex
} from './helpers/ooxml-assertions.mjs';

const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

function findRunWithText(doc, expectedText) {
    return elementsByLocalName(doc, 'r').find(run =>
        elementsByLocalName(run, 't').map(node => node.textContent || '').join('') === expectedText
    ) || null;
}

function assertHeadingFormattingPreserved(xml, headingText) {
    const doc = parseXml(xml);
    const run = findRunWithText(doc, headingText);
    assert(run, `Expected an unchanged run containing "${headingText}"`);

    const rPr = directChildByLocalName(run, 'rPr');
    assert(rPr, 'Expected unchanged heading run properties to survive');
    assert.equal(formatIsEnabled(rPr, 'b'), true, 'Expected unchanged heading to remain bold');

    const fonts = directChildByLocalName(rPr, 'rFonts');
    assert(fonts, 'Expected unchanged font family to survive');
    assert.equal(fonts.getAttribute('w:ascii') || fonts.getAttribute('ascii'), 'Aptos');
    assert.equal(elementsByLocalName(rPr, 'rPrChange').length, 0,
        'Preserving existing formatting must not create a formatting revision');
}

async function testSurgicalInsertionSplitsRunAtCharacterOffset() {
    const xml = `
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:body>
                <w:p><w:r><w:t>Hello world</w:t></w:r></w:p>
            </w:body>
        </w:document>
    `;
    const doc = parseXml(xml);
    const result = applySurgicalMode(doc, 'Hello world', 'Hello brave world', new XMLSerializer(), 'Tester', [], true);

    assert.equal(result.hasChanges, true);
    assert.equal(result.sourceType, 'document');
    assert(result.oxml.includes('<w:ins'), 'Expected insertion redline');
    assert(textIndex(result.oxml, 'Hello ') < textIndex(result.oxml, 'brave '));
    assert(textIndex(result.oxml, 'brave ') < textIndex(result.oxml, 'world'));
}

async function testSurgicalDeletionPreservesNeighboringRunChildren() {
    const xml = `
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:body>
                <w:p><w:r><w:t>A</w:t><w:tab/><w:t>B</w:t></w:r></w:p>
            </w:body>
        </w:document>
    `;
    const doc = parseXml(xml);
    const result = applySurgicalMode(doc, 'A\tB', 'AB', new XMLSerializer(), 'Tester', [], true);

    assert.equal(result.hasChanges, true);
    assert(result.oxml.includes('<w:del'), 'Expected deletion redline for replaced run');
    assert(result.oxml.includes('<w:ins'), 'Expected insertion redline for replacement text');
    assert(result.oxml.includes('>AB</w:t>'), 'Expected replacement text after multi-child run deletion');
}

async function testEmptyDefaultNamespaceTableCellCanReceiveInsertion() {
    const xml = `
        <document xmlns="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <body>
                <tbl>
                    <tr>
                        <tc><p><r><t></t></r></p></tc>
                    </tr>
                </tbl>
            </body>
        </document>
    `;
    const result = await applyRedlineToOxml(xml, '', 'New cell text', {
        author: 'Tester',
        generateRedlines: true
    });
    const parsed = parseXml(result.oxml);

    assert.equal(result.hasChanges, true);
    assert.equal(result.sourceType, 'document');
    assert(serializeXml(parsed).includes('New cell text'), 'Expected inserted text in empty table cell');
    assert(serializeXml(parsed).includes('<w:ins') || serializeXml(parsed).includes('<ins'), 'Expected tracked insertion');
}

async function testApplyRedlineReportsFragmentSourceType() {
    const xml = `
        <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:r><w:t>Hello</w:t></w:r>
        </w:p>
    `;
    const result = await applyRedlineToOxml(xml, 'Hello', 'Hello world', {
        author: 'Tester',
        generateRedlines: true
    });

    assert.equal(result.hasChanges, true);
    assert.equal(result.sourceType, 'fragment');
}

async function testHyperlinkEditKeepsRevisionsInsideHyperlink() {
    const xml = `
        <w:document xmlns:w="${NS_W}" xmlns:r="${NS_R}">
            <w:body>
                <w:p>
                    <w:r><w:t xml:space="preserve">before </w:t></w:r>
                    <w:hyperlink r:id="rId7">
                        <w:r><w:t>link text</w:t></w:r>
                    </w:hyperlink>
                    <w:r><w:t xml:space="preserve"> after</w:t></w:r>
                </w:p>
            </w:body>
        </w:document>
    `;

    const result = await applyRedlineToOxml(xml, 'before link text after', 'before hyperlink text after', {
        author: 'Tester',
        generateRedlines: true
    });
    const doc = parseXml(result.oxml);
    const hyperlinks = doc.getElementsByTagNameNS(NS_W, 'hyperlink');

    assert.equal(result.hasChanges, true);
    assert.ok(hyperlinks.length >= 1, 'Expected hyperlink to survive');
    assert.ok(Array.from(hyperlinks).every(link => link.getAttribute('r:id') === 'rId7'), 'Expected hyperlink relationship id to survive');
    assert.ok(Array.from(hyperlinks).some(link => link.getElementsByTagNameNS(NS_W, 'ins').length > 0), 'Expected insertion inside hyperlink');
    assert.ok(Array.from(hyperlinks).some(link => link.getElementsByTagNameNS(NS_W, 'del').length > 0), 'Expected deletion inside hyperlink');
}

async function testCommentAndBookmarkMarkersSurviveEdit() {
    const xml = `
        <w:document xmlns:w="${NS_W}">
            <w:body>
                <w:p>
                    <w:bookmarkStart w:id="3" w:name="bm"/>
                    <w:commentRangeStart w:id="9"/>
                    <w:r><w:t>target</w:t></w:r>
                    <w:commentRangeEnd w:id="9"/>
                    <w:r><w:commentReference w:id="9"/></w:r>
                    <w:bookmarkEnd w:id="3"/>
                </w:p>
            </w:body>
        </w:document>
    `;

    const result = await applyRedlineToOxml(xml, 'target', 'updated target', {
        author: 'Tester',
        generateRedlines: true
    });
    const doc = parseXml(result.oxml);

    assert.equal(result.hasChanges, true);
    assert.equal(doc.getElementsByTagNameNS(NS_W, 'bookmarkStart').length, 1);
    assert.equal(doc.getElementsByTagNameNS(NS_W, 'bookmarkEnd').length, 1);
    assert.equal(doc.getElementsByTagNameNS(NS_W, 'commentRangeStart').length, 1);
    assert.equal(doc.getElementsByTagNameNS(NS_W, 'commentRangeEnd').length, 1);
    assert.equal(doc.getElementsByTagNameNS(NS_W, 'commentReference').length, 1);
}

async function testTabSurvivesAdjacentEdit() {
    const xml = `
        <w:document xmlns:w="${NS_W}">
            <w:body>
                <w:p><w:r><w:t>A</w:t><w:tab/><w:t>B</w:t></w:r></w:p>
            </w:body>
        </w:document>
    `;

    const result = await applyRedlineToOxml(xml, 'A\tB', 'A\tC', {
        author: 'Tester',
        generateRedlines: true
    });
    const doc = parseXml(result.oxml);

    assert.equal(result.hasChanges, true);
    assert.equal(doc.getElementsByTagNameNS(NS_W, 'tab').length, 1, 'Expected tab to survive adjacent edit');
}

async function testFootnoteReferenceSurvivesEditAfterReference() {
    const xml = `
        <w:document xmlns:w="${NS_W}">
            <w:body>
                <w:p>
                    <w:r><w:t xml:space="preserve">See </w:t></w:r>
                    <w:r><w:footnoteReference w:id="5"/></w:r>
                    <w:r><w:t xml:space="preserve"> clause</w:t></w:r>
                </w:p>
            </w:body>
        </w:document>
    `;

    const result = await applyRedlineToOxml(xml, 'See clause', 'See updated clause', {
        author: 'Tester',
        generateRedlines: true
    });
    const doc = parseXml(result.oxml);
    const refs = doc.getElementsByTagNameNS(NS_W, 'footnoteReference');

    assert.equal(result.hasChanges, true);
    assert.equal(refs.length, 1, 'Expected exactly one footnote reference');
    assert.equal(refs[0].getAttribute('w:id'), '5');
    assert.equal(elementsByLocalName(doc, 'del').some(node => node.getElementsByTagNameNS(NS_W, 'footnoteReference').length > 0), false, 'Footnote reference must not be wrapped in deletion');
}

async function testReconstructionPreservesUnchangedInlineFormatting() {
    const xml = `
        <w:p xmlns:w="${NS_W}">
            <w:r>
                <w:rPr><w:rFonts w:ascii="Aptos"/><w:b/></w:rPr>
                <w:t>Incident Management:</w:t>
            </w:r>
            <w:r><w:t xml:space="preserve"> existing text.</w:t></w:r>
        </w:p>
    `;

    const result = await applyRedlineToOxml(
        xml,
        'Incident Management: existing text.',
        'Incident Management: existing text. Added clause.',
        { author: 'Tester', generateRedlines: true }
    );

    assert.equal(result.hasChanges, true);
    assertHeadingFormattingPreserved(result.oxml, 'Incident Management:');
}

async function testSurgicalModePreservesUnchangedInlineFormatting() {
    const xml = `
        <w:document xmlns:w="${NS_W}">
            <w:body>
                <w:p>
                    <w:r>
                        <w:rPr><w:rFonts w:ascii="Aptos"/><w:b/></w:rPr>
                        <w:t>Incident Management:</w:t>
                    </w:r>
                    <w:r><w:t xml:space="preserve"> existing text.</w:t></w:r>
                </w:p>
            </w:body>
        </w:document>
    `;
    const doc = parseXml(xml);
    const result = applySurgicalMode(
        doc,
        'Incident Management: existing text.',
        'Incident Management: existing text. Added clause.',
        new XMLSerializer(),
        'Tester',
        [],
        true
    );

    assert.equal(result.hasChanges, true);
    assertHeadingFormattingPreserved(result.oxml, 'Incident Management:');
}

await testSurgicalInsertionSplitsRunAtCharacterOffset();
await testSurgicalDeletionPreservesNeighboringRunChildren();
await testEmptyDefaultNamespaceTableCellCanReceiveInsertion();
await testApplyRedlineReportsFragmentSourceType();
await testHyperlinkEditKeepsRevisionsInsideHyperlink();
await testCommentAndBookmarkMarkersSurviveEdit();
await testTabSurvivesAdjacentEdit();
await testFootnoteReferenceSurvivesEditAfterReference();
await testReconstructionPreservesUnchangedInlineFormatting();
await testSurgicalModePreservesUnchangedInlineFormatting();

console.log('engine_reliability_tests.mjs ... PASS');
