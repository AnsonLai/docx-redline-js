import './setup-xml-provider.mjs';

import assert from 'assert/strict';
import { applyRedlineToOxml } from '../engine/oxml-engine.js';
import { extractCanonicalParagraphText } from '../core/paragraph-text.js';
import { applyOperationToDocumentXml } from '../services/standalone-operation-runner.js';
import {
    acceptTrackedChangesInOoxml,
    ingestWordOoxmlToPlainText,
    rejectTrackedChangesInOoxml
} from '../index.js';
import {
    directChildByLocalName,
    elementsByLocalName,
    formatIsEnabled,
    parseXml
} from './helpers/ooxml-assertions.mjs';

const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

function normalizeText(text) {
    return text.replace(/\r\n/g, '\n').replace(/\r/g, '\n').trim();
}

/**
 * 1. Formatting Bleed / Leak Guardrails
 */

async function testFootnoteSuperscriptDoesNotBleedIntoReplacement() {
    // Word visual failure: Replacing text adjacent to a footnote reference can cause
    // the inserted run to inherit w:vertAlign="superscript", visibly shrinking and
    // elevating normal replacement text.
    const xml = `
        <w:document xmlns:w="${NS_W}">
            <w:body>
                <w:p>
                    <w:r><w:t>Review period expires on Friday</w:t></w:r>
                    <w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="1"/></w:r>
                    <w:r><w:t>.</w:t></w:r>
                </w:p>
            </w:body>
        </w:document>
    `;
    const original = 'Review period expires on Friday.';
    const modified = 'Review period expires on Monday afternoon.';
    const result = await applyRedlineToOxml(xml, original, modified, {
        author: 'VisualGuard',
        generateRedlines: true
    });

    assert.equal(result.hasChanges, true);
    const doc = parseXml(result.oxml);
    const insertedRuns = elementsByLocalName(doc, 'ins').flatMap(ins => elementsByLocalName(ins, 'r'));
    assert.ok(insertedRuns.length > 0, 'Expected tracked insertion runs');

    for (const run of insertedRuns) {
        const rPr = directChildByLocalName(run, 'rPr');
        const vertAlign = rPr ? directChildByLocalName(rPr, 'vertAlign') : null;
        assert.equal(
            vertAlign,
            null,
            'Replacement text adjacent to a footnote reference must not inherit vertAlign="superscript"'
        );
    }

    const accepted = acceptTrackedChangesInOoxml(result.oxml, { author: 'VisualGuard' });
    const rejected = rejectTrackedChangesInOoxml(result.oxml, { author: 'VisualGuard' });
    assert.equal(normalizeText(ingestWordOoxmlToPlainText(accepted.oxml)), normalizeText(modified));
    assert.equal(normalizeText(ingestWordOoxmlToPlainText(rejected.oxml)), normalizeText(original));
}

async function testHighlightDoesNotBleedIntoUnstyledInsertion() {
    // Word visual failure: Editing near highlighted text must not cause untargeted
    // plain insertions to become highlighted.
    const xml = `
        <w:document xmlns:w="${NS_W}">
            <w:body>
                <w:p>
                    <w:r><w:rPr><w:highlight w:val="yellow"/></w:rPr><w:t>Confidential Notice:</w:t></w:r>
                    <w:r><w:t xml:space="preserve"> draft for discussion.</w:t></w:r>
                </w:p>
            </w:body>
        </w:document>
    `;
    const original = 'Confidential Notice: draft for discussion.';
    const modified = 'Confidential Notice: draft for formal committee discussion.';
    const result = await applyRedlineToOxml(xml, original, modified, {
        author: 'VisualGuard',
        generateRedlines: true
    });

    assert.equal(result.hasChanges, true);
    const doc = parseXml(result.oxml);
    const insertedRuns = elementsByLocalName(doc, 'ins').flatMap(ins => elementsByLocalName(ins, 'r'));

    // The inserted word "formal committee " is in the unhighlighted region and must not be yellow
    for (const run of insertedRuns) {
        const text = elementsByLocalName(run, 't').map(t => t.textContent).join('');
        if (text.includes('formal')) {
            const rPr = directChildByLocalName(run, 'rPr');
            const highlight = rPr ? directChildByLocalName(rPr, 'highlight') : null;
            assert.equal(highlight, null, 'Inserted unstyled text must not inherit yellow highlight');
        }
    }
}

/**
 * 2. Font & Size Preservation Across Structural Boundaries
 */

async function testReplacementAdjacentToHyperlinkPreservesCustomFontSizeAndFont() {
    // Word visual failure: When an edit crosses between plain text and a hyperlink with custom size/font,
    // a forward-only or default fallback property lookup drops the new text to Word's 12pt Normal default.
    const xml = `
        <w:document xmlns:w="${NS_W}" xmlns:r="${NS_R}">
            <w:body>
                <w:p>
                    <w:r>
                        <w:rPr>
                            <w:rFonts w:ascii="Georgia" w:hAnsi="Georgia"/>
                            <w:sz w:val="28"/>
                            <w:szCs w:val="28"/>
                        </w:rPr>
                        <w:t xml:space="preserve">Refer to </w:t>
                    </w:r>
                    <w:hyperlink r:id="rId1">
                        <w:r>
                            <w:rPr>
                                <w:rFonts w:ascii="Georgia" w:hAnsi="Georgia"/>
                                <w:sz w:val="28"/>
                                <w:szCs w:val="28"/>
                                <w:color w:val="0000FF"/>
                                <w:u w:val="single"/>
                            </w:rPr>
                            <w:t>Exhibit A</w:t>
                        </w:r>
                    </w:hyperlink>
                    <w:r>
                        <w:rPr>
                            <w:rFonts w:ascii="Georgia" w:hAnsi="Georgia"/>
                            <w:sz w:val="28"/>
                            <w:szCs w:val="28"/>
                        </w:rPr>
                        <w:t xml:space="preserve"> for schedules.</w:t>
                    </w:r>
                </w:p>
            </w:body>
        </w:document>
    `;
    const original = 'Refer to Exhibit A for schedules.';
    const modified = 'Refer directly to Exhibit A for detailed schedules.';
    const result = await applyRedlineToOxml(xml, original, modified, {
        author: 'VisualGuard',
        generateRedlines: true
    });

    assert.equal(result.hasChanges, true);
    const doc = parseXml(result.oxml);
    const insertedRuns = elementsByLocalName(doc, 'ins').flatMap(ins => elementsByLocalName(ins, 'r'));
    assert.ok(insertedRuns.length > 0, 'Expected inserted runs');

    for (const run of insertedRuns) {
        const rPr = directChildByLocalName(run, 'rPr');
        assert.ok(rPr, 'Inserted run must retain explicit run properties to avoid dropping to 12pt Normal');
        const sz = directChildByLocalName(rPr, 'sz');
        const szVal = sz?.getAttribute('w:val') || sz?.getAttribute('val');
        assert.equal(szVal, '28', 'Inserted run must preserve 14pt (28 half-points) size');

        const rFonts = directChildByLocalName(rPr, 'rFonts');
        const fontName = rFonts?.getAttribute('w:ascii') || rFonts?.getAttribute('ascii');
        assert.equal(fontName, 'Georgia', 'Inserted run must preserve Georgia font');
    }
}

async function testHeadingReconstructionPreservesHeadingStyleAndRunFormatting() {
    // Word visual failure: Reconstructing a heading line must not strip heading paragraph styling
    // or reset heading typography to Normal body text.
    const xml = `
        <w:document xmlns:w="${NS_W}">
            <w:body>
                <w:p>
                    <w:pPr>
                        <w:pStyle w:val="Heading1"/>
                        <w:spacing w:before="240" w:after="120"/>
                    </w:pPr>
                    <w:r>
                        <w:rPr>
                            <w:rFonts w:ascii="Arial Black" w:hAnsi="Arial Black"/>
                            <w:b/>
                            <w:sz w:val="32"/>
                        </w:rPr>
                        <w:t>Article 1: Scope of Engagement</w:t>
                    </w:r>
                </w:p>
            </w:body>
        </w:document>
    `;
    const original = 'Article 1: Scope of Engagement';
    const modified = 'Article 1: Scope and Duration of Engagement';
    const result = await applyRedlineToOxml(xml, original, modified, {
        author: 'VisualGuard',
        generateRedlines: true
    });

    assert.equal(result.hasChanges, true);
    const doc = parseXml(result.oxml);
    const p = elementsByLocalName(doc, 'p')[0];
    const pPr = directChildByLocalName(p, 'pPr');
    assert.ok(pPr, 'Heading paragraph properties must survive');
    const pStyle = directChildByLocalName(pPr, 'pStyle');
    assert.equal(
        pStyle?.getAttribute('w:val') || pStyle?.getAttribute('val'),
        'Heading1',
        'Paragraph must retain Heading1 style'
    );

    const insertedRuns = elementsByLocalName(doc, 'ins').flatMap(ins => elementsByLocalName(ins, 'r'));
    for (const run of insertedRuns) {
        const rPr = directChildByLocalName(run, 'rPr');
        assert.ok(rPr, 'Inserted heading run must have rPr');
        assert.equal(formatIsEnabled(rPr, 'b'), true, 'Inserted heading run must remain bold');
        const sz = directChildByLocalName(rPr, 'sz');
        const szVal = sz?.getAttribute('w:val') || sz?.getAttribute('val');
        assert.equal(szVal, '32', 'Inserted heading run must remain 16pt (32 half-points)');
    }
}

/**
 * 3. Ghost Marker & List Boundary Invariants
 */

async function testInsertedListItemTracksParagraphMarkToPreventGhostMarker() {
    // Word visual failure: When an inserted list item tracks only its text run,
    // rejecting all revisions or deleting the list item leaves behind an untracked
    // paragraph mark, creating an empty "ghost bullet" or number in Word.
    const xml = `
        <w:document xmlns:w="${NS_W}">
            <w:body>
                <w:p>
                    <w:pPr>
                        <w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>
                    </w:pPr>
                    <w:r><w:t>First requirement</w:t></w:r>
                </w:p>
            </w:body>
        </w:document>
    `;
    const original = 'First requirement';
    const modified = 'First requirement\nSecond requirement';
    const result = await applyRedlineToOxml(xml, original, modified, {
        author: 'VisualGuard',
        generateRedlines: true
    });

    assert.equal(result.hasChanges, true);

    // Rejecting all revisions must cleanly remove the second item AND its paragraph mark
    const rejected = rejectTrackedChangesInOoxml(result.oxml, { author: 'VisualGuard' });
    const rejectedDoc = parseXml(rejected.oxml);
    const paragraphs = elementsByLocalName(rejectedDoc, 'p');
    assert.equal(
        paragraphs.length,
        1,
        'Rejecting inserted list item must leave exactly 1 paragraph; no ghost paragraph marker allowed'
    );
    assert.equal(
        normalizeText(ingestWordOoxmlToPlainText(rejected.oxml)),
        normalizeText(original),
        'Rejected text must restore original single item'
    );
}

async function testSuppressedHeadingExpandsIntoSeparateFormattedBullets() {
    // Word visual failure: a heading with numId="0" is numbering suppression,
    // not a reusable list. Reusing it hides bullets; placing the deletion in
    // the first new paragraph also renders PURPOSEArticle as one run-on line.
    const original = 'A.\tPURPOSE';
    const modified = '* Article A. Purpose and Interagency Alignment\n* Key Focus: Joint Street Outreach & Medical Triage';
    const following = 'The agencies shall coordinate implementation.';
    const xml = `
        <w:document xmlns:w="${NS_W}">
            <w:body>
                <w:p>
                    <w:pPr>
                        <w:pStyle w:val="Level1"/>
                        <w:numPr><w:ilvl w:val="0"/><w:numId w:val="0"/></w:numPr>
                        <w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="24"/><w:u w:val="single"/></w:rPr>
                    </w:pPr>
                    <w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="24"/></w:rPr><w:t>A.</w:t></w:r>
                    <w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="24"/></w:rPr><w:tab/></w:r>
                    <w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="24"/><w:u w:val="single"/></w:rPr><w:t>PURPOSE</w:t></w:r>
                </w:p>
                <w:p><w:r><w:t>${following}</w:t></w:r></w:p>
            </w:body>
        </w:document>
    `;
    const result = await applyOperationToDocumentXml(xml, {
        type: 'redline',
        target: original,
        modified
    }, 'VisualGuard');

    assert.equal(result.hasChanges, true);
    const doc = parseXml(result.documentXml);
    const paragraphs = elementsByLocalName(doc, 'p');
    assert.equal(paragraphs.length, 4, 'Full-document replacement must not add a packaging-only blank paragraph');

    const deletedHeading = paragraphs[0];
    const insertedParagraphs = paragraphs.slice(1, 3);
    assert.match(elementsByLocalName(deletedHeading, 'delText').map(node => node.textContent).join(''), /PURPOSE/);
    assert.doesNotMatch(
        elementsByLocalName(deletedHeading, 't').map(node => node.textContent).join(''),
        /Article A/,
        'Deleted heading and first bullet must occupy different paragraphs'
    );

    for (const paragraph of insertedParagraphs) {
        const pPr = directChildByLocalName(paragraph, 'pPr');
        const numPr = directChildByLocalName(pPr, 'numPr');
        const numId = directChildByLocalName(numPr, 'numId');
        assert.ok(Number.parseInt(numId?.getAttribute('w:val') || numId?.getAttribute('val'), 10) > 0,
            'Inserted bullet must use a positive numbering ID');

        const textInsertion = elementsByLocalName(paragraph, 'ins').find(ins => ins.parentNode === paragraph);
        const run = elementsByLocalName(textInsertion, 'r')[0];
        const rPr = directChildByLocalName(run, 'rPr');
        assert.equal(directChildByLocalName(rPr, 'rFonts')?.getAttribute('w:ascii'), 'Times New Roman');
        assert.equal(directChildByLocalName(rPr, 'sz')?.getAttribute('w:val'), '24');
        assert.equal(directChildByLocalName(rPr, 'b'), null, 'Heading bold must not leak into bullet text');
        assert.equal(directChildByLocalName(rPr, 'u'), null, 'Heading underline must not leak into bullet text');
    }

    const accepted = acceptTrackedChangesInOoxml(result.documentXml, { author: 'VisualGuard' });
    const rejected = rejectTrackedChangesInOoxml(result.documentXml, { author: 'VisualGuard' });
    const acceptedParagraphText = elementsByLocalName(parseXml(accepted.oxml), 'p')
        .map(paragraph => extractCanonicalParagraphText(paragraph));
    const rejectedParagraphText = elementsByLocalName(parseXml(rejected.oxml), 'p')
        .map(paragraph => extractCanonicalParagraphText(paragraph));
    assert.deepEqual(acceptedParagraphText, [
        'Article A. Purpose and Interagency Alignment',
        'Key Focus: Joint Street Outreach & Medical Triage',
        following
    ], 'Accept must produce exactly two list paragraphs without empty separator paragraphs');
    assert.deepEqual(rejectedParagraphText, [original, following],
        'Reject must restore the original heading paragraph exactly');
}

async function testListChangePreservesExistingNumberingIdAndLevel() {
    // Word visual failure: Adding a list item to an existing numbered list must preserve
    // the same numId and ilvl so that Word numbers it sequentially rather than breaking numbering.
    const xml = `
        <w:document xmlns:w="${NS_W}">
            <w:body>
                <w:p>
                    <w:pPr>
                        <w:numPr><w:ilvl w:val="1"/><w:numId w:val="7"/></w:numPr>
                    </w:pPr>
                    <w:r><w:t>Sub-item (a)</w:t></w:r>
                </w:p>
            </w:body>
        </w:document>
    `;
    const original = 'Sub-item (a)';
    const modified = 'Sub-item (a)\nSub-item (b)';
    const result = await applyRedlineToOxml(xml, original, modified, {
        author: 'VisualGuard',
        generateRedlines: true
    });

    assert.equal(result.hasChanges, true);
    const doc = parseXml(result.oxml);
    const allPPr = elementsByLocalName(doc, 'pPr');
    for (const pPr of allPPr) {
        const numPr = directChildByLocalName(pPr, 'numPr');
        if (numPr) {
            const ilvl = directChildByLocalName(numPr, 'ilvl');
            const numId = directChildByLocalName(numPr, 'numId');
            assert.equal(ilvl?.getAttribute('w:val') || ilvl?.getAttribute('val'), '1', 'ilvl must remain 1');
            assert.equal(numId?.getAttribute('w:val') || numId?.getAttribute('val'), '7', 'numId must remain 7');
        }
    }
}

/**
 * 4. Table Cell Layout Preservation
 */

async function testTableCellEditPreservesWidthBordersShadingAndAlignment() {
    // Word visual failure: Editing text inside a table cell must not discard the cell's
    // w:tcPr (cell width, borders, shading/background fill, and vertical alignment).
    const xml = `
        <w:document xmlns:w="${NS_W}">
            <w:body>
                <w:tbl>
                    <w:tblGrid>
                        <w:gridCol w:w="4800"/>
                        <w:gridCol w:w="4800"/>
                    </w:tblGrid>
                    <w:tr>
                        <w:tc>
                            <w:tcPr>
                                <w:tcW w:w="4800" w:type="dxa"/>
                                <w:tcBorders>
                                    <w:top w:val="single" w:sz="12" w:space="0" w:color="003366"/>
                                    <w:bottom w:val="single" w:sz="12" w:space="0" w:color="003366"/>
                                </w:tcBorders>
                                <w:shd w:val="clear" w:color="auto" w:fill="EBF1F5"/>
                                <w:vAlign w:val="center"/>
                            </w:tcPr>
                            <w:p>
                                <w:r><w:t>Project Milestone 1</w:t></w:r>
                            </w:p>
                        </w:tc>
                        <w:tc>
                            <w:tcPr>
                                <w:tcW w:w="4800" w:type="dxa"/>
                            </w:tcPr>
                            <w:p>
                                <w:r><w:t>Complete</w:t></w:r>
                            </w:p>
                        </w:tc>
                    </w:tr>
                </w:tbl>
            </w:body>
        </w:document>
    `;
    const original = 'Project Milestone 1';
    const modified = 'Project Milestone 1: Requirements Signed Off';
    const result = await applyOperationToDocumentXml(
        xml,
        { type: 'redline', target: original, modified: modified },
        'VisualGuard'
    );

    assert.equal(result.hasChanges, true);
    const doc = parseXml(result.documentXml);
    const firstCell = elementsByLocalName(doc, 'tc')[0];
    assert.ok(firstCell, 'Table cell must exist in document XML');
    const tcPr = directChildByLocalName(firstCell, 'tcPr');
    assert.ok(tcPr, 'Table cell properties (w:tcPr) must survive editing');

    const tcW = directChildByLocalName(tcPr, 'tcW');
    assert.equal(tcW?.getAttribute('w:w') || tcW?.getAttribute('w'), '4800', 'Cell width must survive');

    const tcBorders = directChildByLocalName(tcPr, 'tcBorders');
    assert.ok(tcBorders, 'Cell borders must survive');
    const topBorder = directChildByLocalName(tcBorders, 'top');
    assert.equal(topBorder?.getAttribute('w:color') || topBorder?.getAttribute('color'), '003366');

    const shd = directChildByLocalName(tcPr, 'shd');
    assert.equal(shd?.getAttribute('w:fill') || shd?.getAttribute('fill'), 'EBF1F5', 'Cell background shading must survive');

    const vAlign = directChildByLocalName(tcPr, 'vAlign');
    assert.equal(vAlign?.getAttribute('w:val') || vAlign?.getAttribute('val'), 'center', 'Vertical alignment must survive');
}

async function testUnderlineDoesNotBleedIntoSubsequentInsertion() {
    // Word visual failure: When editing text adjacent to underlined text, unstyled
    // trailing text must not inherit underline formatting.
    const xml = `
        <w:document xmlns:w="${NS_W}">
            <w:body>
                <w:p>
                    <w:r><w:rPr><w:u w:val="single"/></w:rPr><w:t>Section 4.1</w:t></w:r>
                    <w:r><w:t xml:space="preserve"> governs termination.</w:t></w:r>
                </w:p>
            </w:body>
        </w:document>
    `;
    const original = 'Section 4.1 governs termination.';
    const modified = 'Section 4.1 and Exhibit B govern termination.';
    const result = await applyRedlineToOxml(xml, original, modified, {
        author: 'VisualGuard',
        generateRedlines: true
    });

    assert.equal(result.hasChanges, true);
    const doc = parseXml(result.oxml);
    const insertedRuns = elementsByLocalName(doc, 'ins').flatMap(ins => elementsByLocalName(ins, 'r'));

    for (const run of insertedRuns) {
        const text = elementsByLocalName(run, 't').map(t => t.textContent).join('');
        if (text.includes('govern')) {
            const rPr = directChildByLocalName(run, 'rPr');
            const u = rPr ? directChildByLocalName(rPr, 'u') : null;
            assert.equal(u, null, 'Inserted plain verb must not inherit underline');
        }
    }
}

async function testListParagraphDeletionTracksParagraphMark() {
    // Word visual failure: Deleting a list item must track the paragraph mark deletion (pPrChange/del)
    // so Word does not leave an empty bullet or number behind when revisions are accepted.
    const xml = `
        <w:document xmlns:w="${NS_W}">
            <w:body>
                <w:p>
                    <w:pPr>
                        <w:numPr><w:ilvl w:val="0"/><w:numId w:val="5"/></w:numPr>
                    </w:pPr>
                    <w:r><w:t>Keep item 1</w:t></w:r>
                </w:p>
                <w:p>
                    <w:pPr>
                        <w:numPr><w:ilvl w:val="0"/><w:numId w:val="5"/></w:numPr>
                    </w:pPr>
                    <w:r><w:t>Delete item 2</w:t></w:r>
                </w:p>
            </w:body>
        </w:document>
    `;
    const original = 'Keep item 1\nDelete item 2';
    const modified = 'Keep item 1';
    const result = await applyRedlineToOxml(xml, original, modified, {
        author: 'VisualGuard',
        generateRedlines: true
    });

    assert.equal(result.hasChanges, true);
    const accepted = acceptTrackedChangesInOoxml(result.oxml, { author: 'VisualGuard' });
    const acceptedDoc = parseXml(accepted.oxml);
    const paragraphs = elementsByLocalName(acceptedDoc, 'p');
    assert.equal(
        paragraphs.length,
        1,
        'Accepted list item deletion must leave exactly 1 paragraph; no ghost bullet'
    );
    assert.equal(normalizeText(ingestWordOoxmlToPlainText(accepted.oxml)), 'Keep item 1');
}

// Execute test suite
await testFootnoteSuperscriptDoesNotBleedIntoReplacement();
await testHighlightDoesNotBleedIntoUnstyledInsertion();
await testUnderlineDoesNotBleedIntoSubsequentInsertion();
await testReplacementAdjacentToHyperlinkPreservesCustomFontSizeAndFont();
await testHeadingReconstructionPreservesHeadingStyleAndRunFormatting();
await testInsertedListItemTracksParagraphMarkToPreventGhostMarker();
await testSuppressedHeadingExpandsIntoSeparateFormattedBullets();
await testListParagraphDeletionTracksParagraphMark();
await testListChangePreservesExistingNumberingIdAndLevel();
await testTableCellEditPreservesWidthBordersShadingAndAlignment();

console.log('PASS: visual_failure_regression_tests.mjs - all semantic visual failure guards passed.');
