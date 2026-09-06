import './setup-xml-provider.mjs';

import assert from 'assert';
import {
    acceptTrackedChangesInOoxml,
    buildTargetReferenceSnapshot,
    getParagraphText,
    ingestWordOoxmlToPlainText,
    rejectTrackedChangesInOoxml
} from '../index.js';
import {
    applyOperationToDocumentXml,
    applyOperationsToDocumentXml,
    orderOperationsForStableTargets
} from '../services/standalone-operation-runner.js';
import { validateRedlineOoxml } from '../core/redline-validation.js';
import { createDynamicNumberingIdState } from '../services/numbering-helpers.js';

const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

function parseXmlStrict(xmlText, label) {
    const parser = new DOMParser();
    const xmlDoc = parser.parseFromString(xmlText, 'application/xml');
    const parseError = xmlDoc.getElementsByTagName('parsererror')[0];
    if (parseError) {
        throw new Error(`[XML parse error] ${label}: ${parseError.textContent || 'Unknown'}`);
    }
    return xmlDoc;
}

function buildDocumentXml(text) {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${NS_W}">
  <w:body>
    <w:p><w:r><w:t>${text}</w:t></w:r></w:p>
    <w:sectPr/>
  </w:body>
</w:document>`;
}

function buildParagraphDocumentXml(paragraphs) {
    const paragraphXml = paragraphs
        .map(text => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`)
        .join('');
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${NS_W}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
  <w:body>${paragraphXml}<w:sectPr/></w:body>
</w:document>`;
}

function buildNumberedListDocumentXml(items, numId = '77') {
    const paragraphs = items
        .map(item => `<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="${numId}"/></w:numPr></w:pPr><w:r><w:t>${item}</w:t></w:r></w:p>`)
        .join('');
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${NS_W}">
  <w:body>
    ${paragraphs}
    <w:sectPr/>
  </w:body>
</w:document>`;
}

function buildTwoColumnTitleTableDocumentXml() {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${NS_W}">
  <w:body>
    <w:tbl>
      <w:tr>
        <w:tc><w:p><w:r><w:t>Title:</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>Title:</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>
    <w:sectPr/>
  </w:body>
</w:document>`;
}

async function testRedlineOperation() {
    const sourceText = 'Alpha target text.';
    const modifiedText = 'Alpha target text updated.';
    const inputXml = buildDocumentXml(sourceText);
    const result = await applyOperationToDocumentXml(
        inputXml,
        {
            type: 'redline',
            target: sourceText,
            modified: modifiedText
        },
        'StandaloneRunnerTest',
        null,
        {
            generateRedlines: false
        }
    );

    assert.strictEqual(result.hasChanges, true, 'redline operation should report changes');
    const resultDoc = parseXmlStrict(result.documentXml, 'redline output');
    const paragraphs = Array.from(resultDoc.getElementsByTagNameNS(NS_W, 'p'));
    const firstParagraphText = getParagraphText(paragraphs[0]).trim();
    assert.strictEqual(firstParagraphText, modifiedText, 'redline operation should rewrite paragraph text');
}

async function testImplicitMultilineTargetReplacesWholeRange() {
    const original = 'The Clerk records the application.\nThe Director reviews the application.';
    const modified = 'The Clerk records the application.\nThe Director approves the application.';
    const paragraphs = original.split('\n')
        .map(text => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`)
        .join('');
    const inputXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`
        + `<w:document xmlns:w="${NS_W}"><w:body>${paragraphs}<w:sectPr/></w:body></w:document>`;

    const result = await applyOperationToDocumentXml(
        inputXml,
        { type: 'redline', target: original, modified },
        'StandaloneRunnerTest',
        null,
        { generateRedlines: true }
    );

    assert.strictEqual(result.hasChanges, true);
    const accepted = acceptTrackedChangesInOoxml(result.documentXml, { allAuthors: true });
    const rejected = rejectTrackedChangesInOoxml(result.documentXml, { allAuthors: true });
    const normalizeParagraphs = text => text.replace(/\n\n/g, '\n');
    assert.strictEqual(normalizeParagraphs(ingestWordOoxmlToPlainText(accepted.oxml)), modified);
    assert.strictEqual(normalizeParagraphs(ingestWordOoxmlToPlainText(rejected.oxml)), original);
}

async function testCommentOperation() {
    const sourceText = 'Comment target paragraph.';
    const inputXml = buildDocumentXml(sourceText);
    const result = await applyOperationToDocumentXml(
        inputXml,
        {
            type: 'comment',
            target: sourceText,
            textToComment: 'target',
            commentContent: 'Please review this term.'
        },
        'StandaloneRunnerTest'
    );

    assert.strictEqual(result.hasChanges, true, 'comment operation should report changes');
    assert.ok(result.commentsXml && result.commentsXml.includes('Please review this term.'), 'comment operation should emit comments xml');
}

async function testRangeListRedlineDoesNotDuplicateExistingItems() {
    const existingItems = [
        'Business plans, strategies, financial information, pricing, and marketing data.',
        'Technical data, specifications, designs, prototypes, software, algorithms, source code, and intellectual property.',
        'Information concerning the Disclosing Party\'s employees, contractors, customers, and suppliers.',
        'Any notes, analyses, compilations, studies, or other materials prepared by the Receiving Party that contain, reflect, or are derived from the foregoing.'
    ];
    const insertedItem = 'Photographs, videos, and other recordings of prototypes and physical hardware.';
    const modifiedText = [
        `1. ${existingItems[0]}`,
        `2. ${insertedItem}`,
        `3. ${existingItems[1]}`,
        `4. ${existingItems[2]}`,
        `5. ${existingItems[3]}`
    ].join('\n');
    const inputXml = buildNumberedListDocumentXml(existingItems);
    const logs = [];
    const result = await applyOperationToDocumentXml(
        inputXml,
        {
            type: 'redline',
            target: existingItems[0],
            targetRef: 'P1',
            targetEndRef: 'P4',
            modified: modifiedText
        },
        'StandaloneRunnerTest',
        null,
        {
            generateRedlines: true,
            onInfo: message => logs.push(String(message)),
            onWarn: message => logs.push(String(message))
        }
    );

    assert.strictEqual(result.hasChanges, true, 'range list redline should report changes');
    assert.strictEqual(
        logs.some(message => message.includes('Applying explicit-range insertion-only heuristic')),
        true,
        'range list redline should use explicit-range insertion-only heuristic'
    );

    const resultDoc = parseXmlStrict(result.documentXml, 'range list redline output');
    const paragraphs = Array.from(resultDoc.getElementsByTagNameNS(NS_W, 'p'));
    const revisionDeletes = resultDoc.getElementsByTagNameNS(NS_W, 'del').length;
    const revisionInserts = resultDoc.getElementsByTagNameNS(NS_W, 'ins').length;
    const paragraphMarkInserts = Array.from(resultDoc.getElementsByTagNameNS(NS_W, 'ins'))
        .filter(node => node.parentNode?.localName === 'rPr'
            && node.parentNode?.parentNode?.localName === 'pPr').length;
    const paragraphTexts = paragraphs.map(paragraph => getParagraphText(paragraph)).filter(Boolean);
    assert.strictEqual(
        paragraphTexts.length,
        5,
        'range list redline should produce five list paragraphs without duplicate tail items'
    );
    assert.strictEqual(
        paragraphTexts.filter(text => text === existingItems[1]).length,
        1,
        'existing item #2 should not be duplicated'
    );
    assert.strictEqual(
        paragraphTexts[1],
        insertedItem,
        'new confidentiality bullet should be inserted at position #2'
    );
    assert.strictEqual(
        revisionDeletes,
        0,
        'surgical insertion should not emit delete revisions for untouched list items'
    );
    assert.strictEqual(
        revisionInserts,
        2,
        'surgical insertion should track both inserted text and its paragraph mark'
    );
    assert.strictEqual(paragraphMarkInserts, 1,
        'range list insertion should track its paragraph mark so rejection removes the marker');

    const listNumIds = paragraphs.map(paragraph => {
        const numPr = paragraph.getElementsByTagNameNS(NS_W, 'numPr')[0];
        const numIdNode = numPr ? numPr.getElementsByTagNameNS(NS_W, 'numId')[0] : null;
        return numIdNode ? (numIdNode.getAttribute('w:val') || numIdNode.getAttribute('val') || null) : null;
    });
    assert.strictEqual(
        listNumIds.filter(numId => numId === '77').length,
        5,
        'surgical insertion should preserve original list numId across all items'
    );

    const rejected = rejectTrackedChangesInOoxml(result.documentXml, { author: 'StandaloneRunnerTest' });
    const rejectedDoc = parseXmlStrict(rejected.oxml, 'rejected range list insertion');
    const rejectedParagraphs = Array.from(rejectedDoc.getElementsByTagNameNS(NS_W, 'p'));
    assert.strictEqual(rejectedParagraphs.length, 4,
        'rejecting a range list insertion should remove its entire paragraph');
    assert.deepStrictEqual(rejectedParagraphs.map(getParagraphText), existingItems,
        'rejecting a range list insertion should restore the original list exactly');
}

async function testSingleParagraphListConcatenationUsesSurgicalInsertion() {
    const existingItems = [
        'Business plans, strategies, financial information, pricing, and marketing data.',
        'Technical data, specifications, designs, prototypes, software, algorithms, source code, and intellectual property.',
        'Information concerning the Disclosing Party\'s employees, contractors, customers, and suppliers.',
        'Any notes, analyses, compilations, studies, or other materials prepared by the Receiving Party that contain, reflect, or are derived from the foregoing.'
    ];
    const insertedItem = 'Photographs, videos, and other recordings of prototypes and physical hardware.';
    const inputXml = buildNumberedListDocumentXml(existingItems);
    const logs = [];
    const result = await applyOperationToDocumentXml(
        inputXml,
        {
            type: 'redline',
            target: existingItems[1],
            targetRef: 'P2',
            modified: `${insertedItem}${existingItems[1]}`
        },
        'StandaloneRunnerTest',
        null,
        {
            generateRedlines: true,
            onInfo: message => logs.push(String(message)),
            onWarn: message => logs.push(String(message))
        }
    );

    assert.strictEqual(result.hasChanges, true, 'single-paragraph concatenation edit should report changes');
    assert.strictEqual(
        logs.some(message => message.includes('single-paragraph list adjacency insertion heuristic')),
        true,
        'single-paragraph concatenation edit should use adjacency insertion heuristic'
    );

    const resultDoc = parseXmlStrict(result.documentXml, 'single paragraph list insertion output');
    const paragraphs = Array.from(resultDoc.getElementsByTagNameNS(NS_W, 'p'));
    const revisionDeletes = resultDoc.getElementsByTagNameNS(NS_W, 'del').length;
    const revisionInserts = resultDoc.getElementsByTagNameNS(NS_W, 'ins').length;
    const paragraphMarkInserts = Array.from(resultDoc.getElementsByTagNameNS(NS_W, 'ins'))
        .filter(node => node.parentNode?.localName === 'rPr'
            && node.parentNode?.parentNode?.localName === 'pPr').length;
    const paragraphTexts = paragraphs.map(paragraph => getParagraphText(paragraph)).filter(Boolean);

    assert.strictEqual(paragraphTexts.length, 5, 'single-paragraph concatenation should become one inserted list item');
    assert.strictEqual(paragraphTexts[1], insertedItem, 'new item should be inserted directly before original target item');
    assert.strictEqual(
        paragraphTexts.filter(text => text === existingItems[1]).length,
        1,
        'original target item should remain a single untouched list item'
    );
    assert.strictEqual(revisionDeletes, 0, 'adjacency insertion should not emit delete revisions');
    assert.strictEqual(revisionInserts, 2,
        'adjacency insertion should track both inserted text and its paragraph mark');
    assert.strictEqual(paragraphMarkInserts, 1,
        'inserted list item should track its paragraph mark so rejection removes the marker');

    const rejected = rejectTrackedChangesInOoxml(result.documentXml, { author: 'StandaloneRunnerTest' });
    const rejectedDoc = parseXmlStrict(rejected.oxml, 'rejected single paragraph list insertion');
    const rejectedParagraphs = Array.from(rejectedDoc.getElementsByTagNameNS(NS_W, 'p'));
    assert.strictEqual(rejectedParagraphs.length, 4,
        'rejecting a list insertion should remove its entire paragraph without a ghost list marker');
    assert.deepStrictEqual(rejectedParagraphs.map(getParagraphText), existingItems,
        'rejecting a list insertion should restore the original list exactly');
}

async function testSingleParagraphListConcatenationWithInlineMarkersDoesNotInsertExtraItem() {
    const existingItems = [
        'The Disclosing Party possesses certain confidential, proprietary, and trade secret information.',
        'The Parties desire to enter into a potential business relationship or transaction (the “Purpose”), which requires the Disclosing Party to disclose certain Confidential Information (as defined below) to the Receiving Party.',
        'The Receiving Party agrees to receive and treat such Confidential Information in confidence, subject to the terms and conditions of this Agreement.'
    ];
    const inputXml = buildNumberedListDocumentXml(existingItems);
    const malformedMergedEdit = [
        'A. The Disclosing Party possesses certain confidential, proprietary, and trade secret information.',
        'B. The Parties desire to enter into a potential business relationship or transaction (the “Purpose”), which requires the Disclosing Party to disclose certain Confidential Information (as defined below) to the Receiving Party.',
        'C. The Receiving Party agrees to receive and treat such Confidential Information in confidence, subject to the terms and conditions of this Agreement.',
        existingItems[0]
    ].join(' ');

    const logs = [];
    const result = await applyOperationToDocumentXml(
        inputXml,
        {
            type: 'redline',
            target: existingItems[0],
            targetRef: 'P1',
            modified: malformedMergedEdit
        },
        'StandaloneRunnerTest',
        null,
        {
            generateRedlines: true,
            onInfo: message => logs.push(String(message)),
            onWarn: message => logs.push(String(message))
        }
    );

    assert.strictEqual(result.hasChanges, true, 'malformed single-paragraph list edit should still report changes');
    assert.strictEqual(
        logs.some(message => message.includes('single-paragraph list adjacency insertion heuristic')),
        false,
        'malformed inline list markers should not trigger adjacency insertion heuristic'
    );

    const resultDoc = parseXmlStrict(result.documentXml, 'single paragraph malformed list edit output');
    const paragraphs = Array.from(resultDoc.getElementsByTagNameNS(NS_W, 'p'));
    const paragraphTexts = paragraphs.map(paragraph => getParagraphText(paragraph)).filter(Boolean);
    assert.strictEqual(
        paragraphTexts.length,
        3,
        'malformed single-paragraph list edit should not insert an extra list paragraph'
    );
}

async function testPlainParagraphInsertionBeforeTargetCreatesSeparateParagraph() {
    const original = 'NON-DISCLOSURE AGREEMENT';
    const insertedMarkdown = '**INSTRUCTIONS:** Please review this Non-Disclosure Agreement carefully.';
    const insertedPlain = 'INSTRUCTIONS: Please review this Non-Disclosure Agreement carefully.';
    const logs = [];
    const inputXml = buildDocumentXml(original);
    const result = await applyOperationToDocumentXml(
        inputXml,
        {
            type: 'redline',
            target: original,
            targetRef: 'P1',
            modified: `${insertedMarkdown}\n${original}`
        },
        'StandaloneRunnerTest',
        null,
        {
            generateRedlines: true,
            onInfo: message => logs.push(String(message)),
            onWarn: message => logs.push(String(message))
        }
    );

    assert.strictEqual(result.hasChanges, true, 'plain insertion-before-target shape should report changes');
    assert.strictEqual(
        logs.some(message => message.includes('plain adjacency insertion heuristic')),
        true,
        'plain insertion-before-target shape should use plain adjacency insertion heuristic'
    );

    const resultDoc = parseXmlStrict(result.documentXml, 'plain insertion-before-target output');
    const paragraphs = Array.from(resultDoc.getElementsByTagNameNS(NS_W, 'p'));
    const paragraphTexts = paragraphs.map(paragraph => getParagraphText(paragraph).trim()).filter(Boolean);

    assert.strictEqual(
        paragraphTexts.length,
        2,
        'plain insertion-before-target should produce two paragraphs'
    );
    assert.strictEqual(
        paragraphTexts[0],
        insertedPlain,
        'plain insertion-before-target should insert new paragraph text before original target paragraph'
    );
    assert.strictEqual(
        paragraphTexts[1],
        original,
        'plain insertion-before-target should preserve original target paragraph as separate paragraph'
    );
    assert.ok(
        result.documentXml.includes('<w:b'),
        'plain insertion-before-target should parse markdown formatting via existing formatter pipeline'
    );
    assert.ok(
        !paragraphTexts[0].includes('**'),
        'plain insertion-before-target should not leave raw markdown markers in output text'
    );
}

async function testFormatOnlyRedlineWithTrackedWrapperStillApplies() {
    const originalText = 'These instructions are for the user to fill out the document. Please replace all bracketed information (e.g., "[Name of Disclosing Party]") with the appropriate details. Ensure all necessary signatures are obtained. NON-DISCLOSURE AGREEMENT';
    const modifiedText = 'These instructions are for the user to fill out the document. Please replace all bracketed information (e.g., "[Name of Disclosing Party]") with the appropriate details. Ensure all necessary signatures are obtained. ++NON-DISCLOSURE AGREEMENT++';
    const inputXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${NS_W}">
  <w:body>
    <w:p>
      <w:ins w:id="1" w:author="Prior" w:date="2026-01-01T00:00:00Z">
        <w:r><w:t>These instructions are for the user to fill out the document. Please replace all bracketed information (e.g., "[Name of Disclosing Party]") with the appropriate details. Ensure all necessary signatures are obtained. </w:t></w:r>
      </w:ins>
      <w:r><w:t>NON-DISCLOSURE AGREEMENT</w:t></w:r>
    </w:p>
    <w:sectPr/>
  </w:body>
</w:document>`;

    const result = await applyOperationToDocumentXml(
        inputXml,
        {
            type: 'redline',
            target: originalText,
            targetRef: 'P1',
            modified: modifiedText
        },
        'StandaloneRunnerTest',
        null,
        {
            generateRedlines: true,
            existingRevisions: 'accept-all-first'
        }
    );

    assert.strictEqual(
        result.hasChanges,
        true,
        'format-only redline should apply when existing tracked-change wrappers are normalized first'
    );
    assert.ok(
        result.documentXml.includes('<w:u'),
        'format-only redline should apply underline markup to the target run'
    );
}

async function testTextToTableWithoutHeaderSeparatorPreservesAllRows() {
    const inputXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${NS_W}">
  <w:body>
    <w:p><w:r><w:t>Disclosing Party: [Name of Disclosing Party]</w:t></w:r></w:p>
    <w:p><w:r><w:t>And</w:t></w:r></w:p>
    <w:p><w:r><w:t>Receiving Party: [Name of Receiving Party]</w:t></w:r></w:p>
    <w:sectPr/>
  </w:body>
</w:document>`;

    const modifiedText = [
        '| Disclosing Party: | [Name of Disclosing Party] [Address of Disclosing Party] (the "Disclosing Party") |',
        '| Receiving Party: | [Name of Receiving Party] [Address of Receiving Party] (the "Receiving Party") |'
    ].join('\n');

    const result = await applyOperationToDocumentXml(
        inputXml,
        {
            type: 'redline',
            targetRef: 'P1',
            targetEndRef: 'P3',
            target: 'Disclosing Party: [Name of Disclosing Party]',
            modified: modifiedText
        },
        'StandaloneRunnerTest',
        null,
        {
            generateRedlines: true
        }
    );

    assert.strictEqual(result.hasChanges, true, 'text-to-table redline should report changes');
    const resultDoc = parseXmlStrict(result.documentXml, 'text-to-table no-header output');
    const tables = Array.from(resultDoc.getElementsByTagNameNS(NS_W, 'tbl'));
    assert.strictEqual(tables.length, 1, 'text-to-table redline should produce one table');
    const tableText = Array.from(tables[0].getElementsByTagNameNS(NS_W, 't'))
        .map(node => String(node.textContent || ''))
        .join(' ');
    assert.ok(
        tableText.includes('Disclosing Party:'),
        'table should include first markdown row text when no header separator is provided'
    );
    assert.ok(
        tableText.includes('Receiving Party:'),
        'table should include second markdown row text'
    );
}

async function testFormatOnlyRedlineSupportsNonWPrefixOoxml() {
    const inputXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<x:document xmlns:x="${NS_W}">
  <x:body>
    <x:p><x:r><x:t>By</x:t></x:r></x:p>
    <x:sectPr/>
  </x:body>
</x:document>`;

    const result = await applyOperationToDocumentXml(
        inputXml,
        {
            type: 'redline',
            targetRef: 'P1',
            target: 'By',
            modified: '**By**'
        },
        'StandaloneRunnerTest',
        null,
        {
            generateRedlines: true
        }
    );

    assert.strictEqual(
        result.hasChanges,
        true,
        'format-only redline should apply even when OOXML uses a non-w namespace prefix'
    );
    const resultDoc = parseXmlStrict(result.documentXml, 'non-w-prefix format-only output');
    assert.ok(
        resultDoc.getElementsByTagNameNS(NS_W, 'b').length > 0,
        'format-only redline should emit bold run properties on non-w-prefix OOXML'
    );
}

async function testFormatOnlyRedlineWithAllTextInsideInsertionWrapper() {
    const inputXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${NS_W}">
  <w:body>
    <w:p>
      <w:ins w:id="9" w:author="Prior" w:date="2026-01-01T00:00:00Z">
        <w:r><w:t>By: [Name]</w:t></w:r>
      </w:ins>
    </w:p>
    <w:sectPr/>
  </w:body>
</w:document>`;

    const result = await applyOperationToDocumentXml(
        inputXml,
        {
            type: 'redline',
            targetRef: 'P1',
            target: 'By: [Name]',
            modified: '**By**: [Name]'
        },
        'StandaloneRunnerTest',
        null,
        {
            generateRedlines: true,
            existingRevisions: 'accept-all-first'
        }
    );

    assert.strictEqual(
        result.hasChanges,
        true,
        'format-only redline should apply when all text is nested inside insertion wrapper and existing revisions are normalized first'
    );
    const resultDoc = parseXmlStrict(result.documentXml, 'format-only insertion wrapper output');
    assert.ok(
        resultDoc.getElementsByTagNameNS(NS_W, 'b').length > 0,
        'format-only redline should emit bold formatting even when text is wrapped in w:ins'
    );
}

async function testFormatOnlyRedlineRematchesWhenRefParagraphDrifts() {
    const inputXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${NS_W}">
  <w:body>
    <w:p><w:r><w:t></w:t></w:r></w:p>
    <w:p><w:r><w:t>By: [Name]</w:t></w:r></w:p>
    <w:sectPr/>
  </w:body>
</w:document>`;

    const result = await applyOperationToDocumentXml(
        inputXml,
        {
            type: 'redline',
            targetRef: 'P1',
            target: 'By: [Name]',
            modified: '**By**: [Name]'
        },
        'StandaloneRunnerTest',
        null,
        {
            generateRedlines: true
        }
    );

    assert.strictEqual(
        result.hasChanges,
        true,
        'format-only redline should rematch by text when targetRef paragraph has drifted'
    );
    const resultDoc = parseXmlStrict(result.documentXml, 'ref drift rematch output');
    assert.ok(
        resultDoc.getElementsByTagNameNS(NS_W, 'b').length > 0,
        'format-only redline should apply bold formatting after ref-drift rematch'
    );
}

async function testFormatOnlyRedlineFallsBackToOoxmlWhenNoSpansAreExtractable() {
    const inputXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${NS_W}">
  <w:body>
    <w:p>
      <w:r><w:fldChar w:fldCharType="begin"/></w:r>
      <w:r><w:instrText xml:space="preserve"> MERGEFIELD  SignatureLine </w:instrText></w:r>
      <w:r><w:fldChar w:fldCharType="end"/></w:r>
    </w:p>
    <w:sectPr/>
  </w:body>
</w:document>`;

    const result = await applyOperationToDocumentXml(
        inputXml,
        {
            type: 'redline',
            targetRef: 'P1',
            target: 'By: [Name]',
            modified: '**By:** [Name]'
        },
        'StandaloneRunnerTest',
        null,
        {
            generateRedlines: true
        }
    );

    assert.strictEqual(
        result.hasChanges,
        true,
        'format-only redline should not no-op when no text spans are extractable from OOXML'
    );
    const resultDoc = parseXmlStrict(result.documentXml, 'no-span format-only output');
    assert.ok(
        resultDoc.getElementsByTagNameNS(NS_W, 'b').length > 0,
        'format-only redline should still emit bold formatting via OOXML fallback'
    );
    const paragraphs = Array.from(resultDoc.getElementsByTagNameNS(NS_W, 'p'));
    const outputText = paragraphs.map(paragraph => getParagraphText(paragraph)).join('\n');
    assert.ok(
        outputText.includes('By: [Name]'),
        'format-only OOXML fallback should preserve rendered target text'
    );
}

async function testRedlinePreservesFieldInstructionsAndRemovesProofingMarkers() {
    const inputXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${NS_W}">
  <w:body>
    <w:p>
      <w:r><w:t>The Supplier must notify the Client where it restricts </w:t></w:r>
      <w:r><w:fldChar w:fldCharType="begin"/></w:r>
      <w:r><w:instrText xml:space="preserve"> REF _Ref12345 \\w \\h </w:instrText></w:r>
      <w:r><w:fldChar w:fldCharType="separate"/></w:r>
      <w:r><w:t>10.6</w:t></w:r>
      <w:r><w:fldChar w:fldCharType="end"/></w:r>
      <w:proofErr w:type="spellStart"/>
      <w:r><w:t>.</w:t></w:r>
      <w:proofErr w:type="spellEnd"/>
    </w:p>
    <w:sectPr/>
  </w:body>
</w:document>`;

    const result = await applyOperationToDocumentXml(
        inputXml,
        {
            type: 'redline',
            targetRef: 'P1',
            target: 'The Supplier must notify the Client where it restricts 10.6.',
            modified: 'The Supplier must, where legally permitted and reasonably practicable, notify the Client where it restricts 10.6.'
        },
        'StandaloneRunnerTest',
        null,
        {
            generateRedlines: true
        }
    );

    assert.strictEqual(result.hasChanges, true, 'field/proofing redline should report changes');
    assert.ok(!result.documentXml.includes('proofErr'), 'proofing markers should be stripped before redlining');

    const resultDoc = parseXmlStrict(result.documentXml, 'field/proofing redline output');
    const fieldChars = Array.from(resultDoc.getElementsByTagNameNS(NS_W, 'fldChar'));
    const instructions = Array.from(resultDoc.getElementsByTagNameNS(NS_W, 'instrText'));
    assert.strictEqual(fieldChars.length, 3, 'complex field begin/separate/end characters should survive redlining');
    assert.strictEqual(instructions.length, 1, 'complex field instruction should survive redlining');
    assert.strictEqual(instructions[0].textContent, ' REF _Ref12345 \\w \\h ', 'field instruction bytes should remain unchanged');
    for (const fieldNode of [...fieldChars, ...instructions]) {
        assert.strictEqual(fieldNode.parentNode?.localName, 'r', `${fieldNode.localName} must remain inside w:r`);
    }
    const paragraphs = Array.from(resultDoc.getElementsByTagNameNS(NS_W, 'p'));
    const outputText = paragraphs.map(paragraph => getParagraphText(paragraph)).join('\n');
    assert.ok(outputText.includes('10.6'), 'visible field result text should remain in the paragraph');
    assert.ok(resultDoc.getElementsByTagNameNS(NS_W, 'ins').length > 0 || resultDoc.getElementsByTagNameNS(NS_W, 'del').length > 0, 'redline should still emit revision markup');
}

async function testDuplicateTableStructuralOpsAreDedupedPerTurn() {
    const inputXml = buildTwoColumnTitleTableDocumentXml();
    const snapshotDoc = parseXmlStrict(inputXml, 'table dedupe snapshot');
    const runtimeContext = {
        targetRefSnapshot: buildTargetReferenceSnapshot(snapshotDoc),
        tableStructuralRedlineKeys: new Set()
    };

    const opA = {
        type: 'redline',
        targetRef: 'P1',
        target: 'Title:',
        modified: 'Title:\nDate:'
    };
    const opB = {
        type: 'redline',
        targetRef: 'P2',
        target: 'Title:',
        modified: 'Title:\nDate:'
    };

    const stepA = await applyOperationToDocumentXml(
        inputXml,
        opA,
        'StandaloneRunnerTest',
        runtimeContext,
        {
            generateRedlines: true
        }
    );
    assert.strictEqual(stepA.hasChanges, true, 'first table-structural op should apply');

    const stepB = await applyOperationToDocumentXml(
        stepA.documentXml,
        opB,
        'StandaloneRunnerTest',
        runtimeContext,
        {
            generateRedlines: true
        }
    );
    assert.strictEqual(
        stepB.hasChanges,
        false,
        'duplicate table-structural op in the same turn should be skipped'
    );
    assert.strictEqual(
        (stepB.warnings || []).some(w => String(w).includes('duplicate table-structural redline')),
        true,
        'dedupe skip should emit a warning'
    );
}

async function testBatchRunsCommentsBeforeTextEditsOnSameParagraph() {
    const inputXml = buildDocumentXml('The Customer may terminate the Agreement by written notice.');
    const operations = [
        {
            type: 'replace',
            targetRef: 'P1',
            target: 'The Customer may terminate the Agreement by written notice.',
            modified: 'The Customer may end this Agreement on written notice.'
        },
        {
            type: 'comment',
            targetRef: 'P1',
            target: 'The Customer may terminate the Agreement by written notice.',
            textToComment: 'terminate the Agreement by written notice',
            commentContent: 'Confirm the termination standard.'
        }
    ];

    const ordered = orderOperationsForStableTargets(operations);
    assert.deepStrictEqual(
        ordered.map(operation => operation.type),
        ['comment', 'replace'],
        'anchor-based comments should execute before text replacements'
    );
    assert.deepStrictEqual(operations.map(operation => operation.type), ['replace', 'comment'],
        'ordering helper must not mutate the caller array');

    const result = await applyOperationsToDocumentXml(
        inputXml,
        operations,
        'BatchRunnerTest',
        null,
        { generateRedlines: true }
    );

    assert.deepStrictEqual(result.executionOrder, [2, 1]);
    assert.deepStrictEqual(result.results.map(entry => entry.status), ['applied', 'applied']);
    assert.ok(result.commentsXml?.includes('Confirm the termination standard.'),
        'batch result should merge the applied comment payload');

    const resultDoc = parseXmlStrict(result.documentXml, 'comment-before-replace batch output');
    assert.ok(resultDoc.getElementsByTagNameNS(NS_W, 'commentRangeStart').length > 0,
        'comment anchors should survive the later replacement');
    assert.ok(resultDoc.getElementsByTagNameNS(NS_W, 'ins').length > 0,
        'replacement should still emit an insertion revision');
    assert.ok(resultDoc.getElementsByTagNameNS(NS_W, 'del').length > 0,
        'replacement should still emit a deletion revision');
}

async function testBatchAtomicRollbackAndLegacyPartialMode() {
    const originalParagraphs = [
        'Operation one source.',
        'Operation two source.',
        'Operation three source.',
        'Operation four source.',
        'Operation five source.'
    ];
    const inputXml = buildParagraphDocumentXml(originalParagraphs);
    const operations = [
        { type: 'replace', target: originalParagraphs[0], modified: 'Operation one applied.' },
        { type: 'replace', target: originalParagraphs[1], modified: 'Operation two applied.' },
        { type: 'replace', target: 'Missing operation three target.', modified: 'Must not apply.' },
        { type: 'replace', target: originalParagraphs[3], modified: 'Operation four applied.' },
        { type: 'replace', target: originalParagraphs[4], modified: 'Operation five applied.' }
    ];
    const atomicContext = {};

    const atomicResult = await applyOperationsToDocumentXml(
        inputXml,
        operations,
        'AtomicBatchTest',
        atomicContext,
        { generateRedlines: false }
    );

    assert.strictEqual(atomicResult.documentXml, inputXml,
        'default atomic batch must return the byte-identical original document');
    assert.strictEqual(atomicResult.hasChanges, false);
    assert.strictEqual(atomicResult.rolledBack, true);
    assert.strictEqual(atomicResult.error?.code, 'BATCH_OPERATION_FAILED');
    assert.deepStrictEqual(
        atomicResult.results.map(entry => entry.status),
        ['applied', 'applied', 'error', 'applied', 'applied'],
        'default continueOnError behavior should describe every attempted operation'
    );
    assert.strictEqual(atomicResult.results[2].error?.code, 'TARGET_NOT_FOUND');
    assert.strictEqual(atomicResult.commentsXml, null);
    assert.deepStrictEqual(atomicResult.numberingXmlParts, []);
    assert.strictEqual(atomicContext.targetRefSnapshot, undefined,
        'rolled-back batches must not commit mutable runtime context');

    const partialContext = {};
    const partialResult = await applyOperationsToDocumentXml(
        inputXml,
        operations,
        'AtomicBatchTest',
        partialContext,
        { generateRedlines: false, atomic: false }
    );

    assert.strictEqual(partialResult.hasChanges, true,
        'atomic:false should retain the established partial-result behavior');
    assert.strictEqual(partialResult.rolledBack, undefined);
    assert.ok(partialContext.targetRefSnapshot instanceof Map,
        'successful partial mode should commit runtime context');
    const partialDoc = parseXmlStrict(partialResult.documentXml, 'non-atomic batch output');
    const partialTexts = Array.from(partialDoc.getElementsByTagNameNS(NS_W, 'p')).map(getParagraphText);
    assert.deepStrictEqual(partialTexts, [
        'Operation one applied.',
        'Operation two applied.',
        originalParagraphs[2],
        'Operation four applied.',
        'Operation five applied.'
    ]);
}

async function testOverlappingBatchAnchorFailsInsteadOfEditingWrongSpan() {
    const originalText = 'The agency shall issue the permit within ten days.';
    const firstModified = 'The agency must issue the permit within ten days.';
    const secondModified = 'The agency shall issue the permit within five days.';
    const inputXml = buildDocumentXml(originalText);

    const result = await applyOperationsToDocumentXml(
        inputXml,
        [
            { type: 'replace', targetRef: 'P1', target: originalText, modified: firstModified },
            { type: 'replace', targetRef: 'P1', target: originalText, modified: secondModified }
        ],
        'OverlapBatchTest',
        null,
        { generateRedlines: false, atomic: false }
    );

    assert.deepStrictEqual(result.results.map(entry => entry.status), ['applied', 'error']);
    assert.strictEqual(result.results[1].error?.code, 'TARGET_NOT_FOUND',
        'stale overlapping anchor should be reported explicitly');
    const resultDoc = parseXmlStrict(result.documentXml, 'overlapping batch output');
    const paragraph = resultDoc.getElementsByTagNameNS(NS_W, 'p')[0];
    assert.strictEqual(getParagraphText(paragraph), firstModified,
        'failed overlapping operation must not silently edit the stale paragraph');
}

async function testNumberedHeadingConversionKeepsBatchRevisionIdsUnique() {
    const paragraphs = [
        'Document title',
        'Opening recital',
        'Second recital',
        'Service description',
        '2.3 Customer Responsibilities'
    ];
    const inputXml = buildParagraphDocumentXml(paragraphs);
    const operations = [
        { type: 'redline', targetRef: 1, target: paragraphs[0], modified: 'Standalone document' },
        { type: 'redline', targetRef: 2, target: paragraphs[1], modified: 'Revised opening recital' },
        { type: 'redline', targetRef: 3, target: paragraphs[2], modified: 'Revised second recital' },
        { type: 'redline', targetRef: 4, target: paragraphs[3], modified: 'Revised service description' },
        { type: 'redline', targetRef: 5, target: paragraphs[4], modified: '2.2 Customer Responsibilities' }
    ];

    const result = await applyOperationsToDocumentXml(
        inputXml,
        operations,
        'BatchAllocatorTest',
        { numberingIdState: createDynamicNumberingIdState() },
        { atomic: true, strictTargets: true }
    );

    assert.notStrictEqual(result.status, 'error');
    assert.strictEqual(result.error, undefined);
    assert.strictEqual(result.hasChanges, true);
    const validation = validateRedlineOoxml(result.documentXml);
    assert.deepStrictEqual(
        validation.issues.filter(issue => issue.code === 'DUPLICATE_REVISION_ID'),
        [],
        'numbered-heading routing must reuse the document-scoped batch allocator'
    );
}

async function run() {
    await testRedlineOperation();
    await testImplicitMultilineTargetReplacesWholeRange();
    await testCommentOperation();
    await testRangeListRedlineDoesNotDuplicateExistingItems();
    await testSingleParagraphListConcatenationUsesSurgicalInsertion();
    await testSingleParagraphListConcatenationWithInlineMarkersDoesNotInsertExtraItem();
    await testPlainParagraphInsertionBeforeTargetCreatesSeparateParagraph();
    await testFormatOnlyRedlineWithTrackedWrapperStillApplies();
    await testTextToTableWithoutHeaderSeparatorPreservesAllRows();
    await testFormatOnlyRedlineSupportsNonWPrefixOoxml();
    await testFormatOnlyRedlineWithAllTextInsideInsertionWrapper();
    await testFormatOnlyRedlineRematchesWhenRefParagraphDrifts();
    await testFormatOnlyRedlineFallsBackToOoxmlWhenNoSpansAreExtractable();
    await testRedlinePreservesFieldInstructionsAndRemovesProofingMarkers();
    await testDuplicateTableStructuralOpsAreDedupedPerTurn();
    await testBatchRunsCommentsBeforeTextEditsOnSameParagraph();
    await testBatchAtomicRollbackAndLegacyPartialMode();
    await testOverlappingBatchAnchorFailsInsteadOfEditingWrongSpan();
    await testNumberedHeadingConversionKeepsBatchRevisionIdsUnique();
    console.log('PASS: standalone operation runner tests');
}

run().catch(err => {
    console.error('FAIL:', err.message);
    process.exit(1);
});



