import './setup-xml-provider.mjs';

import assert from 'assert/strict';

import {
    acceptTrackedChangesInOoxml,
    rejectTrackedChangesInOoxml,
    validateRedlineOoxml
} from '../index.js';
import { getParagraphText } from '../core/paragraph-targeting.js';
import { applyOperationToDocumentXml } from '../services/standalone-operation-runner.js';
import {
    elementsByLocalName,
    extractExactVisibleText,
    parseXmlFragment
} from './helpers/ooxml-assertions.mjs';

const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

function document(paragraphContent) {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${NS_W}"><w:body><w:p>${paragraphContent}</w:p><w:sectPr/></w:body></w:document>`;
}

function textRun(text) {
    const escaped = text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
    const preserve = /^\s|\s$/.test(text) ? ' xml:space="preserve"' : '';
    return `<w:r><w:t${preserve}>${escaped}</w:t></w:r>`;
}

function assertValid(xml, label) {
    const validation = validateRedlineOoxml(xml);
    const errors = validation.issues.filter(issue => issue.severity === 'error');
    assert.deepEqual(errors, [], `${label}: ${errors.map(issue => `${issue.code}: ${issue.message}`).join('; ')}`);
}

function assertResolvedText(resultXml, original, modified, label) {
    const accepted = acceptTrackedChangesInOoxml(resultXml, { allAuthors: true });
    const rejected = rejectTrackedChangesInOoxml(resultXml, { allAuthors: true });
    assert.equal(accepted.status, undefined, `${label}: accept should succeed`);
    assert.equal(rejected.status, undefined, `${label}: reject should succeed`);
    assert.equal(extractExactVisibleText(accepted.oxml), modified, `${label}: accept-all text`);
    assert.equal(extractExactVisibleText(rejected.oxml), original, `${label}: reject-all text`);
}

async function applyReplace(sourceXml, original, modified) {
    return applyOperationToDocumentXml(
        sourceXml,
        { type: 'redline', target: original, modified },
        'StructuralTest',
        null,
        { generateRedlines: true }
    );
}

async function testParagraphReaderMapsTabsInDocumentOrder() {
    const sourceXml = document([
        '<w:r><w:tab/></w:r>',
        textRun('Leading'),
        '<w:r><w:tab/></w:r>',
        textRun('Middle'),
        '<w:r><w:tab/></w:r>'
    ].join(''));
    const parsed = parseXmlFragment(sourceXml);
    const paragraph = elementsByLocalName(parsed, 'p')[0];
    assert.equal(getParagraphText(paragraph), '\tLeading\tMiddle\t');
}

async function testStandaloneTabsAtEveryBoundary() {
    const cases = [
        {
            name: 'leading tab',
            content: '<w:r><w:tab/></w:r>' + textRun('Indented draft'),
            original: '\tIndented draft',
            modified: '\tIndented final'
        },
        {
            name: 'middle tab in a mixed-content run',
            content: '<w:r><w:t>Department</w:t><w:tab/><w:t>Finance - draft</w:t></w:r>',
            original: 'Department\tFinance - draft',
            modified: 'Department\tFinance - final'
        },
        {
            name: 'trailing tab',
            content: textRun('Status draft') + '<w:r><w:tab/></w:r>',
            original: 'Status draft\t',
            modified: 'Status final\t'
        }
    ];

    for (const testCase of cases) {
        const result = await applyReplace(document(testCase.content), testCase.original, testCase.modified);
        assert.equal(result.status, 'ok', `${testCase.name}: ${result.error?.message || result.status}`);
        assert.equal(result.hasChanges, true, `${testCase.name}: should apply`);
        assertValid(result.documentXml, testCase.name);
        assert.equal(elementsByLocalName(parseXmlFragment(result.documentXml), 'tab').length, 1, `${testCase.name}: tab count`);
        assertResolvedText(result.documentXml, testCase.original, testCase.modified, testCase.name);
    }
}

function assertComplexFields(xml, expectedInstructions) {
    const parsed = parseXmlFragment(xml);
    const fieldChars = elementsByLocalName(parsed, 'fldChar');
    const instructions = elementsByLocalName(parsed, 'instrText');
    assert.equal(fieldChars.length, expectedInstructions.length * 3, 'each complex field needs begin/separate/end');
    assert.deepEqual(instructions.map(node => node.textContent), expectedInstructions, 'field instructions must remain byte-exact');

    for (const node of [...fieldChars, ...instructions]) {
        assert.equal(node.parentNode?.localName, 'r', `${node.localName} must be a direct child of w:r`);
    }

    const fieldTypes = fieldChars.map(node => node.getAttribute('w:fldCharType') || node.getAttribute('fldCharType'));
    for (let index = 0; index < fieldTypes.length; index += 3) {
        assert.deepEqual(fieldTypes.slice(index, index + 3), ['begin', 'separate', 'end']);
    }
}

function assertCachedFieldResultsRemainInsideFields(xml, expectedTexts) {
    const parsed = parseXmlFragment(xml);
    const paragraph = elementsByLocalName(parsed, 'p')[0];
    const children = Array.from(paragraph.childNodes).filter(node => node.nodeType === 1);
    const fieldTypeWithin = node => elementsByLocalName(node, 'fldChar')
        .map(field => field.getAttribute('w:fldCharType') || field.getAttribute('fldCharType'))[0];
    const results = [];
    let resultNodes = null;

    for (const child of children) {
        const fieldType = fieldTypeWithin(child);
        if (fieldType === 'separate') {
            resultNodes = [];
        } else if (fieldType === 'end' && resultNodes) {
            results.push(resultNodes);
            resultNodes = null;
        } else if (resultNodes) {
            resultNodes.push(child);
        }
    }

    assert.equal(resultNodes, null, 'every field result must have an end marker');
    assert.deepEqual(
        results.map(nodes => nodes.map(node => node.textContent || '').join('')),
        expectedTexts,
        'cached field result text'
    );
    assert.equal(
        results.flat().some(node => ['ins', 'del'].includes(node.localName)),
        false,
        'unchanged field results must not be revised'
    );
}

async function testComplexFieldSurvivesEditsOnBothSidesAndSplitResultRuns() {
    const instruction = ' REF _Ref12345 \\w \\h ';
    const sourceXml = document([
        textRun('See '),
        '<w:r><w:fldChar w:fldCharType="begin" w:fldLock="true"/></w:r>',
        `<w:r><w:instrText xml:space="preserve">${instruction}</w:instrText></w:r>`,
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r>',
        textRun('10'),
        textRun('.6'),
        '<w:r><w:fldChar w:fldCharType="end"/></w:r>',
        textRun(' in the current clause.')
    ].join(''));
    const original = 'See 10.6 in the current clause.';
    const modified = 'Refer to 10.6 in the amended clause.';

    const result = await applyReplace(sourceXml, original, modified);
    assert.equal(result.status, 'ok', result.error?.message);
    assert.equal(result.hasChanges, true);
    assertValid(result.documentXml, 'complex field');
    assertComplexFields(result.documentXml, [instruction]);
    assertCachedFieldResultsRemainInsideFields(result.documentXml, ['10.6']);
    assertResolvedText(result.documentXml, original, modified, 'complex field');
}

async function testMultipleFieldsRemainOrdered() {
    const firstInstruction = ' REF FirstBookmark \\h ';
    const secondInstruction = ' REF SecondBookmark \\h ';
    const field = (instruction, result) => [
        '<w:r><w:fldChar w:fldCharType="begin"/></w:r>',
        `<w:r><w:instrText xml:space="preserve">${instruction}</w:instrText></w:r>`,
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r>',
        textRun(result),
        '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
    ].join('');
    const sourceXml = document([
        textRun('Sections '),
        field(firstInstruction, '4.1'),
        textRun(' and '),
        field(secondInstruction, '7.2'),
        textRun(' apply today.')
    ].join(''));
    const original = 'Sections 4.1 and 7.2 apply today.';
    const modified = 'Sections 4.1 and 7.2 apply immediately.';

    const result = await applyReplace(sourceXml, original, modified);
    assert.equal(result.status, 'ok', result.error?.message);
    assertValid(result.documentXml, 'multiple fields');
    assertComplexFields(result.documentXml, [firstInstruction, secondInstruction]);
    assertCachedFieldResultsRemainInsideFields(result.documentXml, ['4.1', '7.2']);
    assertResolvedText(result.documentXml, original, modified, 'multiple fields');
}

await testParagraphReaderMapsTabsInDocumentOrder();
await testStandaloneTabsAtEveryBoundary();
await testComplexFieldSurvivesEditsOnBothSidesAndSplitResultRuns();
await testMultipleFieldsRemainOrdered();

console.log('PASS: structural tab and complex field tests');
