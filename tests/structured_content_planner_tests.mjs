import './setup-xml-provider.mjs';

import assert from 'assert/strict';

import {
    acceptTrackedChangesInOoxml,
    analyzeStructuredContent,
    planStructuredReplacement,
    rejectTrackedChangesInOoxml,
    validateRedlineOoxml
} from '../index.js';
import { applyOperationToDocumentXml } from '../services/standalone-operation-runner.js';
import { elementsByLocalName, parseXml } from './helpers/ooxml-assertions.mjs';

const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const SOURCE = `<w:document xmlns:w="${NS_W}"><w:body><w:p><w:r><w:t>Date</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`;
const MIXED = `# ATTACHMENT 4

FOR IMMEDIATE RELEASE

| Agency | Contact |
| --- | --- |
| BCHD | Dr. Jenkins |
| MOHS | Marcus Vance |

## Protocol

1. Joint clearance
2. Rapid escalation

This attachment forms part of the agreement.`;

const missingSeparator = analyzeStructuredContent(`| Agency | Contact |
| BCHD | Dr. Jenkins |`);
assert.equal(missingSeparator.valid, false);
assert.ok(missingSeparator.issues.some(issue => issue.code === 'TABLE_SEPARATOR_REQUIRED'));
assert.equal(planStructuredReplacement('Date', missingSeparator.normalizedMarkdown).operation, null);

const plan = planStructuredReplacement({ exactText: 'Date', index: 1 }, MIXED, { author: 'Planner' });
assert.equal(plan.valid, true);
assert.deepEqual(plan.blocks.map(block => block.type), [
    'heading', 'paragraph', 'table', 'heading', 'list', 'paragraph'
]);
assert.deepEqual(plan.counts, { heading: 2, paragraph: 2, list: 1, table: 1 });
assert.equal(plan.operation.structuredContent, true);

const result = await applyOperationToDocumentXml(SOURCE, plan.operation, 'Planner');
assert.equal(result.status, 'ok');
assert.equal(result.hasChanges, true);
assert.equal(validateRedlineOoxml(result.documentXml).valid, true);

const resultDoc = parseXml(result.documentXml);
assert.equal(elementsByLocalName(resultDoc, 'tbl').length, 1, 'pipe rows must become a real Word table');
const generatedRows = elementsByLocalName(resultDoc, 'tr');
assert.equal(generatedRows.length, 3);
assert.equal(elementsByLocalName(generatedRows[0], 'tblHeader').length, 1,
    'Markdown header rows must repeat on continuation pages');
for (const row of generatedRows) {
    assert.equal(elementsByLocalName(row, 'cantSplit').length, 1,
        'generated table records must not split across pages');
}
assert.equal(elementsByLocalName(resultDoc, 'pStyle').filter(node => /^Heading[12]$/.test(node.getAttribute('w:val'))).length, 2);
const headingInsertions = elementsByLocalName(resultDoc, 'p').filter(paragraph => (
    elementsByLocalName(paragraph, 'pStyle').some(node => /^Heading[12]$/.test(node.getAttribute('w:val')))
));
for (const paragraph of headingInsertions) {
    const insertion = elementsByLocalName(paragraph, 'ins').find(ins => ins.parentNode === paragraph);
    assert.ok(elementsByLocalName(insertion, 'b').length > 0, 'heading runs must be bold explicitly');
    assert.ok(elementsByLocalName(insertion, 'sz').length > 0, 'heading runs must carry an explicit heading size');
}
const insertedListParagraphs = elementsByLocalName(resultDoc, 'p').filter(paragraph => (
    elementsByLocalName(paragraph, 'numPr').length > 0
    && elementsByLocalName(paragraph, 'ins').some(ins => ins.parentNode === paragraph)
));
assert.equal(insertedListParagraphs.length, 2);
for (const paragraph of insertedListParagraphs) {
    const numId = elementsByLocalName(paragraph, 'numId')[0]?.getAttribute('w:val');
    assert.ok(Number.parseInt(numId, 10) > 0);
}
assert.ok(!elementsByLocalName(resultDoc, 't').some(node => node.textContent.includes('| Agency |')),
    'table source syntax must not survive as visible pipe text');

const accepted = acceptTrackedChangesInOoxml(result.documentXml, { author: 'Planner' });
const rejected = rejectTrackedChangesInOoxml(result.documentXml, { author: 'Planner' });
assert.match(accepted.oxml, /ATTACHMENT 4/);
assert.equal(elementsByLocalName(parseXml(accepted.oxml), 'tbl').length, 1);
assert.doesNotMatch(accepted.oxml, /<w:del\b/);
assert.match(rejected.oxml, /<w:t[^>]*>Date<\/w:t>/);
assert.doesNotMatch(rejected.oxml, /ATTACHMENT 4/);
assert.equal(elementsByLocalName(parseXml(rejected.oxml), 'tbl').length, 0);

const invalidResult = await applyOperationToDocumentXml(SOURCE, {
    type: 'replace',
    target: 'Date',
    modified: '| Agency | Contact |\n| BCHD | Dr. Jenkins |',
    structuredContent: true
}, 'Planner');
assert.equal(invalidResult.status, 'error');
assert.equal(invalidResult.error.code, 'STRUCTURED_CONTENT_INVALID');
assert.equal(invalidResult.documentXml, SOURCE);

console.log('PASS: structured content planner validates and applies mixed document blocks');
