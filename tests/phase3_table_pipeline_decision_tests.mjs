import './setup-xml-provider.mjs';

import assert from 'assert/strict';

import { createParser, createSerializer } from '../adapters/xml-adapter.js';
import {
    inferTableReplacementParagraphBlock,
    isLikelyStructuredTableSourceParagraph,
    synthesizeTableMarkdownFromMultilineCellEdit
} from '../core/table-targeting.js';
import { applyTableReconciliation, applyTextToTableTransformation } from '../engine/table-mode.js';
import { ReconciliationPipeline } from '../pipeline/pipeline.js';

const NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const parser = createParser();
const serializer = createSerializer();

const heuristicCases = new Map([
    ['', false],
    ['and', true],
    ['[Name]', true],
    ['(optional)', true],
    ['Title:', true],
    ['Role (optional)', true],
    ['Ordinary sentence.', false],
    ['Ordinary heading', false]
]);
for (const [text, expected] of heuristicCases) {
    assert.equal(isLikelyStructuredTableSourceParagraph(text), expected, text);
}

// Paragraph-block inference skips non-paragraph siblings, tolerates one leading
// blank, stops at prose/second blank, and respects maxScan.
const blockDoc = parser.parseFromString(`<w:body xmlns:w="${NS}">
  <w:p><w:r><w:t>Parties:</w:t></w:r></w:p>
  <!-- ignored -->
  <w:p><w:r><w:t>[Disclosing Party]</w:t></w:r></w:p>
  <w:p><w:r><w:t>and</w:t></w:r></w:p>
  <w:p><w:r><w:t>This sentence ends the block.</w:t></w:r></w:p>
</w:body>`, 'application/xml');
const blockParagraphs = Array.from(blockDoc.getElementsByTagNameNS(NS, 'p'));
assert.deepEqual(inferTableReplacementParagraphBlock(blockParagraphs[0]), blockParagraphs.slice(0, 3));
assert.equal(inferTableReplacementParagraphBlock(blockParagraphs[0], { maxScan: 1 }), null);
assert.equal(inferTableReplacementParagraphBlock(null), null);
assert.equal(inferTableReplacementParagraphBlock(blockParagraphs[3]), null);
assert.equal(inferTableReplacementParagraphBlock(blockParagraphs[0], {
    getParagraphText: paragraph => paragraph === blockParagraphs[1] ? '' : paragraph.textContent
})?.length, 2);

const symmetricTableXml = `<w:tbl xmlns:w="${NS}">
  <w:tr>
    <w:tc><w:p><w:r><w:t>Title:</w:t></w:r></w:p></w:tc>
    <w:tc><w:p><w:r><w:t>Title:</w:t></w:r></w:p></w:tc>
  </w:tr>
  <w:tr>
    <w:tc><w:p><w:r><w:t></w:t></w:r></w:p></w:tc>
    <w:tc><w:p><w:r><w:t></w:t></w:r></w:p></w:tc>
  </w:tr>
</w:tbl>`;
const symmetricDoc = parser.parseFromString(symmetricTableXml, 'application/xml');
const target = symmetricDoc.getElementsByTagNameNS(NS, 'p')[0];
const infos = [];
const warnings = [];
const synthesized = synthesizeTableMarkdownFromMultilineCellEdit(target, 'Title:\nDirector\nOfficer | Secretary', {
    onInfo: message => infos.push(message),
    onWarn: message => warnings.push(message)
});
assert.match(synthesized, /\| Director \| Director \|/);
assert.match(synthesized, /Officer \\| Secretary/);
assert.ok(infos.some(message => message.includes('Symmetric row')));
assert.ok(infos.some(message => message.includes('Synthesized full markdown table')));
assert.deepEqual(warnings, []);

// Ambiguous anchors, already-markdown content, single lines, and detached
// paragraphs fail without mutating the source table.
const before = serializer.serializeToString(symmetricDoc);
assert.equal(synthesizeTableMarkdownFromMultilineCellEdit(target, 'Other:\nDirector', {
    onWarn: message => warnings.push(message)
}), null);
assert.equal(synthesizeTableMarkdownFromMultilineCellEdit(target, 'Other:\nDirector'), null);
assert.ok(warnings.some(message => message.includes('did not anchor')));
assert.equal(synthesizeTableMarkdownFromMultilineCellEdit(target, 'Title:'), null);
assert.equal(synthesizeTableMarkdownFromMultilineCellEdit(target, '| A |\n|---|\n| B |'), null);
const detached = parser.parseFromString(`<w:p xmlns:w="${NS}"><w:r><w:t>A</w:t></w:r></w:p>`, 'application/xml').documentElement;
assert.equal(synthesizeTableMarkdownFromMultilineCellEdit(detached, 'A\nB'), null);
assert.equal(serializer.serializeToString(symmetricDoc), before);

// Nested table content is not mistaken for direct outer rows/cells.
const nestedDoc = parser.parseFromString(`<w:tbl xmlns:w="${NS}"><w:tr><w:tc><w:p><w:r><w:t>Outer:</w:t></w:r></w:p><w:tbl><w:tr><w:tc><w:p><w:r><w:t>Inner</w:t></w:r></w:p></w:tc></w:tr></w:tbl></w:tc></w:tr></w:tbl>`, 'application/xml');
const nestedTarget = nestedDoc.getElementsByTagNameNS(NS, 'p')[0];
const nestedMarkdown = synthesizeTableMarkdownFromMultilineCellEdit(nestedTarget, 'Outer:\nAdded');
assert.match(nestedMarkdown, /Outer:/);
assert.doesNotMatch(nestedMarkdown, /Inner/);

// Table-mode no-change contracts preserve exact input for missing tables,
// malformed/empty markdown, and text-to-table requests with no table data.
const paragraphDoc = parser.parseFromString(`<w:p xmlns:w="${NS}"><w:r><w:t>Keep</w:t></w:r></w:p>`, 'application/xml');
const paragraphXml = serializer.serializeToString(paragraphDoc);
const noExistingTable = applyTableReconciliation(paragraphDoc, '| A |\n|---|\n| B |', serializer, parser, 'Phase 3');
assert.equal(noExistingTable.hasChanges, false);
assert.equal(noExistingTable.oxml, paragraphXml);
const noMarkdown = applyTextToTableTransformation(paragraphDoc, 'not a table', serializer, parser, 'Phase 3', true);
assert.equal(noMarkdown.hasChanges, false);
assert.equal(noMarkdown.oxml, paragraphXml);

// Pipeline routing helpers: validation modes, production-web optimization,
// yielding thresholds, insertion wrapper, indentation, and table generation.
const always = new ReconciliationPipeline({ validationMode: 'always', platform: 'web' });
assert.equal(always.shouldRunValidation(), true);
const never = new ReconciliationPipeline({ validationMode: 'never' });
assert.equal(never.shouldRunValidation(), false);
assert.equal(new ReconciliationPipeline({ validateOutput: false }).shouldRunValidation(), false);
const priorNodeEnv = process.env.NODE_ENV;
process.env.NODE_ENV = 'production';
assert.equal(new ReconciliationPipeline({ platform: 'web' }).shouldRunValidation(), false);
if (priorNodeEnv === undefined) delete process.env.NODE_ENV;
else process.env.NODE_ENV = priorNodeEnv;

const yielding = new ReconciliationPipeline({ enableEventLoopYielding: true, yieldRunThreshold: 0, yieldCharThreshold: 0 });
await yielding.maybeYield(1, 1);
await yielding.maybeYield(0, 0);
await new ReconciliationPipeline({ enableEventLoopYielding: false }).maybeYield(100, 10000);
assert.match(always.wrapForInsertion('<w:p/>'), /<pkg:package/);
assert.equal(always.detectIndentationStep(['- one', '    - nested']), 4);
const invalidTable = always.executeTableGeneration('not a table');
assert.equal(invalidTable.isValid, false);
const validTable = always.executeTableGeneration('| Name | Role |\n|---|---|\n| Ada | Reviewer |');
assert.equal(validTable.isValid, true);
assert.match(validTable.ooxml, /<w:tbl/);

console.log('PASS: Phase 3 table targeting, table mode, and pipeline decision matrix');
