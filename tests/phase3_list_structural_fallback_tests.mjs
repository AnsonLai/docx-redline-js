import './setup-xml-provider.mjs';

import assert from 'assert/strict';

import { parseOoxml, serializeOoxml } from '../engine/oxml-engine.js';
import {
    buildSingleLineListStructuralFallbackPlan,
    clearSingleLineListFallbackExplicitSequence,
    enforceListBindingOnParagraphNodes,
    executeSingleLineListStructuralFallback,
    recordSingleLineListFallbackExplicitSequence,
    resolveSingleLineListFallbackNumberingAction,
    stripSingleLineListMarkerPrefix
} from '../orchestration/list-structural-fallback.js';

const NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const plainParagraph = text => `<w:p xmlns:w="${NS}"><w:r><w:t>${text}</w:t></w:r></w:p>`;
const listedParagraph = text => `<w:p xmlns:w="${NS}"><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="8"/></w:numPr></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`;

assert.equal(stripSingleLineListMarkerPrefix('1. HEADER'), 'HEADER');
assert.equal(stripSingleLineListMarkerPrefix('2.2.1. Clause'), 'Clause');
assert.equal(stripSingleLineListMarkerPrefix('- Item'), 'Item');
assert.equal(stripSingleLineListMarkerPrefix('ordinary text'), 'ordinary text');
assert.equal(stripSingleLineListMarkerPrefix('1. one\n2. two'), '1. one\n2. two');

const decimalPlan = buildSingleLineListStructuralFallbackPlan({
    oxml: plainParagraph('7. Heading'),
    originalText: '7. Heading',
    modifiedText: '7. Heading'
});
assert.deepEqual(decimalPlan, {
    listInput: '7. Heading',
    numberingKey: 'numbered:decimal:single',
    originalText: '7. Heading',
    wasListParagraph: false,
    startAt: 7
});
const formattedPlan = buildSingleLineListStructuralFallbackPlan({
    oxml: plainParagraph('- Heading'),
    originalText: '- Heading',
    modifiedText: '- Heading'
});
assert.equal(formattedPlan.numberingKey, 'bullet:bullet:single');
assert.equal(formattedPlan.startAt, null);
assert.equal(buildSingleLineListStructuralFallbackPlan({
    oxml: plainParagraph('Old'), originalText: 'Old', modifiedText: '1. Different'
}), null);
assert.equal(buildSingleLineListStructuralFallbackPlan({
    oxml: listedParagraph('Heading'), originalText: 'Heading', modifiedText: '1. Heading'
}), null);
assert.equal(buildSingleLineListStructuralFallbackPlan({
    oxml: listedParagraph('Heading'), originalText: '1. Heading', modifiedText: '1. Heading', allowExistingList: true
}).wasListParagraph, true);
assert.equal(buildSingleLineListStructuralFallbackPlan({ oxml: '<broken', originalText: 'A', modifiedText: '1. A' }), null);
assert.equal(buildSingleLineListStructuralFallbackPlan(), null);

const sequenceState = { explicitByNumberingKey: new Map() };
assert.equal(resolveSingleLineListFallbackNumberingAction(null, sequenceState).type, 'none');
assert.equal(resolveSingleLineListFallbackNumberingAction({ numberingKey: 'bullet', startAt: null }, sequenceState).type, 'sharedByStyle');
assert.equal(resolveSingleLineListFallbackNumberingAction({ numberingKey: 'decimal', startAt: 4 }, null).type, 'explicitIsolated');
assert.equal(resolveSingleLineListFallbackNumberingAction({ numberingKey: 'decimal', startAt: 4 }, sequenceState).type, 'explicitStartNew');
recordSingleLineListFallbackExplicitSequence(sequenceState, 'decimal', 12, 3);
assert.deepEqual(resolveSingleLineListFallbackNumberingAction({ numberingKey: 'decimal', startAt: 4 }, sequenceState), {
    type: 'explicitReuse', numberingKey: 'decimal', startAt: 4, numId: '12'
});
clearSingleLineListFallbackExplicitSequence(sequenceState, 'decimal');
assert.equal(sequenceState.explicitByNumberingKey.size, 0);
recordSingleLineListFallbackExplicitSequence(null, 'x', 1, 1);
clearSingleLineListFallbackExplicitSequence(null, 'x');

// Binding creates missing pPr/numPr nodes, replaces existing values, removes
// listPr/pPrChange by default, and can preserve those compatibility nodes.
const bindingDoc = parseOoxml(`<w:body xmlns:w="${NS}">
  <w:p><w:r><w:t>New</w:t></w:r></w:p>
  <w:p><w:pPr><w:listPr/><w:pPrChange/><w:numPr><w:ilvl w:val="1"/><w:numId w:val="2"/></w:numPr></w:pPr><w:r><w:t>Old</w:t></w:r></w:p>
</w:body>`);
const paragraphs = Array.from(bindingDoc.getElementsByTagNameNS('*', 'p'));
assert.equal(enforceListBindingOnParagraphNodes(paragraphs, { numId: 44, ilvl: -3 }), 2);
let boundXml = serializeOoxml(bindingDoc);
assert.equal((boundXml.match(/w:numId w:val="44"/g) || []).length, 2);
assert.equal((boundXml.match(/w:ilvl w:val="0"/g) || []).length, 2);
assert.doesNotMatch(boundXml, /w:listPr|w:pPrChange/);
const preserveDoc = parseOoxml(`<w:p xmlns:w="${NS}"><w:pPr><w:listPr/><w:pPrChange/></w:pPr></w:p>`);
assert.equal(enforceListBindingOnParagraphNodes([preserveDoc.documentElement], {
    numId: 5,
    ilvl: 2,
    clearParagraphPropertyChanges: false,
    removeListPropertyNode: false
}), 1);
boundXml = serializeOoxml(preserveDoc);
assert.match(boundXml, /w:listPr/);
assert.match(boundXml, /w:pPrChange/);
assert.equal(enforceListBindingOnParagraphNodes(null, { numId: 1 }), 0);
assert.equal(enforceListBindingOnParagraphNodes(paragraphs, {}), 0);

// Execute with an injected pipeline so the fallback output is independently
// controlled. Verify trailing blank removal and both numbering start locations.
const generatedOxml = `<w:p xmlns:w="${NS}"><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="9"/></w:numPr></w:pPr><w:r><w:t>Heading</w:t></w:r></w:p><w:p><w:pPr></w:pPr></w:p>`;
const generatedNumbering = `<w:numbering xmlns:w="${NS}"><w:abstractNum w:abstractNumId="3"><w:lvl w:ilvl="0"><w:start w:val="1"/></w:lvl></w:abstractNum><w:num w:numId="9"><w:abstractNumId w:val="3"/></w:num></w:numbering>`;
const pipeline = {
    async executeListGeneration(input, range, context, originalText) {
        assert.equal(input, '7. Heading');
        assert.equal(range, null);
        assert.equal(context, null);
        assert.equal(originalText, '7. Heading');
        return { oxml: generatedOxml, numberingXml: generatedNumbering, isValid: true };
    }
};
const executed = await executeSingleLineListStructuralFallback(decimalPlan, { pipeline, author: 'Phase 3' });
assert.equal(executed.hasChanges, true);
assert.equal(executed.listStructuralFallbackApplied, true);
assert.doesNotMatch(executed.oxml, /<w:p><w:pPr><\/w:pPr><\/w:p>$/);
assert.match(executed.numberingXml, /w:startOverride w:val="7"/);
assert.match(executed.numberingXml, /<w:start w:val="7"/);
const withoutAbstractMutation = await executeSingleLineListStructuralFallback(decimalPlan, {
    pipeline,
    setAbstractStartOverride: false
});
assert.match(withoutAbstractMutation.numberingXml, /w:startOverride w:val="7"/);
assert.match(withoutAbstractMutation.numberingXml, /<w:start w:val="1"/);

const missing = await executeSingleLineListStructuralFallback(null);
assert.equal(missing.hasChanges, false);
assert.match(missing.warnings[0], /plan missing/);
const invalid = await executeSingleLineListStructuralFallback(decimalPlan, {
    pipeline: { async executeListGeneration() { return { oxml: '', isValid: false }; } }
});
assert.equal(invalid.hasChanges, false);
assert.match(invalid.warnings[0], /no valid OOXML/);

console.log('PASS: Phase 3 single-line list structural fallback behavior');
