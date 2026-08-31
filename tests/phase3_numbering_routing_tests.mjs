import './setup-xml-provider.mjs';

import assert from 'assert/strict';

import { parseOoxml, serializeOoxml } from '../engine/oxml-engine.js';
import {
    buildExplicitDecimalMultilevelNumberingXml,
    createDynamicNumberingIdState,
    extractFirstParagraphNumId,
    mergeNumberingXmlBySchemaOrder,
    overwriteParagraphNumIds,
    remapNumberingPayloadForDocument,
    reserveNextNumberingId,
    reserveNextNumberingIdPair
} from '../services/numbering-helpers.js';
import {
    RoutePlanKind,
    buildReconciliationPlan,
    normalizeContentEscapesForRouting
} from '../orchestration/route-plan.js';
import {
    buildListMarkdown,
    inferNumberingStyleFromMarker,
    normalizeListItemsWithLevels
} from '../orchestration/list-markdown.js';

const NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const numbering = (abstractId, numId) => `<w:numbering xmlns:w="${NS}"><w:abstractNum w:abstractNumId="${abstractId}"><w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/></w:lvl></w:abstractNum><w:num w:numId="${numId}"><w:abstractNumId w:val="${abstractId}"/></w:num></w:numbering>`;
const paragraph = numId => `<w:p xmlns:w="${NS}"><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="${numId}"/></w:numPr></w:pPr><w:r><w:t>Item</w:t></w:r></w:p>`;

// Number allocation: empty/malformed parts, collisions, preferred-range wrap,
// independent documents, normalization of caller-provided state, and pairs.
assert.deepEqual([...createDynamicNumberingIdState('').usedNumIds], []);
assert.deepEqual([...createDynamicNumberingIdState('<broken').usedAbstractNumIds], []);
const collisionState = createDynamicNumberingIdState(
    `<w:numbering xmlns:w="${NS}"><w:abstractNum w:abstractNumId="1"/><w:abstractNum w:abstractNumId="3"/><w:num w:numId="1"/><w:num w:numId="2"/></w:numbering>`,
    { minId: 1, maxPreferred: 3 }
);
assert.equal(collisionState.nextNumId, 3);
assert.equal(collisionState.nextAbstractNumId, 2);
assert.equal(reserveNextNumberingId(collisionState), 3);
assert.equal(reserveNextNumberingId(collisionState, 'abstract'), 2);
assert.deepEqual(reserveNextNumberingIdPair(collisionState), { numId: 4, abstractNumId: 4 });
assert.equal(reserveNextNumberingId(null), null);
const normalizedState = { nextNumId: 0, nextAbstractNumId: -1, usedNumIds: [1], maxPreferred: 0 };
assert.equal(reserveNextNumberingId(normalizedState), 1);
assert.ok(normalizedState.usedNumIds instanceof Set);
assert.equal(createDynamicNumberingIdState(numbering(7, 9)).nextNumId, 10);
assert.equal(createDynamicNumberingIdState(numbering(7, 9)).nextNumId, 10, 'allocation must be document-local and deterministic');

// Paragraph references and explicit-start multilevel numbering.
const paragraphDoc = parseOoxml(paragraph(5));
const paragraphs = [paragraphDoc.documentElement];
assert.equal(extractFirstParagraphNumId(paragraphs), '5');
overwriteParagraphNumIds(paragraphs, 12);
assert.equal(extractFirstParagraphNumId(paragraphs), '12');
overwriteParagraphNumIds(null, 3);
assert.equal(extractFirstParagraphNumId([]), null);
const explicit = buildExplicitDecimalMultilevelNumberingXml(20, 21, 7);
assert.match(explicit, /w:startOverride w:val="7"/);
assert.equal((explicit.match(/<w:lvl w:ilvl=/g) || []).length, 9);
assert.match(buildExplicitDecimalMultilevelNumberingXml(1, 2, -4), /w:startOverride w:val="1"/);

// Payload remapping updates definitions and cloned replacement nodes without
// mutating caller nodes.
const replacementDoc = parseOoxml(paragraph(9));
const remapState = createDynamicNumberingIdState(numbering(1, 1));
const remapped = remapNumberingPayloadForDocument(numbering(4, 9), [replacementDoc.documentElement], remapState);
assert.match(remapped.numberingXml, /w:abstractNumId="2"/);
assert.match(remapped.numberingXml, /w:numId="2"/);
assert.equal(extractFirstParagraphNumId(remapped.replacementNodes), '2');
assert.equal(extractFirstParagraphNumId([replacementDoc.documentElement]), '9', 'input node must remain untouched');
const remappedWithoutNodes = remapNumberingPayloadForDocument(numbering(4, 9), null, createDynamicNumberingIdState(''));
assert.deepEqual(remappedWithoutNodes.replacementNodes, []);

// Merge order, duplicate idempotence, missing parts, and malformed input.
const existingWithCleanup = `<w:numbering xmlns:w="${NS}"><w:abstractNum w:abstractNumId="1"/><w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num><w:numIdMacAtCleanup w:val="9"/></w:numbering>`;
const incoming = numbering(2, 2);
const merged = mergeNumberingXmlBySchemaOrder(existingWithCleanup, incoming);
assert.ok(merged.indexOf('w:abstractNumId="2"') < merged.indexOf('w:numId="1"'));
assert.ok(merged.indexOf('w:numId="2"') < merged.indexOf('w:numIdMacAtCleanup'));
assert.equal(mergeNumberingXmlBySchemaOrder(merged, incoming), merged, 'repeated merge must be idempotent');
assert.equal(mergeNumberingXmlBySchemaOrder('', incoming), incoming);
assert.equal(mergeNumberingXmlBySchemaOrder(existingWithCleanup, ''), existingWithCleanup);
assert.equal(mergeNumberingXmlBySchemaOrder(existingWithCleanup, '<broken'), existingWithCleanup);
assert.doesNotThrow(() => parseOoxml(serializeOoxml(parseOoxml(merged))));

// Route decision table: structured lists take precedence; empty input has two
// routes; block content and ordinary OOXML reconciliation remain distinct.
assert.equal(normalizeContentEscapesForRouting(String.raw`a\nb\tc\r`), 'a\nb\tc\r');
assert.equal(normalizeContentEscapesForRouting(null), '');
const routes = [
    [{ originalText: 'old', newContent: '- one\n- two' }, RoutePlanKind.STRUCTURED_LIST_DIRECT],
    [{ originalText: '', newContent: '**bold**' }, RoutePlanKind.EMPTY_FORMATTED_TEXT],
    [{ originalText: '   ', newContent: 'plain' }, RoutePlanKind.EMPTY_HTML],
    [{ originalText: 'old', newContent: '# Heading' }, RoutePlanKind.BLOCK_HTML],
    [{ originalText: 'old', newContent: '| A | B |\n|---|---|\n| 1 | 2 |' }, RoutePlanKind.BLOCK_HTML],
    [{ originalText: 'old', newContent: 'ordinary replacement' }, RoutePlanKind.OOXML_ENGINE]
];
for (const [input, expectedKind] of routes) {
    const plan = buildReconciliationPlan(input);
    assert.equal(plan.kind, expectedKind);
    assert.equal(typeof plan.flags.hasMarkdownTable, 'boolean');
}
const escapedList = buildReconciliationPlan({ originalText: 'old', newContent: String.raw`1. one\n2. two` });
assert.equal(escapedList.kind, RoutePlanKind.STRUCTURED_LIST_DIRECT);
assert.equal(buildReconciliationPlan().kind, RoutePlanKind.EMPTY_HTML);

// List normalization and marker rendering cover decimal, alpha, Roman,
// bullets, nesting, resets after an outdent, and values beyond Z.
assert.equal(inferNumberingStyleFromMarker('1.2.'), 'decimal');
assert.equal(inferNumberingStyleFromMarker('iv.'), 'lowerRoman');
assert.equal(inferNumberingStyleFromMarker('IV.'), 'upperRoman');
assert.equal(inferNumberingStyleFromMarker('a.'), 'lowerAlpha');
assert.equal(inferNumberingStyleFromMarker('A.'), 'upperAlpha');
assert.equal(inferNumberingStyleFromMarker('?'), 'decimal');
const normalized = normalizeListItemsWithLevels(['1. First', '    (a) Nested', '  - Shallow'], { indentSpaces: 2 });
assert.deepEqual(normalized, [
    { text: 'First', level: 0, removedMarker: '1.' },
    { text: 'Nested', level: 2, removedMarker: '(a)' },
    { text: 'Shallow', level: 1, removedMarker: '-' }
]);
assert.deepEqual(normalizeListItemsWithLevels([null], { indentSpaces: 0 }), [{ text: '', level: 0, removedMarker: null }]);
assert.equal(buildListMarkdown(normalized, 'bullet', 'decimal'), '- First\n        - Nested\n    - Shallow');
assert.equal(buildListMarkdown([{ text: 'A', level: 0 }, { text: 'B', level: 0 }], 'numbered', 'decimal'), '1. A\n2. B');
assert.equal(buildListMarkdown(Array.from({ length: 27 }, (_, index) => ({ text: String(index), level: 0 })), 'numbered', 'upperAlpha').split('\n').at(-1), 'AA. 26');
assert.equal(buildListMarkdown([{ text: 'nine', level: 0 }], 'numbered', 'lowerRoman'), 'i. nine');
assert.equal(buildListMarkdown([{ text: 'nine', level: 0 }], 'numbered', 'upperRoman'), 'I. nine');
assert.equal(buildListMarkdown([{ text: 'x', level: 0 }], 'numbered', 'unknown'), '1. x');

console.log('PASS: Phase 3 numbering, routing, and list-markdown behavior matrix');
