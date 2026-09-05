import './setup-xml-provider.mjs';

import assert from 'node:assert/strict';
import { DOMParser } from '@xmldom/xmldom';
import { extractCanonicalParagraphText } from '../core/paragraph-text.js';
import { stripRedundantLeadingListMarkers } from '../core/list-targeting.js';
import { buildSurgicalTextSpans } from '../engine/surgical-spans.js';
import { ingestOoxml } from '../pipeline/ingestion.js';
import {
    classifyListMarker,
    inferNumberingStyleFromMarker,
    parseListItem
} from '../pipeline/list-markers.js';
import { parseListItems } from '../pipeline/content-analysis.js';
import { parseMarkdownListContent } from '../orchestration/list-parsing.js';
import { normalizeListItemsWithLevels } from '../orchestration/list-markdown.js';

const markerCases = [
    ['-', 'bullet', 'bullet'], ['+', 'bullet', 'bullet'], ['•', 'bullet', 'bullet'],
    ['1.', 'numbered', 'decimal'], ['1.2.', 'numbered', 'decimal'],
    ['a.', 'numbered', 'lowerAlpha'], ['A.', 'numbered', 'upperAlpha'],
    ['iv.', 'numbered', 'lowerRoman'], ['IV.', 'numbered', 'upperRoman']
];

for (const [marker, markerType, style] of markerCases) {
    const line = `${marker} Item`;
    const shared = parseListItem(line);
    const pipeline = parseListItems(line)[0];
    const orchestration = parseMarkdownListContent(line).items[0];
    const normalized = normalizeListItemsWithLevels([line])[0];
    assert.equal(classifyListMarker(marker), markerType, marker);
    assert.equal(inferNumberingStyleFromMarker(marker), style, marker);
    assert.equal(shared.markerType, markerType, marker);
    assert.equal(pipeline.listType, markerType, marker);
    assert.equal(orchestration.type, markerType, marker);
    assert.equal(normalized.removedMarker, marker, marker);
    assert.equal(shared.text.trim(), 'Item', marker);
}
assert.equal(stripRedundantLeadingListMarkers('2.1. - Item'), 'Item');

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const paragraphXml = `<w:p xmlns:w="${W}"><w:r><w:t>A</w:t><w:tab/><w:t>B</w:t><w:br/><w:noBreakHyphen/><w:softHyphen/></w:r><w:hyperlink><w:r><w:t>C</w:t></w:r></w:hyperlink><w:r><w:fldChar w:fldCharType="begin"/><w:footnoteReference w:id="2"/><w:endnoteReference w:id="3"/></w:r></w:p>`;
const paragraph = new DOMParser().parseFromString(paragraphXml, 'application/xml').documentElement;
const canonical = extractCanonicalParagraphText(paragraph);
const surgical = buildSurgicalTextSpans([paragraph]).fullText;
const ingested = ingestOoxml(paragraphXml).acceptedText;
assert.equal(canonical, 'A\tB\n‑­C');
assert.equal(surgical, canonical, 'surgical visible projection should match canonical text');
assert.equal(ingested, canonical, 'pipeline ingestion visible projection should match canonical text');

const revisions = new DOMParser().parseFromString(`<w:p xmlns:w="${W}"><w:ins><w:r><w:t>new</w:t></w:r></w:ins><w:del><w:r><w:delText>old</w:delText></w:r></w:del><w:moveTo><w:r><w:t>moved</w:t></w:r></w:moveTo><w:moveFrom><w:r><w:delText>gone</w:delText></w:r></w:moveFrom></w:p>`, 'application/xml').documentElement;
assert.equal(extractCanonicalParagraphText(revisions, { revisionView:'accepted' }), 'newmoved');
assert.equal(extractCanonicalParagraphText(revisions, { revisionView:'rejected' }), 'oldgone');

console.log('PASS: Phase 4 shared list grammar and text-walker parity');
