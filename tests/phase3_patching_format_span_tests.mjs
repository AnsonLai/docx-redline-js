import './setup-xml-provider.mjs';

import assert from 'assert/strict';

import { DiffOp, RunKind } from '../core/types.js';
import {
    applyFormatHintsToSpansRobust,
    splitSpanAtOffset,
    splitSpansAtBoundaries
} from '../engine/format-span-application.js';
import { parseOoxml, serializeOoxml } from '../engine/oxml-engine.js';
import { applyPatches, splitRunsAtDiffBoundaries } from '../pipeline/patching.js';

const NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const run = (text, startOffset, rPrXml = '') => ({
    kind: RunKind.TEXT,
    text,
    rPrXml,
    startOffset,
    endOffset: startOffset + text.length
});
const op = (type, startOffset, endOffset, text = '') => ({ type, startOffset, endOffset, text });

function singleSpan(text, rPr = null) {
    const rPrXml = rPr ? `<w:rPr>${rPr}</w:rPr>` : '';
    const doc = parseOoxml(`<w:p xmlns:w="${NS}"><w:r>${rPrXml}<w:t xml:space="preserve">${text}</w:t></w:r></w:p>`);
    const runElement = doc.getElementsByTagNameNS('*', 'r')[0];
    const textElement = doc.getElementsByTagNameNS('*', 't')[0];
    return {
        doc,
        span: { charStart: 0, charEnd: text.length, textElement, runElement, rPr: runElement.getElementsByTagNameNS('*', 'rPr')[0] || null }
    };
}

// Span boundaries: zero/end are no-ops, interior boundaries split in sorted
// order, detached runs fail cleanly, and whitespace remains explicit.
{
    const { doc, span } = singleSpan(' ab ');
    assert.equal(splitSpanAtOffset(doc, span, 0), null);
    assert.equal(splitSpanAtOffset(doc, span, 4), null);
    const split = splitSpansAtBoundaries(doc, [span], [4, 0, 2, 2]);
    assert.deepEqual(split.map(item => item.textElement.textContent), [' a', 'b ']);
    assert.deepEqual(split.map(item => [item.charStart, item.charEnd]), [[0, 2], [2, 4]]);
    assert.match(serializeOoxml(doc), /xml:space="preserve"> a<\/w:t>/);
}
{
    const { doc, span } = singleSpan('x');
    span.runElement.parentNode.removeChild(span.runElement);
    assert.equal(splitSpanAtOffset(doc, span, 0.5), null);
    assert.deepEqual(splitSpansAtBoundaries(doc, [], [1]), []);
}

// Overlapping hints split once at each boundary and merge formatting. Exercise
// tracked and untracked rPr paths, empty hints, and the final character.
{
    const { doc, span } = singleSpan('abcdef');
    applyFormatHintsToSpansRobust(doc, [span], [
        { start: 0, end: 4, format: { bold: true } },
        { start: 2, end: 6, format: { italic: true } },
        { start: 5, end: 6, format: { underline: true } }
    ], 'Phase 3', true);
    const xml = serializeOoxml(doc);
    assert.equal(doc.getElementsByTagNameNS('*', 't').length, 4);
    assert.match(xml, /<w:b w:val="1"\/>/);
    assert.match(xml, /<w:i w:val="1"\/>/);
    assert.match(xml, /<w:u w:val="single"\/>/);
    assert.equal(doc.getElementsByTagNameNS('*', 'rPrChange').length, 4);
}
{
    const { doc, span } = singleSpan('plain', '<w:b/>');
    applyFormatHintsToSpansRobust(doc, [span], [{ start: 0, end: 5, format: { italic: true } }], 'Phase 3', false);
    assert.equal(doc.getElementsByTagNameNS('*', 'rPrChange').length, 0);
    assert.equal(doc.getElementsByTagNameNS('*', 'b').length, 1);
    assert.equal(doc.getElementsByTagNameNS('*', 'i').length, 1);
    applyFormatHintsToSpansRobust(doc, [], [], 'Phase 3', true);
}

// Run splitting handles text, hyperlinks, untouched structural entries, run
// boundaries, and multiple unsorted diff boundaries.
const structural = { kind: RunKind.BOOKMARK, text: '', startOffset: 2, endOffset: 2, nodeXml: '<w:bookmarkStart/>' };
const hyperlink = { ...run('link', 4, '<w:i/>'), kind: RunKind.HYPERLINK };
const split = splitRunsAtDiffBoundaries(
    [run('abcd', 0, '<w:b/>'), structural, hyperlink],
    [op(DiffOp.DELETE, 3, 6, 'dli'), op(DiffOp.INSERT, 2, 2, 'X')]
);
assert.deepEqual(split.filter(item => item.kind === RunKind.TEXT).map(item => item.text), ['ab', 'c', 'd']);
assert.deepEqual(split.filter(item => item.kind === RunKind.HYPERLINK).map(item => item.text), ['li', 'nk']);
assert.equal(split.includes(structural), true);

// Deletion, insertion, equality, containers, and style choice at run boundaries.
const base = [
    { kind: RunKind.CONTAINER_START, containerId: 'sdt-1', text: '', startOffset: 0, endOffset: 0 },
    { kind: RunKind.PARAGRAPH_START, text: '', pPrXml: '', startOffset: 0, endOffset: 0 },
    run('ab', 0, '<w:b/>'),
    run('cd', 2, '<w:i/>'),
    { kind: RunKind.CONTAINER_END, containerId: 'sdt-1', text: '', startOffset: 4, endOffset: 4 }
];
const patched = applyPatches(base, [
    op(DiffOp.EQUAL, 0, 2, 'ab'),
    op(DiffOp.INSERT, 2, 2, ' X'),
    op(DiffOp.DELETE, 2, 4, 'cd'),
    op(DiffOp.INSERT, 4, 4, '!')
], { generateRedlines: true, author: 'Phase 3', formatHints: [] });
assert.ok(patched.some(item => item.kind === RunKind.INSERTION && item.text === ' X' && item.rPrXml === '<w:b/>'));
assert.ok(patched.some(item => item.kind === RunKind.DELETION && item.text === 'cd' && item.containerContext === 'sdt-1'));
assert.ok(patched.some(item => item.kind === RunKind.INSERTION && item.text === '!'));
const untracked = applyPatches([run('gone', 0)], [op(DiffOp.DELETE, 0, 4, 'gone')], {
    generateRedlines: false,
    author: 'Phase 3'
});
assert.deepEqual(untracked, []);

// Multiline list insertion exercises numbering detection, current-paragraph
// conversion, nested levels, lazy pPr serialization, and inserted paragraphs.
const pPrDoc = parseOoxml(`<w:pPr xmlns:w="${NS}"><w:spacing w:after="120"/></w:pPr>`);
const calls = [];
const numberingService = {
    detectNumberingFormat(marker) {
        calls.push(['detect', marker]);
        return marker.includes('.') && marker.split('.').length > 2
            ? { format: 'outline', depth: 1 }
            : { format: 'decimal', depth: 0 };
    },
    getOrCreateNumId(format, context) {
        calls.push(['num', format.type, context.numId]);
        return '42';
    },
    buildListPPr(numId, ilvl) {
        calls.push(['pPr', numId, ilvl]);
        return `<w:pPr><w:numPr><w:ilvl w:val="${ilvl}"/><w:numId w:val="${numId}"/></w:numPr></w:pPr>`;
    }
};
const listModel = [
    { kind: RunKind.PARAGRAPH_START, text: '', pPrXml: '', pPrElement: pPrDoc.documentElement, startOffset: 0, endOffset: 0 },
    run('seed', 0)
];
const listPatched = applyPatches(listModel, [op(DiffOp.INSERT, 0, 0, '1. First\n    1.1. Nested')], {
    generateRedlines: true,
    author: 'Phase 3',
    numberingService
});
assert.deepEqual(listPatched.filter(item => item.kind === RunKind.INSERTION).map(item => item.text), ['First', 'Nested']);
assert.equal(listPatched.filter(item => item.kind === RunKind.PARAGRAPH_START).length, 2);
assert.ok(listPatched.some(item => item.pPrXml?.includes('w:numId w:val="42"')));
assert.ok(calls.some(call => call[0] === 'detect'));

// No numbering service preserves literal lines; insertion-only empty models and
// leading/trailing-space style selection remain deterministic.
const literalLines = applyPatches([run('x', 0)], [op(DiffOp.INSERT, 0, 0, 'a\nb')], {
    generateRedlines: false,
    author: 'Phase 3'
});
assert.deepEqual(literalLines.filter(item => item.kind === RunKind.TEXT).map(item => item.text), ['a', 'b', 'x']);
const onlyInsert = applyPatches([], [op(DiffOp.INSERT, 0, 0, 'new')], { generateRedlines: true, author: 'Phase 3' });
assert.deepEqual(onlyInsert.map(item => item.text), ['new']);
const nextStyle = applyPatches([run('a', 0, '<w:b/>'), run('b', 1, '<w:i/>')], [op(DiffOp.INSERT, 1, 1, 'x ')], {
    generateRedlines: true,
    author: 'Phase 3'
});
assert.ok(nextStyle.some(item => item.kind === RunKind.INSERTION && item.rPrXml === '<w:i/>'));

console.log('PASS: Phase 3 patching and format-span boundary behavior');
