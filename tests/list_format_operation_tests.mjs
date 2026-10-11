import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';

import './setup-xml-provider.mjs';
import { openDocx } from '../index.js';
import { parseOoxmlSafe, createSerializer } from '../adapters/xml-adapter.js';
import { unzipDocx, zipDocx } from '../document/zip-archive.js';
import { applyOperationsToDocumentXml } from '../services/standalone-operation-runner.js';
import { validateDocumentOperation } from '../services/document-operation-contract.js';
import { canonicalNumberingDefinition, mergeNumberingParts, validateNumberingIntegrity } from '../services/numbering-allocation.js';
import { projectWordRevisionView } from '../core/word-revision-semantics.js';
import { validateRedlineOoxml } from '../core/redline-validation.js';
import { rejectTrackedChangesInOoxml, acceptTrackedChangesInOoxml } from '../services/revision-comment-management.js';
import { parseTypedListMarker, resolveMarkerFormat } from '../services/list-format.js';
import { executeCli } from '../node/cli.js';

/**
 * Canonical list operations WP-4: `list-format` `convert-plain`. Every end-to-end case runs with default options,
 * so the Reject All verifier (enforce) refuses any output whose Reject All differs from the source.
 */

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const AUTHOR = 'List Author';
const dec = new TextDecoder();
const enc = new TextEncoder();
const parse = xml => {
    const parsed = parseOoxmlSafe(xml);
    assert.ok(parsed.doc, parsed.error?.message || 'parseable OOXML');
    return parsed.doc;
};
const el = (node, name) => Array.from(node.getElementsByTagNameNS(W, name));
const attr = (node, name) => node?.getAttribute(`w:${name}`) ?? null;
const innerError = result => result.results?.find(item => item?.error)?.error || result.error;

const p = (text, { pPr = '', id = null, rPr = '' } = {}) => `<w:p${id ? ` w14:paraId="${id}"` : ''}>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}<w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
const documentOf = (...paragraphs) => `<w:document xmlns:w="${W}" xmlns:w14="${W14}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>${paragraphs.join('')}<w:sectPr/></w:body></w:document>`;

async function bare(documentXml, operations, { context = {}, options = {} } = {}) {
    return applyOperationsToDocumentXml(documentXml, operations, AUTHOR, context, options);
}

function views(documentXml) {
    const doc = parse(documentXml);
    const pick = view => projectWordRevisionView(doc, view).map(record => ({ text: record.text, numId: record.numPr?.numId ?? null, ilvl: record.numPr?.ilvl ?? null }));
    return { accepted: pick('accepted'), rejected: pick('rejected') };
}

function definitionOf(numberingXml, numId, ilvl = 0) {
    const doc = parse(numberingXml);
    const num = el(doc, 'num').find(node => attr(node, 'numId') === String(numId));
    assert.ok(num, `w:num ${numId} is defined`);
    const abstractId = attr(el(num, 'abstractNumId')[0], 'val');
    const abstract = el(doc, 'abstractNum').find(node => attr(node, 'abstractNumId') === abstractId);
    const lvl = el(abstract, 'lvl').find(node => attr(node, 'ilvl') === String(ilvl));
    const override = el(num, 'lvlOverride').find(node => attr(node, 'ilvl') === String(ilvl));
    return {
        numFmt: attr(el(lvl, 'numFmt')[0], 'val'),
        lvlText: attr(el(lvl, 'lvlText')[0], 'val'),
        start: attr(el(lvl, 'start')[0], 'val'),
        startOverride: override ? attr(el(override, 'startOverride')[0], 'val') : null
    };
}

// ---------------------------------------------------------------------------------------------------------
// 1. Contract
// ---------------------------------------------------------------------------------------------------------
{
    const ok = list => validateDocumentOperation({ type: 'list-format', target: { exactText: 'A. Scope' }, list });
    assert.equal(ok({ mode: 'convert-plain' }).valid, true);
    assert.equal(ok({ mode: 'convert-plain' }).operation.operationKind, 'list-format');
    assert.equal(ok({ mode: 'convert-plain', format: 'upperRoman', level: 2, strip: 'marker' }).valid, true);
    assert.equal(ok({ mode: 'convert-plain', format: 'decimal', strip: 'none', start: 4 }).valid, true);
    assert.equal(ok({ mode: 'convert-plain', kind: 'bullet', strip: 'none' }).valid, true);
    assert.equal(ok({ mode: 'change-format', format: 'decimal' }).valid, true);
    assert.equal(ok({ mode: 'change-format', kind: 'bullet', scope: 'range' }).valid, true);
    assert.equal(ok({ mode: 'change-format', format: 'upperLetter', lvlText: '(%2)', level: 1, start: 3 }).valid, true);
    assert.equal(ok({ mode: 'join', continueFrom: { paragraphId: '1A2B3C4D' } }).valid, true);
    assert.equal(ok({ mode: 'join', continueFrom: { exactText: 'Item' }, scope: 'range' }).valid, true);
    for (const [label, list, field] of [
        ['missing list', undefined, 'list'],
        ['detach is later', { mode: 'detach' }, 'list.mode'],
        ['materializeMarker is later', { mode: 'convert-plain', materializeMarker: true }, 'list.materializeMarker'],
        ['scope is not a convert-plain key', { mode: 'convert-plain', scope: 'list' }, 'list.scope'],
        ['strip is not a change-format key', { mode: 'change-format', format: 'decimal', strip: 'none' }, 'list.strip'],
        ['change-format needs something to change', { mode: 'change-format' }, 'list.format'],
        ['numbered kind needs a format', { mode: 'change-format', kind: 'numbered' }, 'list.format'],
        ['bad scope', { mode: 'change-format', format: 'decimal', scope: 'document' }, 'list.scope'],
        ['lvlText without placeholder', { mode: 'change-format', format: 'decimal', lvlText: 'x.' }, 'list.lvlText'],
        ['lvlText missing own level', { mode: 'change-format', format: 'decimal', level: 1, lvlText: '%1.' }, 'list.lvlText'],
        ['bullet lvlText with placeholder', { mode: 'change-format', format: 'bullet', lvlText: '%1' }, 'list.lvlText'],
        ['start with scope range', { mode: 'change-format', format: 'decimal', scope: 'range', start: 2 }, 'list.start'],
        ['join needs continueFrom', { mode: 'join' }, 'list.continueFrom'],
        ['join continueFrom needs a strong target', { mode: 'join', continueFrom: {} }, 'list.continueFrom'],
        ['join continueFrom capture', { mode: 'join', continueFrom: { exactText: 'x', captureRef: 'c' } }, 'list.continueFrom'],
        ['join takes no format', { mode: 'join', continueFrom: { exactText: 'x' }, format: 'decimal' }, 'list.format'],
        ['unknown key', { mode: 'convert-plain', colour: 'red' }, 'list.colour'],
        ['bad format', { mode: 'convert-plain', format: 'upperAlpha' }, 'list.format'],
        ['kind contradicts format', { mode: 'convert-plain', kind: 'bullet', format: 'decimal' }, 'list.kind'],
        ['level 9', { mode: 'convert-plain', level: 9 }, 'list.level'],
        ['start comes from the marker', { mode: 'convert-plain', start: 3 }, 'list.start'],
        ['strip none needs a format', { mode: 'convert-plain', strip: 'none' }, 'list.format'],
        ['bad strip', { mode: 'convert-plain', strip: 'all' }, 'list.strip'],
        ['start 0', { mode: 'convert-plain', format: 'decimal', strip: 'none', start: 0 }, 'list.start']
    ]) {
        const result = ok(list);
        assert.equal(result.valid, false, label);
        assert.equal(result.error.code, 'INVALID_OPERATION', label);
        assert.equal(result.error.field, field, label);
        assert.equal(result.error.recovery.action, 'change_request', label);
    }
    const withText = validateDocumentOperation({ type: 'list-format', target: 'A. Scope', modified: 'Scope', list: { mode: 'convert-plain' } });
    assert.equal(withText.error.field, 'modified', 'text changes belong in a separate redline');
    const capture = validateDocumentOperation({ type: 'list-format', target: { captureRef: 'x' }, list: { mode: 'convert-plain' } });
    assert.equal(capture.valid, false);

    const schema = JSON.parse(readFileSync(new URL('../docs/schemas/document-operations.schema.json', import.meta.url), 'utf8'));
    assert.deepEqual(schema.$defs.listFormat.properties.mode, { enum: ['convert-plain', 'change-format', 'join'] });
    assert.ok(schema.$defs.operation.oneOf.some(branch => JSON.stringify(branch).includes('"const":"list-format"')));
    assert.equal(schema.$defs.listFormat.additionalProperties, false);
    for (const key of ['lvlText', 'scope', 'continueFrom']) assert.ok(schema.$defs.listFormat.properties[key], `schema documents list.${key}`);

    const version = await executeCli(['version']);
    assert.equal(version.contractVersion, 9, 'change-format and join are additive modes of list-format: a capability, not a contract bump');
    assert.ok(version.capabilities.includes('list-format-v1'));
    assert.ok(version.capabilities.includes('list-format-change-v1'));
}

// ---------------------------------------------------------------------------------------------------------
// 2. Marker grammar
// ---------------------------------------------------------------------------------------------------------
{
    const formats = text => parseTypedListMarker(text)?.candidates.map(candidate => `${candidate.format}:${candidate.value}`) ?? null;
    assert.deepEqual(formats('A. Scope'), ['upperLetter:1']);
    assert.deepEqual(formats('b) Term'), ['lowerLetter:2']);
    assert.deepEqual(formats('7. Seven'), ['decimal:7']);
    assert.deepEqual(formats('12) Twelve'), ['decimal:12']);
    assert.deepEqual(formats('IV. Four'), ['upperRoman:4']);
    assert.deepEqual(formats('xii. Twelve'), ['lowerRoman:12']);
    assert.deepEqual(formats('I. One'), ['upperLetter:9', 'upperRoman:1'], 'a single roman letter reads both ways');
    assert.deepEqual(formats('C.\tThird'), ['upperLetter:3', 'upperRoman:100']);
    assert.deepEqual(formats('- Item'), ['bullet:null']);
    assert.deepEqual(formats('• Item'), ['bullet:null']);
    assert.equal(parseTypedListMarker('A.\tScope').prefixLength, 3, 'the tab after the marker is stripped too');
    assert.equal(parseTypedListMarker('  A.  Scope').prefixLength, 6, 'leading and trailing marker whitespace');
    for (const prose of ['Mr. Smith', 'Header', '7 days', 'A.Scope', 'Iv. mixed', 'AB. two letters', 'A. ']) {
        assert.equal(parseTypedListMarker(prose), null, prose);
    }
    assert.equal(parseTypedListMarker('2.1. Sub').unsupported, 'multilevel');
    assert.equal(parseTypedListMarker('(a) Paren').unsupported, 'parenthesized');
    assert.equal(parseTypedListMarker('0. Zero').unsupported, 'zero');

    const resolve = (texts, format) => resolveMarkerFormat(texts.map(parseTypedListMarker), format);
    assert.deepEqual(resolve(['I. One']).error.reason, 'ambiguous');
    assert.deepEqual(resolve(['I. One'], 'upperRoman'), { format: 'upperRoman' });
    assert.deepEqual(resolve(['I. One', 'II. Two']), { format: 'upperRoman' });
    assert.deepEqual(resolve(['C. Third', 'D. Fourth']), { format: 'upperLetter' }, 'C, D are consecutive letters, not 100, 500');
    assert.deepEqual(resolve(['H. Eight', 'I. Nine']), { format: 'upperLetter' });
    assert.equal(resolve(['A. One', '1. Two']).error.reason, 'mixed');
    assert.equal(resolve(['A. One', 'B) Two']).error.reason, 'mixed');
    assert.equal(resolve(['- One', '• Two']).error.reason, 'mixed');
    assert.equal(resolve(['A. One'], 'decimal').error.reason, 'format-mismatch');
}

// ---------------------------------------------------------------------------------------------------------
// 3. Each marker family through the bare runner: Accept All shows the bound list with stripped text,
//    Reject All restores the literal marker and no numPr.
// ---------------------------------------------------------------------------------------------------------
for (const [marker, list, numFmt, lvlText, startOverride] of [
    ['1.', {}, 'decimal', '%1.', null],
    ['7)', {}, 'decimal', '%1)', '7'],
    ['A.', {}, 'upperLetter', '%1.', null],
    ['C.', { format: 'upperLetter' }, 'upperLetter', '%1.', '3'],
    ['b)', {}, 'lowerLetter', '%1)', '2'],
    ['I.', { format: 'upperRoman' }, 'upperRoman', '%1.', null],
    ['IV.', {}, 'upperRoman', '%1.', '4'],
    ['iv.', {}, 'lowerRoman', '%1.', '4'],
    ['-', {}, 'bullet', '•', null],
    ['•', {}, 'bullet', '•', null]
]) {
    const sourceText = `${marker} Heading text`;
    const source = documentOf(p('Intro paragraph.'), p(sourceText, { id: '0A0B0C0D' }), p('Closing paragraph.'));
    const result = await bare(source, [{ type: 'list-format', target: { exactText: sourceText }, list: { mode: 'convert-plain', ...list } }]);
    assert.equal(result.status, 'ok', `${marker}: ${JSON.stringify(innerError(result) || null)}`);
    const item = result.results[0];
    assert.equal(item.operationType, 'list-format');
    assert.equal(item.list.format, numFmt, marker);
    assert.equal(item.list.items.length, 1);
    assert.equal(item.list.items[0].marker, marker);
    const numberingXml = mergeNumberingParts('', result.numberingXmlParts);
    const { accepted, rejected } = views(result.documentXml);
    assert.deepEqual(accepted[1], { text: 'Heading text', numId: item.list.items[0].numId, ilvl: '0' }, `${marker}: Accept All`);
    assert.deepEqual(rejected.map(record => record.text), ['Intro paragraph.', sourceText, 'Closing paragraph.'], `${marker}: Reject All text`);
    assert.ok(rejected.every(record => record.numId === null), `${marker}: Reject All has no list`);
    const definition = definitionOf(numberingXml, item.list.items[0].numId);
    assert.equal(definition.numFmt, numFmt, marker);
    assert.equal(definition.lvlText, lvlText, `${marker}: the label keeps the typed punctuation`);
    assert.equal(definition.startOverride, startOverride, `${marker}: start value`);
    assert.equal(definition.start, '1', 'the abstract start is never mutated');
    // In place: same paragraph element (paraId), a tracked pPrChange and a tracked deletion of exactly the marker.
    const doc = parse(result.documentXml);
    const paragraph = el(doc, 'p')[1];
    assert.equal(paragraph.getAttributeNS(W14, 'paraId'), '0A0B0C0D', `${marker}: paragraph identity is kept`);
    assert.equal(el(paragraph, 'pPrChange').length, 1);
    assert.equal(el(paragraph, 'del').length, 1);
    assert.equal(el(paragraph, 'delText').map(node => node.textContent).join(''), `${marker} `);
    assert.equal(el(doc, 'p').length, 3, 'no paragraph is inserted or deleted');
    assert.deepEqual(validateNumberingIntegrity(result.documentXml, numberingXml), []);
    assert.equal(validateRedlineOoxml(result.documentXml).valid, true);
}

// ---------------------------------------------------------------------------------------------------------
// 4. Ranges, start values and sequence sharing
// ---------------------------------------------------------------------------------------------------------
{
    // A., body paragraph, B., C.: one list; the body paragraph stays plain and the count continues across it.
    const source = documentOf(p('A. Scope'), p('Body text about scope.'), p('B. Term'), p(''), p('C. Fees'));
    const result = await bare(source, [{ type: 'list-format', target: { exactText: 'A. Scope' }, targetEnd: { exactText: 'C. Fees' }, list: { mode: 'convert-plain' } }]);
    assert.equal(result.status, 'ok', JSON.stringify(innerError(result) || null));
    const { list } = result.results[0];
    assert.equal(new Set(list.items.map(item => item.numId)).size, 1, 'one list for consecutive markers');
    assert.deepEqual(list.items.map(item => item.value), [1, 2, 3]);
    assert.deepEqual(list.skipped, [{ paragraphIndex: 2, reason: 'no-marker' }, { paragraphIndex: 4, reason: 'empty' }]);
    assert.equal(result.numberingXmlParts.length, 1);
    const { accepted } = views(result.documentXml);
    assert.deepEqual(accepted.map(record => record.text), ['Scope', 'Body text about scope.', 'Term', '', 'Fees']);
    assert.deepEqual(accepted.map(record => record.numId !== null), [true, false, true, false, true]);
}
{
    // A gap in the typed numbers starts a new list at the typed value, so labels match the text.
    const source = documentOf(p('1. Intro'), p('2. Scope'), p('5. Fees'));
    const result = await bare(source, [{ type: 'list-format', target: { exactText: '1. Intro' }, targetEnd: { exactText: '5. Fees' }, list: { mode: 'convert-plain' } }]);
    assert.equal(result.status, 'ok');
    const { items } = result.results[0].list;
    assert.equal(items[0].numId, items[1].numId);
    assert.notEqual(items[1].numId, items[2].numId);
    const numberingXml = mergeNumberingParts('', result.numberingXmlParts);
    assert.equal(definitionOf(numberingXml, items[2].numId).startOverride, '5');
    assert.equal(definitionOf(numberingXml, items[0].numId).startOverride, null);
}
{
    // Separate operations in one batch: "A." then "B." share one numId (sequence state); a later "A." restarts.
    const source = documentOf(p('A. Scope'), p('Body.'), p('B. Term'), p('A. Annex scope'));
    const result = await bare(source, [
        { type: 'list-format', target: { exactText: 'A. Scope' }, list: { mode: 'convert-plain' } },
        { type: 'list-format', target: { exactText: 'B. Term' }, list: { mode: 'convert-plain' } },
        { type: 'list-format', target: { exactText: 'A. Annex scope' }, list: { mode: 'convert-plain' } }
    ]);
    assert.equal(result.status, 'ok', JSON.stringify(innerError(result) || null));
    const [first, second, third] = result.results.map(item => item.list.items[0].numId);
    assert.equal(first, second, 'B. continues the list A. started');
    assert.notEqual(third, first, 'a new A. starts a new list');
    assert.equal(result.numberingXmlParts.length, 2, 'the continued list allocates no definition');
    // Batch numbering never collides with the source numbering part.
    const sourceNumbering = `<w:numbering xmlns:w="${W}"><w:abstractNum w:abstractNumId="4"><w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/></w:lvl></w:abstractNum><w:num w:numId="9"><w:abstractNumId w:val="4"/></w:num></w:numbering>`;
    const packaged = await bare(source, [{ type: 'list-format', target: { exactText: 'A. Scope' }, list: { mode: 'convert-plain' } }], { context: { numberingXml: sourceNumbering } });
    assert.ok(Number(packaged.results[0].list.items[0].numId) > 9);
}
{
    // A stale sequence entry (its numId is not bound in the live document, for example after a refused
    // operation) is never reused.
    const source = documentOf(p('B. Term'));
    const context = { listFallbackSequenceState: { explicitByNumberingKey: new Map([['list-format:upperLetter:0:.', { numId: '77', nextStartAt: 2 }]]) } };
    const result = await bare(source, [{ type: 'list-format', target: { exactText: 'B. Term' }, list: { mode: 'convert-plain' } }], { context });
    assert.equal(result.status, 'ok');
    assert.notEqual(result.results[0].list.items[0].numId, '77');
    assert.equal(definitionOf(mergeNumberingParts('', result.numberingXmlParts), result.results[0].list.items[0].numId).startOverride, '2');
}
{
    // strip: "none" binds without touching the text; start and level are honoured.
    const source = documentOf(p('Scope'), p('Term'));
    const result = await bare(source, [{ type: 'list-format', target: { exactText: 'Scope' }, targetEnd: { exactText: 'Term' }, list: { mode: 'convert-plain', format: 'lowerRoman', strip: 'none', start: 3, level: 1 } }]);
    assert.equal(result.status, 'ok', JSON.stringify(innerError(result) || null));
    const { accepted, rejected } = views(result.documentXml);
    assert.deepEqual(accepted, [{ text: 'Scope', numId: accepted[0].numId, ilvl: '1' }, { text: 'Term', numId: accepted[0].numId, ilvl: '1' }]);
    assert.ok(rejected.every(record => record.numId === null));
    assert.equal(el(parse(result.documentXml), 'del').length, 0);
    const definition = definitionOf(mergeNumberingParts('', result.numberingXmlParts), accepted[0].numId, 1);
    assert.deepEqual([definition.numFmt, definition.lvlText, definition.startOverride], ['lowerRoman', '%2.', '3']);
}
{
    // generateRedlines: false writes the binding and strips the marker directly.
    const source = documentOf(p('A. Scope'));
    const result = await bare(source, [{ type: 'list-format', target: { exactText: 'A. Scope' }, list: { mode: 'convert-plain' }, generateRedlines: false }]);
    assert.equal(result.status, 'ok');
    const doc = parse(result.documentXml);
    assert.equal(el(doc, 'pPrChange').length, 0);
    assert.equal(el(doc, 'del').length, 0);
    assert.deepEqual(views(result.documentXml).rejected, [{ text: 'Scope', numId: views(result.documentXml).accepted[0].numId, ilvl: '0' }]);
}

// ---------------------------------------------------------------------------------------------------------
// 5. Identity and formatting are preserved; a marker split across runs; partial reject coherence
// ---------------------------------------------------------------------------------------------------------
{
    const source = documentOf(
        `<w:p w14:paraId="1234ABCD"><w:pPr><w:pStyle w:val="Heading2"/><w:jc w:val="left"/></w:pPr><w:bookmarkStart w:id="0" w:name="scope"/><w:r><w:rPr><w:b/></w:rPr><w:t>A</w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">. Sco</w:t></w:r><w:r><w:rPr><w:b/><w:i/></w:rPr><w:t>pe</w:t></w:r><w:bookmarkEnd w:id="0"/></w:p>`
    );
    const result = await bare(source, [{ type: 'list-format', target: { exactText: 'A. Scope' }, list: { mode: 'convert-plain' } }]);
    assert.equal(result.status, 'ok', JSON.stringify(innerError(result) || null));
    const doc = parse(result.documentXml);
    const paragraph = el(doc, 'p')[0];
    assert.equal(paragraph.getAttributeNS(W14, 'paraId'), '1234ABCD');
    assert.equal(el(paragraph, 'bookmarkStart').length, 1, 'bookmarks survive');
    assert.equal(attr(el(paragraph, 'pStyle')[0], 'val'), 'Heading2');
    const deletions = el(paragraph, 'del');
    assert.equal(deletions.map(node => node.textContent).join(''), 'A. ');
    assert.ok(deletions.every(node => el(node, 'b').length === 1), 'deleted marker runs keep their formatting');
    const kept = el(paragraph, 'r').filter(run => el(run, 't').length > 0);
    assert.deepEqual(kept.map(run => el(run, 't')[0].textContent), ['Sco', 'pe']);
    assert.equal(el(kept[0], 'b').length, 1, 'the split run keeps its formatting');
    // pPrChange snapshot holds the previous pPr (style and justification, no numPr).
    const snapshot = el(el(paragraph, 'pPrChange')[0], 'pPr')[0];
    assert.equal(el(snapshot, 'numPr').length, 0);
    assert.equal(attr(el(snapshot, 'pStyle')[0], 'val'), 'Heading2');

    // Reject All through the library's own transform reproduces the source paragraph exactly (text, no numPr).
    const rejected = rejectTrackedChangesInOoxml(result.documentXml, { allAuthors: true });
    const rejectedDoc = parse(rejected.oxml);
    assert.equal(el(rejectedDoc, 'numPr').length, 0);
    assert.equal(el(rejectedDoc, 'pPrChange').length, 0);
    assert.equal(el(rejectedDoc, 't').map(node => node.textContent).join(''), 'A. Scope');
    const acceptedDoc = parse(acceptTrackedChangesInOoxml(result.documentXml, { allAuthors: true }).oxml);
    assert.equal(el(acceptedDoc, 't').map(node => node.textContent).join(''), 'Scope');
    assert.equal(el(acceptedDoc, 'numPr').length, 1);

    // Rejecting only the marker deletion (keeping the binding) leaves literal marker plus label: ugly, valid.
    const partial = parse(result.documentXml);
    for (const deletion of el(partial, 'del')) {
        for (const delText of el(deletion, 'delText')) {
            const t = partial.createElementNS(W, 'w:t');
            t.setAttribute('xml:space', 'preserve');
            t.textContent = delText.textContent;
            delText.parentNode.replaceChild(t, delText);
        }
        while (deletion.firstChild) deletion.parentNode.insertBefore(deletion.firstChild, deletion);
        deletion.parentNode.removeChild(deletion);
    }
    const partialXml = createSerializer().serializeToString(partial);
    assert.equal(validateRedlineOoxml(partialXml).valid, true);
    assert.deepEqual(views(partialXml).accepted.map(record => [record.text, record.numId !== null]), [['A. Scope', true]]);
}

// ---------------------------------------------------------------------------------------------------------
// 6. Refusals: recovery-enveloped, nothing written
// ---------------------------------------------------------------------------------------------------------
const NUMBERING = `<w:numbering xmlns:w="${W}"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`;
const STYLES = `<w:styles xmlns:w="${W}"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="ListHeading"><w:name w:val="List Heading"/><w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr></w:style><w:style w:type="paragraph" w:styleId="ListHeadingChild"><w:basedOn w:val="ListHeading"/></w:style></w:styles>`;
for (const [label, paragraphs, operation, code, action, extra] of [
    ['no marker', [p('Header')], { target: { exactText: 'Header' }, list: { mode: 'convert-plain' } }, 'LIST_MARKER_NOT_FOUND', 'change_request'],
    ['ambiguous roman letter', [p('I. Intro')], { target: { exactText: 'I. Intro' }, list: { mode: 'convert-plain' } }, 'LIST_MARKER_AMBIGUOUS', 'set_list_format', error => assert.deepEqual(error.recovery.allowedValues, ['upperLetter', 'upperRoman'])],
    ['format mismatch', [p('A. Scope')], { target: { exactText: 'A. Scope' }, list: { mode: 'convert-plain', format: 'decimal' } }, 'LIST_MARKER_UNSUPPORTED', 'change_request', error => assert.equal(error.reason, 'format-mismatch')],
    ['kind mismatch', [p('A. Scope')], { target: { exactText: 'A. Scope' }, list: { mode: 'convert-plain', kind: 'bullet' } }, 'LIST_MARKER_UNSUPPORTED', 'change_request'],
    ['mixed markers in one run', [p('A. Scope'), p('2. Term')], { target: { exactText: 'A. Scope' }, targetEnd: { exactText: '2. Term' }, list: { mode: 'convert-plain' } }, 'LIST_MARKER_UNSUPPORTED', 'change_request', error => assert.equal(error.reason, 'mixed')],
    ['multi-level marker', [p('2.1. Sub clause')], { target: { exactText: '2.1. Sub clause' }, list: { mode: 'convert-plain' } }, 'LIST_MARKER_UNSUPPORTED', 'convert_in_host', error => assert.equal(error.reason, 'multilevel')],
    ['parenthesized marker', [p('(a) Sub clause')], { target: { exactText: '(a) Sub clause' }, list: { mode: 'convert-plain' } }, 'LIST_MARKER_UNSUPPORTED', 'convert_in_host'],
    ['marker inside a hyperlink', [`<w:p><w:hyperlink r:id="rId9"><w:r><w:t xml:space="preserve">A. </w:t></w:r></w:hyperlink><w:r><w:t>Linked</w:t></w:r></w:p>`], { target: { exactText: 'A. Linked' }, list: { mode: 'convert-plain' } }, 'LIST_MARKER_UNSUPPORTED', 'convert_in_host', error => assert.equal(error.reason, 'marker-shape')],
    ['marker inside a field', [`<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>SEQ</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t xml:space="preserve">A. </w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r><w:r><w:t>Field</w:t></w:r></w:p>`], { target: { exactText: 'A. Field' }, list: { mode: 'convert-plain' } }, 'LIST_MARKER_UNSUPPORTED', 'convert_in_host'],
    ['already a list item', [p('A. Listed', { pPr: '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>' })], { target: { exactText: 'A. Listed' }, list: { mode: 'convert-plain' } }, 'LIST_FORMAT_TARGET_IS_LIST', 'use_paragraph_numbering'],
    ['style-linked numbering', [p('A. Styled', { pPr: '<w:pStyle w:val="ListHeadingChild"/>' })], { target: { exactText: 'A. Styled' }, list: { mode: 'convert-plain' } }, 'LIST_FORMAT_STYLE_NUMBERED', 'convert_in_host', error => assert.equal(error.styleId, 'ListHeading')],
    ['another author pending pPrChange', [p('A. Pending', { pPr: '<w:jc w:val="center"/><w:pPrChange w:id="9" w:author="Other" w:date="2026-01-01T00:00:00Z"><w:pPr/></w:pPrChange>' })], { target: { exactText: 'A. Pending' }, list: { mode: 'convert-plain' } }, 'PARAGRAPH_FORMAT_REVISION_CONFLICT', 'resolve_pending_revision']
]) {
    const source = documentOf(...paragraphs);
    const result = await bare(source, [{ type: 'list-format', ...operation }], { context: { numberingXml: NUMBERING, stylesXml: STYLES } });
    assert.equal(result.status, 'error', label);
    const error = innerError(result);
    assert.equal(error.code, code, `${label}: ${error.code} ${error.message}`);
    assert.equal(error.recovery?.action, action, label);
    assert.equal(error.recovery.sameArgumentsSafe, false, label);
    assert.equal(result.documentXml, source, `${label}: nothing written`);
    assert.equal(result.numberingXmlParts.length, 0, label);
    extra?.(error);
}
{
    // A paragraph whose style numbering is suppressed with numId 0 is plain and converts.
    const source = documentOf(p('A. Suppressed', { pPr: '<w:pStyle w:val="ListHeading"/><w:numPr><w:numId w:val="0"/></w:numPr>' }));
    const result = await bare(source, [{ type: 'list-format', target: { exactText: 'A. Suppressed' }, list: { mode: 'convert-plain' } }], { context: { numberingXml: NUMBERING, stylesXml: STYLES } });
    assert.equal(result.status, 'ok', JSON.stringify(innerError(result) || null));
    const rejectedRecord = views(result.documentXml).rejected[0];
    assert.deepEqual([rejectedRecord.text, rejectedRecord.numId], ['A. Suppressed', '0'], 'Reject All restores the numId 0 suppression');
}
{
    // Ranges cannot leave their container; two list-format writes to one paragraph conflict at compile time.
    const table = `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:tcPr/>${p('B. In cell')}</w:tc></w:tr></w:tbl>`;
    const crossing = await bare(documentOf(p('A. Body'), table), [{ type: 'list-format', target: { exactText: 'A. Body' }, targetEnd: { exactText: 'B. In cell' }, list: { mode: 'convert-plain' } }]);
    assert.equal(innerError(crossing).code, 'TARGET_RANGE_INVALID');
    const twice = await bare(documentOf(p('A. Body')), [
        { type: 'list-format', target: { exactText: 'A. Body' }, list: { mode: 'convert-plain' } },
        { type: 'redline', target: { exactText: 'A. Body' }, modified: 'A. Body text' }
    ]);
    assert.equal(twice.status, 'error');
    assert.equal(twice.conflicts?.[0]?.code || innerError(twice).code, 'OVERLAPPING_SOURCE_TARGETS');
}

// ---------------------------------------------------------------------------------------------------------
// 7. Packaged openDocx path (numbering part, inspect labels, Accept All / Reject All, part refusal)
// ---------------------------------------------------------------------------------------------------------
const TEMPLATE = new URL('./fixtures/agentic-lists/nested-lists-source.docx', import.meta.url);
function buildDocx(bodyXml, { stylesTransform = null } = {}) {
    const entries = unzipDocx(readFileSync(TEMPLATE));
    const original = dec.decode(entries.get('word/document.xml'));
    const sectPr = original.match(/<w:sectPr[\s\S]*<\/w:sectPr>\s*<\/w:body>/)[0];
    const head = original.slice(0, original.indexOf('<w:body>') + '<w:body>'.length);
    entries.set('word/document.xml', enc.encode(`${head}${bodyXml}${sectPr}</w:document>`));
    if (stylesTransform) entries.set('word/styles.xml', enc.encode(stylesTransform(dec.decode(entries.get('word/styles.xml')))));
    return new Uint8Array(zipDocx(entries));
}
const labelsOf = bytes => openDocx(bytes).inspect({}).paragraphs.map(item => [item.text, item.list?.label ?? null]);
{
    const bytes = buildDocx([p('A. Scope'), p('Explanatory body text.'), p('B. Term'), p('C. Fees'), p('Closing.')].join(''));
    const before = labelsOf(bytes);
    const result = await openDocx(bytes).applyOperations([
        { type: 'list-format', target: { exactText: 'A. Scope' }, list: { mode: 'convert-plain' } },
        { type: 'list-format', target: { exactText: 'B. Term' }, targetEnd: { exactText: 'C. Fees' }, list: { mode: 'convert-plain' } }
    ], { author: AUTHOR });
    assert.equal(result.status, 'ok', JSON.stringify(innerError(result) || null));
    assert.equal(result.written, true);
    const output = result.toUint8Array();
    const entries = unzipDocx(output);
    const numberingXml = dec.decode(entries.get('word/numbering.xml'));
    const numIds = new Set(result.results.flatMap(item => item.list.items.map(entry => entry.numId)));
    assert.equal(numIds.size, 1, 'the batch shares one list');
    assert.ok(numberingXml.includes(`w:numId="${[...numIds][0]}"`), 'the facade merged the new definition');
    assert.ok(numberingXml.includes('w:abstractNumId="0"'), 'existing definitions are kept');
    assert.deepEqual(validateNumberingIntegrity(dec.decode(entries.get('word/document.xml')), numberingXml), []);

    const accepted = await openDocx(output).resolveRevisions('accept', { allAuthors: true });
    assert.deepEqual(labelsOf(accepted.toUint8Array()), [['Scope', 'A.'], ['Explanatory body text.', null], ['Term', 'B.'], ['Fees', 'C.'], ['Closing.', null]], 'Accept All shows the bound list with stripped text');
    const rejected = await openDocx(output).resolveRevisions('reject', { allAuthors: true });
    assert.deepEqual(labelsOf(rejected.toUint8Array()), before, 'Reject All restores the literal markers and no list');

    const partResult = await openDocx(bytes).applyOperations([{ type: 'list-format', part: 'word/header1.xml', target: { exactText: 'A. Scope' }, list: { mode: 'convert-plain' } }], { author: AUTHOR });
    assert.equal(partResult.status, 'error');
    assert.equal(partResult.written, false);

    // Style-linked numbering is detected through the package's styles part.
    const styled = buildDocx(p('A. Styled', { pPr: '<w:pStyle w:val="NumberedHeadingWp4"/>' }), {
        stylesTransform: styles => styles.replace('</w:styles>', '<w:style w:type="paragraph" w:styleId="NumberedHeadingWp4"><w:name w:val="Numbered Heading WP4"/><w:pPr><w:numPr><w:numId w:val="2"/></w:numPr></w:pPr></w:style></w:styles>')
    });
    const styledResult = await openDocx(styled).applyOperations([{ type: 'list-format', target: { exactText: 'A. Styled' }, list: { mode: 'convert-plain' } }], { author: AUTHOR });
    assert.equal(innerError(styledResult).code, 'LIST_FORMAT_STYLE_NUMBERED');
    assert.equal(styledResult.written, false);
}

// ---------------------------------------------------------------------------------------------------------
// 8. Text-equal redline conversions now convert in place (letters, roman, bullets); the old whole-paragraph
//    delete/insert form is kept only for shapes the in-place core refuses, and it stays reject-exact.
// ---------------------------------------------------------------------------------------------------------
{
    const quiet = console.log;
    console.log = () => { };
    try {
        for (const marker of ['B.', 'iv.', '-']) {
            const text = `${marker} Header`;
            const source = documentOf(p(text, { id: '5A5B5C5D' }));
            const result = await bare(source, [{ type: 'redline', target: { exactText: text }, modified: text }]);
            assert.equal(result.status, 'ok', marker);
            const doc = parse(result.documentXml);
            assert.equal(el(doc, 'p').length, 1, `${marker}: converted in place, not deleted and re-inserted`);
            assert.equal(el(doc, 'p')[0].getAttributeNS(W14, 'paraId'), '5A5B5C5D');
            assert.equal(el(doc, 'delText').map(node => node.textContent).join(''), `${marker} `);
            const { accepted, rejected } = views(result.documentXml);
            assert.equal(accepted[0].text, 'Header');
            assert.ok(accepted[0].numId);
            assert.deepEqual(rejected, [{ text, numId: null, ilvl: null }], `${marker}: Reject All is exact`);
        }
        // Old form: the in-place core declines a marker followed by a symbol run, so the whole-paragraph
        // delete/insert fallback runs; it is still exact under Reject All.
        const legacySource = documentOf(`<w:p w14:paraId="6A6B6C6D"><w:r><w:t xml:space="preserve">B.</w:t></w:r><w:r><w:sym w:font="Symbol" w:char="F020"/></w:r><w:r><w:t xml:space="preserve"> Header</w:t></w:r></w:p>`);
        const legacy = await bare(legacySource, [{ type: 'redline', target: { exactText: 'B. Header' }, modified: 'B. Header' }]);
        assert.equal(legacy.status, 'ok', JSON.stringify(innerError(legacy) || null));
        const legacyViews = views(legacy.documentXml);
        assert.deepEqual(legacyViews.rejected.filter(record => record.text).map(record => [record.text, record.numId]), [['B. Header', null]], 'the legacy fallback Reject All restores the original');
        assert.ok(legacyViews.accepted.some(record => record.numId !== null), 'the legacy fallback still binds a list');
        assert.equal(el(parse(legacy.documentXml), 'p').length, 2, 'the legacy form is one deleted plus one inserted paragraph');
        // The legacy form was never safe for markers inside hyperlinks, smart tags or the author's own insertion:
        // its Reject All drops the wrapper, so the verifier refuses it (and the in-place core refuses the shape).
        for (const body of [
            `<w:p><w:hyperlink r:id="rId9"><w:r><w:t xml:space="preserve">B. </w:t></w:r></w:hyperlink><w:r><w:t>Header</w:t></w:r></w:p>`,
            `<w:p><w:smartTag w:uri="urn:x" w:element="place"><w:r><w:t xml:space="preserve">B. </w:t></w:r></w:smartTag><w:r><w:t>Header</w:t></w:r></w:p>`
        ]) {
            const refusedLegacy = await bare(documentOf(body), [{ type: 'redline', target: { exactText: 'B. Header' }, modified: 'B. Header' }]);
            assert.equal(innerError(refusedLegacy).code, 'REVISION_FIDELITY_VIOLATION');
            const direct = await bare(documentOf(body), [{ type: 'list-format', target: { exactText: 'B. Header' }, list: { mode: 'convert-plain' } }]);
            assert.equal(innerError(direct).code, 'LIST_MARKER_UNSUPPORTED');
        }
    } finally {
        console.log = quiet;
    }
}

// ---------------------------------------------------------------------------------------------------------
// 9. WP-1 diagnostics now point at list-format
// ---------------------------------------------------------------------------------------------------------
{
    const source = documentOf(p('Header'));
    const refused = await bare(source, [{ type: 'redline', target: { exactText: 'Header' }, modified: 'C. Header' }], { options: { atomic: true } });
    const error = innerError(refused);
    assert.equal(error.code, 'UNSUPPORTED_LIST_CONVERSION');
    assert.equal(error.recovery.action, 'use_list_format_convert_plain');
    assert.deepEqual(error.recovery.suggestedOperation, { type: 'list-format', list: { mode: 'convert-plain', format: 'upperLetter', strip: 'none', start: 3 } });
    // Following the suggestion (plus the target) produces the requested list item.
    const followed = await bare(source, [{ ...error.recovery.suggestedOperation, target: { exactText: 'Header' } }]);
    assert.equal(followed.status, 'ok');
    const definition = definitionOf(mergeNumberingParts('', followed.numberingXmlParts), followed.results[0].list.items[0].numId);
    assert.deepEqual([definition.numFmt, definition.startOverride], ['upperLetter', '3']);
    assert.equal(views(followed.documentXml).accepted[0].text, 'Header');
}

// ---------------------------------------------------------------------------------------------------------
// 10. WP-5 change-format and join. Fixture numbering: every case runs with the Reject All verifier on.
// ---------------------------------------------------------------------------------------------------------
const lv = (ilvl, numFmt, lvlText, extra = '') => `<w:lvl w:ilvl="${ilvl}"><w:start w:val="1"/><w:numFmt w:val="${numFmt}"/>${extra}<w:lvlText w:val="${lvlText}"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="${720 * (ilvl + 1)}" w:hanging="360"/></w:pPr>${numFmt === 'bullet' ? '<w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/></w:rPr>' : ''}</w:lvl>`;
const abstractNum = (id, levels, extra = '') => `<w:abstractNum w:abstractNumId="${id}"><w:nsid w:val="0000AB${String(id).padStart(2, '0')}"/><w:multiLevelType w:val="hybridMultilevel"/>${extra}${levels.join('')}</w:abstractNum>`;
const numEl = (numId, abstractId, overrides = '') => `<w:num w:numId="${numId}"><w:abstractNumId w:val="${abstractId}"/>${overrides}</w:num>`;
const LISTS = `<w:numbering xmlns:w="${W}">${[
    abstractNum(10, [lv(0, 'bullet', ''), lv(1, 'bullet', 'o')]),
    abstractNum(11, [lv(0, 'decimal', '%1.'), lv(1, 'lowerLetter', '%2.'), lv(2, 'lowerRoman', '%3.')]),
    abstractNum(12, [lv(0, 'decimal', '%1)', '<w:isLgl/>')]),
    abstractNum(13, [lv(0, 'decimal', '%1.')]),
    abstractNum(14, [lv(0, 'decimal', '%1.')]),
    abstractNum(15, [lv(0, 'decimal', '%1.').replace('<w:numFmt', '<w:pStyle w:val="ListNumber"/><w:numFmt')]),
    abstractNum(16, [lv(0, 'decimal', '%1.')]),
    abstractNum(17, [lv(0, 'decimal', '%1.'), lv(1, 'lowerLetter', '%2.', '<w:lvlRestart w:val="0"/>')]),
    abstractNum(18, [lv(0, 'decimal', '%1.')])
].join('')}${[
    numEl(1, 10), numEl(2, 11), numEl(3, 12), numEl(4, 13, '<w:lvlOverride w:ilvl="0"><w:startOverride w:val="5"/></w:lvlOverride>'),
    numEl(5, 14), numEl(6, 14), numEl(7, 15), numEl(8, 16, `<w:lvlOverride w:ilvl="0">${lv(0, 'upperRoman', '%1.')}</w:lvlOverride>`), numEl(9, 17),
    numEl(11, 18)
].join('')}</w:numbering>`;
const LIST_STYLES = `<w:styles xmlns:w="${W}"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="Clause"><w:name w:val="Clause"/><w:pPr><w:numPr><w:numId w:val="11"/></w:numPr></w:pPr></w:style></w:styles>`;
const li = (text, numId, ilvl = 0, extraPPr = '') => p(text, { pPr: `<w:numPr><w:ilvl w:val="${ilvl}"/><w:numId w:val="${numId}"/></w:numPr>${extraPPr}` });
const listContext = { context: { numberingXml: LISTS, stylesXml: LIST_STYLES } };
const numPrRecords = records => records.map(record => [record.text, record.numId, record.ilvl]);
const resultNumbering = result => mergeNumberingParts(LISTS, result.numberingXmlParts);
const levelOf = (numberingXml, numId, ilvl) => {
    const doc = parse(numberingXml);
    const num = el(doc, 'num').find(node => attr(node, 'numId') === String(numId));
    const abstract = el(doc, 'abstractNum').find(node => attr(node, 'abstractNumId') === attr(el(num, 'abstractNumId')[0], 'val'));
    return el(abstract, 'lvl').find(node => attr(node, 'ilvl') === String(ilvl));
};
{
    // (a) bullet -> decimal, scope "list" (default): every live item of numId 1 is rebound to ONE new definition with a
    // tracked pPrChange; the shared definition is never edited; Reject All restores numId 1 exactly, Accept All the new one.
    const source = documentOf(p('Intro.'), li('Bullet one', 1), li('Bullet two', 1), li('Bullet nested', 1, 1), li('Bullet three', 1), p('Outro.'), li('Other list', 2));
    const result = await bare(source, [{ type: 'list-format', target: { exactText: 'Bullet two' }, list: { mode: 'change-format', format: 'decimal' } }], listContext);
    assert.equal(result.status, 'ok', JSON.stringify(innerError(result) || null));
    const { list } = result.results[0];
    assert.equal(list.mode, 'change-format');
    assert.equal(list.scope, 'list');
    assert.equal(list.sourceNumId, '1');
    assert.deepEqual([list.format, list.level, list.lvlText], ['decimal', 0, '%1.']);
    assert.deepEqual(list.items.map(item => [item.paragraphIndex, item.level, item.inTarget]), [[2, 0, false], [3, 0, true], [4, 1, false], [5, 0, false]]);
    assert.ok(Number(list.numId) > 11, 'a fresh numId');
    const before = views(source).accepted;
    const { accepted, rejected } = views(result.documentXml);
    assert.deepEqual(numPrRecords(rejected), numPrRecords(before), 'Reject All restores every numPr exactly');
    assert.deepEqual(numPrRecords(accepted), numPrRecords(before).map(([text, numId, ilvl]) => [text, numId === '1' ? list.numId : numId, ilvl]), 'Accept All binds the new list, levels kept');
    assert.equal(el(parse(result.documentXml), 'pPrChange').length, 4);
    const numberingXml = resultNumbering(result);
    const level0 = levelOf(numberingXml, list.numId, 0);
    assert.equal(attr(el(level0, 'numFmt')[0], 'val'), 'decimal');
    assert.equal(el(level0, 'rFonts').length, 0, 'the bullet symbol font is not kept for digits');
    assert.equal(attr(el(level0, 'ind')[0], 'left'), '720', 'level geometry is copied');
    const level1 = levelOf(numberingXml, list.numId, 1);
    assert.deepEqual([attr(el(level1, 'numFmt')[0], 'val'), attr(el(level1, 'lvlText')[0], 'val'), el(level1, 'rFonts').length], ['bullet', 'o', 1], 'other levels are copied unchanged');
    assert.equal(canonicalNumberingDefinition(numberingXml, 1), canonicalNumberingDefinition(LISTS, 1), 'the source definition is untouched');
    assert.deepEqual(validateNumberingIntegrity(result.documentXml, numberingXml), []);
    // The library's own Reject All / Accept All transforms agree.
    assert.deepEqual(numPrRecords(views(rejectTrackedChangesInOoxml(result.documentXml, { allAuthors: true }).oxml).accepted), numPrRecords(before));
    assert.deepEqual(numPrRecords(views(acceptTrackedChangesInOoxml(result.documentXml, { allAuthors: true }).oxml).accepted), numPrRecords(accepted));
}
{
    // (b) decimal -> upperLetter keeps the level text; isLgl (forces arabic numbers) is dropped; the source num's own
    // startOverride and list.start become the start of the NEW abstractNum's level (never a shared definition).
    const source = documentOf(li('Legal one', 3), li('Legal two', 3), li('Restarted one', 4), li('Restarted two', 4));
    const result = await bare(source, [
        { type: 'list-format', target: { exactText: 'Legal one' }, list: { mode: 'change-format', format: 'upperLetter' } },
        { type: 'list-format', target: { exactText: 'Restarted one' }, list: { mode: 'change-format', format: 'lowerRoman' } }
    ], listContext);
    assert.equal(result.status, 'ok', JSON.stringify(innerError(result) || null));
    const [legal, restarted] = result.results.map(item => item.list);
    const numberingXml = resultNumbering(result);
    const legalLevel = levelOf(numberingXml, legal.numId, 0);
    assert.deepEqual([attr(el(legalLevel, 'numFmt')[0], 'val'), attr(el(legalLevel, 'lvlText')[0], 'val'), el(legalLevel, 'isLgl').length], ['upperLetter', '%1)', 0]);
    assert.deepEqual(restarted.levelStarts, { 0: 5 }, 'the source instance restart value is carried over');
    assert.deepEqual([definitionOf(numberingXml, restarted.numId).start, definitionOf(numberingXml, restarted.numId).startOverride], ['5', null]);
    const withStart = await bare(source, [{ type: 'list-format', target: { exactText: 'Legal two' }, list: { mode: 'change-format', start: 3 } }], listContext);
    assert.equal(withStart.status, 'ok', JSON.stringify(innerError(withStart) || null));
    assert.equal(definitionOf(resultNumbering(withStart), withStart.results[0].list.numId).start, '3');
    assert.equal(canonicalNumberingDefinition(resultNumbering(withStart), 3), canonicalNumberingDefinition(LISTS, 3), 'the source definition is never mutated');
    assert.equal(levelOf(resultNumbering(withStart), withStart.results[0].list.numId, 0).getElementsByTagNameNS(W, 'isLgl').length, 1, 'start-only change keeps the format');
}
{
    // (c) nested level: the target's level is redefined, the other levels are copied, every item keeps its level.
    const source = documentOf(li('Top one', 2), li('Sub a', 2, 1), li('Sub b', 2, 1), li('Top two', 2), li('Sub c', 2, 1));
    const result = await bare(source, [{ type: 'list-format', target: { exactText: 'Sub b' }, list: { mode: 'change-format', format: 'upperRoman', lvlText: '(%2)' } }], listContext);
    assert.equal(result.status, 'ok', JSON.stringify(innerError(result) || null));
    const { list } = result.results[0];
    assert.equal(list.level, 1);
    const numberingXml = resultNumbering(result);
    assert.deepEqual([attr(el(levelOf(numberingXml, list.numId, 1), 'numFmt')[0], 'val'), attr(el(levelOf(numberingXml, list.numId, 1), 'lvlText')[0], 'val')], ['upperRoman', '(%2)']);
    assert.equal(attr(el(levelOf(numberingXml, list.numId, 0), 'numFmt')[0], 'val'), 'decimal');
    assert.deepEqual(numPrRecords(views(result.documentXml).rejected), numPrRecords(views(source).accepted));
    // A mixed-level target without list.level is refused with the levels to choose from; list.level resolves it.
    const mixed = await bare(source, [{ type: 'list-format', target: { exactText: 'Top one' }, targetEnd: { exactText: 'Sub a' }, list: { mode: 'change-format', format: 'upperLetter' } }], listContext);
    assert.equal(innerError(mixed).code, 'LIST_FORMAT_MIXED_RANGE');
    assert.equal(innerError(mixed).reason, 'levels');
    assert.deepEqual(innerError(mixed).recovery.allowedValues, [0, 1]);
    const chosen = await bare(source, [{ type: 'list-format', target: { exactText: 'Top one' }, targetEnd: { exactText: 'Sub a' }, list: { mode: 'change-format', format: 'upperLetter', level: 0 } }], listContext);
    assert.equal(chosen.status, 'ok', JSON.stringify(innerError(chosen) || null));
}
{
    // (d) scope "range" splits the list at the range, which must run to the end of the list; the new instance
    // continues the count with per-level start values on the new definition (level 0 in use at 2, level 1 at b).
    const source = documentOf(li('Item one', 2), li('Item two', 2), li('Item two a', 2, 1), li('Item two b', 2, 1), li('Item three', 2), li('Item four', 2));
    const suffix = await bare(source, [{ type: 'list-format', target: { exactText: 'Item two b' }, targetEnd: { exactText: 'Item four' }, list: { mode: 'change-format', format: 'upperLetter', level: 0, scope: 'range' } }], listContext);
    assert.equal(suffix.status, 'ok', JSON.stringify(innerError(suffix) || null));
    const { list } = suffix.results[0];
    assert.deepEqual(list.items.map(item => item.paragraphIndex), [4, 5, 6]);
    assert.deepEqual(list.levelStarts, { 0: 2, 1: 2 }, 'level 0 is implicitly in use at its current value 2; level 1 continues at b');
    assert.equal(definitionOf(resultNumbering(suffix), list.numId).startOverride, null, 'no per-instance override (Word renumbers those after Accept All)');
    const { accepted, rejected } = views(suffix.documentXml);
    assert.deepEqual(accepted.map(record => record.numId === list.numId), [false, false, false, true, true, true]);
    assert.deepEqual(numPrRecords(rejected), numPrRecords(views(source).accepted));
    // Mid-sequence: later items would stay on the old list and renumber. Refused, naming the last item.
    const middle = await bare(source, [{ type: 'list-format', target: { exactText: 'Item two' }, targetEnd: { exactText: 'Item three' }, list: { mode: 'change-format', format: 'upperLetter', level: 0, scope: 'range' } }], listContext);
    const error = innerError(middle);
    assert.equal(error.code, 'LIST_FORMAT_SPLIT_UNSUPPORTED');
    assert.equal(error.reason, 'mid-sequence');
    assert.deepEqual(error.followingParagraphIndexes, [6]);
    assert.equal(error.lastListParagraphIndex, 6);
    assert.equal(error.recovery.action, 'extend_range_or_use_list_scope');
    assert.equal(middle.documentXml, source);
    // A level whose start must be overridden and that restarts inside the range cannot be continued exactly.
    const restart = await bare(source, [{ type: 'list-format', target: { exactText: 'Item two b' }, targetEnd: { exactText: 'Item four' }, list: { mode: 'change-format', format: 'upperLetter', level: 1, scope: 'range' } }], listContext);
    assert.equal(restart.status, 'ok', 'level 1 does not occur again after the restart, so the count is still exact');
    const nested = documentOf(li('Item one', 2), li('Item one a', 2, 1), li('Item one b', 2, 1), li('Item two', 2), li('Item two a', 2, 1));
    const refusedNested = await bare(nested, [{ type: 'list-format', target: { exactText: 'Item one b' }, targetEnd: { exactText: 'Item two a' }, list: { mode: 'change-format', format: 'upperRoman', level: 1, scope: 'range' } }], listContext);
    assert.equal(innerError(refusedNested).code, 'LIST_FORMAT_SPLIT_UNSUPPORTED');
    assert.equal(innerError(refusedNested).reason, 'nested-restart');
    // Starting at a top-level item is exact.
    const topLevel = await bare(nested, [{ type: 'list-format', target: { exactText: 'Item two' }, targetEnd: { exactText: 'Item two a' }, list: { mode: 'change-format', format: 'upperRoman', level: 1, scope: 'range' } }], listContext);
    assert.equal(topLevel.status, 'ok', JSON.stringify(innerError(topLevel) || null));
    assert.deepEqual(topLevel.results[0].list.levelStarts, { 0: 2 });
}
{
    // (e) join: the later list continues the earlier one; levels kept; nothing allocated.
    const source = documentOf(li('First list one', 2), li('First list two', 2), p('Between the lists.'), li('Second list one', 3), li('Second list two', 3));
    const result = await bare(source, [{ type: 'list-format', target: { exactText: 'Second list one' }, list: { mode: 'join', continueFrom: { exactText: 'First list two' } } }], listContext);
    assert.equal(result.status, 'ok', JSON.stringify(innerError(result) || null));
    const { list } = result.results[0];
    assert.deepEqual([list.mode, list.sourceNumId, list.numId, list.continueFrom.paragraphIndex], ['join', '3', '2', 2]);
    assert.equal(result.numberingXmlParts.length, 0);
    const { accepted, rejected } = views(result.documentXml);
    assert.deepEqual(accepted.map(record => record.numId), ['2', '2', null, '2', '2']);
    assert.deepEqual(numPrRecords(rejected), numPrRecords(views(source).accepted));
    // Same list: nothing to do (ok, no change, warning).
    const same = await bare(source, [{ type: 'list-format', target: { exactText: 'First list two' }, list: { mode: 'join', continueFrom: { exactText: 'First list one' } } }], listContext);
    assert.equal(same.status, 'ok');
    assert.equal(same.hasChanges, false);
    assert.ok(same.results[0].warnings.some(warning => warning.startsWith('LIST_FORMAT_UNCHANGED: ')));
    for (const [label, documentXml, operation, code, reason] of [
        ['continueFrom after the target', source, { target: { exactText: 'First list one' }, list: { mode: 'join', continueFrom: { exactText: 'Second list one' } } }, 'LIST_FORMAT_JOIN_INVALID', 'after-target'],
        ['continueFrom is not a list item', source, { target: { exactText: 'Second list one' }, list: { mode: 'join', continueFrom: { exactText: 'Between the lists.' } } }, 'LIST_FORMAT_JOIN_INVALID', 'not-a-list'],
        ['continued list has later items', documentOf(li('First one', 2), li('Second one', 3), li('First two', 2)), { target: { exactText: 'Second one' }, list: { mode: 'join', continueFrom: { exactText: 'First one' } } }, 'LIST_FORMAT_SPLIT_UNSUPPORTED', 'interleaved'],
        ['joined range mid-sequence', source, { target: { exactText: 'Second list one' }, list: { mode: 'join', continueFrom: { exactText: 'First list one' }, scope: 'range' } }, 'LIST_FORMAT_SPLIT_UNSUPPORTED', 'mid-sequence'],
        ['continueFrom not found', source, { target: { exactText: 'Second list one' }, list: { mode: 'join', continueFrom: { exactText: 'No such paragraph' } } }, 'TARGET_NOT_FOUND', undefined]
    ]) {
        const refused = await bare(documentXml, [{ type: 'list-format', ...operation }], listContext);
        assert.equal(refused.status, 'error', label);
        const error = innerError(refused);
        assert.equal(error.code, code, `${label}: ${error.code} ${error.message}`);
        assert.equal(error.reason, reason, label);
        assert.equal(refused.documentXml, documentXml, `${label}: nothing written`);
    }
    // A joined suffix range is allowed: the earlier items of the second list keep their numbering.
    const suffix = await bare(source, [{ type: 'list-format', target: { exactText: 'Second list two' }, list: { mode: 'join', continueFrom: { exactText: 'First list one' }, scope: 'range' } }], listContext);
    assert.equal(suffix.status, 'ok', JSON.stringify(innerError(suffix) || null));
    assert.deepEqual(views(suffix.documentXml).accepted.map(record => record.numId), ['2', '2', null, '3', '2']);
}
{
    // (f) Refusals: recovery-enveloped, nothing written.
    const deletedMark = `<w:rPr><w:del w:id="70" w:author="Other" w:date="2026-01-01T00:00:00Z"/></w:rPr>`;
    for (const [label, paragraphs, operation, code, reason, action] of [
        ['plain target', [p('Plain paragraph'), li('Item', 2)], { target: { exactText: 'Plain paragraph' }, list: { mode: 'change-format', format: 'upperLetter' } }, 'LIST_FORMAT_TARGET_NOT_LIST', undefined, 'use_list_format_convert_plain'],
        ['list and non-list paragraphs', [li('Item', 2), p('Body text'), li('Item two', 2)], { target: { exactText: 'Item' }, targetEnd: { exactText: 'Item two' }, list: { mode: 'change-format', format: 'upperLetter' } }, 'LIST_FORMAT_MIXED_RANGE', 'non-list', 'narrow_range'],
        ['two lists in one range', [li('Item', 2), li('Other item', 3)], { target: { exactText: 'Item' }, targetEnd: { exactText: 'Other item' }, list: { mode: 'change-format', format: 'upperLetter' } }, 'LIST_FORMAT_MIXED_RANGE', 'several-lists', 'narrow_range'],
        ['style-linked level', [li('Styled', 7)], { target: { exactText: 'Styled' }, list: { mode: 'change-format', format: 'upperLetter' } }, 'LIST_FORMAT_STYLE_LINKED', undefined, 'convert_in_host'],
        ['a paragraph style binds the list', [li('Clause direct', 11)], { target: { exactText: 'Clause direct' }, list: { mode: 'change-format', format: 'upperLetter' } }, 'LIST_FORMAT_STYLE_LINKED', undefined, 'convert_in_host'],
        ['numbered through its style only', [p('Clause by style', { pPr: '<w:pStyle w:val="Clause"/>' })], { target: { exactText: 'Clause by style' }, list: { mode: 'change-format', format: 'upperLetter' } }, 'LIST_FORMAT_STYLE_NUMBERED', undefined, 'convert_in_host'],
        ['shared abstract definition', [li('Shared one', 5), li('Shared two', 6)], { target: { exactText: 'Shared one' }, list: { mode: 'change-format', format: 'upperLetter' } }, 'LIST_FORMAT_SPLIT_UNSUPPORTED', 'shared-definition', 'convert_in_host'],
        ['instance level override', [li('Overridden', 8)], { target: { exactText: 'Overridden' }, list: { mode: 'change-format', format: 'upperLetter' } }, 'LIST_FORMAT_DEFINITION_UNSUPPORTED', 'level-override', 'convert_in_host'],
        ['lvlRestart in a split', [li('Restart one', 9), li('Restart one a', 9, 1), li('Restart two', 9)], { target: { exactText: 'Restart two' }, list: { mode: 'change-format', format: 'upperLetter', scope: 'range' } }, 'LIST_FORMAT_DEFINITION_UNSUPPORTED', 'level-restart', 'extend_range_or_use_list_scope'],
        ['undefined numId', [li('Dangling', 42)], { target: { exactText: 'Dangling' }, list: { mode: 'change-format', format: 'upperLetter' } }, 'LIST_FORMAT_DEFINITION_UNSUPPORTED', 'missing', 'inspect_numbering'],
        ['bullet text on a numbered format', [li('Item', 2)], { target: { exactText: 'Item' }, list: { mode: 'change-format', lvlText: '*' } }, 'INVALID_OPERATION', undefined, 'change_request'],
        ['another author pending pPrChange outside the target', [li('Item', 2), li('Item two', 2, 0, '<w:pPrChange w:id="9" w:author="Other" w:date="2026-01-01T00:00:00Z"><w:pPr/></w:pPrChange>')], { target: { exactText: 'Item' }, list: { mode: 'change-format', format: 'upperLetter' } }, 'PARAGRAPH_FORMAT_REVISION_CONFLICT', undefined, 'resolve_pending_revision']
    ]) {
        const source = documentOf(...paragraphs);
        const result = await bare(source, [{ type: 'list-format', ...operation }], listContext);
        assert.equal(result.status, 'error', label);
        const error = innerError(result);
        assert.equal(error.code, code, `${label}: ${error.code} ${error.message}`);
        assert.equal(error.reason, reason, label);
        assert.equal(error.recovery?.action, action, label);
        assert.equal(error.recovery.sameArgumentsSafe, false, label);
        assert.equal(result.documentXml, source, `${label}: nothing written`);
        assert.equal(result.numberingXmlParts.length, 0, label);
    }
    // R1: a tracked-deleted list paragraph keeps its numPr (scope list skips it; in a range it is reported).
    const source = documentOf(li('Kept one', 2), p('Deleted item', { pPr: `<w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr>${deletedMark}` }), li('Kept two', 2));
    const result = await bare(source, [{ type: 'list-format', target: { exactText: 'Kept one' }, targetEnd: { exactText: 'Kept two' }, list: { mode: 'change-format', format: 'upperLetter' } }], listContext);
    assert.equal(result.status, 'ok', JSON.stringify(innerError(result) || null));
    assert.deepEqual(result.results[0].list.skipped, [{ paragraphIndex: 2, reason: 'tracked-deleted' }]);
    const deleted = el(parse(result.documentXml), 'p')[1];
    assert.equal(el(deleted, 'pPrChange').length, 0);
    assert.equal(attr(el(deleted, 'numId')[0], 'val'), '2', 'the deleted paragraph keeps numId 2');
    // Same-author earlier snapshot (the paragraph was bound earlier by the same author) is kept: Reject All returns to
    // the unnumbered state, not to the intermediate list.
    const earlier = documentOf(li('Bound earlier', 2, 0, `<w:pPrChange w:id="9" w:author="${AUTHOR}" w:date="2026-01-01T00:00:00Z"><w:pPr/></w:pPrChange>`), li('Item two', 2));
    const kept = await bare(earlier, [{ type: 'list-format', target: { exactText: 'Item two' }, list: { mode: 'change-format', format: 'upperLetter' } }], listContext);
    assert.equal(kept.status, 'ok', JSON.stringify(innerError(kept) || null));
    assert.deepEqual(views(kept.documentXml).rejected.map(record => record.numId), [null, '2']);
    assert.equal(el(parse(kept.documentXml), 'pPrChange').length, 2);
    // Unchanged: the list already has the requested format.
    const unchanged = await bare(documentOf(li('Item', 2)), [{ type: 'list-format', target: { exactText: 'Item' }, list: { mode: 'change-format', format: 'decimal' } }], listContext);
    assert.equal(unchanged.status, 'ok');
    assert.equal(unchanged.hasChanges, false);
    assert.ok(unchanged.results[0].warnings.some(warning => warning.startsWith('LIST_FORMAT_UNCHANGED: ')));
}
{
    // (g) Batch: a second change-format on a list the first one already rebound reads the batch's own definition;
    // the earliest snapshot is kept, so Reject All still restores the source binding.
    const source = documentOf(li('Top one', 2), li('Sub a', 2, 1), li('Top two', 2));
    const result = await bare(source, [
        { type: 'list-format', target: { exactText: 'Top one' }, list: { mode: 'change-format', format: 'upperRoman' } },
        { type: 'list-format', target: { exactText: 'Sub a' }, list: { mode: 'change-format', format: 'bullet' } }
    ], listContext);
    assert.equal(result.status, 'ok', JSON.stringify(innerError(result) || null));
    const [first, second] = result.results.map(item => item.list);
    assert.equal(second.sourceNumId, first.numId, 'the second operation starts from the first one\'s definition');
    const numberingXml = resultNumbering(result);
    assert.equal(attr(el(levelOf(numberingXml, second.numId, 0), 'numFmt')[0], 'val'), 'upperRoman', 'the first change is kept');
    assert.deepEqual([attr(el(levelOf(numberingXml, second.numId, 1), 'numFmt')[0], 'val'), attr(el(levelOf(numberingXml, second.numId, 1), 'lvlText')[0], 'val')], ['bullet', '•']);
    assert.equal(el(parse(result.documentXml), 'pPrChange').length, 3);
    assert.deepEqual(numPrRecords(views(result.documentXml).rejected), numPrRecords(views(source).accepted));
    assert.deepEqual(validateNumberingIntegrity(result.documentXml, numberingXml), []);
    // generateRedlines: false rebinds directly, but only inside the target: an untracked edit has no Reject All.
    const outside = await bare(source, [{ type: 'list-format', target: { exactText: 'Top one' }, list: { mode: 'change-format', format: 'upperRoman' }, generateRedlines: false }], listContext);
    assert.equal(innerError(outside).code, 'INVALID_OPERATION');
    assert.equal(innerError(outside).field, 'generateRedlines');
    const direct = await bare(source, [{ type: 'list-format', target: { exactText: 'Top one' }, targetEnd: { exactText: 'Top two' }, list: { mode: 'change-format', format: 'upperRoman', level: 0 }, generateRedlines: false }], listContext);
    assert.equal(direct.status, 'ok', JSON.stringify(innerError(direct) || null));
    assert.equal(el(parse(direct.documentXml), 'pPrChange').length, 0);
    assert.ok(views(direct.documentXml).rejected.every(record => record.numId === direct.results[0].list.numId));
}
{
    // (h) Packaged openDocx path (Word-authored template numbering): labels after Accept All / Reject All.
    const listItem = (text, numId, ilvl = 0) => p(text, { pPr: `<w:numPr><w:ilvl w:val="${ilvl}"/><w:numId w:val="${numId}"/></w:numPr>` });
    const bytes = buildDocx([listItem('Bullet alpha', 1), listItem('Bullet beta', 1), p('Between.'), listItem('Step one', 2), listItem('Step two', 2), p('Gap.'), listItem('Step three', 2)].join(''));
    const before = labelsOf(bytes);
    const result = await openDocx(bytes).applyOperations([
        { type: 'list-format', target: { exactText: 'Bullet beta' }, list: { mode: 'change-format', format: 'decimal' } },
        { type: 'list-format', target: { exactText: 'Step one' }, list: { mode: 'change-format', format: 'upperLetter' } }
    ], { author: AUTHOR });
    assert.equal(result.status, 'ok', JSON.stringify(innerError(result) || null));
    const output = result.toUint8Array();
    const entries = unzipDocx(output);
    const numberingXml = dec.decode(entries.get('word/numbering.xml'));
    const sourceNumbering = dec.decode(unzipDocx(bytes).get('word/numbering.xml'));
    for (const numId of [1, 2]) assert.equal(canonicalNumberingDefinition(numberingXml, numId), canonicalNumberingDefinition(sourceNumbering, numId), `numId ${numId} is untouched`);
    assert.deepEqual(validateNumberingIntegrity(dec.decode(entries.get('word/document.xml')), numberingXml), []);
    const accepted = await openDocx(output).resolveRevisions('accept', { allAuthors: true });
    const acceptedLabels = labelsOf(accepted.toUint8Array());
    assert.deepEqual(acceptedLabels.map(([text]) => text), before.map(([text]) => text));
    assert.deepEqual(acceptedLabels.filter(([text]) => text.startsWith('Step')).map(([, label]) => label), ['A.', 'B.', 'C.'], 'Accept All shows the new labels');
    assert.deepEqual(acceptedLabels.filter(([text]) => text.startsWith('Bullet')).map(([, label]) => label), ['1.', '2.']);
    const rejected = await openDocx(output).resolveRevisions('reject', { allAuthors: true });
    assert.deepEqual(labelsOf(rejected.toUint8Array()), before, 'Reject All restores the original labels');

    const joined = await openDocx(bytes).applyOperations([{ type: 'list-format', target: { exactText: 'Step one' }, list: { mode: 'join', continueFrom: { exactText: 'Bullet beta' } } }], { author: AUTHOR });
    assert.equal(joined.status, 'ok', JSON.stringify(innerError(joined) || null));
    const joinedAccepted = labelsOf((await openDocx(joined.toUint8Array()).resolveRevisions('accept', { allAuthors: true })).toUint8Array());
    assert.ok(joinedAccepted.filter(([text]) => text.startsWith('Step')).every(([, label]) => label === before[0][1]), 'joined items take the continued list\'s format');
    assert.deepEqual(labelsOf((await openDocx(joined.toUint8Array()).resolveRevisions('reject', { allAuthors: true })).toUint8Array()), before);
}

console.log('PASS: list-format operation');
