import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';

import './setup-xml-provider.mjs';
import { openDocx } from '../index.js';
import { parseOoxmlSafe, createSerializer } from '../adapters/xml-adapter.js';
import { unzipDocx, zipDocx } from '../document/zip-archive.js';
import { applyOperationsToDocumentXml } from '../services/standalone-operation-runner.js';
import { validateDocumentOperation } from '../services/document-operation-contract.js';
import { mergeNumberingParts, validateNumberingIntegrity } from '../services/numbering-allocation.js';
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
    for (const [label, list, field] of [
        ['missing list', undefined, 'list'],
        ['change-format is WP-5', { mode: 'change-format', format: 'decimal' }, 'list.mode'],
        ['join is WP-5', { mode: 'join' }, 'list.mode'],
        ['scope is WP-5', { mode: 'convert-plain', scope: 'list' }, 'list.scope'],
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
    assert.deepEqual(schema.$defs.listFormat.properties.mode, { const: 'convert-plain' });
    assert.ok(schema.$defs.operation.oneOf.some(branch => JSON.stringify(branch).includes('"const":"list-format"')));
    assert.equal(schema.$defs.listFormat.additionalProperties, false);

    const version = await executeCli(['version']);
    assert.equal(version.contractVersion, 9);
    assert.ok(version.capabilities.includes('list-format-v1'));
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

console.log('PASS: list-format operation');
