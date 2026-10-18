import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';

import './setup-xml-provider.mjs';
import { openDocx } from '../index.js';
import { parseOoxmlSafe } from '../adapters/xml-adapter.js';
import { unzipDocx } from '../document/zip-archive.js';
import { applyOperationsToDocumentXml } from '../services/standalone-operation-runner.js';
import { mergeNumberingParts } from '../services/numbering-allocation.js';
import { projectWordRevisionView } from '../core/word-revision-semantics.js';
import {
    PPR_SCHEMA_ORDER,
    applyParagraphPropertiesToPPr,
    checkParagraphPropertiesChanged,
    insertPPrChildInOrder
} from '../engine/rpr-helpers.js';
import { snapshotAndAttachPPrChange } from '../engine/run-builders.js';
import { parseMarkerStart } from '../orchestration/list-structural-fallback.js';
import { createWordElement } from '../core/word-xml.js';
import { generateOperation, makeRng } from './helpers/operation-generators.mjs';

/**
 * Canonical list operations WP-3: `paragraph-format.numbering` binding recorded as a tracked w:pPrChange.
 * Every end-to-end case runs with default options (the Reject All verifier refuses by default) through both
 * the bare runner and the packaged `openDocx` facade.
 */

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const AUTHOR = 'Binding Author';
const dec = new TextDecoder();
const parse = xml => {
    const parsed = parseOoxmlSafe(xml);
    assert.ok(parsed.doc, parsed.error?.message || 'parseable OOXML');
    return parsed.doc;
};
const el = (doc, name) => Array.from(doc.getElementsByTagNameNS(W, name));
const pPrOf = p => Array.from(p.childNodes).find(c => c.nodeType === 1 && c.localName === 'pPr');
const childNames = node => Array.from(node.childNodes).filter(c => c.nodeType === 1).map(c => c.nodeName);

// ---------------------------------------------------------------------------------------------------------
// 1. pPr child ordering (CT_PPr sequence) and the pPrChange snapshot
// ---------------------------------------------------------------------------------------------------------
{
    assert.ok(PPR_SCHEMA_ORDER.indexOf('w:numPr') > PPR_SCHEMA_ORDER.indexOf('w:widowControl'));
    assert.ok(PPR_SCHEMA_ORDER.indexOf('w:numPr') < PPR_SCHEMA_ORDER.indexOf('w:suppressLineNumbers'));
    assert.ok(PPR_SCHEMA_ORDER.indexOf('w:numPr') < PPR_SCHEMA_ORDER.indexOf('w:ind'));

    const xmlDoc = parse(`<w:p xmlns:w="${W}"><w:pPr><w:pStyle w:val="Body"/><w:keepNext/><w:keepLines/><w:pageBreakBefore/><w:widowControl/><w:spacing w:after="120"/><w:ind w:left="720"/><w:jc w:val="left"/><w:outlineLvl w:val="1"/><w:rPr><w:b/></w:rPr></w:pPr></w:p>`);
    const pPr = el(xmlDoc, 'pPr')[0];
    applyParagraphPropertiesToPPr(xmlDoc, pPr, { numbering: { numId: 3, level: 2 } });
    assert.deepEqual(childNames(pPr), [
        'w:pStyle', 'w:keepNext', 'w:keepLines', 'w:pageBreakBefore', 'w:widowControl', 'w:numPr',
        'w:spacing', 'w:ind', 'w:jc', 'w:outlineLvl', 'w:rPr'
    ], 'numPr follows widowControl and precedes spacing/ind');
    assert.deepEqual(childNames(el(xmlDoc, 'numPr')[0]), ['w:ilvl', 'w:numId'], 'CT_NumPr order is ilvl, numId');
    assert.equal(el(xmlDoc, 'numId')[0].getAttribute('w:val'), '3');
    assert.equal(el(xmlDoc, 'ilvl')[0].getAttribute('w:val'), '2');

    // Re-applying replaces the binding (no duplicate numPr), null removes it.
    applyParagraphPropertiesToPPr(xmlDoc, pPr, { numbering: { numId: 4 } });
    assert.equal(el(xmlDoc, 'numPr').length, 1);
    assert.equal(el(xmlDoc, 'numId')[0].getAttribute('w:val'), '4');
    assert.equal(el(xmlDoc, 'ilvl')[0].getAttribute('w:val'), '0', 'level defaults to 0');
    applyParagraphPropertiesToPPr(xmlDoc, pPr, { numbering: null });
    assert.equal(el(xmlDoc, 'numPr').length, 0);

    // numPr lands before an existing pPrChange / sectPr and after pStyle on a sparse pPr.
    const sparse = parse(`<w:p xmlns:w="${W}"><w:pPr><w:pStyle w:val="X"/><w:rPr/><w:sectPr/><w:pPrChange w:id="1" w:author="A" w:date="2026-01-01T00:00:00Z"><w:pPr/></w:pPrChange></w:pPr></w:p>`);
    const sparsePPr = el(sparse, 'pPr')[0];
    const numPr = createWordElement(sparse, 'w:numPr');
    insertPPrChildInOrder(sparsePPr, numPr);
    assert.deepEqual(childNames(sparsePPr), ['w:pStyle', 'w:numPr', 'w:rPr', 'w:sectPr', 'w:pPrChange']);
}

// checkParagraphPropertiesChanged: only a real change counts, and a pPrChange snapshot is never "current".
{
    const xmlDoc = parse(`<w:p xmlns:w="${W}"><w:pPr><w:numPr><w:ilvl w:val="1"/><w:numId w:val="5"/></w:numPr><w:pPrChange w:id="1" w:author="A" w:date="2026-01-01T00:00:00Z"><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="9"/></w:numPr></w:pPr></w:pPrChange></w:pPr></w:p>`);
    const pPr = el(xmlDoc, 'pPr')[0];
    assert.equal(checkParagraphPropertiesChanged(pPr, { numbering: { numId: 5, level: 1 } }), false, 'same binding is a no-op');
    assert.equal(checkParagraphPropertiesChanged(pPr, { numbering: { numId: '5', level: 1 } }), false, 'string numId equals integer');
    assert.equal(checkParagraphPropertiesChanged(pPr, { numbering: { numId: 9, level: 0 } }), true, 'the snapshot binding is not current');
    assert.equal(checkParagraphPropertiesChanged(pPr, { numbering: { numId: 5 } }), true, 'level differs');
    assert.equal(checkParagraphPropertiesChanged(pPr, { numbering: null }), true);
    const plain = parse(`<w:p xmlns:w="${W}"><w:pPr/></w:p>`);
    assert.equal(checkParagraphPropertiesChanged(el(plain, 'pPr')[0], { numbering: null }), false, 'removing nothing is a no-op');
}

// Snapshot: the previous pPr (without numPr, or with the old numPr) is what pPrChange holds.
{
    const xmlDoc = parse(`<w:p xmlns:w="${W}"><w:pPr><w:ind w:left="720"/></w:pPr></w:p>`);
    const pPr = el(xmlDoc, 'pPr')[0];
    snapshotAndAttachPPrChange(xmlDoc, pPr, AUTHOR);
    applyParagraphPropertiesToPPr(xmlDoc, pPr, { numbering: { numId: 1, level: 0 } });
    assert.deepEqual(childNames(pPr), ['w:numPr', 'w:ind', 'w:pPrChange']);
    const snapshot = el(xmlDoc, 'pPrChange')[0].getElementsByTagNameNS(W, 'pPr')[0];
    assert.deepEqual(childNames(snapshot), ['w:ind'], 'snapshot has the previous pPr and no numPr');
}

// ---------------------------------------------------------------------------------------------------------
// 2. Contract: shape validation
// ---------------------------------------------------------------------------------------------------------
const NUMBERING = `<w:numbering xmlns:w="${W}"><w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="multilevel"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl><w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%2."/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`;
const BODY = `<w:document xmlns:w="${W}"><w:body>`
    + `<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>Listed item</w:t></w:r></w:p>`
    + `<w:p><w:pPr><w:ind w:left="720"/><w:jc w:val="left"/></w:pPr><w:r><w:t>Plain paragraph</w:t></w:r></w:p>`
    + `<w:p><w:r><w:t>Another plain paragraph</w:t></w:r></w:p>`
    + `</w:body></w:document>`;

async function bare(operations, options = {}, documentXml = BODY) {
    return applyOperationsToDocumentXml(documentXml, operations, AUTHOR, { numberingXml: NUMBERING }, options);
}
const innerError = result => result.results?.find(item => item?.error)?.error || result.error;

for (const [label, numbering] of [
    ['zero numId', { numId: 0 }],
    ['negative level', { numId: 1, level: -1 }],
    ['level above 8', { numId: 1, level: 9 }],
    ['missing numId', { level: 0 }],
    ['unknown key', { numId: 1, start: 4 }],
    ['non-object', 'list']
]) {
    const result = await bare([{ type: 'paragraph-format', target: { exactText: 'Plain paragraph' }, properties: { numbering } }]);
    assert.equal(result.status, 'error', `${label} is rejected`);
    const error = innerError(result);
    assert.equal(error.code, 'INVALID_OPERATION', label);
    assert.equal(error.field, 'properties.numbering');
}

// ---------------------------------------------------------------------------------------------------------
// 3. Bare runner, default options
// ---------------------------------------------------------------------------------------------------------
{
    const result = await bare([{ type: 'paragraph-format', target: { exactText: 'Plain paragraph' }, properties: { numbering: { numId: 1, level: 1 } } }]);
    assert.equal(result.status, 'ok');
    assert.equal(result.hasChanges, true);
    const xmlDoc = parse(result.documentXml);
    const accepted = projectWordRevisionView(xmlDoc, 'accepted');
    const rejected = projectWordRevisionView(xmlDoc, 'rejected');
    assert.deepEqual(accepted[1].numPr && { numId: accepted[1].numPr.numId, ilvl: accepted[1].numPr.ilvl }, { numId: '1', ilvl: '1' }, 'Accept All shows the new binding');
    assert.equal(rejected[1].numPr, null, 'Reject All restores the plain paragraph');
    assert.deepEqual(rejected.map(r => r.props), projectWordRevisionView(parse(BODY), 'rejected').map(r => r.props), 'Reject All restores every pPr');
    const paragraph = el(xmlDoc, 'p')[1];
    assert.deepEqual(childNames(pPrOf(paragraph)), ['w:numPr', 'w:ind', 'w:jc', 'w:pPrChange'], 'schema order, snapshot last');
    assert.equal(result.numberingXmlParts?.length || 0, 0, 'a binding never allocates numbering');
}

// Rebinding a list item to another level, and removal.
{
    const rebind = await bare([{ type: 'paragraph-format', target: { exactText: 'Listed item' }, properties: { numbering: { numId: 1, level: 1 } } }]);
    assert.equal(rebind.status, 'ok');
    const xmlDoc = parse(rebind.documentXml);
    const rejected = projectWordRevisionView(xmlDoc, 'rejected');
    assert.deepEqual({ numId: rejected[0].numPr.numId, ilvl: rejected[0].numPr.ilvl }, { numId: '1', ilvl: '0' }, 'Reject All restores the previous numPr exactly');
    const accepted = projectWordRevisionView(xmlDoc, 'accepted');
    assert.equal(accepted[0].numPr.ilvl, '1');

    const remove = await bare([{ type: 'paragraph-format', target: { exactText: 'Listed item' }, properties: { numbering: null } }]);
    assert.equal(remove.status, 'ok');
    const removed = parse(remove.documentXml);
    assert.equal(projectWordRevisionView(removed, 'accepted')[0].numPr, null, 'Accept All shows the paragraph out of the list');
    const restored = projectWordRevisionView(removed, 'rejected')[0].numPr;
    assert.deepEqual({ numId: restored.numId, ilvl: restored.ilvl }, { numId: '1', ilvl: '0' }, 'Reject All restores list membership');
    assert.ok(el(removed, 'pPrChange')[0].getElementsByTagNameNS(W, 'numId').length === 1, 'snapshot holds the old numPr');
}

// No-op: already bound.
{
    const result = await bare([{ type: 'paragraph-format', target: { exactText: 'Listed item' }, properties: { numbering: { numId: 1, level: 0 } } }]);
    assert.equal(result.hasChanges, false);
    assert.ok(!result.documentXml.includes('pPrChange'));
}

// Untracked mode: written directly, no revision.
{
    const result = await bare([{ type: 'paragraph-format', target: { exactText: 'Plain paragraph' }, properties: { numbering: { numId: 1 } } }], { generateRedlines: false });
    assert.equal(result.status, 'ok');
    assert.ok(!result.documentXml.includes('pPrChange'));
    assert.equal(el(parse(result.documentXml), 'numPr').length, 2);
}

// Combined properties share one snapshot.
{
    const result = await bare([{ type: 'paragraph-format', target: { exactText: 'Plain paragraph' }, properties: { alignment: 'center', numbering: { numId: 1 } } }]);
    assert.equal(result.status, 'ok');
    const xmlDoc = parse(result.documentXml);
    assert.equal(el(xmlDoc, 'pPrChange').length, 1);
    const rejected = projectWordRevisionView(xmlDoc, 'rejected');
    assert.equal(rejected[1].numPr, null);
}

// ---------------------------------------------------------------------------------------------------------
// 4. Refusals carry recovery envelopes and leave the document untouched
// ---------------------------------------------------------------------------------------------------------
for (const [label, numbering, reason] of [
    ['unknown numId', { numId: 42 }, 'unknown-num'],
    ['undefined level', { numId: 1, level: 5 }, 'undefined-level']
]) {
    const result = await bare([{ type: 'paragraph-format', target: { exactText: 'Plain paragraph' }, properties: { numbering } }]);
    assert.equal(result.status, 'error', label);
    const error = innerError(result);
    assert.equal(error.code, 'PARAGRAPH_NUMBERING_NOT_FOUND', label);
    assert.equal(error.reason, reason, label);
    assert.equal(error.recovery.action, 'inspect_numbering');
    assert.equal(error.recovery.sameArgumentsSafe, false, 'unchanged retry would fail again');
    assert.equal(result.documentXml, BODY, 'refusal leaves the document unchanged');
}
{
    const noNumbering = await applyOperationsToDocumentXml(BODY, [{ type: 'paragraph-format', target: { exactText: 'Plain paragraph' }, properties: { numbering: { numId: 1 } } }], AUTHOR, {}, {});
    assert.equal(innerError(noNumbering).code, 'PARAGRAPH_NUMBERING_NOT_FOUND', 'no numbering part: nothing to bind to');
}

// ---------------------------------------------------------------------------------------------------------
// 5. Pending pPrChange: same author keeps the earliest snapshot, another author is refused
// ---------------------------------------------------------------------------------------------------------
{
    const first = await bare([{ type: 'paragraph-format', target: { exactText: 'Plain paragraph' }, properties: { alignment: 'center' } }]);
    assert.equal(first.status, 'ok');
    const second = await bare([{ type: 'paragraph-format', target: { exactText: 'Plain paragraph' }, properties: { numbering: { numId: 1 } } }], {}, first.documentXml);
    assert.equal(second.status, 'ok');
    const xmlDoc = parse(second.documentXml);
    assert.equal(el(xmlDoc, 'pPrChange').length, 1, 'one snapshot per paragraph');
    const rejected = projectWordRevisionView(xmlDoc, 'rejected');
    const original = projectWordRevisionView(parse(BODY), 'rejected');
    assert.deepEqual(rejected[1].props, original[1].props, 'Reject All restores the earliest state, not the intermediate one');
    assert.equal(rejected[1].numPr, null);

    const other = await applyOperationsToDocumentXml(first.documentXml, [{ type: 'paragraph-format', target: { exactText: 'Plain paragraph' }, properties: { numbering: { numId: 1 } } }], 'Second Reviewer', { numberingXml: NUMBERING }, {});
    assert.equal(other.status, 'error');
    assert.equal(innerError(other).code, 'PARAGRAPH_FORMAT_REVISION_CONFLICT');
    assert.equal(innerError(other).recovery.requiresUserAuthorization, true);
}

// ---------------------------------------------------------------------------------------------------------
// 6. Packaged facade (openDocx, default options) on a Word-authored package
// ---------------------------------------------------------------------------------------------------------
{
    const template = unzipDocx(new Uint8Array(readFileSync(new URL('./fixtures/agentic-lists/nested-lists-source.docx', import.meta.url))));
    const original = dec.decode(template.get('word/document.xml'));
    const inspected = openDocx(new Uint8Array(readFileSync(new URL('./fixtures/agentic-lists/nested-lists-source.docx', import.meta.url)))).inspect();
    const listed = inspected.paragraphs.find(p => p.list && p.list.numId != null);
    const plain = inspected.paragraphs.find(p => !p.list && p.exactText && p.exactText.trim() && !p.inTable);
    assert.ok(listed && plain, 'fixture has a list item and a plain paragraph');
    const bytes = new Uint8Array(readFileSync(new URL('./fixtures/agentic-lists/nested-lists-source.docx', import.meta.url)));
    const target = p => ({ paragraphId: p.paragraphId, fingerprint: p.fingerprint, exactText: p.exactText });
    const numbering = { numId: Number(listed.list.numId), level: Number(listed.list.level) || 0 };

    const result = await openDocx(bytes).applyOperations(
        [{ type: 'paragraph-format', target: target(plain), properties: { numbering } }],
        { author: AUTHOR }
    );
    assert.equal(result.status, 'ok', JSON.stringify(result.error));
    const outEntries = unzipDocx(result.toUint8Array());
    const outXml = dec.decode(outEntries.get('word/document.xml'));
    const xmlDoc = parse(outXml);
    const accepted = projectWordRevisionView(xmlDoc, 'accepted');
    const rejected = projectWordRevisionView(xmlDoc, 'rejected');
    const sourceRejected = projectWordRevisionView(parse(original), 'rejected');
    const index = rejected.findIndex((r, i) => r.text === plain.exactText && !sourceRejected[i]?.numPr && accepted[i]?.numPr);
    assert.ok(index >= 0, 'the plain paragraph is in the list in the Accept All view');
    assert.equal(accepted[index].numPr.numId, String(numbering.numId));
    assert.deepEqual(rejected.map(r => [r.text, r.props, r.numPr]), sourceRejected.map(r => [r.text, r.props, r.numPr]), 'Reject All equals the source');
    assert.equal(dec.decode(outEntries.get('word/numbering.xml')), dec.decode(template.get('word/numbering.xml')), 'numbering part untouched');

    const refused = await openDocx(bytes).applyOperations(
        [{ type: 'paragraph-format', target: target(plain), properties: { numbering: { numId: 9999, level: 0 } } }],
        { author: AUTHOR }
    );
    assert.equal(refused.status, 'error');
    assert.equal(refused.results?.find(item => item?.error)?.error?.code ?? refused.error?.code, 'PARAGRAPH_NUMBERING_NOT_FOUND');
}

// ---------------------------------------------------------------------------------------------------------
// 7. Generator kind
// ---------------------------------------------------------------------------------------------------------
{
    const paragraphs = [
        { paragraphId: 'A', fingerprint: 'f', exactText: 'One', list: { numId: '1', level: 1 } },
        { paragraphId: 'B', fingerprint: 'f', exactText: 'Two' }
    ];
    const seen = new Set();
    for (let seed = 1; seed <= 20; seed++) {
        const generated = generateOperation(makeRng(seed), paragraphs, 'paragraph-format-numbering', seed);
        assert.equal(generated.operation.type, 'paragraph-format');
        seen.add(JSON.stringify(generated.operation.properties));
    }
    assert.ok(seen.has('{"numbering":null}') && seen.has('{"numbering":{"numId":1,"level":1}}'), 'binds and removes');
    assert.equal(generateOperation(makeRng(1), [paragraphs[1]], 'paragraph-format-numbering', 0), null, 'no list, no site');
}

// ---------------------------------------------------------------------------------------------------------
// 8. Roman/letter start values (parseMarkerStart) and the single-line conversion fallback
// ---------------------------------------------------------------------------------------------------------
{
    assert.equal(parseMarkerStart('7.', 'decimal'), 7);
    assert.equal(parseMarkerStart('1.', 'decimal'), 1);
    assert.equal(parseMarkerStart('C.', 'upperAlpha'), 3);
    assert.equal(parseMarkerStart('c.', 'lowerAlpha'), 3);
    assert.equal(parseMarkerStart('A.', 'upperAlpha'), null, 'default start needs no override');
    assert.equal(parseMarkerStart('iv.', 'lowerRoman'), 4);
    assert.equal(parseMarkerStart('IX.', 'upperRoman'), 9);
    assert.equal(parseMarkerStart('XIV.', 'upperRoman'), 14);
    assert.equal(parseMarkerStart('i.', 'lowerRoman'), null);
    assert.equal(parseMarkerStart('-', 'bullet'), null);

    const quiet = console.log;
    console.log = () => { };
    try {
        for (const [marker, start] of [['C. Header', '3'], ['iv. Header', '4']]) {
            const doc = `<w:document xmlns:w="${W}"><w:body><w:p><w:r><w:t>${marker}</w:t></w:r></w:p></w:body></w:document>`;
            const result = await applyOperationsToDocumentXml(doc, [{ type: 'redline', target: { exactText: marker }, modified: marker }], AUTHOR, {}, {});
            assert.equal(result.status, 'ok', marker);
            const numbering = mergeNumberingParts('', result.numberingXmlParts);
            const numId = parse(result.documentXml).getElementsByTagNameNS(W, 'numId')[0].getAttribute('w:val');
            const numDoc = parse(numbering);
            const num = el(numDoc, 'num').find(n => n.getAttribute('w:numId') === numId);
            assert.equal(el(num, 'startOverride')[0]?.getAttribute('w:val'), start, `${marker} starts at ${start}`);
            const abstractId = el(num, 'abstractNumId')[0].getAttribute('w:val');
            const abstract = el(numDoc, 'abstractNum').find(a => a.getAttribute('w:abstractNumId') === abstractId);
            const fmt = el(abstract, 'numFmt')[0].getAttribute('w:val');
            assert.ok(/Letter|Roman/.test(fmt), `${marker} keeps its letter/roman format (got ${fmt})`);
        }
    } finally {
        console.log = quiet;
    }
}

console.log('PASS: paragraph-format numbering binding');
