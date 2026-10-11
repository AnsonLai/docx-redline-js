import './setup-xml-provider.mjs';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';

import { openDocx } from '../index.js';
import { parseOoxml, serializeOoxml } from '../engine/oxml-engine.js';
import { MemoryZip, unzipDocx, zipDocx } from '../document/zip-archive.js';
import { applyOperationsToDocumentXml } from '../services/standalone-operation-runner.js';
import { ensureNumberingArtifactsInZip } from '../services/standalone-docx-plumbing.js';
import {
    createDynamicNumberingIdState,
    overwriteParagraphNumIds
} from '../services/numbering-helpers.js';
import {
    allocateNumberingDefinition,
    allocateNumberingForReplacement,
    canonicalNumberingDefinition,
    mergeNumberingParts,
    mergeNumberingXmlBySchemaOrder,
    validateNumberingIntegrity
} from '../services/numbering-allocation.js';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const decoder = new TextDecoder();
const encoder = new TextEncoder();
const els = (node, name) => Array.from(node.getElementsByTagNameNS('*', name));
const para = (numId, ilvl = 0, text = 'Item') =>
    `<w:p xmlns:w="${W}"><w:pPr><w:numPr><w:ilvl w:val="${ilvl}"/><w:numId w:val="${numId}"/></w:numPr></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`;
const deletedPara = (numId, text = 'Gone') =>
    `<w:p xmlns:w="${W}"><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="${numId}"/></w:numPr><w:rPr><w:del w:id="9" w:author="A" w:date="2026-01-01T00:00:00Z"/></w:rPr></w:pPr><w:del w:id="10" w:author="A" w:date="2026-01-01T00:00:00Z"><w:r><w:delText>${text}</w:delText></w:r></w:del></w:p>`;
const bodyOf = (...paragraphs) => `<w:document xmlns:w="${W}"><w:body>${paragraphs.join('').replace(/ xmlns:w="[^"]+"/g, '')}<w:sectPr/></w:body></w:document>`;
const sourceNumbering = `<w:numbering xmlns:w="${W}">`
    + '<w:abstractNum w:abstractNumId="4"><w:nsid w:val="ABCDEF01"/><w:multiLevelType w:val="hybridMultilevel"/>'
    + '<w:lvl w:ilvl="0" w:tplc="X"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:pStyle w:val="ListBullet"/><w:lvlText w:val="o"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr><w:rPr><w:rFonts w:ascii="Courier New" w:hAnsi="Courier New"/></w:rPr></w:lvl>'
    + '<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%2)"/><w:lvlJc w:val="right"/><w:pPr><w:ind w:left="2160" w:hanging="180"/></w:pPr></w:lvl>'
    + '</w:abstractNum><w:num w:numId="7"><w:abstractNumId w:val="4"/></w:num></w:numbering>';

// ---- 1. Allocation: fresh ids, copied geometry, per-level startOverride -------------------------------------
{
    const state = createDynamicNumberingIdState(sourceNumbering);
    const first = allocateNumberingDefinition({
        numberingIdState: state,
        sourceNumberingXml: sourceNumbering,
        sourceNumId: 7,
        levels: [{ ilvl: 0, numFmt: 'decimal', lvlText: '%1.' }],
        startOverrides: { 0: 5, 2: 3 }
    });
    assert.ok(first.numId > 7 && first.abstractNumId > 4, 'ids come after every id the document owns');
    const doc = parseOoxml(first.numberingXml);
    const lvl0 = els(doc, 'lvl').find(l => l.getAttribute('w:ilvl') === '0');
    assert.equal(els(lvl0, 'numFmt')[0].getAttribute('w:val'), 'decimal', 'redefined level takes the new format');
    assert.equal(els(lvl0, 'lvlText')[0].getAttribute('w:val'), '%1.');
    assert.equal(els(lvl0, 'ind')[0].getAttribute('w:left'), '1440', 'indentation copied from the source level');
    assert.equal(els(lvl0, 'rFonts')[0].getAttribute('w:ascii'), 'Courier New', 'level fonts copied');
    assert.equal(els(lvl0, 'pStyle').length, 0, 'style link is not copied');
    const lvl1 = els(doc, 'lvl').find(l => l.getAttribute('w:ilvl') === '1');
    assert.equal(els(lvl1, 'numFmt')[0].getAttribute('w:val'), 'decimal');
    assert.equal(els(lvl1, 'lvlJc')[0].getAttribute('w:val'), 'right', 'untouched level keeps its geometry');
    assert.equal(els(doc, 'lvl').length, 9, 'nine levels (source levels plus defaults)');
    assert.equal(els(doc, 'nsid').length, 0, 'nsid is not duplicated');
    const overrides = els(els(doc, 'num')[0], 'lvlOverride')
        .map(o => [o.getAttribute('w:ilvl'), els(o, 'startOverride')[0].getAttribute('w:val')]);
    assert.deepEqual(overrides, [['0', '5'], ['2', '3']], 'per-level startOverride on the num');
    assert.ok(!els(doc, 'abstractNum').some(a => els(a, 'start').some(s => s.getAttribute('w:val') === '5')), 'abstract start is never mutated');

    const second = allocateNumberingDefinition({ numberingIdState: state, levels: [] });
    assert.notEqual(second.numId, first.numId);
    assert.notEqual(second.abstractNumId, first.abstractNumId);
    assert.equal(els(els(parseOoxml(second.numberingXml), 'lvl')[0], 'ind')[0].getAttribute('w:left'), '720', 'no source: default geometry');

    const merged = mergeNumberingParts(sourceNumbering, [first.numberingXml, second.numberingXml]);
    assert.deepEqual(validateNumberingIntegrity(bodyOf(para(first.numId, 2), para(second.numId), para(7, 1)), merged), []);
    assert.equal(canonicalNumberingDefinition(merged, 7), canonicalNumberingDefinition(sourceNumbering, 7), 'old definition untouched');
}

// ---- 2. Merge: first writer wins by id, conflicts reported, schema order, idempotent --------------------------
{
    const a = `<w:numbering xmlns:w="${W}"><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num></w:numbering>`;
    const sameNoisy = `<w:numbering xmlns:w="${W}">\n <w:abstractNum w:abstractNumId="1">\n  <w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/></w:lvl></w:abstractNum>\n <w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num></w:numbering>`;
    const different = `<w:numbering xmlns:w="${W}"><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num></w:numbering>`;
    const conflicts = [];
    const noConflict = mergeNumberingXmlBySchemaOrder(a, sameNoisy, { onConflict: c => conflicts.push(c) });
    assert.deepEqual(conflicts, [], 'canonically equal definitions do not conflict');
    assert.equal(noConflict, mergeNumberingXmlBySchemaOrder(a, a));
    const merged = mergeNumberingXmlBySchemaOrder(a, different, { onConflict: c => conflicts.push(c) });
    assert.deepEqual(conflicts, [{ kind: 'abstractNum', id: 1 }], 'differing definition under an existing id is reported, never overwritten');
    assert.ok(merged.includes('w:val="decimal"') && !merged.includes('w:val="bullet"'), 'existing definition kept');
    assert.ok(/w:numId="2"/.test(merged), 'non-colliding num still added');
    assert.ok(merged.indexOf('<w:abstractNum') < merged.indexOf('<w:num '), 'abstractNum before num');
    assert.equal(mergeNumberingParts(merged, [different]), merged, 'repeat merge is idempotent');
    assert.equal(mergeNumberingParts('', [a, '', null]), a);
}

// ---- 3. Validator ------------------------------------------------------------------------------------------------
{
    const codes = (body, numbering) => validateNumberingIntegrity(body, numbering).map(i => i.code).sort();
    assert.deepEqual(codes(bodyOf(para(7), para(7, 1), para(0)), sourceNumbering), []);
    assert.deepEqual(codes(bodyOf(para(9)), sourceNumbering), ['NUMBERING_NUM_UNDEFINED']);
    assert.deepEqual(codes(bodyOf(para(7)), null), ['NUMBERING_NUM_UNDEFINED'], 'numbering part missing');
    assert.deepEqual(codes(bodyOf(para(7, 3)), sourceNumbering), ['NUMBERING_LEVEL_UNDEFINED']);
    assert.deepEqual(codes(bodyOf(para(7)), sourceNumbering.replace('<w:num w:numId="7"><w:abstractNumId w:val="4"/>', '<w:num w:numId="7"><w:abstractNumId w:val="5"/>')), ['NUMBERING_ABSTRACT_UNDEFINED']);
    assert.deepEqual(codes(bodyOf(para(7)), sourceNumbering.replace('</w:numbering>', '<w:num w:numId="7"><w:abstractNumId w:val="4"/></w:num></w:numbering>')), ['NUMBERING_DUPLICATE_ID']);
    // A pPrChange snapshot is a binding Reject All can restore: it must resolve too.
    const snapshot = '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="7"/></w:numPr><w:pPrChange w:id="3" w:author="A" w:date="2026-01-01T00:00:00Z"><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="42"/></w:numPr></w:pPr></w:pPrChange></w:pPr><w:r><w:t>x</w:t></w:r></w:p>';
    assert.deepEqual(codes(bodyOf(snapshot), sourceNumbering), ['NUMBERING_NUM_UNDEFINED']);
}

// ---- 4. R1 / R3: policy ------------------------------------------------------------------------------------------------
{
    const node = xml => parseOoxml(xml).documentElement;
    const valOfFirstNumId = n => els(n, 'numId')[0].getAttribute('w:val');
    const gone = node(deletedPara(7));
    const keep = node(para(7, 0, 'Keep'));
    overwriteParagraphNumIds([gone, keep], 99);
    assert.equal(valOfFirstNumId(gone), '7', 'R1: tracked-deleted paragraph keeps numPr');
    assert.equal(valOfFirstNumId(keep), '99');

    const payload = (id, format) => `<w:numbering xmlns:w="${W}"><w:abstractNum w:abstractNumId="${id}"><w:lvl w:ilvl="0"><w:numFmt w:val="${format}"/></w:lvl></w:abstractNum><w:num w:numId="${id}"><w:abstractNumId w:val="${id}"/></w:num></w:numbering>`;
    const source = [node(para(7, 0, 'Src'))];
    const generated = () => [node(deletedPara(7)), node(para(1, 0, 'New'))];

    // sourceNumbering level 0 is a bullet: a bullet payload is the same kind, a decimal payload a format change.
    let result = allocateNumberingForReplacement({
        numberingXml: payload(1, 'bullet'), replacementNodes: generated(), sourceParagraphs: source,
        numberingIdState: createDynamicNumberingIdState(sourceNumbering)
    });
    assert.equal(result.policy, 'continue-source', 'same-kind replacement continues the source list');
    assert.equal(result.numberingXml, null);
    assert.deepEqual(result.replacementNodes.map(valOfFirstNumId), ['7', '7'], 'deleted keeps 7; new binds to the source list');

    const state = createDynamicNumberingIdState(sourceNumbering);
    result = allocateNumberingForReplacement({
        numberingXml: payload(1, 'decimal'), replacementNodes: generated(), sourceParagraphs: source, numberingIdState: state
    });
    assert.equal(result.policy, 'allocate');
    const ids = result.replacementNodes.map(valOfFirstNumId);
    assert.equal(ids[0], '7', 'R1: tracked-deleted source paragraph is not remapped');
    assert.ok(Number(ids[1]) > 7, 'new paragraph gets a fresh numId');
    const allocatedBody = bodyOf(...result.replacementNodes.map(n => serializeOoxml(n)));
    assert.deepEqual(validateNumberingIntegrity(allocatedBody, mergeNumberingParts(sourceNumbering, [result.numberingXml])), []);

    result = allocateNumberingForReplacement({
        numberingXml: payload(1, 'bullet'), replacementNodes: generated(), sourceParagraphs: source,
        numberingIdState: createDynamicNumberingIdState(sourceNumbering), allowContinueSource: false
    });
    assert.equal(result.policy, 'allocate', 'table scope never continues a source list');
    result = allocateNumberingForReplacement({ numberingXml: null, replacementNodes: generated(), sourceParagraphs: source, numberingIdState: state });
    assert.equal(result.policy, 'passthrough');
}

// ---- 5. R2: one merge for the plumbing and every caller ------------------------------------------------------------------
{
    const read = path => readFileSync(new URL(`../${path}`, import.meta.url), 'utf8');
    for (const file of ['document/docx-document.js', 'services/standalone-docx-plumbing.js', 'services/standalone-operation-runner.js']) {
        assert.ok(!/mergeNumberingXml\s*:/.test(read(file)) && !/insertNumberingNodeInSchemaOrder/.test(read(file)), `${file} keeps no merge of its own`);
    }
    assert.ok(!/mergeNumberingXmlBySchemaOrder|insertNumberingNodeInSchemaOrder|reuseSourceListNumbering/.test(read('services/numbering-helpers.js')), 'helpers hold no merge or reuse heuristic');
    assert.ok(read('document/docx-document.js').includes('ensureNumberingArtifactsInZip(zip, outcome.numberingXmlParts)'));
    assert.ok(read('document/docx-document.js').includes('ensureNumberingArtifactsInZip(zip, result.numberingXmlParts)'), 'facade call sites pass the same arguments');

    // The plumbing merges into an existing part even without a caller-supplied callback (no replacement).
    const entries = new Map([['word/numbering.xml', encoder.encode(sourceNumbering)]]);
    const zip = new MemoryZip(entries);
    const incoming = allocateNumberingDefinition({ numberingIdState: createDynamicNumberingIdState(sourceNumbering), levels: [] });
    await ensureNumberingArtifactsInZip(zip, [incoming.numberingXml]);
    const merged = decoder.decode(entries.get('word/numbering.xml'));
    assert.ok(merged.includes('w:numId="7"') && merged.includes(`w:numId="${incoming.numId}"`), 'existing and incoming definitions both present');
    const clash = `<w:numbering xmlns:w="${W}"><w:abstractNum w:abstractNumId="4"><w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/></w:lvl></w:abstractNum></w:numbering>`;
    await assert.rejects(() => ensureNumberingArtifactsInZip(zip, [clash]), error => error.code === 'NUMBERING_ID_CONFLICT');
}

// ---- 6. Fresh contexts: a call that is not told about the document's numbering reuses ids; the merge reports it ----------
{
    const doc = `<w:document xmlns:w="${W}"><w:body><w:p><w:r><w:t>Plain provision one.</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`;
    const existing = `<w:numbering xmlns:w="${W}"><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:numFmt w:val="upperRoman"/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num></w:numbering>`;
    const run = context => applyOperationsToDocumentXml(doc, [{
        type: 'redline',
        target: { index: 1, exactText: 'Plain provision one.' },
        modified: 'Rewritten:' + String.fromCharCode(10) + '- Alpha first.' + String.fromCharCode(10) + '- Alpha second.',
        structuredContent: true
    }], 'Tester', context, { atomic: true, generateRedlines: true, structuredContent: true });
    const blind = await run({});
    assert.equal(blind.status, 'ok');
    const conflicts = [];
    mergeNumberingParts(existing, blind.numberingXmlParts, { onConflict: c => conflicts.push(c) });
    assert.ok(conflicts.length > 0, 'ids allocated without the document numbering collide; the merge reports it instead of binding silently');
    // Told about the numbering, the same call allocates after every existing id and merges cleanly.
    const informed = await run({ numberingXml: existing });
    const clean = [];
    const mergedInformed = mergeNumberingParts(existing, informed.numberingXmlParts, { onConflict: c => clean.push(c) });
    assert.deepEqual(clean, []);
    assert.deepEqual(validateNumberingIntegrity(informed.documentXml, mergedInformed), []);
}

// ---- 7. Package without a numbering part gets one (part, content type, relationship) ------------------------------------
{
    const xmlHeader = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
    const entries = new Map([
        ['[Content_Types].xml', encoder.encode(`${xmlHeader}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`)],
        ['_rels/.rels', encoder.encode(`${xmlHeader}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`)],
        ['word/_rels/document.xml.rels', encoder.encode(`${xmlHeader}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`)],
        ['word/document.xml', encoder.encode(`${xmlHeader}<w:document xmlns:w="${W}"><w:body><w:p><w:r><w:t>Plain provision.</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`)]
    ]);
    const docx = await openDocx(zipDocx(entries));
    const result = await docx.applyOperations([{
        type: 'redline',
        target: { index: 1, exactText: 'Plain provision.' },
        modified: 'Plain provision:\n- First bullet.\n- Second bullet.',
        structuredContent: true
    }], { author: 'Tester', atomic: true, generateRedlines: true, structuredContent: true });
    assert.equal(result.status, 'ok', JSON.stringify(result.error || result.issues));
    const out = unzipDocx(result.uint8Array);
    assert.ok(out.has('word/numbering.xml'), 'numbering part created');
    assert.match(decoder.decode(out.get('[Content_Types].xml')), /\/word\/numbering\.xml/);
    assert.match(decoder.decode(out.get('word/_rels/document.xml.rels')), /relationships\/numbering/);
    assert.deepEqual(validateNumberingIntegrity(decoder.decode(out.get('word/document.xml')), decoder.decode(out.get('word/numbering.xml'))), []);
}

console.log('PASS: numbering allocation tests');
