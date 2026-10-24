import './setup-xml-provider.mjs';

// The last paragraph mark of a container (the body's final mark, a table cell's last mark) must never carry a
// tracked insertion: Word's Reject All cannot remove it and leaves an empty paragraph behind. Regression for
// structured list regeneration of the container's last paragraph ("Last.\n1. Item", a rewrite into a list),
// which ended with an inserted-mark paragraph after the deleted original. The rewrite keeps Word's equivalent
// shape: the original's deleted text joins the last new paragraph, which keeps the untracked mark and records
// the original properties in a tracked w:pPrChange.
import assert from 'assert/strict';
import { readFileSync } from 'node:fs';
import {
    acceptTrackedChangesInOoxml, inspectDocumentParts, openDocx, rejectTrackedChangesInOoxml, validateRedlineOoxml
} from '../index.js';
import { applyOperationsToDocumentXml } from '../services/standalone-operation-runner.js';
import { parseXml } from '../adapters/xml-adapter.js';
import { unzipDocx, zipDocx } from '../document/zip-archive.js';

const NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const W = `xmlns:w="${NS}"`;
const p = text => `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
const cellTable = cellBody => `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol/></w:tblGrid><w:tr><w:tc>${cellBody}</w:tc></w:tr></w:tbl>`;
const docXml = body => `<w:document ${W}><w:body>${body}<w:sectPr/></w:body></w:document>`;
const strong = q => ({ index: q.index, exactText: q.exactText, paragraphId: q.paragraphId, fingerprint: q.fingerprint, ...(q.inTable ? { inTable: true } : {}) });
const texts = xml => inspectDocumentParts({ documentXml: xml }).paragraphs.map(q => q.exactText);
const directChild = (node, name) => Array.from(node?.childNodes || []).find(child => child.nodeType === 1 && child.localName === name) || null;
const markOf = paragraph => {
    const rPr = directChild(directChild(paragraph, 'pPr'), 'rPr');
    return directChild(rPr, 'ins') ? 'ins' : (directChild(rPr, 'del') ? 'del' : null);
};
const lastParagraph = container => Array.from(container.childNodes).filter(node => node.nodeType === 1 && node.localName === 'p').pop();
/** Tracked marks on every container's last paragraph (body and table cells). Empty means Word can reject cleanly. */
function trackedContainerFinals(xml) {
    const doc = parseXml(xml);
    const containers = [['body', doc.getElementsByTagNameNS(NS, 'body')[0]], ...Array.from(doc.getElementsByTagNameNS(NS, 'tc')).map(tc => ['cell', tc])];
    return containers.map(([name, container]) => [name, markOf(lastParagraph(container))]).filter(([, mark]) => mark).map(([name, mark]) => `${name}:${mark}`);
}

async function apply(body, targetIndex, modified, extra = {}) {
    const documentXml = docXml(body);
    const source = inspectDocumentParts({ documentXml }).paragraphs;
    const result = await applyOperationsToDocumentXml(documentXml, [{ type: 'redline', target: strong(source[targetIndex]), modified, ...extra }],
        'Tester', null, { atomic: true, strictTargets: true });
    assert.equal(result.status, 'ok', JSON.stringify({ modified, extra, error: result.results?.find(item => item.error)?.error ?? result.error }).slice(0, 600));
    const issues = validateRedlineOoxml(result.documentXml).issues.filter(issue => issue.severity === 'error');
    assert.deepEqual(issues, [], 'valid redline markup');
    const ids = [...result.documentXml.matchAll(/<w:(?:ins|del|pPrChange|rPrChange)\b[^>]*\bw:id="([^"]+)"/g)].map(m => m[1]);
    assert.equal(new Set(ids).size, ids.length, 'revision ids are unique');
    return {
        xml: result.documentXml,
        accepted: texts(acceptTrackedChangesInOoxml(result.documentXml, { allAuthors: true }).oxml),
        rejected: texts(rejectTrackedChangesInOoxml(result.documentXml, { allAuthors: true }).oxml),
        source: source.map(q => q.exactText),
        receipt: result.results?.[0]?.receipt
    };
}

// ---------------------------------------------------------------------------------------------------------
// 1. Body end: structured list regeneration keeps the final mark untracked
// ---------------------------------------------------------------------------------------------------------
for (const [name, modified, accepted] of [
    ['typed item appended (listMarkers literal)', 'Last.\n1. Item', ['Intro', 'Last.', 'Item']],
    ['changed target plus typed item', 'Last changed.\n1. Item', ['Intro', 'Last changed.', 'Item']],
    ['rewrite into a two-item list', '1. Last.\n2. Next', ['Intro', 'Last.', 'Next']],
    ['two typed items appended', 'Last.\n1. One\n2. Two', ['Intro', 'Last.', 'One', 'Two']],
]) {
    // An explicit structuredContent: true routes these through structured list generation.
    const extra = { structuredContent: true, ...(name.includes('literal') ? { listMarkers: 'literal' } : {}) };
    const view = await apply(p('Intro') + p('Last.'), 1, modified, extra);
    assert.deepEqual(trackedContainerFinals(view.xml), [], `${name}: the body's final mark is untracked`);
    assert.deepEqual(view.rejected, view.source, `${name}: Reject All restores the source`);
    assert.deepEqual(view.accepted.map(text => text.replace(/^\d+\.\s+/, '')), accepted, `${name}: Accept All`);
    const body = parseXml(view.xml).getElementsByTagNameNS(NS, 'body')[0];
    const final = lastParagraph(body);
    assert.ok(directChild(directChild(final, 'pPr'), 'pPrChange'), `${name}: the final paragraph records the original properties`);
    assert.ok(directChild(directChild(final, 'pPr'), 'numPr'), `${name}: the final paragraph is a list item`);
    if (view.receipt) {
        const present = new Set([...view.xml.matchAll(/\bw:id="([^"]+)"/g)].map(m => m[1]));
        for (const item of view.receipt.revisionItems || []) assert.ok(present.has(item.id), `${name}: receipt id ${item.id} is in the output`);
    }
}

// A source whose last paragraph already had a tracked property change: Reject All restores the oldest properties.
{
    const seeded = '<w:p><w:pPr><w:jc w:val="center"/><w:pPrChange w:id="50" w:author="Tester" w:date="2026-01-01T00:00:00Z"><w:pPr><w:jc w:val="right"/></w:pPr></w:pPrChange></w:pPr>'
        + '<w:r><w:t>Last.</w:t></w:r></w:p>';
    const view = await apply(p('Intro') + seeded, 1, '1. Last.\n2. Next', { structuredContent: true });
    assert.deepEqual(trackedContainerFinals(view.xml), []);
    const rejectedXml = rejectTrackedChangesInOoxml(view.xml, { allAuthors: true }).oxml;
    const rejectedFinal = lastParagraph(parseXml(rejectedXml).getElementsByTagNameNS(NS, 'body')[0]);
    assert.equal(directChild(directChild(rejectedFinal, 'pPr'), 'jc')?.getAttribute('w:val'), 'right', 'Reject All restores the oldest snapshot');
    assert.equal(view.xml.match(/<w:pPrChange\b/g).length, 1, 'one property change, no duplicate history');
}

// ---------------------------------------------------------------------------------------------------------
// 2. Mid-document edits are unchanged (the rewrite only applies at a container's end)
// ---------------------------------------------------------------------------------------------------------
{
    const view = await apply(p('Intro') + p('Last.') + p('After'), 1, '1. Last.\n2. Next');
    assert.deepEqual(view.rejected, view.source);
    assert.deepEqual(view.accepted.slice(-1), ['After']);
    const body = parseXml(view.xml).getElementsByTagNameNS(NS, 'body')[0];
    const paragraphs = Array.from(body.childNodes).filter(node => node.localName === 'p');
    assert.equal(markOf(paragraphs[paragraphs.length - 1]), null, 'the untouched final paragraph stays untracked');
}

// ---------------------------------------------------------------------------------------------------------
// 3. A table cell's last paragraph
// ---------------------------------------------------------------------------------------------------------
{
    const view = await apply(cellTable(p('Cell')) + p('Tail'), 0, '1. Cell\n2. More');
    assert.deepEqual(trackedContainerFinals(view.xml), [], 'the cell\'s last mark is untracked');
    assert.deepEqual(view.rejected, view.source);
}

// ---------------------------------------------------------------------------------------------------------
// 4. Every body-end append shape, in both structuredContent modes
// ---------------------------------------------------------------------------------------------------------
for (const sc of [true, false]) {
    for (const modified of ['Last.\nAppended.', 'Last.\nOne.\nTwo.', 'Last changed.\nAppended.', 'Last.\n- Bullet', '1. Last.\n2. Next']) {
        const view = await apply(p('Intro') + p('Last.'), 1, modified, { structuredContent: sc, listMarkers: 'literal' });
        assert.deepEqual(trackedContainerFinals(view.xml), [], `structuredContent ${sc}, ${JSON.stringify(modified)}: final mark untracked`);
        assert.deepEqual(view.rejected, view.source, `structuredContent ${sc}, ${JSON.stringify(modified)}: Reject All`);
    }
}

// ---------------------------------------------------------------------------------------------------------
// 5. openDocx
// ---------------------------------------------------------------------------------------------------------
{
    const entries = new Map(unzipDocx(new Uint8Array(readFileSync(new URL('./fixtures/sample_doc_test.docx', import.meta.url)))));
    entries.set('word/document.xml', new TextEncoder().encode(docXml(p('Intro') + p('Last.'))));
    const doc = openDocx(zipDocx(entries));
    const target = strong(doc.inspect().paragraphs[1]);
    const result = await doc.applyOperations([{ type: 'redline', target, modified: '1. Last.\n2. Next' }], { author: 'Tester' });
    assert.equal(result.status, 'ok', JSON.stringify(result.error ?? result.results));
    const written = new TextDecoder().decode(openDocx(result.toUint8Array()).entries.get('word/document.xml'));
    assert.deepEqual(trackedContainerFinals(written), []);
    assert.deepEqual(texts(rejectTrackedChangesInOoxml(written, { allAuthors: true }).oxml), ['Intro', 'Last.']);
}

console.log('container_final_mark_tests: all passed');
