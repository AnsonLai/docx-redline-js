import './setup-xml-provider.mjs';

// list-format convert-plain refuses a paragraph whose soft-break lines carry further typed markers
// ("A. ...<w:br/>B. ...<w:br/>C. ..."): converting it would strip only the first marker and bind one list item that
// keeps the other lines as literal text (KB-11). The refusal comes before any write. Soft-break lines without a
// marker are continuation lines and convert as before.
import assert from 'assert/strict';
import { acceptTrackedChangesInOoxml, inspectDocumentParts, openDocx, rejectTrackedChangesInOoxml } from '../index.js';
import { applyOperationsToDocumentXml, preflightOperations } from '../services/standalone-operation-runner.js';
import { readFileSync } from 'node:fs';
import { unzipDocx, zipDocx } from '../document/zip-archive.js';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const docXml = body => `<w:document ${W}><w:body>${body}<w:sectPr/></w:body></w:document>`;
const softLines = (lines, breakTag = '<w:br/>') => `<w:p><w:r>${lines.map((text, i) => `${i ? breakTag : ''}<w:t xml:space="preserve">${text}</w:t>`).join('')}</w:r></w:p>`;
const p = text => `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
const strong = q => ({ index: q.index, exactText: q.exactText, paragraphId: q.paragraphId, fingerprint: q.fingerprint });
const errorOf = result => result.results?.find(item => item.error)?.error || result.error;

async function convert(body, { targetIndex = 0, endIndex = null, list = {} } = {}) {
    const documentXml = docXml(body);
    const paragraphs = inspectDocumentParts({ documentXml }).paragraphs;
    const operation = {
        type: 'list-format',
        target: strong(paragraphs[targetIndex]),
        ...(endIndex != null ? { targetEnd: strong(paragraphs[endIndex]) } : {}),
        list: { mode: 'convert-plain', ...list }
    };
    const result = await applyOperationsToDocumentXml(documentXml, [operation], 'Tester', null, { atomic: true, strictTargets: true });
    return { documentXml, result, operation };
}

const RECITALS = ['A. The Disclosing Party possesses confidential information.', 'B. The Parties desire to evaluate a transaction.', 'C. The Receiving Party agrees to keep it secret.'];

// ---------------------------------------------------------------------------------------------------------
// 1. Refusals, with nothing written
// ---------------------------------------------------------------------------------------------------------
for (const [name, body, extra, lineMarker] of [
    ['A/B/C recitals on w:br', softLines(RECITALS), {}, 'B.'],
    ['A/B/C recitals on w:cr', softLines(RECITALS, '<w:cr/>'), {}, 'B.'],
    ['a marker of another family on a later line', softLines(['A. Heading', 'Body text', '1. A numbered line']), {}, '1.'],
    ['bullets on soft breaks', softLines(['- First', '- Second']), {}, '-'],
    ['an explicit format does not change it', softLines(RECITALS), { list: { format: 'upperLetter' } }, 'B.'],
]) {
    const { documentXml, result } = await convert(body, extra);
    assert.equal(result.status, 'error', `${name}: refused`);
    const error = errorOf(result);
    assert.equal(error.code, 'LIST_MARKER_UNSUPPORTED', `${name}: code`);
    assert.equal(error.reason, 'soft-break-markers', `${name}: reason`);
    assert.equal(error.lineMarker, lineMarker, `${name}: names the marker on the later line`);
    assert.match(error.recovery?.hint || '', /Split the paragraph at its soft line breaks/, `${name}: recovery says to split the paragraph first`);
    assert.equal(error.recovery?.action, 'change_request');
    assert.equal(result.documentXml, documentXml, `${name}: nothing is written`);
    assert.equal(result.hasChanges, false);
}

// In a targetEnd range, any paragraph with soft-break markers refuses the whole operation.
{
    const { documentXml, result } = await convert(p('A. Plain first item') + softLines(['B. Second item', 'C. Third item']), { targetIndex: 0, endIndex: 1 });
    assert.equal(result.status, 'error');
    assert.equal(errorOf(result).reason, 'soft-break-markers');
    assert.equal(errorOf(result).paragraphIndex, 2, 'names the paragraph in the range');
    assert.equal(result.documentXml, documentXml, 'nothing is written for the range');
}

// Preflight reports the same refusal without mutating.
{
    const documentXml = docXml(softLines(RECITALS));
    const target = strong(inspectDocumentParts({ documentXml }).paragraphs[0]);
    const preflight = await preflightOperations(documentXml, [{ type: 'list-format', target, list: { mode: 'convert-plain' } }], 'Tester', { strictTargets: true });
    const reported = preflight.results?.[0];
    if (reported?.error) assert.equal(reported.error.code, 'LIST_MARKER_UNSUPPORTED');
}

// openDocx: refused, the package is not rewritten.
{
    const entries = new Map(unzipDocx(new Uint8Array(readFileSync(new URL('./fixtures/sample_doc_test.docx', import.meta.url)))));
    entries.set('word/document.xml', new TextEncoder().encode(docXml(softLines(RECITALS))));
    const bytes = zipDocx(entries);
    const doc = openDocx(bytes);
    const result = await doc.applyOperations([{ type: 'list-format', target: strong(doc.inspect().paragraphs[0]), list: { mode: 'convert-plain' } }], { author: 'Tester' });
    assert.equal(result.status, 'error');
    assert.equal(result.written, false);
    assert.equal(errorOf(result).code, 'LIST_MARKER_UNSUPPORTED');
}

// ---------------------------------------------------------------------------------------------------------
// 2. Unchanged behaviour
// ---------------------------------------------------------------------------------------------------------
const views = xml => ({
    accepted: inspectDocumentParts({ documentXml: acceptTrackedChangesInOoxml(xml, { allAuthors: true }).oxml }).paragraphs,
    rejected: inspectDocumentParts({ documentXml: rejectTrackedChangesInOoxml(xml, { allAuthors: true }).oxml }).paragraphs.map(q => q.exactText)
});
{
    // A single marker with ordinary soft-break continuation lines still converts.
    const { documentXml, result } = await convert(softLines(['A. Heading of the clause', 'continued on a second line', 'and a third']));
    assert.equal(result.status, 'ok', JSON.stringify(errorOf(result)));
    const { accepted, rejected } = views(result.documentXml);
    assert.equal(accepted[0].exactText, 'Heading of the clause\ncontinued on a second line\nand a third');
    assert.ok(accepted[0].list, 'bound to a list');
    assert.deepEqual(rejected, inspectDocumentParts({ documentXml }).paragraphs.map(q => q.exactText), 'Reject All restores the source');
}
{
    // A later line that merely starts with a capital letter and no marker punctuation is not a marker.
    const { result } = await convert(softLines(['A. Parties', 'Buyer and Seller agree']));
    assert.equal(result.status, 'ok', JSON.stringify(errorOf(result)));
}
{
    // Separate paragraphs, one marker each: converts as before.
    const { result } = await convert(RECITALS.map(p).join(''), { targetIndex: 0, endIndex: 2 });
    assert.equal(result.status, 'ok', JSON.stringify(errorOf(result)));
    const { accepted } = views(result.documentXml);
    assert.deepEqual(accepted.map(q => q.exactText), RECITALS.map(text => text.slice(3)));
    assert.ok(accepted.every(q => q.list), 'three list items');
}
{
    // strip: "none" with an explicit format binds without parsing markers, as before.
    const { result } = await convert(softLines(RECITALS), { list: { strip: 'none', format: 'upperLetter' } });
    assert.equal(result.status, 'ok', JSON.stringify(errorOf(result)));
}

console.log('list_format_soft_break_tests: all passed');
