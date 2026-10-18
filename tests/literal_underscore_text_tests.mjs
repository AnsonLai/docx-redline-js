import './setup-xml-provider.mjs';

// Literal underscore runs (signature and date lines) and snake_case text are not
// Markdown emphasis. A caller's structuredContent: false keeps every inline marker
// literal. Regression for the 0.9.0 underscore-collapse report.
import assert from 'assert/strict';
import { acceptTrackedChangesInOoxml, inspectDocumentParts } from '../index.js';
import { applyOperationsToDocumentXml } from '../services/standalone-operation-runner.js';
import { preprocessMarkdown } from '../pipeline/markdown-processor.js';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const line = '_'.repeat(25);

// Delimiter rules.
for (const [text, clean, hints] of [
    [`Signed: ${line} (print name)`, `Signed: ${line} (print name)`, 0],
    ['file_name_here', 'file_name_here', 0],
    ['______Date______', '______Date______', 0],
    ['a __bold__ b', 'a bold b', 1],
    ['an _italic_ word', 'an italic word', 1],
    ['**x** and __y__ _z_', 'x and y z', 3],
]) {
    const result = preprocessMarkdown(text);
    assert.equal(result.cleanText, clean, `clean text for ${JSON.stringify(text)}`);
    assert.equal(result.formatHints.length, hints, `format hints for ${JSON.stringify(text)}`);
}

const cell = text => `<w:tc><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;
const row = cells => `<w:tr>${cells.map(cell).join('')}</w:tr>`;
const documentXml = `<w:document ${W}><w:body>`
    + `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol/><w:gridCol/></w:tblGrid>${row(['DISCLOSING PARTY:', 'RECEIVING PARTY:'])}${row([line, line])}</w:tbl>`
    + `<w:p><w:r><w:t>Signed: ${line}</w:t></w:r></w:p>`
    + `<w:p><w:r><w:t>Plain</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`;
const source = inspectDocumentParts({ documentXml }).paragraphs;
const strong = p => ({ index: p.index, exactText: p.exactText, paragraphId: p.paragraphId, fingerprint: p.fingerprint, inTable: p.inTable });

async function accepted(operation) {
    const result = await applyOperationsToDocumentXml(documentXml, [operation], 'Tester', null,
        { atomic: true, strictTargets: true, generateRedlines: true });
    assert.equal(result.status, 'ok', JSON.stringify(result.error ?? result.results));
    const xml = acceptTrackedChangesInOoxml(result.documentXml, { allAuthors: true }).oxml;
    return inspectDocumentParts({ documentXml: xml }).paragraphs.map(p => p.exactText);
}

// No-change table reconcile keeps the signature lines.
{
    const texts = await accepted({
        type: 'table-reconciliation', target: strong(source[0]), structuredContent: true,
        modified: `| DISCLOSING PARTY: | RECEIVING PARTY: |\n| --- | --- |\n| ${line} | ${line} |`
    });
    assert.deepEqual(texts.slice(2, 4), [line, line], 'table reconcile keeps underscore runs');
}

// structuredContent: false rewrite keeps the run and every literal marker.
{
    const target = source.find(p => p.exactText.startsWith('Signed:'));
    const texts = await accepted({
        type: 'redline', target: strong(target), structuredContent: false,
        modified: `${target.exactText} (print name)`
    });
    assert.equal(texts.find(t => t.startsWith('Signed:')), `Signed: ${line} (print name)`);
}
{
    const target = source.find(p => p.exactText === 'Plain');
    const texts = await accepted({
        type: 'redline', target: strong(target), structuredContent: false,
        modified: 'Use **literal** __markers__ and _here_'
    });
    assert.equal(texts.at(-1), 'Use **literal** __markers__ and _here_', 'plain text keeps inline markers literal');
}

// More delimiter rules (CommonMark flanking for underscores; other markers unchanged).
for (const [text, clean, hints] of [
    ['__bold__.', 'bold.', 1],
    ['(_aside_)', '(aside)', 1],
    ['_ spaced _', '_ spaced _', 0],
    ['__ spaced __', '__ spaced __', 0],
    ['a_b_c', 'a_b_c', 0],
    ['x__y__z', 'x__y__z', 0],
    ['MAX_VALUE and min_value', 'MAX_VALUE and min_value', 0],
    ['Date: ________ Time: ______', 'Date: ________ Time: ______', 0],
    ['___', '___', 0],
    ['**strong** ~~gone~~ ++under++ *it*', 'strong gone under it', 4],
    ['_one_ and _two_', 'one and two', 2],
]) {
    const result = preprocessMarkdown(text);
    assert.equal(result.cleanText, clean, `clean text for ${JSON.stringify(text)}`);
    assert.equal(result.formatHints.length, hints, `format hints for ${JSON.stringify(text)}`);
}

// End to end through the operation applier.
const plainXml = `<w:document ${W}><w:body><w:p><w:r><w:t>Plain</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`;
const plainTarget = (() => {
    const p = inspectDocumentParts({ documentXml: plainXml }).paragraphs[0];
    return { index: p.index, exactText: p.exactText, paragraphId: p.paragraphId, fingerprint: p.fingerprint };
})();
async function applyPlain(operation, options = {}) {
    const result = await applyOperationsToDocumentXml(plainXml, [{ target: plainTarget, ...operation }], 'Tester', null,
        { atomic: true, strictTargets: true, ...options });
    assert.equal(result.status, 'ok', JSON.stringify(result.error ?? result.results));
    const acceptedXml = acceptTrackedChangesInOoxml(result.documentXml, { allAuthors: true }).oxml;
    return { xml: acceptedXml, text: inspectDocumentParts({ documentXml: acceptedXml }).paragraphs[0].exactText };
}
const markers = 'Use __bold__ and _italic_ here';
{
    // structuredContent: true (the default) still formats emphasis.
    const formatted = await applyPlain({ type: 'redline', modified: markers });
    assert.equal(formatted.text, 'Use bold and italic here');
    assert.match(formatted.xml, /<w:b\/>|<w:b w:val="(?:1|true)"\/>/, 'bold run');
    assert.match(formatted.xml, /<w:i\/>|<w:i w:val="(?:1|true)"\/>/, 'italic run');
}
{
    // The batch option applies when the operation does not say.
    const literal = await applyPlain({ type: 'redline', modified: markers }, { structuredContent: false });
    assert.equal(literal.text, markers, 'batch structuredContent: false keeps markers');
    assert.doesNotMatch(literal.xml, /<w:b\/>|<w:i\/>/);
    // The operation overrides the batch in both directions.
    assert.equal((await applyPlain({ type: 'redline', modified: markers, structuredContent: true }, { structuredContent: false })).text, 'Use bold and italic here');
    assert.equal((await applyPlain({ type: 'redline', modified: markers, structuredContent: false }, { structuredContent: true })).text, markers);
}
{
    // replace, untracked edits and block-level Markdown are plain text too.
    assert.equal((await applyPlain({ type: 'replace', modified: markers, structuredContent: false })).text, markers);
    assert.equal((await applyPlain({ type: 'redline', modified: markers, structuredContent: false }, { generateRedlines: false })).text, markers);
    const heading = await applyPlain({ type: 'redline', modified: '# Not a heading', structuredContent: false });
    assert.equal(heading.text, '# Not a heading');
    assert.doesNotMatch(heading.xml, /w:pStyle w:val="Heading/);
}
{
    // openDocx facade.
    const { openDocx } = await import('../index.js');
    const { unzipDocx, zipDocx } = await import('../document/zip-archive.js');
    const { readFileSync } = await import('node:fs');
    const entries = new Map(unzipDocx(new Uint8Array(readFileSync(new URL('./fixtures/sample_doc_test.docx', import.meta.url)))));
    entries.set('word/document.xml', new TextEncoder().encode(`<w:document ${W}><w:body><w:p><w:r><w:t>Signed: ${line}</w:t></w:r></w:p><w:p><w:r><w:t>Plain</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`));
    const doc = openDocx(zipDocx(entries));
    const [signed, plain] = doc.inspect().paragraphs;
    const handle = p => ({ index: p.index, exactText: p.exactText, paragraphId: p.paragraphId, fingerprint: p.fingerprint });
    const result = await doc.applyOperations([
        { type: 'redline', target: handle(signed), modified: `Signed: ${line} (name)` },
        { type: 'redline', target: handle(plain), modified: markers, structuredContent: false }
    ], { author: 'Tester' });
    assert.equal(result.status, 'ok', JSON.stringify(result.error ?? result.results));
    const reopened = openDocx(result.toUint8Array());
    const acceptedXml = acceptTrackedChangesInOoxml(new TextDecoder().decode(reopened.entries.get('word/document.xml')), { allAuthors: true }).oxml;
    assert.deepEqual(inspectDocumentParts({ documentXml: acceptedXml }).paragraphs.map(p => p.exactText), [`Signed: ${line} (name)`, markers]);
}

console.log('literal_underscore_text_tests: all passed');
