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

console.log('literal_underscore_text_tests: all passed');
