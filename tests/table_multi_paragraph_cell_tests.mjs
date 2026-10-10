import assert from 'node:assert/strict';

import './setup-xml-provider.mjs';
import { parseOoxmlSafe } from '../adapters/xml-adapter.js';
import { projectWordRevisionView } from '../core/word-revision-semantics.js';
import { applyOperationsToDocumentXml } from '../services/standalone-operation-runner.js';

// Regression: a Markdown table reconciliation that edits one cell must not collapse a sibling cell's
// multi-paragraph content (a cell `LS` + empty paragraph lost its empty paragraph, untracked).
// Reject All of the output must equal Reject All of the source, with default (enforce) fidelity options.

const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const rPr = '<w:pPr><w:rPr><w:rFonts w:cstheme="minorHAnsi"/></w:rPr></w:pPr>';
const para = text => `<w:p>${rPr}${text ? `<w:r><w:t>${text}</w:t></w:r>` : ''}</w:p>`;
const cell = (...paras) => `<w:tc><w:tcPr><w:tcW w:w="0" w:type="auto"/></w:tcPr>${paras.map(para).join('')}</w:tc>`;
const row = cells => `<w:tr>${cells.join('')}</w:tr>`;
const table = rows => `<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/></w:tblPr><w:tblGrid><w:gridCol/><w:gridCol/><w:gridCol/></w:tblGrid>${rows.join('')}</w:tbl>`;
const SOURCE = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="${NS_W}"><w:body>`
    + table([
        row([cell('Item'), cell('Owner'), cell('Status')]),
        row([cell('Plan'), cell('LS', ''), cell('Complete')]),
        row([cell('Survey'), cell('AJ', 'BK', ''), cell('Complete')])
    ])
    + `<w:p/><w:sectPr/></w:body></w:document>`;

const TABLE_MD = rows => `| Item | Owner | Status |\n| --- | --- | --- |\n${rows.join('\n')}`;
const rejectedTexts = xml => projectWordRevisionView(parseOoxmlSafe(xml).doc, 'rejected').map(r => r.text);
const count = (xml, tag) => (xml.match(new RegExp(`<w:${tag}[ >/]`, 'g')) || []).length;

async function edit(rows) {
    const result = await applyOperationsToDocumentXml(SOURCE, [{
        type: 'redline',
        target: 'Item',
        modified: TABLE_MD(rows),
        structuredContent: true
    }], 'Fidelity Test', {}, { atomic: true, generateRedlines: true });
    assert.ok(result.documentXml && result.hasChanges, JSON.stringify(result.results || result.error || result));
    return result.documentXml;
}

const sourceRejected = rejectedTexts(SOURCE);
assert.equal(sourceRejected.length, 13);

// Edit one cell: sibling multi-paragraph cells keep every paragraph, untracked.
const edited = await edit(['| Plan | LS | In progress |', '| Survey | AJ\nBK | Complete |']);
assert.deepEqual(rejectedTexts(edited), sourceRejected, 'Reject All restores every source paragraph');
assert.equal(count(edited, 'p'), count(SOURCE, 'p'), 'no paragraph is dropped or added');
assert.ok(/<w:delText[^>]*>Complete<\/w:delText>/.test(edited), 'edited cell tracks the deleted text');

// Edit inside a multi-paragraph cell: other paragraphs of that cell stay as-is.
const inner = await edit(['| Plan | LS | Complete |', '| Survey | AJ\nBK edited | Complete |']);
assert.deepEqual(rejectedTexts(inner), sourceRejected, 'Reject All restores multi-paragraph cell');
assert.equal(count(inner, 'p'), count(SOURCE, 'p'));

console.log('PASS: table multi-paragraph cell reconciliation keeps untouched paragraphs');
