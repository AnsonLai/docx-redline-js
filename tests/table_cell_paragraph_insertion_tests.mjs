import './setup-xml-provider.mjs';

// Multi-line redlines of a table cell insert rows below the target row (the
// synthesized-table heuristic). Regression: a one-column Markdown separator ("| --- |")
// was not recognized, so a one-column table gained a stray "---" row
// (structuredContent: false) or was refused TABLE_SEPARATOR_REQUIRED (true). Also: a
// caller's structuredContent: false keeps cell text literal on the table route.
import assert from 'assert/strict';
import { readFileSync } from 'node:fs';
import {
    acceptTrackedChangesInOoxml, analyzeStructuredContent, inspectDocumentParts, openDocx,
    rejectTrackedChangesInOoxml, validateRedlineOoxml
} from '../index.js';
import { applyOperationsToDocumentXml } from '../services/standalone-operation-runner.js';
import { parseTable } from '../pipeline/content-analysis.js';
import { deriveParagraphThenTableAppend } from '../services/operation-heuristics.js';
import { unzipDocx, zipDocx } from '../document/zip-archive.js';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const p = text => `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
const table = (rows) => `<w:tbl><w:tblPr/><w:tblGrid>${'<w:gridCol/>'.repeat(rows[0].length)}</w:tblGrid>`
    + rows.map(row => `<w:tr>${row.map(cell => `<w:tc>${p(cell)}</w:tc>`).join('')}</w:tr>`).join('') + '</w:tbl>';
const docXml = body => `<w:document ${W}><w:body>${body}<w:sectPr/></w:body></w:document>`;
const strong = q => ({ index: q.index, exactText: q.exactText, paragraphId: q.paragraphId, fingerprint: q.fingerprint, ...(q.inTable ? { inTable: true } : {}) });
const texts = xml => inspectDocumentParts({ documentXml: xml }).paragraphs.map(q => q.exactText);
const rowsOf = xml => (xml.match(/<w:tr\b/g) || []).length;

async function apply(body, buildOperation, options = {}) {
    const documentXml = docXml(body);
    const operation = { type: 'redline', ...buildOperation(inspectDocumentParts({ documentXml }).paragraphs) };
    const result = await applyOperationsToDocumentXml(documentXml, [operation], 'Tester', null, { atomic: true, strictTargets: true, ...options });
    assert.equal(result.status, 'ok', JSON.stringify(result.error ?? result.results));
    const validation = validateRedlineOoxml(result.documentXml);
    assert.equal(validation.valid, true, JSON.stringify(validation.issues));
    const accepted = acceptTrackedChangesInOoxml(result.documentXml, { allAuthors: true }).oxml;
    const rejected = rejectTrackedChangesInOoxml(result.documentXml, { allAuthors: true }).oxml;
    return { accepted, rejected, source: documentXml };
}
const cellTarget = text => ps => strong(ps.find(q => q.exactText === text));

// ---------------------------------------------------------------------------------------------------------
// 1. One-column separators are separators; a bare "---" is not
// ---------------------------------------------------------------------------------------------------------
{
    assert.deepEqual(parseTable('| A |\n| --- |\n| B |'), { headers: ['A'], rows: [['B']], hasHeader: true });
    assert.equal(parseTable('| A |\n| :---: |\n| B |').hasHeader, true, 'alignment colons');
    assert.equal(parseTable('|A|\n|---|\n|B|').hasHeader, true, 'no spaces');
    assert.equal(parseTable('| A | B |\n| --- | --- |\n| 1 | 2 |').hasHeader, true, 'two columns unchanged');
    assert.equal(parseTable('| A |\n| B |').hasHeader, false, 'no separator');
    assert.equal(parseTable('| A |\n|---\n| B |').hasHeader, false, 'a one-column separator needs both pipes');
    assert.equal(parseTable('| A |\n| -- |\n| B |').hasHeader, false, 'three dashes minimum');

    const issues = markdown => analyzeStructuredContent(markdown).issues.map(issue => issue.code);
    assert.ok(!issues('| A |\n| --- |\n| B |').includes('TABLE_SEPARATOR_REQUIRED'), 'one-column table is valid structured content');
    assert.ok(issues('| A |\n| B |').includes('TABLE_SEPARATOR_REQUIRED'), 'a missing separator is still refused');
    assert.ok(!issues('Intro\n\n---\n\nOutro').includes('TABLE_SEPARATOR_REQUIRED'), 'a bare rule is not a table');

    assert.deepEqual(deriveParagraphThenTableAppend('Intro\n| A |\n| --- |\n| B |'), { paragraphMarkdown: 'Intro', tableText: '| A |\n| --- |\n| B |' });
}

// ---------------------------------------------------------------------------------------------------------
// 2. Multi-line cell edits insert rows, in every structuredContent mode
// ---------------------------------------------------------------------------------------------------------
for (const mode of [false, true, undefined]) {
    const label = `structuredContent ${mode}`;
    const sc = mode === undefined ? {} : { structuredContent: mode };

    const one = await apply(table([['Cell text.']]), ps => ({ target: cellTarget('Cell text.')(ps), modified: 'Cell text.\nSecond line', ...sc }));
    assert.deepEqual(texts(one.accepted), ['Cell text.', 'Second line'], `${label}: one column`);
    assert.equal(rowsOf(one.accepted), 2, `${label}: one row added`);
    assert.doesNotMatch(one.accepted, />---</, `${label}: no separator row`);
    assert.deepEqual(texts(one.rejected), texts(one.source), `${label}: Reject All restores the source`);

    const several = await apply(table([['Cell text.']]), ps => ({ target: cellTarget('Cell text.')(ps), modified: 'Cell text.\nTwo\nThree', ...sc }));
    assert.deepEqual(texts(several.accepted), ['Cell text.', 'Two', 'Three'], `${label}: several rows`);
    assert.equal(rowsOf(several.accepted), 3);

    const below = await apply(table([['Name'], ['Cell text.'], ['Last']]), ps => ({ target: cellTarget('Cell text.')(ps), modified: 'Cell text.\nInserted', ...sc }));
    assert.deepEqual(texts(below.accepted), ['Name', 'Cell text.', 'Inserted', 'Last'], `${label}: the row goes below the target row`);
    assert.deepEqual(texts(below.rejected), ['Name', 'Cell text.', 'Last']);

    const wide = await apply(table([['Name', 'Role'], ['Cell text.', 'Other']]), ps => ({ target: cellTarget('Cell text.')(ps), modified: 'Cell text.\nSecond line', ...sc }));
    assert.deepEqual(texts(wide.accepted), ['Name', 'Role', 'Cell text.', 'Other', 'Second line', ''], `${label}: two columns`);
    assert.deepEqual(texts(wide.rejected), ['Name', 'Role', 'Cell text.', 'Other']);
}

// ---------------------------------------------------------------------------------------------------------
// 3. Caller-written one-column tables
// ---------------------------------------------------------------------------------------------------------
{
    // Reconcile an existing one-column table.
    const reconciled = await apply(table([['Head'], ['One']]), ps => ({ target: cellTarget('One')(ps), modified: '| Head |\n| --- |\n| One |\n| Two |' }));
    assert.deepEqual(texts(reconciled.accepted), ['Head', 'One', 'Two']);
    assert.doesNotMatch(reconciled.accepted, />---</);
    assert.deepEqual(texts(reconciled.rejected), ['Head', 'One']);

    // A paragraph becomes a one-column table.
    const created = await apply(p('Replace me.'), ps => ({ target: strong(ps[0]), modified: '| Head |\n| --- |\n| Value |' }));
    assert.equal(rowsOf(created.accepted), 2);
    assert.deepEqual(texts(created.accepted), ['Head', 'Value']);
    assert.doesNotMatch(created.accepted, />---</);
    assert.deepEqual(texts(created.rejected), ['Replace me.'], 'Reject All removes the table');
}

// ---------------------------------------------------------------------------------------------------------
// 4. Plain-text contract on the table route
// ---------------------------------------------------------------------------------------------------------
{
    const markers = 'Use __literal__ _x_ and **y**';
    for (const shape of [[['Cell text.']], [['Cell text.', 'Other']]]) {
        const literal = await apply(table(shape), ps => ({ target: cellTarget('Cell text.')(ps), modified: `Cell text.\n${markers}`, structuredContent: false }));
        assert.ok(texts(literal.accepted).includes(markers), `structuredContent: false keeps cell markers literal (${shape[0].length} columns)`);
        const formatted = await apply(table(shape), ps => ({ target: cellTarget('Cell text.')(ps), modified: `Cell text.\n${markers}`, structuredContent: true }));
        assert.ok(texts(formatted.accepted).includes('Use literal x and y'), `structuredContent: true parses cell markers (${shape[0].length} columns)`);
    }
    const reconciled = await apply(table([['Head'], ['One']]), ps => ({ target: cellTarget('One')(ps), modified: '| Head |\n| --- |\n| One __a__ |', structuredContent: false }));
    assert.deepEqual(texts(reconciled.accepted), ['Head', 'One __a__'], 'a caller-written table keeps markers with structuredContent: false');
}

// ---------------------------------------------------------------------------------------------------------
// 5. openDocx
// ---------------------------------------------------------------------------------------------------------
{
    const entries = new Map(unzipDocx(new Uint8Array(readFileSync(new URL('./fixtures/sample_doc_test.docx', import.meta.url)))));
    entries.set('word/document.xml', new TextEncoder().encode(docXml(table([['Cell text.']]))));
    const doc = openDocx(zipDocx(entries));
    const target = strong(doc.inspect().paragraphs.find(q => q.exactText === 'Cell text.'));
    const result = await doc.applyOperations([{ type: 'redline', target, modified: 'Cell text.\nSecond line' }], { author: 'Tester' });
    assert.equal(result.status, 'ok', JSON.stringify(result.error ?? result.results));
    const written = new TextDecoder().decode(openDocx(result.toUint8Array()).entries.get('word/document.xml'));
    assert.deepEqual(texts(acceptTrackedChangesInOoxml(written, { allAuthors: true }).oxml), ['Cell text.', 'Second line']);
}

console.log('table_cell_paragraph_insertion_tests: all passed');
