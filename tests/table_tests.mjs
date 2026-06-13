import './setup-xml-provider.mjs';

import assert from 'assert/strict';
import fs from 'fs/promises';
import path from 'path';
import { fileURLToPath } from 'url';
import { applyRedlineToOxml } from '../engine/oxml-engine.js';
import { ingestOoxml, ingestTableToVirtualGrid } from '../pipeline/ingestion.js';
import { parseTable } from '../pipeline/pipeline.js';
import { diffTablesWithVirtualGrid, serializeVirtualGridToOoxml } from '../services/table-reconciliation.js';
import {
    elementsByLocalName,
    parseXml,
    parseXmlFragment,
    textContentByLocalName
} from './helpers/ooxml-assertions.mjs';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const DOC_PATH = path.join(__dirname, './sample_doc/word/document.xml');

async function loadSampleDocumentXml() {
    return fs.readFile(DOC_PATH, 'utf-8');
}

async function testReproTableIssue() {
    const originalOoxml = await loadSampleDocumentXml();
    const { acceptedText: originalText } = ingestOoxml(originalOoxml);

    const start = originalText.indexOf('Disclosing Party:');
    const end = originalText.indexOf('RECITALS');
    assert.notEqual(start, -1, 'Expected sample document to contain Disclosing Party section');
    assert.notEqual(end, -1, 'Expected sample document to contain RECITALS section');
    assert(start < end, 'Expected Disclosing Party section to appear before RECITALS');

    const tableMd = `
| Disclosing Party | Receiving Party |
| --- | --- |
| [Name of Disclosing Party] | [Name of Receiving Party] |
| [Address of Disclosing Party] | [Address of Receiving Party] |
`;
    const newText = originalText.slice(0, start) + tableMd + '\n' + originalText.slice(end);

    const result = await applyRedlineToOxml(originalOoxml, originalText, newText, {
        author: 'DebugUser',
        generateRedlines: true
    });
    const parsed = parseXml(result.oxml);

    assert.equal(result.hasChanges, true);
    assert(elementsByLocalName(parsed, 'tbl').length > 0, 'Expected output to preserve or create a table');
    assert(result.oxml.includes('[Name of Disclosing Party]'), 'Expected new table content in output');
    assert(result.oxml.includes('<w:ins') || result.oxml.includes('<ins'), 'Expected tracked insertions in output');
}

async function testTableDiffDiagnostic() {
    const originalOoxml = await loadSampleDocumentXml();
    const xmlDoc = parseXml(originalOoxml);
    const tables = elementsByLocalName(xmlDoc, 'tbl');
    assert(tables.length > 0, 'Expected sample document to contain a table');

    const oldGrid = ingestTableToVirtualGrid(tables[0]);
    const newTableMd = `
| DISCLOSING PARTY: | RECEIVING PARTY: |
| --- | --- |
| _________________________ | _________________________ |
| By: [Name] | By: [Name] |
| Title: | Title: |
| Date: ___________________ | Date: ___________________ |
`;

    const newTableData = parseTable(newTableMd);
    const operations = diffTablesWithVirtualGrid(oldGrid, newTableData);
    const reconciled = serializeVirtualGridToOoxml(oldGrid, operations, {
        generateRedlines: true,
        author: 'TestUser'
    });

    parseXmlFragment(reconciled);
    assert(operations.length > 0, 'Expected table diff to produce at least one operation');
    assert(reconciled.includes('Date:'), 'Expected reconciled table to include inserted Date row');
    assert(reconciled.includes('<w:ins') || reconciled.includes('<ins'), 'Expected reconciled table to include redlines');
}

async function testTableReconciliationFlow() {
    const originalOoxml = await loadSampleDocumentXml();
    const xmlDoc = parseXml(originalOoxml);
    const tables = elementsByLocalName(xmlDoc, 'tbl');
    assert(tables.length > 0, 'Expected sample document to contain a table');

    const tableOoxml = new XMLSerializer().serializeToString(tables[0]);
    const tableText = textContentByLocalName(tables[0], 't').trim();
    const newTableMd = `
| DISCLOSING PARTY: | RECEIVING PARTY: |
| --- | --- |
| _________________________ | _________________________ |
| By: [Name] | By: [Name] |
| Title: | Title: |
| Date: ___________________ | Date: ___________________ |
`;

    const result = await applyRedlineToOxml(tableOoxml, tableText, newTableMd, {
        author: 'TestUser',
        generateRedlines: true
    });

    parseXml(result.oxml);
    assert.equal(result.hasChanges, true);
    assert(result.oxml.includes('Date:'), 'Expected table update output to include Date row');
    assert(result.oxml.includes('<w:ins') || result.oxml.includes('<ins'), 'Expected table update to include redlines');
}

async function testDefaultNamespaceTableReconciliationFlow() {
    const tableOoxml = `
<tbl xmlns="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <tr>
    <tc><p><r><t>Name</t></r></p></tc>
    <tc><p><r><t>Role</t></r></p></tc>
  </tr>
  <tr>
    <tc><p><r><t>Ada</t></r></p></tc>
    <tc><p><r><t>Reviewer</t></r></p></tc>
  </tr>
</tbl>`;
    const newTableMd = `
| Name | Role |
| --- | --- |
| Ada | Reviewer |
| Grace | Approver |
`;

    const result = await applyRedlineToOxml(tableOoxml, 'NameRoleAdaReviewer', newTableMd, {
        author: 'NamespaceTester',
        generateRedlines: true
    });

    parseXml(result.oxml);
    assert.equal(result.hasChanges, true);
    assert(result.oxml.includes('Grace'), 'Expected default-namespace table update to include inserted row content');
}

await testReproTableIssue();
await testTableDiffDiagnostic();
await testTableReconciliationFlow();
await testDefaultNamespaceTableReconciliationFlow();

console.log('table_tests.mjs ... PASS');
