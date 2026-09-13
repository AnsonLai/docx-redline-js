import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { buildZip } from '../scripts/lib/minimal-zip.mjs';
import { executeCli } from '../node/cli.js';
import { openDocx } from '../node/index.js';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const contentTypes = `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`;
const rels = `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`;
const paragraph = (id, text, properties = '') => `<w:p w:paraId="${id}">${properties}<w:r><w:t>${text}</w:t></w:r></w:p>`;
const documentXml = `<w:document xmlns:w="${W}"><w:body>
${paragraph('00000001', '4.2 Prohibited Use', '<w:pPr><w:pStyle w:val="Heading1"/></w:pPr>')}
${paragraph('00000002', 'Intro context.')}
${paragraph('00000003', 'Email notice alpha.')}
${paragraph('00000004', 'Bridge context.')}
${paragraph('00000005', 'EMAIL NOTICE beta.')}
${paragraph('00000006', 'Closing context.')}
${paragraph('00000007', '5. Notices', '<w:pPr><w:pStyle w:val="Heading1"/></w:pPr>')}
<w:p w:paraId="00000008"><w:del w:id="7" w:author="Prior" w:date="2026-01-01T00:00:00Z"><w:r><w:delText>Deleted email clause.</w:delText></w:r></w:del></w:p>
${paragraph('00000009', 'Accepted tail.')}
<w:sectPr/></w:body></w:document>`;
const fixture = buildZip([
    { name: '[Content_Types].xml', data: contentTypes },
    { name: 'word/document.xml', data: documentXml },
    { name: 'word/_rels/document.xml.rels', data: rels }
]);
const directory = await mkdtemp(path.join(tmpdir(), 'docx-cli-context-'));

try {
    const input = path.join(directory, 'context.docx');
    await writeFile(input, fixture);

    const upper = await executeCli(['extract', input, '--search', 'EMAIL NOTICE', '--around', '1']);
    const lower = await executeCli(['extract', input, '--search', 'email notice', '--around', '1']);
    assert.deepEqual(upper.paragraphs, lower.paragraphs);
    assert.deepEqual(lower.paragraphs.map(item => item.index), [2, 3, 4, 5, 6]);
    assert.deepEqual(lower.paragraphs.map(item => item.selectionRole), ['context', 'match', 'context', 'match', 'context']);
    assert.deepEqual(lower.paragraphs.find(item => item.index === 4).contextFor, [3, 5]);
    assert.equal(lower.selection.caseSensitive, false);
    assert.equal(lower.selection.totalMatches, 2);
    assert.equal(lower.selection.returnedMatches, 2);
    assert.equal(lower.selection.returnedParagraphs, 5);
    assert.equal(lower.machineReferencesAreNotUserLocations, true);
    assert.match(lower.paragraphs.find(item => item.index === 3).humanReference, /4\.2 Prohibited Use/);
    assert.equal(lower.paragraphs[0].revisionView, 'accepted');
    assert.deepEqual(Object.keys(lower.paragraphs[1]).slice(0, 3), ['humanReference', 'provision', 'nearestHeading']);

    const limited = await executeCli(['extract', input, '--search', 'email notice', '--around', '1', '--limit', '1']);
    assert.deepEqual(limited.paragraphs.map(item => item.index), [2, 3, 4]);
    assert.equal(limited.selection.totalMatches, 2);
    assert.equal(limited.selection.returnedMatches, 1);
    assert.equal(limited.selection.truncated, true);
    assert.equal(limited.selection.nextAfter, 3);

    const continued = await executeCli(['extract', input, '--search', 'email notice', '--context', '1', '--limit', '1', '--after', '3']);
    assert.deepEqual(continued.paragraphs.map(item => item.index), [4, 5, 6]);
    assert.equal(continued.selection.returnedMatches, 1);
    assert.equal(continued.selection.truncated, false);
    assert.equal(continued.selection.nextAfter, null);

    const shortAlias = await executeCli(['extract', input, '--search', 'email notice', '-C', '1']);
    assert.deepEqual(shortAlias.paragraphs.map(item => item.index), [2, 3, 4, 5, 6]);

    const clamped = await executeCli(['extract', input, '--search', 'alpha', '--around', '5', '--range', '2:4']);
    assert.deepEqual(clamped.paragraphs.map(item => item.index), [2, 3, 4]);

    const acceptedDeleted = await executeCli(['extract', input, '--search', 'deleted email', '--around', '1']);
    assert.equal(acceptedDeleted.selection.totalMatches, 0);
    const rejectedDeleted = await executeCli(['extract', input, '--search', 'deleted email', '--view', 'rejected', '--around', '1']);
    assert.deepEqual(rejectedDeleted.paragraphs.map(item => item.index), [7, 8, 9]);
    assert.equal(rejectedDeleted.paragraphs[1].revisionView, 'rejected');
    assert.match(rejectedDeleted.paragraphs[1].humanReference, /5\. Notices/);

    const directInspection = openDocx(fixture).inspect({ search: 'email notice', around: 1, limit: 1 });
    assert.deepEqual(directInspection.paragraphs.map(item => item.index), [2, 3, 4]);
    assert.equal(directInspection.selection.returnedMatches, 1);

    const invalidAround = await executeCli(['extract', input, '--around', '2']);
    assert.equal(invalidAround.error.code, 'INVALID_FILTER');
    assert.match(invalidAround.error.message, /requires --search/);
} finally {
    await rm(directory, { recursive: true, force: true });
}

console.log('agent CLI inspection context tests passed');
