import assert from 'assert/strict';
import { execFileSync } from 'child_process';
import { copyFileSync, mkdirSync, readFileSync, writeFileSync } from 'fs';
import { dirname } from 'path';
import { fileURLToPath } from 'url';

import { buildDashboardData, renderDashboardHtml } from '../scripts/generate-test-dashboard.mjs';

const data = buildDashboardData();
assert.equal(data.cases.length, 109);
assert.equal(data.cases.filter(item => item.lane === 'synthetic').length, 49);
assert.equal(data.cases.filter(item => item.lane === 'superdoc').length, 60);
assert.equal(data.cases.filter(item => item.visualEligible).length, 92);
assert.equal(data.priorities.emptyCellDispositions.length, 1);
assert.ok(Array.isArray(data.lane1Cases));

const embeddedFixtureUrl = new URL('../tmp/dashboard-report-test/simple-redline.docx', import.meta.url);
mkdirSync(dirname(fileURLToPath(embeddedFixtureUrl)), { recursive: true });
for (const suffix of ['', '.source', '.accepted', '.rejected']) {
    copyFileSync(
        new URL('./fixtures/sample_doc_test.docx', import.meta.url),
        new URL(`../tmp/dashboard-report-test/simple-redline${suffix}.docx`, import.meta.url)
    );
}
writeFileSync(new URL('../tmp/dashboard-report-test/simple-redline.expected.json', import.meta.url), JSON.stringify({
    sourceText: 'The old sentence.',
    expectedAcceptedText: 'The new sentence.',
    expectedRejectedText: 'The old sentence.'
}));
writeFileSync(new URL('../tmp/dashboard-report-test/simple-redline.document.xml', import.meta.url),
    '<w:document><w:ins/><w:del/><w:rPrChange/></w:document>');
const embeddedData = buildDashboardData(dirname(fileURLToPath(embeddedFixtureUrl)));
assert.equal(embeddedData.cases.filter(item => item.docxVariants?.tracked).length, 1);
const embeddedCase = embeddedData.cases.find(item => item.identity === 'synthetic:simple-redline');
assert.deepEqual(Object.keys(embeddedCase.docxVariants), ['source', 'tracked', 'accepted', 'rejected']);
assert.ok(Object.values(embeddedCase.docxVariants).every(value => value.length > 100));
assert.deepEqual(embeddedCase.revisions, { insertions: 1, deletions: 1, formatting: 1 });
assert.equal(embeddedCase.expectations.accepted, 'The new sentence.');

const lane1FixturesUrl = new URL('../tmp/dashboard-lane1-report-test/', import.meta.url);
mkdirSync(fileURLToPath(lane1FixturesUrl), { recursive: true });
const lane1Name = 'interagency-agreement-multi-author';
const lane1Bytes = {
    source: 'current-source-copy',
    tracked: 'current-structured-tracked-copy',
    accepted: 'current-accepted-copy',
    rejected: 'current-rejected-copy'
};
for (const [state, contents] of Object.entries(lane1Bytes)) {
    const suffix = state === 'tracked' ? '' : `.${state}`;
    writeFileSync(new URL(`${lane1Name}${suffix}.docx`, lane1FixturesUrl), contents);
}
writeFileSync(new URL('manifest.json', lane1FixturesUrl), JSON.stringify({
    generatedAt: '2026-09-05T07:54:54.621Z',
    cases: [{ identity: 'lane1:interagency', name: lane1Name, title: 'Interagency structured copy' }]
}));
const lane1EmbeddedData = buildDashboardData(
    dirname(fileURLToPath(embeddedFixtureUrl)),
    null,
    fileURLToPath(lane1FixturesUrl)
);
assert.equal(lane1EmbeddedData.lane1Cases.length, 1);
assert.equal(lane1EmbeddedData.lane1Cases[0].fixtureGeneratedAt, '2026-09-05T07:54:54.621Z');
for (const [state, contents] of Object.entries(lane1Bytes)) {
    assert.equal(
        Buffer.from(lane1EmbeddedData.lane1Cases[0].docxVariants[state], 'base64').toString('utf8'),
        contents,
        `Lane 1 ${state} must embed the exact selected fixture bytes`
    );
}

const realCase = data.cases.find(item => item.lane === 'superdoc');
const realSourceId = realCase.identity.slice('superdoc:'.length);
const corpusFixturesUrl = new URL('../tmp/dashboard-real-report-test/', import.meta.url);
mkdirSync(fileURLToPath(corpusFixturesUrl), { recursive: true });
const realName = '01-real-dashboard-case';
for (const suffix of ['', '.source', '.accepted', '.rejected']) {
    copyFileSync(
        new URL('./fixtures/sample_doc_test.docx', import.meta.url),
        new URL(`${realName}${suffix}.docx`, corpusFixturesUrl)
    );
}
writeFileSync(new URL('suite.json', corpusFixturesUrl), JSON.stringify({
    cases: [{ name: realName, sourceId: realSourceId }]
}));
writeFileSync(new URL(`${realName}.expected.json`, corpusFixturesUrl), JSON.stringify({
    originalTarget: 'Original real-document target',
    modifiedTarget: 'Modified real-document target'
}));
writeFileSync(new URL(`${realName}.document.xml`, corpusFixturesUrl),
    '<w:document><w:ins/><w:del/></w:document>');
const realEmbeddedData = buildDashboardData(
    dirname(fileURLToPath(embeddedFixtureUrl)),
    fileURLToPath(corpusFixturesUrl)
);
const embeddedRealCase = realEmbeddedData.cases.find(item => item.identity === realCase.identity);
assert.ok(Object.values(embeddedRealCase.docxVariants).every(value => value.length > 100));
assert.equal(embeddedRealCase.expectations.source, 'Original real-document target');
assert.equal(embeddedRealCase.expectations.accepted, 'Modified real-document target');
assert.equal(embeddedRealCase.visualEligible, true);
assert.match(embeddedRealCase.displayName, /Original real-document target/);

const html = renderDashboardHtml(data);
assert.match(html, /<!doctype html>/i);
assert.match(html, /Task × structure matrix/);
assert.match(html, /Oracle comparison/);
assert.match(html, /Planned high-priority gaps/);
assert.match(html, /DOCX comparison/);
assert.match(html, /renderDocxComparison/);
assert.match(html, /Source ↔ tracked/);
assert.match(html, /Accepted ↔ rejected/);
assert.match(html, /downloadView/);
assert.match(html, /sync-scroll/);
assert.match(html, /id="sidebar-toggle"/);
assert.match(html, /aria-controls="dashboard-sidebar"/);
assert.match(html, /sidebar-hidden \.shell\{max-width:none\}/);
assert.match(html, /docx-dashboard-sidebar-hidden/);
assert.match(html, /reviewed real legal\/administrative documents/);
assert.match(html, /Reviewed real documents/);
assert.match(html, /synthetic:simple-redline/);
assert.match(html, /Lane 1 Visual Inspection/);
assert.match(html, /id="lane1-workbench"/);
assert.match(html, /id="lane1-case"/);
assert.match(html, /renderLane1Comparison/);
assert.match(html, /item\.name\+'-STRUCTURED-FIX\.docx'/);
assert.doesNotMatch(html, /<script[^>]+src=/i);
assert.doesNotMatch(html, /fetch\s*\(/);
assert.doesNotMatch(html, /[\u0000-\u0008\u000B\u000C\u000E-\u001F]/);
assert.match(html, /replace\(\/\\b\\w\/g/);

const output = new URL('../tmp/test-comparison-dashboard.html', import.meta.url);
execFileSync(process.execPath, [
    'scripts/generate-test-dashboard.mjs',
    '--output',
    decodeURIComponent(output.pathname).replace(/^\/([A-Za-z]:)/, '$1')
], { cwd: new URL('../', import.meta.url), encoding: 'utf8' });
const written = readFileSync(output, 'utf8');
assert.match(written, /DOCX Redline Test Comparison Dashboard/);
assert.match(written, /"cases":\[/);

console.log('PASS: self-contained HTML test comparison dashboard');
