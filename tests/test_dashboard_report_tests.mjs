import assert from 'assert/strict';
import { execFileSync } from 'child_process';
import { copyFileSync, mkdirSync, readFileSync, writeFileSync } from 'fs';
import { dirname } from 'path';
import { fileURLToPath } from 'url';

import { buildDashboardData, renderDashboardHtml } from '../scripts/generate-test-dashboard.mjs';

const data = buildDashboardData();
assert.equal(data.cases.length, 64);
assert.equal(data.cases.filter(item => item.lane === 'synthetic').length, 33);
assert.equal(data.cases.filter(item => item.lane === 'superdoc').length, 31);
assert.equal(data.cases.filter(item => item.visualEligible).length, 47);
assert.equal(data.priorities.emptyCellDispositions.length, 7);

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
assert.match(html, /reviewed real legal\/administrative documents/);
assert.match(html, /Reviewed real documents/);
assert.match(html, /synthetic:simple-redline/);
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
