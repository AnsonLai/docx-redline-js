import './setup-xml-provider.mjs';

import assert from 'assert/strict';
import { readFileSync } from 'fs';
import { join } from 'path';

import { validateRedlineOoxml } from '../core/redline-validation.js';
import { generateReredlineStressFixtures } from '../scripts/export-reredline-stress-fixtures.mjs';

const outputDir = join(process.cwd(), 'tmp', 'reredline-stress-test');
const manifest = await generateReredlineStressFixtures(outputDir);

assert.equal(manifest.scenarios.length, 2);
for (const scenario of manifest.scenarios) {
    for (const stage of ['source', 'round1', 'rerelined']) {
        const stageInfo = scenario.stages[stage];
        const xml = readFileSync(join(outputDir, stageInfo.documentXml), 'utf8');
        const validation = validateRedlineOoxml(xml);
        assert.equal(validation.valid, true, `${scenario.name}/${stage}: ${JSON.stringify(validation.issues)}`);
        assert.equal((xml.match(/<w:tbl\b/g) || []).length, 3, `${scenario.name}/${stage} should contain three tables`);
        assert((xml.match(/<w:numPr\b/g) || []).length >= 6, `${scenario.name}/${stage} should contain six list paragraphs`);
    }

    const sourceXml = readFileSync(join(outputDir, scenario.stages.source.documentXml), 'utf8');
    const round1Xml = readFileSync(join(outputDir, scenario.stages.round1.documentXml), 'utf8');
    const rerelinedXml = readFileSync(join(outputDir, scenario.stages.rerelined.documentXml), 'utf8');

    assert(!/<w:(?:ins|del|rPrChange)\b/.test(sourceXml), 'source must be clean');
    assert.match(round1Xml, /w:author="Round One Reviewer"/);
    assert.match(rerelinedXml, /w:author="Round Two Reviewer"/);
    assert.doesNotMatch(rerelinedXml, /w:author="Round One Reviewer"/, 'prior revisions should be accepted before re-redlining');

    const round1Total = Object.values(scenario.stages.round1.revisions).reduce((sum, count) => sum + count, 0);
    const rerelinedTotal = Object.values(scenario.stages.rerelined.revisions).reduce((sum, count) => sum + count, 0);
    assert(round1Total >= 20, `${scenario.name} round1 should be heavily redlined`);
    assert(rerelinedTotal >= 20, `${scenario.name} rerelined should be heavily redlined`);
}

console.log('reredline_stress_fixture_tests.mjs ... PASS');
