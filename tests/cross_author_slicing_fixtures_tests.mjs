import assert from 'assert/strict';
import { existsSync, readFileSync } from 'fs';
import { dirname, join } from 'path';
import { fileURLToPath } from 'url';

import './setup-xml-provider.mjs';
import {
    acceptTrackedChangesInOoxml,
    applyRedlineToOxml,
    ingestWordOoxmlToPlainText,
    rejectTrackedChangesInOoxml,
    validateRedlineOoxml
} from '../index.js';
import { parseOoxmlSafe } from '../adapters/xml-adapter.js';

const testDir = dirname(fileURLToPath(import.meta.url));
const repoRoot = dirname(testDir);
const fixturesDir = join(repoRoot, 'tests', 'fixtures', 'cross-author-slicing');

const SCENARIOS = [
    'insert-interior',
    'delete-interior',
    'delete-boundary-start',
    'delete-boundary-end',
    'delete-straddle-baseline-insertion',
    'multi-author-stacked'
];

function loadFixtureXml(scenario, state) {
    const xmlPath = join(fixturesDir, `${scenario}-${state}.xml`);
    assert(existsSync(xmlPath), `Missing fixture: ${xmlPath}`);
    let xml = readFileSync(xmlPath, 'utf8');
    if (xml.charCodeAt(0) === 0xFEFF) {
        xml = xml.slice(1);
    }
    return xml;
}

function getLocalName(node) {
    return (node.localName || node.nodeName || '').replace(/^.*:/, '');
}

function findChildrenByLocalName(parent, localName) {
    return Array.from(parent.childNodes || []).filter(
        child => child.nodeType === 1 && getLocalName(child) === localName
    );
}

function findDescendantsByLocalName(element, localName) {
    return Array.from(element.getElementsByTagName('*')).filter(
        el => getLocalName(el) === localName
    );
}

// ----------------------------------------------------------------------------
// 1. Fixture Inventory Verification
// ----------------------------------------------------------------------------
console.log('--- 1. Verifying fixture file inventory ---');
for (const scenario of SCENARIOS) {
    for (const state of ['pending', 'accepted', 'rejected']) {
        const docxPath = join(fixturesDir, `${scenario}-${state}.docx`);
        const xmlPath = join(fixturesDir, `${scenario}-${state}.xml`);
        assert(existsSync(docxPath), `DOCX fixture missing: ${docxPath}`);
        assert(existsSync(xmlPath), `XML fixture missing: ${xmlPath}`);
    }
}
console.log('All 18 .docx and 18 .xml fixtures present.');

// ----------------------------------------------------------------------------
// 2. Structural Inspection: Word Desktop Native OOXML Patterns
// ----------------------------------------------------------------------------
console.log('--- 2. Verifying Word Desktop native OOXML structures ---');

// 2.1 insert-interior: Sibling <w:ins> splitting
{
    const xml = loadFixtureXml('insert-interior', 'pending');
    const { doc } = parseOoxmlSafe(xml, 'application/xml');
    const paragraphs = findDescendantsByLocalName(doc, 'p');
    assert.equal(paragraphs.length, 1);
    const p = paragraphs[0];

    const topIns = findChildrenByLocalName(p, 'ins');
    assert.equal(topIns.length, 3, 'Expected 3 sibling <w:ins> elements at <w:p> level');
    assert.equal(topIns[0].getAttribute('w:author'), 'Barry Plasteras');
    assert.equal(topIns[1].getAttribute('w:author'), 'Anson Lai');
    assert.equal(topIns[2].getAttribute('w:author'), 'Barry Plasteras');

    assert.equal(topIns[0].textContent.trim(), 'amended by this');
    assert.equal(topIns[1].textContent.trim(), 'MASTER');
    assert.equal(topIns[2].textContent.trim(), 'Agreement.');
    console.log('  PASS: insert-interior produces 3 sibling <w:ins> nodes');
}

// 2.2 delete-interior: Nested <w:del> inside <w:ins>
{
    const xml = loadFixtureXml('delete-interior', 'pending');
    const { doc } = parseOoxmlSafe(xml, 'application/xml');
    const p = findDescendantsByLocalName(doc, 'p')[0];

    const topIns = findChildrenByLocalName(p, 'ins');
    assert.equal(topIns.length, 1, 'Expected 1 top-level <w:ins> by Barry Plasteras');
    assert.equal(topIns[0].getAttribute('w:author'), 'Barry Plasteras');

    const nestedDels = findDescendantsByLocalName(topIns[0], 'del');
    assert.equal(nestedDels.length, 1, 'Expected <w:del> nested inside Barry\'s <w:ins>');
    assert.equal(nestedDels[0].getAttribute('w:author'), 'Anson Lai');

    const delTexts = findDescendantsByLocalName(nestedDels[0], 'delText');
    assert.equal(delTexts.length, 1);
    assert.equal(delTexts[0].textContent.trim(), 'generate');
    console.log('  PASS: delete-interior nests <w:del> inside <w:ins> with <w:delText>');
}

// 2.3 delete-boundary-start: Leading nested <w:del>
{
    const xml = loadFixtureXml('delete-boundary-start', 'pending');
    const { doc } = parseOoxmlSafe(xml, 'application/xml');
    const p = findDescendantsByLocalName(doc, 'p')[0];

    const topIns = findChildrenByLocalName(p, 'ins');
    assert.equal(topIns.length, 1);
    assert.equal(topIns[0].getAttribute('w:author'), 'Barry Plasteras');

    const firstChild = topIns[0].firstChild;
    assert.equal(getLocalName(firstChild), 'del', 'First child of <w:ins> should be <w:del>');
    assert.equal(firstChild.getAttribute('w:author'), 'Anson Lai');
    assert.equal(firstChild.textContent.trim(), 'Notwithstanding the foregoing,');
    console.log('  PASS: delete-boundary-start places <w:del> at insertion head');
}

// 2.4 delete-boundary-end: Trailing nested <w:del>
{
    const xml = loadFixtureXml('delete-boundary-end', 'pending');
    const { doc } = parseOoxmlSafe(xml, 'application/xml');
    const p = findDescendantsByLocalName(doc, 'p')[0];

    const topIns = findChildrenByLocalName(p, 'ins');
    assert.equal(topIns.length, 1);
    assert.equal(topIns[0].getAttribute('w:author'), 'Barry Plasteras');

    const lastChild = topIns[0].lastChild;
    assert.equal(getLocalName(lastChild), 'del', 'Last child of <w:ins> should be <w:del>');
    assert.equal(lastChild.getAttribute('w:author'), 'Anson Lai');
    assert.equal(lastChild.textContent.trim(), 'and applicable law.');
    console.log('  PASS: delete-boundary-end places <w:del> at insertion tail');
}

// 2.5 delete-straddle-baseline-insertion: Split top-level del and nested del
{
    const xml = loadFixtureXml('delete-straddle-baseline-insertion', 'pending');
    const { doc } = parseOoxmlSafe(xml, 'application/xml');
    const p = findDescendantsByLocalName(doc, 'p')[0];

    const topDels = findChildrenByLocalName(p, 'del');
    assert.equal(topDels.length, 1, 'Expected top-level <w:del> for baseline portion');
    assert.equal(topDels[0].getAttribute('w:author'), 'Anson Lai');
    assert.equal(topDels[0].textContent.trim(), 'start');

    const topIns = findChildrenByLocalName(p, 'ins');
    assert.equal(topIns.length, 1, 'Expected top-level <w:ins> for insertion');
    assert.equal(topIns[0].getAttribute('w:author'), 'Barry Plasteras');

    const nestedDels = findDescendantsByLocalName(topIns[0], 'del');
    assert.equal(nestedDels.length, 1, 'Expected nested <w:del> for insertion portion');
    assert.equal(nestedDels[0].getAttribute('w:author'), 'Anson Lai');
    assert.equal(nestedDels[0].textContent.trim(), 'inserted');
    console.log('  PASS: delete-straddle splits into top-level del and nested del');
}

// 2.6 multi-author-stacked: Multiple nested dels inside single ins
{
    const xml = loadFixtureXml('multi-author-stacked', 'pending');
    const { doc } = parseOoxmlSafe(xml, 'application/xml');
    const p = findDescendantsByLocalName(doc, 'p')[0];

    const topIns = findChildrenByLocalName(p, 'ins');
    assert.equal(topIns.length, 1);
    assert.equal(topIns[0].getAttribute('w:author'), 'Barry Plasteras');

    const nestedDels = findDescendantsByLocalName(topIns[0], 'del');
    assert.equal(nestedDels.length, 2, 'Expected 2 nested <w:del> elements from distinct authors');
    assert.equal(nestedDels[0].getAttribute('w:author'), 'Anson Lai');
    assert.equal(nestedDels[0].textContent.trim(), 'of the proposal');
    assert.equal(nestedDels[1].getAttribute('w:author'), 'Chris Davis');
    assert.equal(nestedDels[1].textContent.trim(), 'initial');
    console.log('  PASS: multi-author-stacked hosts multiple foreign dels inside ins');
}

// ----------------------------------------------------------------------------
// 3. Lifecycle Oracle Tests: Word Desktop vs docx-redline-js accept/reject
// ----------------------------------------------------------------------------
console.log('--- 3. Verifying Lifecycle Oracles (Accept/Reject Parity) ---');

for (const scenario of SCENARIOS) {
    const pendingXml = loadFixtureXml(scenario, 'pending');
    const wordAcceptedXml = loadFixtureXml(scenario, 'accepted');
    const wordRejectedXml = loadFixtureXml(scenario, 'rejected');

    const wordAcceptedText = ingestWordOoxmlToPlainText(wordAcceptedXml).trim();
    const wordRejectedText = ingestWordOoxmlToPlainText(wordRejectedXml).trim();

    // 3.1 Accept All
    const acceptAllResult = acceptTrackedChangesInOoxml(pendingXml, { allAuthors: true });
    assert(acceptAllResult.hasChanges, `${scenario}: acceptAll should report changes`);
    const acceptedText = ingestWordOoxmlToPlainText(acceptAllResult.oxml).trim();
    assert.equal(acceptedText, wordAcceptedText, `${scenario}: acceptAll text parity mismatch`);

    // 3.2 Reject All
    const rejectAllResult = rejectTrackedChangesInOoxml(pendingXml, { allAuthors: true });
    assert(rejectAllResult.hasChanges, `${scenario}: rejectAll should report changes`);
    const rejectedText = ingestWordOoxmlToPlainText(rejectAllResult.oxml).trim();
    assert.equal(rejectedText, wordRejectedText, `${scenario}: rejectAll text parity mismatch`);

    console.log(`  PASS: ${scenario} lifecycle parity verified for AcceptAll and RejectAll`);
}

// ----------------------------------------------------------------------------
// 4. Selective Author Accept/Reject Mechanics on delete-interior
// ----------------------------------------------------------------------------
console.log('--- 4. Verifying Selective Author Accept/Reject Mechanics ---');
{
    const pendingXml = loadFixtureXml('delete-interior', 'pending');

    // Case A: Reject Author A (Barry Plasteras) only
    // Barry's insertion is rejected; Anson's internal deletion is cascaded and removed.
    const rejectBarry = rejectTrackedChangesInOoxml(pendingXml, { author: 'Barry Plasteras' });
    const rejectBarryText = ingestWordOoxmlToPlainText(rejectBarry.oxml).trim();
    assert.equal(rejectBarryText, 'Background.', 'Rejecting Barry should revert to baseline');

    // Case B: Reject Author B (Anson Lai) only
    // Anson's deletion is unwrapped; Barry's insertion retains full text "generate"
    const rejectAnson = rejectTrackedChangesInOoxml(pendingXml, { author: 'Anson Lai' });
    const rejectAnsonParsed = parseOoxmlSafe(rejectAnson.oxml, 'application/xml');
    const delsRemaining = findDescendantsByLocalName(rejectAnsonParsed.doc, 'del');
    assert.equal(delsRemaining.length, 0, 'Anson del should be unwrapped');
    const insRemaining = findDescendantsByLocalName(rejectAnsonParsed.doc, 'ins');
    assert.equal(insRemaining.length, 1, 'Barry ins should remain');
    assert.equal(insRemaining[0].textContent.includes('generate'), true, 'Barry ins should contain generate');

    // Case C: Accept Author A (Barry Plasteras) only
    // Barry's insertion unwraps into baseline text; Anson's deletion remains pending
    const acceptBarry = acceptTrackedChangesInOoxml(pendingXml, { author: 'Barry Plasteras' });
    const acceptBarryParsed = parseOoxmlSafe(acceptBarry.oxml, 'application/xml');
    const topInsAfter = findChildrenByLocalName(findDescendantsByLocalName(acceptBarryParsed.doc, 'p')[0], 'ins');
    assert.equal(topInsAfter.length, 0, 'Barry ins should be unwrapped');
    const delsAfter = findDescendantsByLocalName(acceptBarryParsed.doc, 'del');
    assert.equal(delsAfter.length, 1, 'Anson del should remain pending');
    assert.equal(delsAfter[0].getAttribute('w:author'), 'Anson Lai');

    console.log('  PASS: Selective accept/reject on delete-interior passes all assertions');
}

// ----------------------------------------------------------------------------
// 5. Engine Reproduction: Apply WP03-WP05 to Word Desktop Baselines
// ----------------------------------------------------------------------------
console.log('--- 5. Reproducing Word Desktop fixtures with slice-cross-author ---');

for (const scenario of SCENARIOS.slice(0, 5)) {
    const wordPendingXml = loadFixtureXml(scenario, 'pending');
    const sourceResult = rejectTrackedChangesInOoxml(wordPendingXml, { author: 'Anson Lai' });
    assert.notEqual(sourceResult.status, 'error', `${scenario}: failed to reconstruct pre-Anson source`);
    assert.equal(sourceResult.hasChanges, true, `${scenario}: Anson revision was not removed`);

    const originalText = ingestWordOoxmlToPlainText(sourceResult.oxml).trim();
    const modifiedText = ingestWordOoxmlToPlainText(wordPendingXml).trim();
    const engineResult = await applyRedlineToOxml(sourceResult.oxml, originalText, modifiedText, {
        author: 'Anson Lai',
        existingRevisions: 'slice-cross-author'
    });

    assert.equal(engineResult.status, 'ok', `${scenario}: ${engineResult.error?.message || 'engine failed'}`);
    const validation = validateRedlineOoxml(engineResult.oxml);
    assert.equal(validation.valid, true, `${scenario}: ${JSON.stringify(validation.issues)}`);

    const accepted = acceptTrackedChangesInOoxml(engineResult.oxml, { allAuthors: true });
    const rejected = rejectTrackedChangesInOoxml(engineResult.oxml, { allAuthors: true });
    assert.equal(
        ingestWordOoxmlToPlainText(accepted.oxml).trim(),
        ingestWordOoxmlToPlainText(loadFixtureXml(scenario, 'accepted')).trim(),
        `${scenario}: generated AcceptAll text differs from Word Desktop`
    );
    assert.equal(
        ingestWordOoxmlToPlainText(rejected.oxml).trim(),
        ingestWordOoxmlToPlainText(loadFixtureXml(scenario, 'rejected')).trim(),
        `${scenario}: generated RejectAll text differs from Word Desktop`
    );

    const { doc } = parseOoxmlSafe(engineResult.oxml, 'application/xml');
    const paragraph = findDescendantsByLocalName(doc, 'p')[0];
    if (scenario === 'insert-interior') {
        const topInsertions = findChildrenByLocalName(paragraph, 'ins');
        assert.deepEqual(
            topInsertions.map(node => node.getAttribute('w:author')),
            ['Barry Plasteras', 'Anson Lai', 'Barry Plasteras']
        );
    } else {
        const foreignCarrier = findChildrenByLocalName(paragraph, 'ins')[0];
        assert(foreignCarrier, `${scenario}: expected Barry insertion carrier`);
        const nestedDels = findDescendantsByLocalName(foreignCarrier, 'del');
        assert(nestedDels.some(node => node.getAttribute('w:author') === 'Anson Lai'),
            `${scenario}: expected nested Anson deletion`);
    }
    console.log(`  PASS: ${scenario} reproduced with engine slicing`);
}

console.log('\nAll cross-author slicing fixture tests PASSED!');
