import assert from 'assert/strict';

import './setup-xml-provider.mjs';
import {
    acceptTrackedChangesInOoxml,
    applyRedlineToOxml,
    ingestWordOoxmlToPlainText,
    rejectTrackedChangesInOoxml,
    validateRedlineOoxml
} from '../index.js';
import { NS_W } from '../core/types.js';
import { parseOoxmlSafe } from '../adapters/xml-adapter.js';
import { applyOperationsToDocumentXml } from '../services/standalone-operation-runner.js';
import { inspectDocumentParts } from '../services/document-inspection.js';

const NS_W16DU = 'http://schemas.microsoft.com/office/word/2023/wordml/word16du';
const ALICE_DATE = '2026-09-08T09:00:00Z';
const ALICE_DATE_UTC = '2026-09-08T16:00:00Z';

function parse(xml) {
    const parsed = parseOoxmlSafe(xml, 'application/xml');
    assert.equal(parsed.error, null, `XML parse error: ${parsed.error?.message}`);
    return parsed.doc;
}

function elements(node, localName) {
    return Array.from(node.getElementsByTagNameNS(NS_W, localName));
}

function directChildren(node, localName = null) {
    return Array.from(node.childNodes || []).filter(child => {
        return child.nodeType === 1 && (!localName || child.localName === localName);
    });
}

function attr(node, localName) {
    return node.getAttributeNS(NS_W, localName) || node.getAttribute(`w:${localName}`);
}

function getParagraphIlvl(p) {
    const pPr = directChildren(p, 'pPr')[0];
    if (!pPr) return null;
    const numPr = directChildren(pPr, 'numPr')[0];
    if (!numPr) return null;
    const ilvl = directChildren(numPr, 'ilvl')[0];
    return ilvl ? (ilvl.getAttribute('w:val') || ilvl.getAttribute('val')) : null;
}

function getParagraphNumId(p) {
    const pPr = directChildren(p, 'pPr')[0];
    if (!pPr) return null;
    const numPr = directChildren(pPr, 'numPr')[0];
    if (!numPr) return null;
    const numId = directChildren(numPr, 'numId')[0];
    return numId ? (numId.getAttribute('w:val') || numId.getAttribute('val')) : null;
}

function bulletParagraph(content, { ilvl = '0', numId = '1' } = {}) {
    return `<w:p xmlns:w="${NS_W}" xmlns:w16du="${NS_W16DU}">`
        + `<w:pPr><w:numPr><w:ilvl w:val="${ilvl}"/><w:numId w:val="${numId}"/></w:numPr></w:pPr>`
        + `${content}</w:p>`;
}

function insertion(content, id = 1, author = 'Alice') {
    return `<w:ins w:id="${id}" w:author="${author}" w:date="${ALICE_DATE}" `
        + `w16du:dateUtc="${ALICE_DATE_UTC}">${content}</w:ins>`;
}

function assertValidAndUnique(xml, label = '') {
    const validation = validateRedlineOoxml(xml);
    assert.equal(validation.valid, true, `${label} validation failed: ${JSON.stringify(validation.issues)}`);

    const doc = parse(xml);
    const insNodes = elements(doc, 'ins');
    const delNodes = elements(doc, 'del');
    const allNodes = insNodes.concat(delNodes);
    const ids = allNodes.map(node => attr(node, 'id')).filter(Boolean);
    assert.equal(new Set(ids).size, ids.length, `${label}: duplicate revision IDs found: ${ids.join(', ')}`);

    for (const ins of insNodes) {
        assert.equal(elements(ins, 'ins').length, 0, `${label}: illegal <w:ins> nested inside <w:ins>`);
        assert.equal(elements(ins, 'moveFrom').length, 0, `${label}: illegal <w:moveFrom> inside <w:ins>`);
        assert.equal(elements(ins, 'moveTo').length, 0, `${label}: illegal <w:moveTo> inside <w:ins>`);
    }
}

// ============================================================================
// Test 1: Interior Deletion Inside a Bullet Item
// ============================================================================
{
    console.log('--- Test 1: Interior Deletion Inside a Bullet Item ---');
    const source = bulletParagraph(
        insertion('<w:r><w:t>Submit monthly safety inspection report by the 15th.</w:t></w:r>', 1, 'Alice'),
        { ilvl: '0', numId: '10' }
    );

    const result = await applyRedlineToOxml(
        source,
        'Submit monthly safety inspection report by the 15th.',
        'Submit monthly safety report by the 15th.',
        { author: 'Bob', existingRevisions: 'slice-cross-author' }
    );

    assert.equal(result.status, 'ok');
    assert.equal(result.hasChanges, true);
    assertValidAndUnique(result.oxml, 'Test 1: interior deletion in bullet');

    const doc = parse(result.oxml);
    const p = elements(doc, 'p')[0];
    assert.equal(getParagraphIlvl(p), '0', 'Must preserve ilvl="0"');
    assert.equal(getParagraphNumId(p), '10', 'Must preserve numId="10"');

    // Structural checks: 1 top-level ins by Alice, containing 1 nested del by Bob
    const topIns = directChildren(p, 'ins');
    assert.equal(topIns.length, 1);
    assert.equal(attr(topIns[0], 'author'), 'Alice');

    const nestedDel = elements(topIns[0], 'del');
    assert.equal(nestedDel.length, 1);
    assert.equal(attr(nestedDel[0], 'author'), 'Bob');
    assert.equal(nestedDel[0].textContent.trim(), 'inspection');

    // Lifecycle: Accept All
    const accAll = acceptTrackedChangesInOoxml(result.oxml, { allAuthors: true });
    assert.equal(ingestWordOoxmlToPlainText(accAll.oxml).trim(), 'Submit monthly safety report by the 15th.');
    const accP = elements(parse(accAll.oxml), 'p')[0];
    assert.equal(getParagraphIlvl(accP), '0', 'Accepted bullet must keep ilvl="0"');
    assert.equal(getParagraphNumId(accP), '10', 'Accepted bullet must keep numId="10"');

    // Lifecycle: Reject All
    const rejAll = rejectTrackedChangesInOoxml(result.oxml, { allAuthors: true });
    assert.equal(ingestWordOoxmlToPlainText(rejAll.oxml).trim(), '');
    const rejP = elements(parse(rejAll.oxml), 'p')[0];
    assert.equal(getParagraphIlvl(rejP), '0', 'Rejected bullet must keep ilvl="0"');
    assert.equal(getParagraphNumId(rejP), '10', 'Rejected bullet must keep numId="10"');

    // Selective Author: Reject Bob only (restores Alice's full insertion)
    const rejBob = rejectTrackedChangesInOoxml(result.oxml, { author: 'Bob' });
    assert.equal(ingestWordOoxmlToPlainText(rejBob.oxml).trim(), 'Submit monthly safety inspection report by the 15th.');
    const rejBobP = elements(parse(rejBob.oxml), 'p')[0];
    assert.equal(elements(rejBobP, 'del').length, 0, 'Bob del should be removed');
    assert.equal(elements(rejBobP, 'ins').length, 1, 'Alice ins should remain intact');

    // Selective Author: Accept Alice only (unwraps Alice into baseline, Bob del remains pending)
    const accAlice = acceptTrackedChangesInOoxml(result.oxml, { author: 'Alice' });
    const accAliceDoc = parse(accAlice.oxml);
    assert.equal(elements(accAliceDoc, 'ins').length, 0, 'Alice ins should be unwrapped');
    const delsAfter = elements(accAliceDoc, 'del');
    assert.equal(delsAfter.length, 1, 'Bob del should remain pending');
    assert.equal(attr(delsAfter[0], 'author'), 'Bob');

    console.log('  PASS: Interior deletion inside bullet item');
}

// ============================================================================
// Test 2: Interior Insertion Inside a Bullet Item
// ============================================================================
{
    console.log('--- Test 2: Interior Insertion Inside a Bullet Item ---');
    const source = bulletParagraph(
        insertion('<w:r><w:t>Review quarterly financial report.</w:t></w:r>', 1, 'Alice'),
        { ilvl: '1', numId: '5' }
    );

    const result = await applyRedlineToOxml(
        source,
        'Review quarterly financial report.',
        'Review quarterly internal financial report.',
        { author: 'Bob', existingRevisions: 'slice-cross-author' }
    );

    assert.equal(result.status, 'ok');
    assert.equal(result.hasChanges, true);
    assertValidAndUnique(result.oxml, 'Test 2: interior insertion in bullet');

    const doc = parse(result.oxml);
    const p = elements(doc, 'p')[0];
    assert.equal(getParagraphIlvl(p), '1', 'Must preserve ilvl="1"');
    assert.equal(getParagraphNumId(p), '5', 'Must preserve numId="5"');

    // Sibling splitting: Alice, Bob, Alice
    const siblingIns = directChildren(p, 'ins');
    assert.equal(siblingIns.length, 3, 'Expected 3 sibling <w:ins> elements');
    assert.equal(attr(siblingIns[0], 'author'), 'Alice');
    assert.equal(attr(siblingIns[1], 'author'), 'Bob');
    assert.equal(attr(siblingIns[2], 'author'), 'Alice');
    assert.equal(siblingIns[1].textContent.trim(), 'internal');

    // Lifecycle Accept All
    const acc = acceptTrackedChangesInOoxml(result.oxml, { allAuthors: true });
    assert.equal(ingestWordOoxmlToPlainText(acc.oxml).trim(), 'Review quarterly internal financial report.');
    const accP = elements(parse(acc.oxml), 'p')[0];
    assert.equal(getParagraphIlvl(accP), '1');
    assert.equal(getParagraphNumId(accP), '5');

    // Selective Author: Reject Bob only
    const rejBob = rejectTrackedChangesInOoxml(result.oxml, { author: 'Bob' });
    assert.equal(ingestWordOoxmlToPlainText(rejBob.oxml).trim(), 'Review quarterly financial report.');

    // Selective Author: Reject Alice only
    const rejAlice = rejectTrackedChangesInOoxml(result.oxml, { author: 'Alice' });
    assert.equal(ingestWordOoxmlToPlainText(rejAlice.oxml).trim(), 'internal');

    console.log('  PASS: Interior insertion inside bullet item');
}

// ============================================================================
// Test 3: Boundary Deletions in a Bullet Item (Start & End)
// ============================================================================
{
    console.log('--- Test 3: Boundary Deletions in a Bullet Item ---');

    // 3A: Deletion at the beginning of bullet insertion
    const sourceStart = bulletParagraph(
        insertion('<w:r><w:t>Mandatory notice: all workers must wear helmets.</w:t></w:r>', 1, 'Alice'),
        { ilvl: '0', numId: '3' }
    );
    const resStart = await applyRedlineToOxml(
        sourceStart,
        'Mandatory notice: all workers must wear helmets.',
        'all workers must wear helmets.',
        { author: 'Bob', existingRevisions: 'slice-cross-author' }
    );
    assert.equal(resStart.status, 'ok');
    assertValidAndUnique(resStart.oxml, 'Test 3A: boundary start deletion');

    const docStart = parse(resStart.oxml);
    const pStart = elements(docStart, 'p')[0];
    assert.equal(getParagraphIlvl(pStart), '0');
    assert.equal(getParagraphNumId(pStart), '3');
    const topInsStart = directChildren(pStart, 'ins')[0];
    const firstChild = topInsStart.firstChild;
    assert.equal(firstChild.localName, 'del');
    assert.equal(attr(firstChild, 'author'), 'Bob');
    assert.equal(firstChild.textContent.trim(), 'Mandatory notice:');

    // 3B: Deletion at the end of bullet insertion
    const sourceEnd = bulletParagraph(
        insertion('<w:r><w:t>Comply with safety rules and regional ordinances</w:t></w:r>', 1, 'Alice'),
        { ilvl: '0', numId: '3' }
    );
    const resEnd = await applyRedlineToOxml(
        sourceEnd,
        'Comply with safety rules and regional ordinances',
        'Comply with safety rules',
        { author: 'Bob', existingRevisions: 'slice-cross-author' }
    );
    assert.equal(resEnd.status, 'ok');
    assertValidAndUnique(resEnd.oxml, 'Test 3B: boundary end deletion');

    const docEnd = parse(resEnd.oxml);
    const pEnd = elements(docEnd, 'p')[0];
    assert.equal(getParagraphIlvl(pEnd), '0');
    assert.equal(getParagraphNumId(pEnd), '3');
    const topInsEnd = directChildren(pEnd, 'ins')[0];
    const lastChild = topInsEnd.lastChild;
    assert.equal(lastChild.localName, 'del');
    assert.equal(attr(lastChild, 'author'), 'Bob');
    assert.equal(lastChild.textContent.trim(), 'and regional ordinances');

    console.log('  PASS: Boundary deletions inside bullet item');
}

// ============================================================================
// Test 4: Straddling Deletion Across Baseline Bullet Prefix and Insertion
// ============================================================================
{
    console.log('--- Test 4: Straddling Deletion Across Baseline Prefix and Insertion ---');
    const source = bulletParagraph(
        '<w:r><w:t xml:space="preserve">Phase 1: </w:t></w:r>'
        + insertion('<w:r><w:t>Preliminary infrastructure assessment and survey.</w:t></w:r>', 1, 'Alice'),
        { ilvl: '1', numId: '7' }
    );

    const result = await applyRedlineToOxml(
        source,
        'Phase 1: Preliminary infrastructure assessment and survey.',
        'assessment and survey.',
        { author: 'Bob', existingRevisions: 'slice-cross-author' }
    );

    assert.equal(result.status, 'ok');
    assert.equal(result.hasChanges, true);
    assertValidAndUnique(result.oxml, 'Test 4: straddling deletion');

    const doc = parse(result.oxml);
    const p = elements(doc, 'p')[0];
    assert.equal(getParagraphIlvl(p), '1');
    assert.equal(getParagraphNumId(p), '7');

    // Top-level del for baseline "Phase 1: "
    const topDel = directChildren(p, 'del');
    assert.equal(topDel.length, 1);
    assert.equal(attr(topDel[0], 'author'), 'Bob');
    assert.equal(topDel[0].textContent.trim(), 'Phase 1:');

    // Top-level ins by Alice containing nested del by Bob for "Preliminary infrastructure "
    const topIns = directChildren(p, 'ins');
    assert.equal(topIns.length, 1);
    assert.equal(attr(topIns[0], 'author'), 'Alice');
    const nestedDel = elements(topIns[0], 'del');
    assert.equal(nestedDel.length, 1);
    assert.equal(attr(nestedDel[0], 'author'), 'Bob');
    assert.equal(nestedDel[0].textContent.trim(), 'Preliminary infrastructure');

    // Lifecycle parity
    const acc = acceptTrackedChangesInOoxml(result.oxml, { allAuthors: true });
    assert.equal(ingestWordOoxmlToPlainText(acc.oxml).trim(), 'assessment and survey.');
    const rej = rejectTrackedChangesInOoxml(result.oxml, { allAuthors: true });
    assert.equal(ingestWordOoxmlToPlainText(rej.oxml).trim(), 'Phase 1:');

    console.log('  PASS: Straddling deletion across baseline prefix and insertion');
}

// ============================================================================
// Test 5: Multi-Level Outline Slicing (ilvl=0, 1, 2)
// ============================================================================
{
    console.log('--- Test 5: Multi-Level Outline Slicing ---');

    for (const level of ['0', '1', '2']) {
        const source = bulletParagraph(
            insertion(`<w:r><w:t>Level ${level} checklist item pending validation.</w:t></w:r>`, 1, 'Alice'),
            { ilvl: level, numId: '12' }
        );

        const result = await applyRedlineToOxml(
            source,
            `Level ${level} checklist item pending validation.`,
            `Level ${level} checklist item approved.`,
            { author: 'Bob', existingRevisions: 'slice-cross-author' }
        );

        assert.equal(result.status, 'ok');
        assertValidAndUnique(result.oxml, `Test 5: ilvl=${level}`);

        const doc = parse(result.oxml);
        const p = elements(doc, 'p')[0];
        assert.equal(getParagraphIlvl(p), level, `ilvl must remain "${level}"`);
        assert.equal(getParagraphNumId(p), '12', 'numId must remain "12"');

        const acc = acceptTrackedChangesInOoxml(result.oxml, { allAuthors: true });
        assert.equal(ingestWordOoxmlToPlainText(acc.oxml).trim(), `Level ${level} checklist item approved.`);
    }

    console.log('  PASS: Multi-level outline slicing across ilvl=0, 1, 2');
}

// ============================================================================
// Test 6: Multi-Author Stacked Revisions in a Numbered List Item
// ============================================================================
{
    console.log('--- Test 6: Multi-Author Stacked Revisions in a Numbered List Item ---');

    // Author A (Alice) inserts initial numbered bullet
    const source = bulletParagraph(
        insertion('<w:r><w:t>Draft specification of the proposed protocol with initial metrics.</w:t></w:r>', 1, 'Alice'),
        { ilvl: '0', numId: '9' }
    );

    // Round 1: Author B (Bob) deletes "of the proposed protocol "
    const round1 = await applyRedlineToOxml(
        source,
        'Draft specification of the proposed protocol with initial metrics.',
        'Draft specification with initial metrics.',
        { author: 'Bob', existingRevisions: 'slice-cross-author' }
    );
    assert.equal(round1.status, 'ok');
    assertValidAndUnique(round1.oxml, 'Round 1');

    // Round 2: Author C (Charlie) deletes "initial " from the remaining text
    const round2 = await applyRedlineToOxml(
        round1.oxml,
        'Draft specification with initial metrics.',
        'Draft specification with metrics.',
        { author: 'Charlie', existingRevisions: 'slice-cross-author' }
    );
    assert.equal(round2.status, 'ok');
    assertValidAndUnique(round2.oxml, 'Round 2');

    const doc = parse(round2.oxml);
    const p = elements(doc, 'p')[0];
    assert.equal(getParagraphIlvl(p), '0');
    assert.equal(getParagraphNumId(p), '9');

    // Structure: 1 top-level ins by Alice, containing 2 nested dels by Bob and Charlie
    const topIns = directChildren(p, 'ins');
    assert.equal(topIns.length, 1);
    assert.equal(attr(topIns[0], 'author'), 'Alice');

    const nestedDels = elements(topIns[0], 'del');
    assert.equal(nestedDels.length, 2, 'Must contain 2 nested dels from distinct authors');
    assert.equal(attr(nestedDels[0], 'author'), 'Bob');
    assert.equal(nestedDels[0].textContent.trim(), 'of the proposed protocol');
    assert.equal(attr(nestedDels[1], 'author'), 'Charlie');
    assert.equal(nestedDels[1].textContent.trim(), 'initial');

    // Lifecycle: Accept All
    const acc = acceptTrackedChangesInOoxml(round2.oxml, { allAuthors: true });
    assert.equal(ingestWordOoxmlToPlainText(acc.oxml).trim(), 'Draft specification with metrics.');
    const accP = elements(parse(acc.oxml), 'p')[0];
    assert.equal(getParagraphIlvl(accP), '0');
    assert.equal(getParagraphNumId(accP), '9');

    // Lifecycle: Reject All
    const rej = rejectTrackedChangesInOoxml(round2.oxml, { allAuthors: true });
    assert.equal(ingestWordOoxmlToPlainText(rej.oxml).trim(), '');

    console.log('  PASS: Multi-author stacked revisions in a numbered list item');
}

// ============================================================================
// Test 7: Document-Level Multi-Bullet Operations via applyOperationsToDocumentXml
// ============================================================================
{
    console.log('--- Test 7: Document-Level Multi-Bullet Operations ---');

    const docXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${NS_W}" xmlns:w16du="${NS_W16DU}">
  <w:body>
    <w:p><w:r><w:t>Meeting Agenda Items:</w:t></w:r></w:p>
    <w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="4"/></w:numPr></w:pPr><w:r><w:t>Review previous meeting minutes.</w:t></w:r></w:p>
    <w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="4"/></w:numPr></w:pPr><w:ins w:id="1" w:author="Alice" w:date="${ALICE_DATE}" w16du:dateUtc="${ALICE_DATE_UTC}"><w:r><w:t>Discuss proposed budget allocations for Q4.</w:t></w:r></w:ins></w:p>
    <w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="4"/></w:numPr></w:pPr><w:r><w:t>Adjourn meeting.</w:t></w:r></w:p>
    <w:sectPr/>
  </w:body>
</w:document>`;

    const ops = [
        {
            type: 'replace',
            target: 'Discuss proposed budget allocations for Q4.',
            modified: 'Discuss revised budget allocations for Q4.',
            author: 'Bob',
            existingRevisions: 'slice-cross-author'
        }
    ];

    const result = await applyOperationsToDocumentXml(docXml, ops, 'Bob');
    assert.equal(result.hasChanges, true);
    assert.equal(result.results.length, 1);
    assert.equal(result.results[0].status, 'applied');
    assertValidAndUnique(result.documentXml, 'Test 7: document-level multi-bullet');

    const doc = parse(result.documentXml);
    const paragraphs = elements(doc, 'p');
    assert.equal(paragraphs.length, 4, 'Must have 4 paragraphs (intro + 3 bullets)');

    // Bullet 2 (index 2) must preserve numPr
    const bullet2 = paragraphs[2];
    assert.equal(getParagraphIlvl(bullet2), '0');
    assert.equal(getParagraphNumId(bullet2), '4');

    // Slicing inside Bullet 2: replacing "proposed" with "revised"
    const topIns = directChildren(bullet2, 'ins');
    assert.equal(topIns.length, 3, 'Expected 3 sibling <w:ins> elements for word replacement');
    assert.equal(attr(topIns[0], 'author'), 'Alice');
    assert.equal(attr(topIns[1], 'author'), 'Bob');
    assert.equal(attr(topIns[2], 'author'), 'Alice');
    assert.equal(topIns[1].textContent.trim(), 'revised');

    const nestedDel = elements(topIns[0], 'del');
    assert.equal(nestedDel.length, 1);
    assert.equal(attr(nestedDel[0], 'author'), 'Bob');
    assert.equal(nestedDel[0].textContent.trim(), 'proposed');

    // Document inspection reports both authors
    const inspection = inspectDocumentParts({ documentXml: result.documentXml });
    assert.deepEqual(inspection.revisionAuthors, ['Alice', 'Bob']);

    // Accept All
    const acc = acceptTrackedChangesInOoxml(result.documentXml, { allAuthors: true });
    const accText = ingestWordOoxmlToPlainText(acc.oxml);
    assert.ok(accText.includes('Discuss revised budget allocations for Q4.'));
    assert.ok(!accText.includes('proposed'));

    console.log('  PASS: Document-level multi-bullet operations');
}

console.log('\nAll cross-author slicing list/bullet tests PASSED!');
