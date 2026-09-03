import './setup-xml-provider.mjs';

import assert from 'assert/strict';
import { applyOperationToDocumentXml, applyOperationsToDocumentXml } from '../services/standalone-operation-runner.js';
import {
    acceptTrackedChangesInOoxml,
    rejectTrackedChangesInOoxml
} from '../index.js';
import {
    directChildByLocalName,
    elementsByLocalName,
    parseXml
} from './helpers/ooxml-assertions.mjs';

const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

function getParagraphIlvl(p) {
    const pPr = directChildByLocalName(p, 'pPr');
    if (!pPr) return null;
    const numPr = directChildByLocalName(pPr, 'numPr');
    if (!numPr) return null;
    const ilvl = directChildByLocalName(numPr, 'ilvl');
    return ilvl ? (ilvl.getAttribute('w:val') || ilvl.getAttribute('val')) : null;
}

function getParagraphNumId(p) {
    const pPr = directChildByLocalName(p, 'pPr');
    if (!pPr) return null;
    const numPr = directChildByLocalName(pPr, 'numPr');
    if (!numPr) return null;
    const numId = directChildByLocalName(numPr, 'numId');
    return numId ? (numId.getAttribute('w:val') || numId.getAttribute('val')) : null;
}

function getParagraphText(p) {
    return elementsByLocalName(p, 't').map(t => t.textContent).join('');
}

function createThreeLevelListDocXml(numId = '11') {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${NS_W}">
  <w:body>
    <w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="${numId}"/></w:numPr></w:pPr><w:r><w:t>Governance Framework.</w:t></w:r></w:p>
    <w:p><w:pPr><w:numPr><w:ilvl w:val="1"/><w:numId w:val="${numId}"/></w:numPr></w:pPr><w:r><w:t>Board Responsibilities.</w:t></w:r></w:p>
    <w:p><w:pPr><w:numPr><w:ilvl w:val="2"/><w:numId w:val="${numId}"/></w:numPr></w:pPr><w:r><w:t>Approve annual operating budget.</w:t></w:r></w:p>
    <w:p><w:pPr><w:numPr><w:ilvl w:val="2"/><w:numId w:val="${numId}"/></w:numPr></w:pPr><w:r><w:t>Appoint interim audit committee.</w:t></w:r></w:p>
    <w:p><w:pPr><w:numPr><w:ilvl w:val="1"/><w:numId w:val="${numId}"/></w:numPr></w:pPr><w:r><w:t>Executive Authority.</w:t></w:r></w:p>
    <w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="${numId}"/></w:numPr></w:pPr><w:r><w:t>Operational Controls.</w:t></w:r></w:p>
    <w:sectPr/>
  </w:body>
</w:document>`;
}

async function testNestedChildInsertionAtMultipleLevels() {
    // Scenario 1: Inserting a Level 2 child under Level 1 "Board Responsibilities."
    // Target: Level 1 item (ilvl=1).
    // Modified: Indented bullet that should be placed at child depth (ilvl=2).
    const inputXml = createThreeLevelListDocXml('21');
    const op = {
        type: 'list-change',
        target: 'Board Responsibilities.',
        modified: '1. Board Responsibilities.\n  - Review enterprise risk register.'
    };

    const result = await applyOperationToDocumentXml(inputXml, op, 'AuthorA');
    assert.equal(result.hasChanges, true, 'List insertion must report changes');

    // Verify tracked changes in raw result
    const doc = parseXml(result.documentXml);
    const insElements = elementsByLocalName(doc, 'ins');
    assert.ok(insElements.length > 0, 'Must contain tracked insertion for the new child item');

    // Test ACCEPT ALL:
    const accepted = acceptTrackedChangesInOoxml(result.documentXml, { author: 'AuthorA' });
    const acceptedDoc = parseXml(accepted.oxml);
    const acceptedParagraphs = elementsByLocalName(acceptedDoc, 'p');

    assert.equal(acceptedParagraphs.length, 7, 'Must have 7 paragraphs after inserting 1 child');
    const levels = acceptedParagraphs.map(p => getParagraphIlvl(p));
    // Expected: [0, 1, 2 (inserted), 2, 2, 1, 0]
    assert.equal(levels[0], '0', 'Level 0 parent intact');
    assert.equal(levels[1], '1', 'Level 1 anchor intact');
    assert.equal(levels[2], '2', 'Inserted child must be promoted to child depth (ilvl=2)');
    assert.equal(levels[3], '2', 'Existing Level 2 child intact');
    assert.equal(levels[4], '2', 'Existing Level 2 child intact');
    assert.equal(levels[5], '1', 'Subsequent Level 1 intact');
    assert.equal(levels[6], '0', 'Subsequent Level 0 intact');

    assert.ok(getParagraphText(acceptedParagraphs[2]).includes('Review enterprise risk register.'));

    // Test REJECT ALL:
    const rejected = rejectTrackedChangesInOoxml(result.documentXml, { author: 'AuthorA' });
    const rejectedDoc = parseXml(rejected.oxml);
    const rejectedParagraphs = elementsByLocalName(rejectedDoc, 'p');
    assert.equal(rejectedParagraphs.length, 6, 'Rejected state must return to 6 paragraphs');
    const rejectedLevels = rejectedParagraphs.map(p => getParagraphIlvl(p));
    assert.deepEqual(rejectedLevels, ['0', '1', '2', '2', '1', '0'], 'Original levels restored');
}

async function testConcurrentEditsAcrossAllThreeLevels() {
    // Scenario 2: Multiple concurrent changes across levels 0, 1, and 2 in a single document:
    // - Edit Level 0: "Governance Framework." -> "Corporate Governance Framework."
    // - Edit Level 1: "Executive Authority." -> "Executive Authority and Delegations."
    // - Edit Level 2: "Approve annual operating budget." -> "Approve annual operating and capital budget."
    // - Edit Level 2: "Appoint interim audit committee." -> "Appoint permanent standing audit committee."
    const inputXml = createThreeLevelListDocXml('22');

    const operations = [
        {
            type: 'redline',
            target: 'Governance Framework.',
            modified: 'Corporate Governance Framework.'
        },
        {
            type: 'redline',
            target: 'Executive Authority.',
            modified: 'Executive Authority and Delegations.'
        },
        {
            type: 'redline',
            target: 'Approve annual operating budget.',
            modified: 'Approve annual operating and capital budget.'
        },
        {
            type: 'redline',
            target: 'Appoint interim audit committee.',
            modified: 'Appoint permanent standing audit committee.'
        }
    ];

    const result = await applyOperationsToDocumentXml(inputXml, operations, 'AuthorB');
    assert.equal(result.hasChanges, true, 'Batch multi-level changes must apply');

    const accepted = acceptTrackedChangesInOoxml(result.documentXml, { author: 'AuthorB' });
    const acceptedDoc = parseXml(accepted.oxml);
    const paragraphs = elementsByLocalName(acceptedDoc, 'p');

    assert.equal(paragraphs.length, 6, 'Paragraph count must remain 6');
    const levels = paragraphs.map(p => getParagraphIlvl(p));
    assert.deepEqual(levels, ['0', '1', '2', '2', '1', '0'], 'All 3 levels must be preserved without level drift');

    assert.equal(getParagraphText(paragraphs[0]), 'Corporate Governance Framework.');
    assert.equal(getParagraphText(paragraphs[1]), 'Board Responsibilities.');
    assert.equal(getParagraphText(paragraphs[2]), 'Approve annual operating and capital budget.');
    assert.equal(getParagraphText(paragraphs[3]), 'Appoint permanent standing audit committee.');
    assert.equal(getParagraphText(paragraphs[4]), 'Executive Authority and Delegations.');
    assert.equal(getParagraphText(paragraphs[5]), 'Operational Controls.');

    // Reject all must restore original texts exactly
    const rejected = rejectTrackedChangesInOoxml(result.documentXml, { author: 'AuthorB' });
    const rejectedDoc = parseXml(rejected.oxml);
    const rejParagraphs = elementsByLocalName(rejectedDoc, 'p');
    assert.equal(getParagraphText(rejParagraphs[0]), 'Governance Framework.');
    assert.equal(getParagraphText(rejParagraphs[2]), 'Approve annual operating budget.');
    assert.equal(getParagraphText(rejParagraphs[3]), 'Appoint interim audit committee.');
    assert.equal(getParagraphText(rejParagraphs[4]), 'Executive Authority.');
}

async function testLevelFlatteningGuard() {
    // Scenario 3: Guard against the visual defect of "level flattening"
    // An edit inside a Level 2 bullet must NEVER drop w:ilvl or reset it to 0.
    const inputXml = createThreeLevelListDocXml('23');
    const op = {
        type: 'redline',
        target: 'Appoint interim audit committee.',
        modified: 'Appoint interim audit and compliance committee.'
    };

    const result = await applyOperationToDocumentXml(inputXml, op, 'AuthorC');
    assert.equal(result.hasChanges, true);

    const doc = parseXml(result.documentXml);
    const paragraphs = elementsByLocalName(doc, 'p');
    const targetP = paragraphs[3]; // The fourth paragraph (Appoint interim...)

    const ilvl = getParagraphIlvl(targetP);
    assert.equal(ilvl, '2', 'Level 2 item must retain ilvl="2" after redline (no level flattening to 0)');
}

async function testFormattingPreservationAcrossLevels() {
    // Scenario 4: Formatting changes on a nested item must not leak or destroy levels
    const inputXml = createThreeLevelListDocXml('24');
    const op = {
        type: 'redline',
        target: 'Approve annual operating budget.',
        modified: '**Approve annual operating budget.**'
    };

    const result = await applyOperationToDocumentXml(inputXml, op, 'AuthorD');
    assert.equal(result.hasChanges, true);

    const doc = parseXml(result.documentXml);
    const paragraphs = elementsByLocalName(doc, 'p');

    // Verify Level 2 item retained ilvl="2"
    assert.equal(getParagraphIlvl(paragraphs[2]), '2', 'Formatted item must retain ilvl="2"');

    // Verify Level 1 parent has NO bold run property
    const parentRPrs = elementsByLocalName(paragraphs[1], 'rPr');
    for (const rPr of parentRPrs) {
        assert.equal(directChildByLocalName(rPr, 'b'), null, 'Parent must not inherit child bold');
    }

    // Verify Level 2 sibling has NO bold run property
    const siblingRPrs = elementsByLocalName(paragraphs[3], 'rPr');
    for (const rPr of siblingRPrs) {
        assert.equal(directChildByLocalName(rPr, 'b'), null, 'Sibling must not inherit bold');
    }
}

async function testMultiLevelRangeDeletionAndAcceptance() {
    // Scenario 5: Multi-item range replacement in multi-level list:
    // Replacing two adjacent Level 2 items with a single consolidated Level 2 item
    const inputXml = createThreeLevelListDocXml('25');
    const op = {
        type: 'redline',
        target: 'Approve annual operating budget.\nAppoint interim audit committee.',
        modified: 'Approve annual operating budget and designate audit committee.'
    };

    const result = await applyOperationToDocumentXml(inputXml, op, 'AuthorE');
    assert.equal(result.hasChanges, true);

    const accepted = acceptTrackedChangesInOoxml(result.documentXml, { author: 'AuthorE' });
    const acceptedDoc = parseXml(accepted.oxml);
    const paragraphs = elementsByLocalName(acceptedDoc, 'p');

    // Was 6 paragraphs, consolidated 2 into 1 -> 5 paragraphs
    assert.equal(paragraphs.length, 5, 'Consolidated range must yield 5 paragraphs');
    const levels = paragraphs.map(p => getParagraphIlvl(p));
    assert.deepEqual(levels, ['0', '1', '2', '1', '0'], 'Consolidated item must stay at Level 2 without breaking subsequent levels');
    assert.ok(getParagraphText(paragraphs[2]).includes('designate audit committee'));

    // Reject restores both Level 2 items
    const rejected = rejectTrackedChangesInOoxml(result.documentXml, { author: 'AuthorE' });
    const rejectedDoc = parseXml(rejected.oxml);
    const rejParagraphs = elementsByLocalName(rejectedDoc, 'p');
    assert.equal(rejParagraphs.length, 6);
    assert.deepEqual(rejParagraphs.map(p => getParagraphIlvl(p)), ['0', '1', '2', '2', '1', '0']);
}

async function testAgendaMultiLevelSubBulletsDozenItems() {
    // Scenario 6: Test 47 Agenda document multi-level sub-bullets with > 12 changes
    // Synthetic creation of 6 Level 1 sub-bullets and 12 Level 2 sub-sub-bullets (18 items total)
    // under Item 3 in the Town of Prattsville Board Agenda.
    const { readFileSync } = await import('fs');
    const { unzipEntries } = await import('../scripts/lib/zip-reader.mjs');
    const zip = unzipEntries(readFileSync('tmp/superdoc-corpus/8b0f4f46292113aeb389445b035d62621f28a8075fe26b46d339a5658801d087.docx'));
    const docXml = zip.get('word/document.xml').toString('utf8');

    const modifiedText = [
        '3. Knotweed',
        '  a. Survey and containment',
        '    i. Spring boundary mapping',
        '    ii. Growth rate assessment',
        '  b. Eradication treatment protocols',
        '    i. Targeted herbicide application',
        '    ii. Manual root excavation',
        '  c. Grant funding and budget',
        '    i. Laura Jane Musser grant',
        '    ii. Municipal matching appropriation',
        '  d. Community notification',
        '    i. Landowner outreach letters',
        '    ii. Public informational workshop',
        '  e. Disposal and monitoring',
        '    i. Controlled transport protocol',
        '    ii. Post-treatment site monitoring',
        '  f. Inter-agency coordination',
        '    i. County soil and water district',
        '    ii. State conservation review'
    ].join('\n');

    const op = {
        type: 'list-change',
        target: 'Knotweed',
        modified: modifiedText
    };

    const result = await applyOperationToDocumentXml(docXml, op, 'AgendaAuthor');
    assert.equal(result.hasChanges, true);

    // Verify all 18 items are present in accepted document
    const accepted = acceptTrackedChangesInOoxml(result.documentXml, { allAuthors: true });
    const acceptedDoc = parseXml(accepted.oxml);
    const paragraphs = elementsByLocalName(acceptedDoc, 'p');

    const level1Titles = [
        'Survey and containment',
        'Eradication treatment protocols',
        'Grant funding and budget',
        'Community notification',
        'Disposal and monitoring',
        'Inter-agency coordination'
    ];
    const level2Titles = [
        'Spring boundary mapping',
        'Growth rate assessment',
        'Targeted herbicide application',
        'Manual root excavation',
        'Laura Jane Musser grant',
        'Municipal matching appropriation',
        'Landowner outreach letters',
        'Public informational workshop',
        'Controlled transport protocol',
        'Post-treatment site monitoring',
        'County soil and water district',
        'State conservation review'
    ];

    // Assert every Level 1 item has ilvl="1" and numId="9"
    for (const title of level1Titles) {
        const p = paragraphs.find(node => getParagraphText(node).includes(title));
        assert.ok(p, `Missing Level 1 item: ${title}`);
        assert.equal(getParagraphIlvl(p), '1', `Item ${title} must be at ilvl="1"`);
        assert.equal(getParagraphNumId(p), '9', `Item ${title} must retain numId="9"`);
    }

    // Assert every Level 2 item has ilvl="2" and numId="9"
    for (const title of level2Titles) {
        const p = paragraphs.find(node => getParagraphText(node).includes(title));
        assert.ok(p, `Missing Level 2 item: ${title}`);
        assert.equal(getParagraphIlvl(p), '2', `Item ${title} must be at ilvl="2"`);
        assert.equal(getParagraphNumId(p), '9', `Item ${title} must retain numId="9"`);
    }

    // Assert parent Knotweed retained ilvl="0"
    const knotweedP = paragraphs.find(node => getParagraphText(node) === 'Knotweed');
    assert.ok(knotweedP, 'Parent Knotweed paragraph must exist');
    assert.equal(getParagraphIlvl(knotweedP), '0', 'Parent item must retain ilvl="0"');

    // Assert subsequent items like UTV law retain ilvl="0"
    const utvP = paragraphs.find(node => getParagraphText(node).includes('UTV law'));
    assert.ok(utvP, 'Subsequent UTV law paragraph must exist');
    assert.equal(getParagraphIlvl(utvP), '0', 'Subsequent list item must retain ilvl="0"');

    // Assert rejectTrackedChanges cleanly restores original document
    const rejected = rejectTrackedChangesInOoxml(result.documentXml, { allAuthors: true });
    const rejectedDoc = parseXml(rejected.oxml);
    const rejParagraphs = elementsByLocalName(rejectedDoc, 'p');
    for (const title of [...level1Titles, ...level2Titles]) {
        assert.ok(!rejParagraphs.some(node => getParagraphText(node).includes(title)), `Rejected document must not contain: ${title}`);
    }
}

async function testAgendaFebMultiLevelSubBulletsDozenItems() {
    // Scenario 7: Test 37 Agenda 2.10.25 multi-level sub-bullets with > 12 changes
    // Synthetic creation of 6 Level 1 sub-bullets and 12 Level 2 sub-sub-bullets (18 items total)
    // under Item 3 in New Business ("Greene County EMS Contract").
    const { readFileSync } = await import('fs');
    const { unzipEntries } = await import('../scripts/lib/zip-reader.mjs');
    const zip = unzipEntries(readFileSync('tmp/superdoc-corpus/77c33ab2ca57efd0d4afb87f08c45c88f3984256c3fd5fdf3e81f87a24644760.docx'));
    const docXml = zip.get('word/document.xml').toString('utf8');

    const modifiedText = [
        '3. Greene County EMS Contract',
        '  a. Service coverage and dispatch standards',
        '    i. Priority 1 response time benchmark',
        '    ii. Advanced life support staffing minimums',
        '  b. Municipal subsidy and quarterly cost allocation',
        '    i. Base annual contribution schedule',
        '    ii. Fuel surcharge adjustment formula',
        '  c. Equipment maintenance and vehicle readiness',
        '    i. Defibrillator telemetry calibration',
        '    ii. Winter tire and chain certification',
        '  d. Quality assurance and performance reporting',
        '    i. Monthly clinical run reviews',
        '    ii. Patient transport outcome metrics',
        '  e. Mutual aid and emergency surge protocol',
        '    i. Backup ambulance dispatch routing',
        '    ii. Mass casualty incident coordination',
        '  f. Inter-municipal legal compliance',
        '    i. County attorney indemnification clause',
        '    ii. Board resolution adoption timeline'
    ].join('\n');

    const op = {
        type: 'list-change',
        target: 'Greene County EMS Contract',
        modified: modifiedText
    };

    const result = await applyOperationToDocumentXml(docXml, op, 'AgendaFebAuthor');
    assert.equal(result.hasChanges, true);

    const accepted = acceptTrackedChangesInOoxml(result.documentXml, { allAuthors: true });
    const acceptedDoc = parseXml(accepted.oxml);
    const paragraphs = elementsByLocalName(acceptedDoc, 'p');

    const level1Titles = [
        'Service coverage and dispatch standards',
        'Municipal subsidy and quarterly cost allocation',
        'Equipment maintenance and vehicle readiness',
        'Quality assurance and performance reporting',
        'Mutual aid and emergency surge protocol',
        'Inter-municipal legal compliance'
    ];
    const level2Titles = [
        'Priority 1 response time benchmark',
        'Advanced life support staffing minimums',
        'Base annual contribution schedule',
        'Fuel surcharge adjustment formula',
        'Defibrillator telemetry calibration',
        'Winter tire and chain certification',
        'Monthly clinical run reviews',
        'Patient transport outcome metrics',
        'Backup ambulance dispatch routing',
        'Mass casualty incident coordination',
        'County attorney indemnification clause',
        'Board resolution adoption timeline'
    ];

    for (const title of level1Titles) {
        const p = paragraphs.find(node => getParagraphText(node).includes(title));
        assert.ok(p, `Missing Level 1 item: ${title}`);
        assert.equal(getParagraphIlvl(p), '1', `Item ${title} must be at ilvl="1"`);
        assert.equal(getParagraphNumId(p), '9', `Item ${title} must retain numId="9"`);
    }

    for (const title of level2Titles) {
        const p = paragraphs.find(node => getParagraphText(node).includes(title));
        assert.ok(p, `Missing Level 2 item: ${title}`);
        assert.equal(getParagraphIlvl(p), '2', `Item ${title} must be at ilvl="2"`);
        assert.equal(getParagraphNumId(p), '9', `Item ${title} must retain numId="9"`);
    }

    // Sibling continuation
    const wasteP = paragraphs.find(node => getParagraphText(node).includes('County Waste Agreement'));
    assert.ok(wasteP, 'Subsequent County Waste Agreement paragraph must exist');
    assert.equal(getParagraphIlvl(wasteP), '0', 'Subsequent list item must retain ilvl="0"');

    // Rejection verification
    const rejected = rejectTrackedChangesInOoxml(result.documentXml, { allAuthors: true });
    const rejectedDoc = parseXml(rejected.oxml);
    const rejParagraphs = elementsByLocalName(rejectedDoc, 'p');
    for (const title of [...level1Titles, ...level2Titles]) {
        assert.ok(!rejParagraphs.some(node => getParagraphText(node).includes(title)), `Rejected document must not contain: ${title}`);
    }
}

async function testPpgMeetingMultiLevelSubBulletsDozenItems() {
    // Scenario 8: Test 40 PPG Meeting Minutes multi-level sub-bullets with > 12 changes
    // Synthetic creation of 6 Level 1 sub-bullets and 12 Level 2 sub-sub-bullets (18 items total)
    // under Item 3 in Agenda ("Staff changes – Liz").
    const { readFileSync } = await import('fs');
    const { unzipEntries } = await import('../scripts/lib/zip-reader.mjs');
    const zip = unzipEntries(readFileSync('tmp/superdoc-corpus/a3b94b62e9505655d150eae21dc72bb99b37eda901bcb50516b7e9fded5ed33b.docx'));
    const docXml = zip.get('word/document.xml').toString('utf8');

    const modifiedText = [
        '3. Staff changes – Liz',
        '  a. General practitioner clinical staffing',
        '    i. Dr Adisa recruitment and onboarding',
        '    ii. Dr Oxley return from maternity leave',
        '  b. Advanced nursing and allied health roles',
        '    i. Nurse Practitioner appointment schedule',
        '    ii. Clinical pharmacist consultation hours',
        '  c. Reception and patient support team',
        '    i. Telephone triage training curriculum',
        '    ii. Front desk customer service standards',
        '  d. Phlebotomy and diagnostic services',
        '    i. Morning blood clinic capacity expansion',
        '    ii. Specimen courier collection schedule',
        '  e. Locum coverage and leave management',
        '    i. Holiday season locum booking protocol',
        '    ii. Emergency sickness cover rota',
        '  f. Patient communication and directory updates',
        '    i. Website clinician profile revisions',
        '    ii. Practice leaflet reprint distribution'
    ].join('\n');

    const op = {
        type: 'list-change',
        target: 'Staff changes – Liz',
        modified: modifiedText
    };

    const result = await applyOperationToDocumentXml(docXml, op, 'PpgAuthor');
    assert.equal(result.hasChanges, true);

    const accepted = acceptTrackedChangesInOoxml(result.documentXml, { allAuthors: true });
    const acceptedDoc = parseXml(accepted.oxml);
    const paragraphs = elementsByLocalName(acceptedDoc, 'p');

    const level1Titles = [
        'General practitioner clinical staffing',
        'Advanced nursing and allied health roles',
        'Reception and patient support team',
        'Phlebotomy and diagnostic services',
        'Locum coverage and leave management',
        'Patient communication and directory updates'
    ];
    const level2Titles = [
        'Dr Adisa recruitment and onboarding',
        'Dr Oxley return from maternity leave',
        'Nurse Practitioner appointment schedule',
        'Clinical pharmacist consultation hours',
        'Telephone triage training curriculum',
        'Front desk customer service standards',
        'Morning blood clinic capacity expansion',
        'Specimen courier collection schedule',
        'Holiday season locum booking protocol',
        'Emergency sickness cover rota',
        'Website clinician profile revisions',
        'Practice leaflet reprint distribution'
    ];

    for (const title of level1Titles) {
        const p = paragraphs.find(node => getParagraphText(node).includes(title));
        assert.ok(p, `Missing Level 1 item: ${title}`);
        assert.equal(getParagraphIlvl(p), '1', `Item ${title} must be at ilvl="1"`);
        assert.equal(getParagraphNumId(p), '8', `Item ${title} must retain numId="8"`);
    }

    for (const title of level2Titles) {
        const p = paragraphs.find(node => getParagraphText(node).includes(title));
        assert.ok(p, `Missing Level 2 item: ${title}`);
        assert.equal(getParagraphIlvl(p), '2', `Item ${title} must be at ilvl="2"`);
        assert.equal(getParagraphNumId(p), '8', `Item ${title} must retain numId="8"`);
    }

    // Subsequent item retention
    const surgeryP = paragraphs.find(node => getParagraphText(node).includes('Update on New Surgery – Liz'));
    assert.ok(surgeryP, 'Subsequent Surgery paragraph must exist');
    assert.equal(getParagraphIlvl(surgeryP), '0', 'Subsequent list item must retain ilvl="0"');

    // Rejection verification
    const rejected = rejectTrackedChangesInOoxml(result.documentXml, { allAuthors: true });
    const rejectedDoc = parseXml(rejected.oxml);
    const rejParagraphs = elementsByLocalName(rejectedDoc, 'p');
    for (const title of [...level1Titles, ...level2Titles]) {
        assert.ok(!rejParagraphs.some(node => getParagraphText(node).includes(title)), `Rejected document must not contain: ${title}`);
    }
}

// Run all tests
await testNestedChildInsertionAtMultipleLevels();
await testConcurrentEditsAcrossAllThreeLevels();
await testLevelFlatteningGuard();
await testFormattingPreservationAcrossLevels();
await testMultiLevelRangeDeletionAndAcceptance();
await testAgendaMultiLevelSubBulletsDozenItems();
await testAgendaFebMultiLevelSubBulletsDozenItems();
await testPpgMeetingMultiLevelSubBulletsDozenItems();

console.log('PASS: multilevel_bullet_tests.mjs - all 8 multi-level bullet tests passed.');


