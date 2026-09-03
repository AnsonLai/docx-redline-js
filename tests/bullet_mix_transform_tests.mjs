import assert from 'assert';
import { readFileSync } from 'fs';
import {
    applyOperationToDocumentXml,
    applyOperationsToDocumentXml
} from '../services/standalone-operation-runner.js';
import {
    applyRedlineToOxml,
    acceptTrackedChangesInOoxml,
    rejectTrackedChangesInOoxml
} from '../index.js';
import {
    elementsByLocalName,
    parseXml
} from './helpers/ooxml-assertions.mjs';
import { unzipEntries } from '../scripts/lib/zip-reader.mjs';
import './setup-xml-provider.mjs';


function getParagraphText(p) {
    const texts = [];
    const walk = (node) => {
        if (!node) return;
        if (node.nodeType === 1 && node.localName === 't') {
            texts.push(node.textContent || '');
        }
        for (let i = 0; i < (node.childNodes?.length || 0); i++) {
            walk(node.childNodes[i]);
        }
    };
    walk(p);
    return texts.join('');
}

function getParagraphIlvl(p) {
    const numPr = Array.from(p.childNodes || []).find(n => n.localName === 'pPr')?.childNodes;
    if (!numPr) return null;
    const np = Array.from(numPr).find(n => n.localName === 'numPr');
    if (!np) return null;
    const ilvlNode = Array.from(np.childNodes).find(n => n.localName === 'ilvl');
    return ilvlNode?.getAttribute('w:val') || ilvlNode?.getAttribute('val') || null;
}

function getParagraphNumId(p) {
    const numPr = Array.from(p.childNodes || []).find(n => n.localName === 'pPr')?.childNodes;
    if (!numPr) return null;
    const np = Array.from(numPr).find(n => n.localName === 'numPr');
    if (!np) return null;
    const numIdNode = Array.from(np.childNodes).find(n => n.localName === 'numId');
    return numIdNode?.getAttribute('w:val') || numIdNode?.getAttribute('val') || null;
}

// ----------------------------------------------------------------------------
// Test 1: Inline Phrase Edit Inside an Existing Bullet Item
// Verifies surgical text replacement inside a bullet run without altering list properties.
// ----------------------------------------------------------------------------
async function testInlinePhraseEditInsideExistingBullet() {
    const pOxml = [
        '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">',
        '  <w:pPr><w:numPr><w:ilvl w:val="1"/><w:numId w:val="5"/></w:numPr></w:pPr>',
        '  <w:r><w:t>Submit monthly safety inspection report by the 15th of each calendar month.</w:t></w:r>',
        '</w:p>'
    ].join('\n');

    const res = await applyRedlineToOxml(
        pOxml,
        'Submit monthly safety inspection report by the 15th of each calendar month.',
        'Submit monthly safety and hazard inspection report by the 20th of each calendar month.',
        { generateRedlines: true, author: 'Inspector' }
    );

    assert.equal(res.hasChanges, true);
    assert.ok(res.oxml.includes('<w:ins'), 'Must contain tracked insertion');
    assert.ok(res.oxml.includes('<w:del'), 'Must contain tracked deletion');

    const doc = parseXml(res.oxml);
    const p = elementsByLocalName(doc, 'p')[0];
    assert.equal(getParagraphIlvl(p), '1', 'Must retain ilvl="1"');
    assert.equal(getParagraphNumId(p), '5', 'Must retain numId="5"');

    // Accept restores modified sentence
    const accepted = acceptTrackedChangesInOoxml(res.oxml, { allAuthors: true });
    const accDoc = parseXml(accepted.oxml);
    const accP = elementsByLocalName(accDoc, 'p')[0];
    assert.equal(getParagraphText(accP), 'Submit monthly safety and hazard inspection report by the 20th of each calendar month.');

    // Reject restores original sentence
    const rejected = rejectTrackedChangesInOoxml(res.oxml, { allAuthors: true });
    const rejDoc = parseXml(rejected.oxml);
    const rejP = elementsByLocalName(rejDoc, 'p')[0];
    assert.equal(getParagraphText(rejP), 'Submit monthly safety inspection report by the 15th of each calendar month.');
}

// ----------------------------------------------------------------------------
// Test 2: Inline Formatting Changes Inside Bullet Items
// Verifies bold / italic markdown formatting conversion within bullet text.
// ----------------------------------------------------------------------------
async function testInlineFormattingInsideBulletItem() {
    const pOxml = [
        '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">',
        '  <w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="3"/></w:numPr></w:pPr>',
        '  <w:r><w:t>Permits must be displayed prominently at the job site.</w:t></w:r>',
        '</w:p>'
    ].join('\n');

    const res = await applyRedlineToOxml(
        pOxml,
        'Permits must be displayed prominently at the job site.',
        '**Mandatory Notice**: Permits must be displayed *prominently* at the job site.',
        { generateRedlines: true, author: 'CodeOfficer' }
    );

    assert.equal(res.hasChanges, true);
    assert.ok(res.oxml.includes('<w:b'), 'Must contain bold tag <w:b');
    assert.ok(res.oxml.includes('<w:rPrChange') || res.oxml.includes('<w:ins'), 'Must track format or text changes');

    const doc = parseXml(res.oxml);
    const p = elementsByLocalName(doc, 'p')[0];
    assert.equal(getParagraphIlvl(p), '0', 'Must preserve ilvl="0"');
    assert.equal(getParagraphNumId(p), '3', 'Must preserve numId="3"');

    const accepted = acceptTrackedChangesInOoxml(res.oxml);
    assert.ok(accepted.oxml.includes('Mandatory Notice'), 'Accepted must include bold prefix text');
}

// ----------------------------------------------------------------------------
// Test 3: Tracked Deletion of Entire Bullet Items
// Verifies that deleting a bullet marks it with w:del and cleans up on Accept.
// ----------------------------------------------------------------------------
async function testTrackedDeletionOfBulletItem() {
    const zip = unzipEntries(readFileSync('tmp/superdoc-corpus/77c33ab2ca57efd0d4afb87f08c45c88f3984256c3fd5fdf3e81f87a24644760.docx'));
    const docXml = zip.get('word/document.xml').toString('utf8');

    const op = {
        type: 'delete',
        target: 'Tax Collector Report'
    };

    const res = await applyOperationToDocumentXml(docXml, op, 'Auditor');
    assert.equal(res.hasChanges, true);
    assert.ok(res.documentXml.includes('<w:del'), 'Must contain tracked deletion markup');
    assert.ok(res.documentXml.includes('Tax Collector Report'), 'Text must remain in w:del prior to acceptance');

    // Accept cleans it up
    const accepted = acceptTrackedChangesInOoxml(res.documentXml, { allAuthors: true });
    assert.ok(!accepted.oxml.includes('Tax Collector Report'), 'Accepted document must not contain deleted bullet text');

    // Reject restores it
    const rejected = rejectTrackedChangesInOoxml(res.documentXml, { allAuthors: true });
    assert.ok(rejected.oxml.includes('Tax Collector Report'), 'Rejected document must cleanly restore original bullet');
}

// ----------------------------------------------------------------------------
// Test 4: Splitting a Compound Bullet Item into Sub-Bullets
// ----------------------------------------------------------------------------
async function testBulletSplittingIntoMultipleSubBullets() {
    const zip = unzipEntries(readFileSync('tmp/superdoc-corpus/77c33ab2ca57efd0d4afb87f08c45c88f3984256c3fd5fdf3e81f87a24644760.docx'));
    const docXml = zip.get('word/document.xml').toString('utf8');

    const op = {
        type: 'list-change',
        target: 'Water District Report',
        modified: [
            '2. Water District Report',
            '  - Infrastructure assessment and leak detection',
            '    - Complete acoustic pipe survey for District 1',
            '    - Prioritize main valve replacements on Main Street',
            '  - Meter replacement and telemetry upgrades',
            '    - Install 120 smart radio meters by fiscal year end',
            '    - Calibrate master flow meter at pump station',
            '  - Water quality testing and state compliance reporting',
            '    - Submit quarterly lead and copper sample results',
            '    - Publish annual drinking water consumer confidence report'
        ].join('\n')
    };

    const res = await applyOperationToDocumentXml(docXml, op, 'WaterEngineer');
    assert.equal(res.hasChanges, true);

    const accepted = acceptTrackedChangesInOoxml(res.documentXml, { allAuthors: true });
    const accDoc = parseXml(accepted.oxml);
    const paragraphs = elementsByLocalName(accDoc, 'p');

    const l1Items = [
        'Infrastructure assessment and leak detection',
        'Meter replacement and telemetry upgrades',
        'Water quality testing and state compliance reporting'
    ];
    const l2Items = [
        'Complete acoustic pipe survey for District 1',
        'Prioritize main valve replacements on Main Street',
        'Install 120 smart radio meters by fiscal year end',
        'Calibrate master flow meter at pump station',
        'Submit quarterly lead and copper sample results',
        'Publish annual drinking water consumer confidence report'
    ];

    for (const text of l1Items) {
        const p = paragraphs.find(node => getParagraphText(node).includes(text));
        assert.ok(p, `Missing Level 1 split item: ${text}`);
        assert.equal(getParagraphIlvl(p), '1', `Item "${text}" must be at ilvl="1"`);
        assert.equal(getParagraphNumId(p), '2', 'Must retain Reports numId="2"');
    }

    for (const text of l2Items) {
        const p = paragraphs.find(node => getParagraphText(node).includes(text));
        assert.ok(p, `Missing Level 2 split item: ${text}`);
        assert.equal(getParagraphIlvl(p), '2', `Item "${text}" must be at ilvl="2"`);
        assert.equal(getParagraphNumId(p), '2', 'Must retain Reports numId="2"');
    }
}

// ----------------------------------------------------------------------------
// Test 5: Atomic Compound Batch on Bullet Hierarchy (Text Edit + Comment + Sub-bullets + Deletion)
// ----------------------------------------------------------------------------
async function testAtomicMixedBatchOnBulletHierarchy() {
    const zip = unzipEntries(readFileSync('tmp/superdoc-corpus/77c33ab2ca57efd0d4afb87f08c45c88f3984256c3fd5fdf3e81f87a24644760.docx'));
    const docXml = zip.get('word/document.xml').toString('utf8');

    const batch = [
        {
            type: 'comment',
            target: 'Highway',
            textToComment: 'Highway',
            commentContent: 'Verify snowplow route schedule before first frost.'
        },
        {
            type: 'replace',
            target: 'Water District Report',
            modified: 'Municipal Water and Sewer District Annual Report'
        },
        {
            type: 'list-change',
            target: 'Code Enforcement Monthly Report',
            modified: [
                '3. Code Enforcement Monthly Report',
                '  - Building Permits and Certificates of Occupancy',
                '    - Fourteen residential building permits issued',
                '    - Three commercial certificates of occupancy granted',
                '  - Zoning Violations and Stop Work Orders',
                '    - Two stop work orders issued on unpermitted grading',
                '    - Five setback violation notifications served'
            ].join('\n')
        },
        {
            type: 'comment',
            target: 'WWTP Report',
            textToComment: 'WWTP',
            commentContent: 'Confirm state wastewater permit renewal date.'
        },
        {
            type: 'delete',
            target: 'Tax Collector Report'
        }
    ];

    const res = await applyOperationsToDocumentXml(docXml, batch, 'CouncilBoard');
    assert.equal(res.hasChanges, true);
    assert.equal(res.results.length, 5);
    for (const r of res.results) {
        assert.equal(r.status, 'applied', `Op ${r.index} (${r.type}) must be applied`);
    }
    assert.ok(res.commentsXml, 'Must generate commentsXml for the two bullet comments');
    assert.ok(res.commentsXml.includes('Verify snowplow route schedule'));
    assert.ok(res.commentsXml.includes('Confirm state wastewater permit renewal date'));

    // Verify Accept All
    const accepted = acceptTrackedChangesInOoxml(res.documentXml, { allAuthors: true });
    const accDoc = parseXml(accepted.oxml);
    const accWaterP = elementsByLocalName(accDoc, 'p').find(p => getParagraphText(p).includes('Municipal Water and Sewer District Annual Report'));
    assert.ok(accWaterP, 'Modified Water District Report present in accepted');
    assert.ok(accepted.oxml.includes('Fourteen residential building permits issued'));
    assert.ok(!accepted.oxml.includes('Tax Collector Report'), 'Tax Collector Report must be removed on accept');

    // Verify Reject All
    const rejected = rejectTrackedChangesInOoxml(res.documentXml, { allAuthors: true });
    const rejDoc = parseXml(rejected.oxml);
    const rejWaterP = elementsByLocalName(rejDoc, 'p').find(p => getParagraphText(p).includes('Water District Report'));
    assert.ok(rejWaterP, 'Original Water District Report restored');
    assert.ok(!rejected.oxml.includes('Fourteen residential building permits issued'), 'Sub-bullets removed on reject');
    assert.ok(rejected.oxml.includes('Tax Collector Report'), 'Tax Collector Report restored on reject');
}

// ----------------------------------------------------------------------------
// Test 6: Symbol Bullets Mixed with Numbered Clauses (Bylaws)
// ----------------------------------------------------------------------------
async function testSymbolBulletsMixedWithNumberedClauses() {
    const zip = unzipEntries(readFileSync('tmp/superdoc-corpus/d27fe5513ca474bcd6c02136d2e4b8179203f5855c4333af81bc964755c7abe4.docx'));
    const docXml = zip.get('word/document.xml').toString('utf8');

    const ops = [
        {
            type: 'replace',
            target: 'One (1) representative from the Board of Supervisors;',
            modified: 'Two (2) voting representatives from the County Board of Supervisors;'
        },
        {
            type: 'comment',
            target: 'Five (5) representatives from the incorporated cities/town;',
            textToComment: 'incorporated cities/town',
            commentContent: 'Confirm municipal apportionment following 2024 census.'
        },
        {
            type: 'replace',
            target: 'Three (3) representatives from the community at large.',
            modified: 'Three (3) representatives from the community at large, including at least one youth advocate.'
        }
    ];

    const res = await applyOperationsToDocumentXml(docXml, ops, 'BylawsCommittee');
    assert.equal(res.hasChanges, true);
    assert.equal(res.results.length, 3);
    for (const r of res.results) {
        assert.equal(r.status, 'applied');
    }

    const accepted = acceptTrackedChangesInOoxml(res.documentXml, { allAuthors: true });
    const accDoc = parseXml(accepted.oxml);
    const accP = elementsByLocalName(accDoc, 'p').find(p => getParagraphText(p).includes('voting representatives from the County Board'));
    assert.ok(accP, 'Must find accepted modified supervisor bullet');
    assert.ok(elementsByLocalName(accDoc, 'p').some(p => getParagraphText(p).includes('youth advocate')));

    const rejected = rejectTrackedChangesInOoxml(res.documentXml, { allAuthors: true });
    const rejDoc = parseXml(rejected.oxml);
    const rejP = elementsByLocalName(rejDoc, 'p').find(p => getParagraphText(p).includes('One (1) representative from the Board of Supervisors;'));
    assert.ok(rejP, 'Must restore original supervisor bullet');
    assert.ok(!elementsByLocalName(rejDoc, 'p').some(p => getParagraphText(p).includes('youth advocate')));
}

// Execute all tests
await testInlinePhraseEditInsideExistingBullet();
await testInlineFormattingInsideBulletItem();
await testTrackedDeletionOfBulletItem();
await testBulletSplittingIntoMultipleSubBullets();
await testAtomicMixedBatchOnBulletHierarchy();
await testSymbolBulletsMixedWithNumberedClauses();

console.log('PASS: bullet_mix_transform_tests.mjs - all 6 bullet mix and transform tests passed.');
