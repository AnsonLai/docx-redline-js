import assert from 'assert/strict';
import { createHash } from 'crypto';
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'fs';
import { dirname, join } from 'path';
import { fileURLToPath } from 'url';

import './setup-xml-provider.mjs';
import {
    acceptTrackedChangesInOoxml,
    ingestWordOoxmlToPlainText,
    rejectTrackedChangesInOoxml
} from '../index.js';
import { openDocx } from '../node/index.js';
import { MemoryZip } from '../node/zip-archive.js';
import { unzipEntries } from '../scripts/lib/zip-reader.mjs';
import { preflightOperations } from '../services/operation-preflight.js';
import { validateDocxPackage } from '../services/standalone-docx-plumbing.js';
import { applyOperationsToDocumentXml } from '../services/standalone-operation-runner.js';

const testDir = dirname(fileURLToPath(import.meta.url));
const repoRoot = dirname(testDir);

// ----------------------------------------------------------------------------
// Corpus Document Identity & Provenance
// ----------------------------------------------------------------------------
export const CORPUS_ID = '89d6188c5d95a0b718a80627bd5da7c1b0afccf0198971ba23244db75bb38fe7';
export const EXPECTED_SHA256 = '0f16166ed9f42df70cb61904e612713ece8eb7ba935fe5555bb1a079fadee4b0';
export const DOWNLOAD_URL = `https://docxcorp.us/documents/${CORPUS_ID}.docx`;
export const LOCAL_CORPUS_PATH = join(repoRoot, 'tmp', 'superdoc-corpus', `${CORPUS_ID}.docx`);

/**
 * Loads the source DOCX buffer from the local corpus cache, or fetches it
 * from the SuperDoc docx-corpus CDN if absent, verifying its SHA-256 digest.
 */
export async function loadSourceDocxBuffer() {
    if (existsSync(LOCAL_CORPUS_PATH)) {
        const buffer = readFileSync(LOCAL_CORPUS_PATH);
        const digest = createHash('sha256').update(buffer).digest('hex');
        if (digest === EXPECTED_SHA256) {
            return buffer;
        }
    }

    // Auto-fetch if not cached or hash mismatch
    const response = await fetch(DOWNLOAD_URL);
    if (!response.ok) {
        throw new Error(`Failed to fetch corpus document ${CORPUS_ID}: HTTP ${response.status}`);
    }
    const buffer = Buffer.from(await response.arrayBuffer());
    const digest = createHash('sha256').update(buffer).digest('hex');
    assert.equal(digest, EXPECTED_SHA256, `SHA-256 mismatch for fetched document ${CORPUS_ID}`);

    mkdirSync(dirname(LOCAL_CORPUS_PATH), { recursive: true });
    writeFileSync(LOCAL_CORPUS_PATH, buffer);
    return buffer;
}

// ----------------------------------------------------------------------------
// Real-World Legal Markup Suite (Healthcare Transaction Counsel Review)
//
// Models a comprehensive contract markup by healthcare counsel acting for a
// community pharmacy consortium negotiating the Southampton City Council
// Smoking Cessation Agreement (10,770 words, 533 paragraphs, 5 tables).
// ----------------------------------------------------------------------------
export const LEGAL_MARKUP_OPERATIONS = [
    // --- 1. Parties & Recitals ---
    {
        type: 'comment',
        target: { exactText: 'Southampton City Council of Civic Centre, Southampton, SO14 7LT (“the Council”)' },
        textToComment: 'SO14 7LT',
        commentContent: 'Verify official registered office postcode (Council constitution lists SO14 7LY).',
        author: 'Healthcare Counsel'
    },
    {
        type: 'replace',
        target: { exactText: '[INSERT NAME/BRANCH OF PHARMACY]' },
        modified: 'Southampton Community Pharmacy Consortium Limited (Company No. 08492011) of High Street, Southampton, SO14 2DF (“the Service Provider”)',
        author: 'Healthcare Counsel'
    },
    {
        type: 'comment',
        target: { exactText: 'Whereas:' },
        textToComment: 'Whereas:',
        commentContent: 'Confirm recital incorporates explicit statutory reference to Section 2B National Health Service Act 2006 for local authority public health commissioning.',
        author: 'Healthcare Counsel'
    },

    // --- 2. Relationship between Parties: Catching Boilerplate Error ---
    {
        type: 'replace',
        target: {
            exactText: 'The Service Provider is an independent provider of general practitioner services and is not an employee, partner or agent of the Council. The Service Provider must not allow its employees or agents to represent or conduct its activities in a manner so as to give the impression that the Service Provider is an employee, partner or agent of the Council.'
        },
        modified: 'The Service Provider is an independent provider of community pharmacy services and is not an employee, partner or agent of the Council. The Service Provider must not allow its employees or agents to represent or conduct its activities in a manner so as to give the impression that the Service Provider is an employee, partner or agent of the Council.',
        author: 'Healthcare Counsel'
    },
    {
        type: 'comment',
        target: {
            exactText: 'The Service Provider is an independent provider of general practitioner services and is not an employee, partner or agent of the Council. The Service Provider must not allow its employees or agents to represent or conduct its activities in a manner so as to give the impression that the Service Provider is an employee, partner or agent of the Council.'
        },
        textToComment: 'general practitioner services',
        commentContent: 'Drafting defect corrected: the draft inadvertently imported standard GP medical services wording instead of community pharmacy services.',
        author: 'Healthcare Counsel'
    },

    // --- 3. Clause 1: Commencement and Duration (Multi-Paragraph Term & Extension) ---
    {
        type: 'replace',
        target: {
            exactText: 'This Agreement shall take effect on (1/04/2020) (the “Commencement Date”) and shall continue in full force until (31/03/2021) (the “End Date”) unless otherwise terminated in accordance with the provisions of this Agreement.'
        },
        modified: 'This Agreement shall take effect on 1st April 2020 (the “Commencement Date”) and shall continue in full force until 31st March 2022 (the “End Date”) unless otherwise extended or terminated in accordance with the provisions of this Agreement.',
        author: 'Healthcare Counsel'
    },
    {
        type: 'comment',
        target: {
            exactText: 'This Agreement shall take effect on (1/04/2020) (the “Commencement Date”) and shall continue in full force until (31/03/2021) (the “End Date”) unless otherwise terminated in accordance with the provisions of this Agreement.'
        },
        textToComment: 'Commencement Date',
        commentContent: 'Two-year initial term aligns with CCG public health funding ring-fence cycle. Break costs clause recommended if council seeks early break option.',
        author: 'Healthcare Counsel'
    },
    {
        type: 'replace',
        target: {
            exactText: 'The Council may extend the term of this Agreement at its absolute discretion by up to 12 months on giving the Service Provider no less than 3 months’ written notice of such intention before the End Date, such notice to be issued separately for each service under this Agreement, as detailed in Schedule 2 (Location of and services to be provided).'
        },
        modified: 'The Council may, subject to mutual written agreement with the Service Provider, extend the term of this Agreement by up to twelve (12) months on giving no less than three (3) months’ prior written notice before the End Date, such extension to be confirmed in writing by both Parties.',
        author: 'Healthcare Counsel'
    },

    // --- 4. Indemnity & Liability Cap ---
    {
        type: 'comment',
        target: {
            exactText: 'Without prejudice to any rights or remedies of the Council, the Service Provider shall indemnify the Council and keep the Council indemnified in full against any expense, liability, loss, claim, fine, cost or proceeding whatsoever incurred by or made against the Council arising directly or indirectly out of the wrongful act, default, breach of contract or negligence of the Service Provider, its subcontractors, employees or agents in the course of the provision of the services detailed in Schedule 1, or otherwise in connection with this Agreement.'
        },
        textToComment: 'indemnify the Council',
        commentContent: 'Provider cannot agree to open-ended indemnity without an aggregate liability cap. Propose capping aggregate liability at 150% of annual fees paid hereunder, excluding death/personal injury.',
        author: 'Healthcare Counsel'
    },

    // --- 5. Clause 13.2: Insurance (Numbered List Items & Statutory Cover) ---
    {
        type: 'replace',
        target: { exactText: '13.2.1.1 employer\'s liability (of at least £2,000,000 in respect of any one claim); and' },
        modified: '13.2.1.1 employer\'s liability (of at least £10,000,000 in respect of any one claim); and',
        author: 'Healthcare Counsel'
    },
    {
        type: 'replace',
        target: { exactText: '13.2.1.2 public liability (up to £5,000,000 in respect of any one claim),' },
        modified: '13.2.1.2 public liability (of at least £5,000,000 in respect of any one claim); and',
        author: 'Healthcare Counsel'
    },
    {
        type: 'comment',
        target: { exactText: '13.2.1.2 public liability (up to £5,000,000 in respect of any one claim),' },
        textToComment: 'up to £5,000,000',
        commentContent: 'Replaced "up to" with "of at least" to eliminate ambiguity regarding minimum cover obligations versus indemnity ceilings.',
        author: 'Healthcare Counsel'
    },
    {
        type: 'replace',
        target: { exactText: '13.2.1.3 Professional indemnity (of at least £2,000,000 in respect of any one claim),' },
        modified: '13.2.1.3 professional indemnity (of at least £5,000,000 in respect of any one claim and in the annual aggregate, with six (6) years retroactive run-off cover); and',
        author: 'Healthcare Counsel'
    },

    // --- 6. Clause 14: Termination & Statutory Compliance (GDPR Update) ---
    {
        type: 'replace',
        target: {
            exactText: '14.8 Where the Council has been served with a notice of discontinuation of funding by the funder or its agents the Council may terminate this Agreement by giving written notice to the Service Provider. The period of notice given by the Council to the Service Provider shall be no more than 5 Working Days less than the period of notice given to the Council by the funder or three months whichever is less.'
        },
        modified: '14.8 Where the Council has been served with a notice of discontinuation of funding by the funder or its agents the Council may terminate this Agreement by giving written notice to the Service Provider. The period of notice given by the Council to the Service Provider shall be no more than 5 Working Days less than the period of notice given to the Council by the funder or three months whichever is less, provided that the Council shall provide not less than thirty (30) Working Days prior written notice to allow orderly completion or handover of active client quit attempts.',
        author: 'Healthcare Counsel'
    },
    {
        type: 'comment',
        target: {
            exactText: '14.8 Where the Council has been served with a notice of discontinuation of funding by the funder or its agents the Council may terminate this Agreement by giving written notice to the Service Provider. The period of notice given by the Council to the Service Provider shall be no more than 5 Working Days less than the period of notice given to the Council by the funder or three months whichever is less.'
        },
        textToComment: 'discontinuation of funding',
        commentContent: '5 days notice would force abrupt cessation of patient medication courses and leave unconsumed pharmacy NRT stock stranded without reimbursement. 30 days floor is commercially essential.',
        author: 'Healthcare Counsel'
    },
    {
        type: 'replace',
        target: {
            exactText: '14.11 The Service Provider must not commit any breach of the Employment Relations Act 1999 (Blacklists) Regulations 2010 or section 137 of the Trade Union and Labour Relations (Consolidation) Act 1992, or commit any breach of the Data Protection Act 1998 by unlawfully processing personal data in connection with any blacklisting activities. The Council may terminate this Agreement with immediate effect in the event of any breach by the Service Provider of this clause 14.'
        },
        modified: '14.11 The Service Provider must not commit any breach of the Employment Relations Act 1999 (Blacklists) Regulations 2010 or section 137 of the Trade Union and Labour Relations (Consolidation) Act 1992, or commit any breach of the UK GDPR (as defined in the Data Protection Act 2018) or the Data Protection Act 2018 by unlawfully processing personal data in connection with any blacklisting activities. The Council may terminate this Agreement with immediate effect in the event of any material breach by the Service Provider of this clause 14.',
        author: 'Healthcare Counsel'
    },

    // --- 7. Schedule 1: Table 3 (Service Schedule Term & Annual Review Date) ---
    {
        type: 'replace',
        target: { exactText: '1st April 2020 – 31st March 2021' },
        modified: '1st April 2020 – 31st March 2022 (with option to extend to 31st March 2023)',
        author: 'Healthcare Counsel'
    },
    {
        type: 'replace',
        target: { exactText: 'By March 2021' },
        modified: 'Annually, no later than 31st January of each contract year',
        author: 'Healthcare Counsel'
    },
    {
        type: 'comment',
        target: { exactText: 'By March 2021' },
        textToComment: 'By March 2021',
        commentContent: 'Review schedule moved to January to allow timely input into CCG joint commissioning intentions before council financial year-end.',
        author: 'Healthcare Counsel'
    },

    // --- 8. Schedule 1: Clinical Protocols (Treatment Session List Items) ---
    {
        type: 'replace',
        target: { exactText: 'complete a carbon monoxide (CO) test and an explanation of its use as a motivational aid;' },
        modified: 'complete a validated carbon monoxide (CO) breath test (recording calibrated ppm values in Pharmoutcomes) and provide a structured explanation of its use as a motivational aid;',
        author: 'Healthcare Counsel'
    },
    {
        type: 'comment',
        target: { exactText: 'complete a carbon monoxide (CO) test and an explanation of its use as a motivational aid;' },
        textToComment: 'carbon monoxide (CO) test',
        commentContent: 'Request explicit provision for telephone/video consultations where client cannot attend pharmacy premises due to shielding or infection control guidelines.',
        author: 'Healthcare Counsel'
    },
    {
        type: 'comment',
        target: {
            exactText: 'identify if the person is on antipsychotic medication and, if so, ask them for consent to let their prescriber know that they are stopping smoking and, if applicable starting NRT, so their medication can be kept under review. Subsequently communicate with the prescriber. This will be through Pharmoutcomes where possible. This is because some people need their doses changed during quit attempts. Being on an antipsychotic is not a barrier to this service or to stopping smoking.'
        },
        textToComment: 'ask them for consent',
        commentContent: 'Ensure client consent mechanism satisfies UK GDPR Article 9 explicit consent standards for special category healthcare data. Prescriber notice SLA should be within 48 hours.',
        author: 'Healthcare Counsel'
    },

    // --- 9. Schedule 1: Table 5 (Unit Costs, Fee Tariffs & Reimbursement Terms) ---
    {
        type: 'replace',
        target: {
            exactText: 'For a Department of Health and Social Care 4 week quit of a pregnant woman, person on anti-psychotic medication or person under the care of homeless healthcare or specialist substance misuse services(Higher payment rate)'
        },
        modified: 'For a Department of Health and Social Care 4 week quit of a pregnant woman, person on anti-psychotic medication or person under the care of homeless healthcare or specialist substance misuse services (Higher payment rate of £85.00 per validated client)',
        author: 'Healthcare Counsel'
    },
    {
        type: 'comment',
        target: {
            exactText: 'For a Department of Health and Social Care 4 week quit of a pregnant woman, person on anti-psychotic medication or person under the care of homeless healthcare or specialist substance misuse services(Higher payment rate)'
        },
        textToComment: 'Higher payment rate',
        commentContent: 'Clarify whether the higher payment rate (£85.00) applies if a client enters the cohort mid-way through the 4-week window (e.g. newly confirmed pregnancy).',
        author: 'Healthcare Counsel'
    },
    {
        type: 'replace',
        target: { exactText: 'For a Department of Health and Social Care 4 week quit for anyone else (Standard payment rate)' },
        modified: 'For a Department of Health and Social Care 4 week quit for anyone else (Standard payment rate of £45.00 per validated client)',
        author: 'Healthcare Counsel'
    },
    {
        type: 'replace',
        target: { exactText: 'For completing an initial consultation and setting a quit date' },
        modified: 'For completing an initial consultation and setting a recorded quit date (£25.00 per client)',
        author: 'Healthcare Counsel'
    },
    {
        type: 'replace',
        target: { exactText: 'SCC will also reimburse the cost of NRT at C+D cost price plus 5% VAT on the basis of information supplied through Pharmoutcomes.' },
        modified: 'SCC will also reimburse the cost of NRT at current Chemist and Druggist (C+D) trade cost price plus 5% VAT on the basis of claims submitted monthly via Pharmoutcomes.',
        author: 'Healthcare Counsel'
    },
    {
        type: 'comment',
        target: { exactText: 'SCC will also reimburse the cost of NRT at C+D cost price plus 5% VAT on the basis of information supplied through Pharmoutcomes.' },
        textToComment: 'C+D cost price',
        commentContent: 'Where C+D list price fluctuates or product shortages require proprietary alternative sourcing, provider should be reimbursed at actual invoice acquisition cost upon submission of wholesaler receipt.',
        author: 'Healthcare Counsel'
    }
];

// ----------------------------------------------------------------------------
// Test 1: Preflight Validation of Complex Legal Markup Batch
// Verifies that strict target resolution uniquely disambiguates all 28 targets
// across the 10,770-word document without ambiguities or collision issues.
// ----------------------------------------------------------------------------
async function testPreflightLegalBatch(docXml) {
    const preflight = preflightOperations(docXml, LEGAL_MARKUP_OPERATIONS, 'Healthcare Counsel');
    assert.equal(preflight.status, 'ok', 'Preflight must succeed for all 28 operations');
    assert.equal(preflight.issues?.length ?? 0, 0, 'Preflight must report zero issues');
}

// ----------------------------------------------------------------------------
// Test 2: Engine Atomic Batch Application & Comments Generation
// Verifies applying multi-paragraph edits, table changes, list items, and
// inline comments to document XML produces valid tracked revisions and commentsXml.
// ----------------------------------------------------------------------------
async function testEngineAtomicBatch(docXml) {
    const res = await applyOperationsToDocumentXml(docXml, LEGAL_MARKUP_OPERATIONS, 'Healthcare Counsel');
    assert.equal(res.hasChanges, true, 'Batch must produce changes');
    assert.equal(res.results.length, LEGAL_MARKUP_OPERATIONS.length, 'Must report result for every operation');

    for (const r of res.results) {
        assert.equal(r.status, 'applied', `Operation ${r.index} (${r.type}) must be applied`);
    }

    // Verify commentsXml artifact was produced
    assert.ok(res.commentsXml, 'Batch must generate commentsXml for the 12 comment operations');
    assert.ok(res.commentsXml.includes('Verify official registered office postcode'));
    assert.ok(res.commentsXml.includes('Section 2B National Health Service Act 2006'));
    assert.ok(res.commentsXml.includes('Drafting defect corrected'));
    assert.ok(res.commentsXml.includes('UK GDPR Article 9 explicit consent'));
    assert.ok(res.commentsXml.includes('higher payment rate (£85.00)'));
    assert.ok(res.commentsXml.includes('C+D list price'));

    // Verify tracked changes markup in documentXml
    assert.ok(res.documentXml.includes('<w:ins'), 'documentXml must contain w:ins elements');
    assert.ok(res.documentXml.includes('<w:del'), 'documentXml must contain w:del elements');
    assert.ok(res.documentXml.includes('<w:commentRangeStart'), 'documentXml must contain comment start markers');
    assert.ok(res.documentXml.includes('<w:commentRangeEnd'), 'documentXml must contain comment end markers');
    assert.ok(res.documentXml.includes('<w:commentReference'), 'documentXml must contain comment references');

    return res;
}

// ----------------------------------------------------------------------------
// Test 3: Differential Accept All on Complete Legal Markup
// Verifies that accepting all tracked changes produces the clean amended text
// with all lawyer redlines integrated and all deleted original text removed.
// ----------------------------------------------------------------------------
async function testDifferentialAcceptAll(trackedDocXml) {
    const accepted = acceptTrackedChangesInOoxml(trackedDocXml, { allAuthors: true });
    assert.equal(accepted.hasChanges, true, 'Accepting revisions must report changes from tracked state');

    const plainText = ingestWordOoxmlToPlainText(accepted.oxml);

    // Additions must be present
    assert.ok(plainText.includes('Southampton Community Pharmacy Consortium Limited'), 'Must include consortium entity');
    assert.ok(plainText.includes('community pharmacy services'), 'Must include corrected pharmacy scope');
    assert.ok(plainText.includes('1st April 2020'), 'Must include formal commencement date');
    assert.ok(plainText.includes('31st March 2022'), 'Must include 2-year term end date');
    assert.ok(plainText.includes('31st March 2023'), 'Must include optional extension date');
    assert.ok(plainText.includes('£10,000,000'), 'Must include upgraded employer liability');
    assert.ok(plainText.includes('thirty (30) Working Days'), 'Must include 30-day notice floor');
    assert.ok(plainText.includes('UK GDPR (as defined in the Data Protection Act 2018)'), 'Must include UK GDPR update');
    assert.ok(plainText.includes('£85.00 per validated client'), 'Must include higher rate tariff');
    assert.ok(plainText.includes('£45.00 per validated client'), 'Must include standard rate tariff');
    assert.ok(plainText.includes('£25.00 per client'), 'Must include initial consultation fee');
    assert.ok(plainText.includes('calibrated ppm values in Pharmoutcomes'), 'Must include CO ppm protocol');

    // Deletions must be completely absent
    assert.ok(!plainText.includes('[INSERT NAME/BRANCH OF PHARMACY]'), 'Old placeholder must be excised');
    assert.ok(!plainText.includes('general practitioner services'), 'Old GP services text must be excised');
    assert.ok(!plainText.includes('(1/04/2020)'), 'Old bracketed date must be excised');
    assert.ok(!plainText.includes('(31/03/2021)'), 'Old bracketed end date must be excised');
    assert.ok(!plainText.includes('Data Protection Act 1998'), 'Repealed 1998 Act must be excised');
    assert.ok(!plainText.includes('up to £5,000,000 in respect of any one claim'), 'Ambiguous public liability must be excised');
}

// ----------------------------------------------------------------------------
// Test 4: Differential Reject All on Complete Legal Markup
// Verifies that rejecting all tracked changes cleanly restores the untouched
// source text, leaving no trace of the proposed revisions.
// ----------------------------------------------------------------------------
async function testDifferentialRejectAll(trackedDocXml) {
    const rejected = rejectTrackedChangesInOoxml(trackedDocXml, { allAuthors: true });
    assert.equal(rejected.hasChanges, true, 'Rejecting revisions must report changes from tracked state');

    const plainText = ingestWordOoxmlToPlainText(rejected.oxml);

    // Original text must be restored
    assert.ok(plainText.includes('[INSERT NAME/BRANCH OF PHARMACY]'), 'Original placeholder must be restored');
    assert.ok(plainText.includes('general practitioner services'), 'Original GP text must be restored');
    assert.ok(plainText.includes('(1/04/2020)'), 'Original commencement date must be restored');
    assert.ok(plainText.includes('(31/03/2021)'), 'Original end date must be restored');
    assert.ok(plainText.includes('Data Protection Act 1998'), 'Original 1998 Act reference must be restored');
    assert.ok(plainText.includes('up to £5,000,000 in respect of any one claim'), 'Original public liability text must be restored');

    // Proposed lawyer additions must be completely absent
    assert.ok(!plainText.includes('Southampton Community Pharmacy Consortium Limited'), 'Consortium name must not appear');
    assert.ok(!plainText.includes('community pharmacy services'), 'Amended pharmacy services text must not appear');
    assert.ok(!plainText.includes('£85.00 per validated client'), 'Higher tariff must not appear');
    assert.ok(!plainText.includes('UK GDPR'), 'UK GDPR reference must not appear');
    assert.ok(!plainText.includes('thirty (30) Working Days'), 'Extended notice floor must not appear');
}

// ----------------------------------------------------------------------------
// Test 5: Full Package Pipeline via openDocx & Untouched Parts Preservation
// Verifies opening the raw .docx buffer, applying the mixed batch, generating
// the output package with comments.xml and relationship updates, validating OPC
// integrity, and checking that untouched parts remain byte-identical.
// ----------------------------------------------------------------------------
async function testPackagePipeline(sourceBuffer) {
    const originalEntries = unzipEntries(sourceBuffer);
    const doc = openDocx(sourceBuffer);

    // Verify inspection
    const inspection = doc.inspect();
    assert.equal(inspection.status, 'ok');
    assert.equal(inspection.paragraphs.length, 533, 'Source contract must contain 533 paragraphs');

    const result = await doc.applyOperations(LEGAL_MARKUP_OPERATIONS, {
        author: 'Healthcare Counsel',
        atomic: true,
        validate: true
    });
    assert.equal(result.written, true, 'Package operation must succeed and be written');
    assert.ok(!result.rolledBack, 'Successful package operation must not roll back');

    // Verify artifacts modified
    assert.ok(result.artifactsChanged.includes('word/document.xml'));
    assert.ok(result.artifactsChanged.includes('word/comments.xml'));
    assert.ok(result.artifactsChanged.includes('word/_rels/document.xml.rels'));
    assert.ok(result.artifactsChanged.includes('[Content_Types].xml'));

    // Verify output ZIP structure
    const outputBuffer = result.toBuffer();
    assert.ok(outputBuffer.length > 0, 'Output buffer must not be empty');
    const outputEntries = unzipEntries(outputBuffer);

    assert.ok(outputEntries.has('word/comments.xml'), 'Output package must contain word/comments.xml');
    assert.ok(outputEntries.get('word/comments.xml').length > 3000, 'Comments XML must contain substantial content');

    // Structural OPC Package Validation
    await validateDocxPackage(new MemoryZip(outputEntries));

    // Byte-for-byte preservation of untouched parts
    const untouchedParts = [
        'word/header1.xml',
        'word/footer1.xml',
        'word/footnotes.xml',
        'word/endnotes.xml',
        'word/settings.xml',
        'word/theme/theme1.xml',
        'word/webSettings.xml'
    ];

    for (const partName of untouchedParts) {
        if (originalEntries.has(partName)) {
            assert.ok(outputEntries.has(partName), `Output package must preserve ${partName}`);
            assert.deepEqual(
                outputEntries.get(partName),
                originalEntries.get(partName),
                `Untouched part ${partName} must remain byte-identical`
            );
        }
    }
}

// ----------------------------------------------------------------------------
// Test 6: Atomic Rollback Guarantee on Long Legal Document
// Verifies that if any single operation in the batch fails (e.g. an unmatchable
// comment anchor), the entire 10,770-word package rolls back byte-for-byte.
// ----------------------------------------------------------------------------
async function testAtomicRollback(sourceBuffer) {
    const doc = openDocx(sourceBuffer);

    const operationsWithFailingAnchor = [
        ...LEGAL_MARKUP_OPERATIONS.slice(0, 5),
        {
            type: 'comment',
            target: { exactText: 'Southampton City Council of Civic Centre, Southampton, SO14 7LT (“the Council”)' },
            textToComment: 'NON_EXISTENT_ANCHOR_TEXT_FOR_FAILURE_TEST',
            commentContent: 'This comment anchor does not exist and must trigger atomic rollback.',
            author: 'Healthcare Counsel'
        },
        ...LEGAL_MARKUP_OPERATIONS.slice(5)
    ];

    const result = await doc.applyOperations(operationsWithFailingAnchor, {
        author: 'Healthcare Counsel',
        atomic: true
    });

    assert.equal(result.written, false, 'Failed batch must not be written');
    assert.equal(result.rolledBack, true, 'Failed batch must trigger atomic rollback');
    assert.equal(result.status, 'error', 'Result status must be error');

    // Buffer must be byte-for-byte identical to the original input buffer
    assert.deepEqual(result.toBuffer(), sourceBuffer, 'Rolled-back package buffer must be byte-identical to source');
}

// ----------------------------------------------------------------------------
// Main Execution
// ----------------------------------------------------------------------------
async function main() {
    console.log('Loading SuperDoc Smoking Cessation Contract fixture...');
    const sourceBuffer = await loadSourceDocxBuffer();
    const sourceZip = unzipEntries(sourceBuffer);
    const docXml = sourceZip.get('word/document.xml').toString('utf8');

    console.log('  1. Running preflight on 28-operation legal markup suite...');
    await testPreflightLegalBatch(docXml);

    console.log('  2. Running engine atomic batch application & commentsXml generation...');
    const engineResult = await testEngineAtomicBatch(docXml);

    console.log('  3. Running differential Accept All verification...');
    await testDifferentialAcceptAll(engineResult.documentXml);

    console.log('  4. Running differential Reject All verification...');
    await testDifferentialRejectAll(engineResult.documentXml);

    console.log('  5. Running complete package pipeline via openDocx with untouched-part preservation...');
    await testPackagePipeline(sourceBuffer);

    console.log('  6. Running atomic rollback guarantee verification on long legal document...');
    await testAtomicRollback(sourceBuffer);

    console.log('PASS: smoking_cessation_mixed_batch_tests.mjs - all 6 legal mixed batch tests passed.');
}

await main();
