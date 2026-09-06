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
export const CORPUS_ID = 'c5bb43ede56775b6d469103dc389a6bcb14b0265460b7cec7528564621d25af1';
export const EXPECTED_SHA256 = '665a9f4ead98f89b091430ba9ef67e39533ae3362952fc20938039dc7590094d';
export const DOWNLOAD_URL = `https://docxcorp.us/documents/${CORPUS_ID}.docx`;
export const LOCAL_CORPUS_PATH = join(repoRoot, 'tmp', 'superdoc-corpus', `${CORPUS_ID}.docx`);

export const AUTHOR_1 = 'BCHD Lead Agency Counsel';
export const AUTHOR_2 = 'MOHS Counterparty Counsel';

/**
 * Loads the source DOCX buffer from local corpus cache or downloads it from CDN.
 */
export async function loadSourceDocxBuffer() {
    if (existsSync(LOCAL_CORPUS_PATH)) {
        const buffer = readFileSync(LOCAL_CORPUS_PATH);
        const digest = createHash('sha256').update(buffer).digest('hex');
        if (digest === EXPECTED_SHA256) {
            return buffer;
        }
    }

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
// Attachment 4 Content: Joint Agency Public Announcement & Communications Protocol
// Complete standalone document with 2 markdown tables, 4-item policy list,
// and executive statements.
// ----------------------------------------------------------------------------
export const ATTACHMENT_4_BODY = `Date

# ATTACHMENT 4 – JOINT AGENCY PUBLIC ANNOUNCEMENT & COMMUNICATIONS PROTOCOL

## FOR IMMEDIATE RELEASE: BALTIMORE CITY LAUNCHES COMPREHENSIVE INTERAGENCY HEALTH & HOUSING COLLABORATIVE

BALTIMORE, MD — The Baltimore City Health Department (BCHD) and the Mayor's Office of Homeless Services (MOHS) today announced the formal launch of a comprehensive interagency collaborative designed to unify street-level clinical medicine, behavioral health intervention, and rapid housing navigation across Baltimore City.

Under this historic partnership, multidisciplinary clinical outreach teams will be paired directly with homeless service coordinators, ensuring that individuals experiencing unsheltered homelessness receive immediate healthcare triage and a direct pathway to permanent supportive housing.

"Health equity begins by meeting our most vulnerable residents where they are," stated the Baltimore City Health Commissioner. "By uniting public health clinical services with supportive housing resources, we are removing long-standing structural silos and establishing a compassionate, coordinated system of care."

"Housing is healthcare," stated the Director of the Mayor's Office of Homeless Services. "This agreement provides the stable framework and clinical resources required to assist residents in moving rapidly from street-level vulnerability to lasting stabilization."

| Agency / Role | Designated Contact | Title / Division | Embargo Deadline | Release Channel |
| --- | --- | --- | --- | --- |
| BCHD Lead PIO | Dr. Sarah Jenkins | Chief Medical Communications Officer | Oct 1, 2026 09:00 EST | City Press Wire |
| MOHS Liaison | Marcus Vance | Director of Public Information | Oct 1, 2026 09:00 EST | Mayoral Briefing |
| City Solicitor | Elena Rostova | Special Assistant City Solicitor | Oct 1, 2026 08:30 EST | Legal Review |

## PROTOCOL FOR PUBLIC STATEMENTS AND MEDIA INQUIRIES:

1. Joint Clearance Protocol: All official press releases, public announcements, and media advisories must receive written concurrence from both BCHD and MOHS public information officers prior to public dissemination.
2. Designee Restrictions: No employee, agent, or contracted provider may deliver on-the-record public statements regarding the collaborative without prior authorization from the designated agency liaisons.
3. Rapid Crisis Escalation: Inquiries regarding emergency health incidents, cold-weather shelter activations, or public encampments must be coordinated through the City Joint Information Center (JIC) within two (2) hours.
4. Annual Performance Transparency: The collaborative shall publish an annual joint outcomes dashboard documenting client engagements, medical encounters, and permanent housing placements.

| Milestone Phase | Target Implementation Period | Key Deliverables & Clinical Outputs | Responsible Agency | Target KPI / Success Benchmark |
| --- | --- | --- | --- | --- |
| Phase 1: Operational Launch | Q3 2026 (July – Sept) | Mobilize 4 street medicine clinical units and establish shared EHR platform | BCHD / MOHS | 500 unduplicated individuals engaged |
| Phase 2: Housing Navigation | Q4 2026 (Oct – Dec) | Transition eligible shelter residents into permanent supportive housing | MOHS | 150 placements with 90-day retention |
| Phase 3: Annual Audit & Review | Q2 2027 (April – June) | Complete comprehensive federal grant compliance review and KPI audit | Joint Oversight Panel | 100% compliance with Uniform Guidance |

This Attachment 4 constitutes an integral component of the Interagency Agreement between BCHD and MOHS, governing all public announcements, external relations, and performance reporting obligations throughout the Agreement Term.`;

// ----------------------------------------------------------------------------
// Pass 1: Author 1 (BCHD Lead Agency Counsel)
// 24 operations: Section headers fiddled into lists, Attachment 3 subbullets,
// Attachment 4 appended document with tables & lists, contract terms & comments.
// ----------------------------------------------------------------------------
// ----------------------------------------------------------------------------
// Pass 1: Author 1 (BCHD Lead Agency Counsel) Initial Proposals
// Formal section headers are maintained as clean headers (NOT bullets).
// Section C Term initial proposal: 36 months.
// Section D.3 Disbursement initial proposal: 45 calendar days.
// Section E Termination initial proposal: 90 calendar days.
// ----------------------------------------------------------------------------
export const AUTHOR_1_PASS_1_OPERATIONS = [
    // --- 1. Parties & Recitals ---
    {
        type: 'replace',
        target: { exactText: 'NAME OF DEPARTMENT', index: 5 },
        modified: 'Baltimore City Health Department',
        author: AUTHOR_1
    },
    {
        type: 'replace',
        target: { exactText: 'NAME OF DEPARTMENT', index: 7 },
        modified: "Mayor's Office of Homeless Services",
        author: AUTHOR_1
    },
    {
        type: 'replace',
        target: {
            exactText: 'THIS INTER-AGENCY AGREEMENT (this “Agreement”) made this ______ day of ____________, 20__ between two agencies of the Mayor and City Council of Baltimore (hereinafter called the “City”), namely the NAME OF DEPARTMENT (hereinafter called the "DEPARTMENT 1") and the NAME OF DEPARTMENT (hereinafter called “DEPARTMENT 2”).',
            index: 10
        },
        modified: 'THIS INTER-AGENCY AGREEMENT (this “Agreement”) made this 1st day of July, 2026 between two agencies of the Mayor and City Council of Baltimore (hereinafter called the “City”), namely the Baltimore City Health Department (hereinafter called "BCHD" or "DEPARTMENT 1") and the Mayor\'s Office of Homeless Services (hereinafter called "MOHS" or "DEPARTMENT 2").',
        author: AUTHOR_1
    },
    {
        type: 'replace',
        target: {
            exactText: 'WHEREAS, this Agreement represents a cooperative effort to utilize funds from a grant/appropriation to DEPARTMENT 1 as identified in Attachment 1 for the purposes of SPECIFY PURPOSE; and ',
            index: 12
        },
        modified: 'WHEREAS, this Agreement represents a cooperative effort to utilize funds from a grant/appropriation to DEPARTMENT 1 as identified in Attachment 1 for the purposes of establishing an integrated street medicine and housing navigation collaborative for vulnerable unsheltered Baltimore residents; and',
        author: AUTHOR_1
    },
    {
        type: 'replace',
        target: {
            exactText: 'WHEREAS, DEPARTMENT 1 shall provide to DEPARTMENT 2 funding as described in this Agreement to facilitate the payment of DOLLAR AMOUNT for the DEPARTMENT 2’s SPECIFY PURPOSE; and',
            index: 14
        },
        modified: 'WHEREAS, DEPARTMENT 1 shall provide to DEPARTMENT 2 funding as described in this Agreement to facilitate the payment of TWO MILLION FOUR HUNDRED FIFTY THOUSAND DOLLARS ($2,450,000.00) for the DEPARTMENT 2’s coordinated shelter outreach and clinical stabilization services; and',
        author: AUTHOR_1
    },
    {
        type: 'comment',
        target: {
            exactText: 'WHEREAS, DEPARTMENT 1 shall provide to DEPARTMENT 2 funding as described in this Agreement to facilitate the payment of DOLLAR AMOUNT for the DEPARTMENT 2’s SPECIFY PURPOSE; and',
            index: 14
        },
        textToComment: 'DOLLAR AMOUNT',
        commentContent: 'Funding allocation reflects combined SAMHSA federal award ($1.8M) plus local matching municipal funds ($650K).',
        author: AUTHOR_1
    },

    // --- 2. Formal Section Headers (Maintained as Clean Headers, NOT Bullets) ---
    {
        type: 'replace',
        target: { exactText: 'A.\tPURPOSE', index: 18 },
        modified: 'A.\tPURPOSE AND INTERAGENCY ALIGNMENT',
        author: AUTHOR_1
    },
    {
        type: 'replace',
        target: {
            exactText: 'DEPARTMENT 1 hereby agrees to provide funds, in the amounts set forth herein, to DEPARTMENT 2 in order to SPECIFY PURPOSE.',
            index: 20
        },
        modified: 'DEPARTMENT 1 hereby agrees to provide funds, in the amounts set forth herein, to DEPARTMENT 2 in order to deliver multidisciplinary field healthcare, trauma-informed clinical casework, and rapid permanent supportive housing navigation.',
        author: AUTHOR_1
    },
    {
        type: 'replace',
        target: { exactText: 'B.\tSCOPE OF SERVICES', index: 22 },
        modified: 'B.\tSCOPE OF COORDINATED SERVICES AND RESPONSIBILITIES',
        author: AUTHOR_1
    },
    {
        type: 'replace',
        target: { exactText: 'C.\tTERM ', index: 26 },
        modified: 'C.\tTERM AND PERFORMANCE DURATION',
        author: AUTHOR_1
    },
    // BCHD initial proposal: 36 months (altered in Pass 1b via same-author merge)
    {
        type: 'replace',
        target: {
            exactText: 'The term of this Agreement shall be DURATION beginning on      , 20   and ending on                     , 20   , unless terminated earlier in accordance with this Agreement.',
            index: 28
        },
        modified: 'The term of this Agreement shall be thirty-six (36) months beginning on July 1, 2026 and ending on June 30, 2029, unless terminated earlier in accordance with this Agreement.',
        author: AUTHOR_1
    },
    {
        type: 'replace',
        target: { exactText: 'D. \tCOMPENSATION AND METHOD OF PAYMENT ', index: 30 },
        modified: 'D. \tCOMPENSATION, FISCAL ACCOUNTING, AND INVOICING',
        author: AUTHOR_1
    },
    {
        type: 'replace',
        target: {
            exactText: '1.\tSubject to the availability of funding, DEPARTMENT 1 shall provide funds by journal entry to the DEPARTMENT 2 for payment of SPECIFY PURPOSE described hereunder, in an amount not to exceed        DOLLARS ($    .   ) for the term.',
            index: 32
        },
        modified: '1.\tSubject to the availability of funding, DEPARTMENT 1 shall provide funds by journal entry to the DEPARTMENT 2 for payment of coordinated outreach and stabilization services described hereunder, in an amount not to exceed TWO MILLION FOUR HUNDRED FIFTY THOUSAND DOLLARS ($2,450,000.00) for the term.',
        author: AUTHOR_1
    },
    // BCHD initial proposal: 45 calendar days (altered in Pass 1b via same-author merge)
    {
        type: 'replace',
        target: {
            exactText: '3.\tDEPARTMENT 1 shall only pay invoices to if the required supporting documentation accompanies each invoice. ',
            index: 36
        },
        modified: '3.\tDEPARTMENT 1 shall process journal entries and disburse funds to DEPARTMENT 2 within forty-five (45) calendar days of receipt of verified invoices and required supporting documentation.',
        author: AUTHOR_1
    },
    {
        type: 'replace',
        target: { exactText: 'E.\tTERMINATION', index: 38 },
        modified: 'E.\tEARLY TERMINATION PROTOCOLS AND TRANSITION',
        author: AUTHOR_1
    },
    // BCHD initial proposal on Section E: 90 calendar days
    {
        type: 'replace',
        target: {
            exactText: 'Either party may terminate this Agreement by giving to the other party written notification thereof at least sixty (60) days prior to termination.  DEPARTMENT 1 shall provide funding to the DEPARTMENT 2 for services under this Agreement through the date of termination.  Upon termination, the parties hereto agree that all reports and supporting documentation required for services rendered shall be provided by the DEPARTMENT 2 to DEPARTMENT 1.  ',
            index: 40
        },
        modified: 'Either party may terminate this Agreement without cause by giving to the other party written notification thereof at least ninety (90) calendar days prior to termination.  DEPARTMENT 1 shall provide funding to the DEPARTMENT 2 for services under this Agreement through the date of termination.  Upon termination, the parties hereto agree that all reports, clinical care transition summaries, and supporting documentation required for services rendered shall be provided by the DEPARTMENT 2 to DEPARTMENT 1 within fifteen (15) business days.',
        author: AUTHOR_1
    },
    {
        type: 'comment',
        target: {
            exactText: 'Either party may terminate this Agreement by giving to the other party written notification thereof at least sixty (60) days prior to termination.  DEPARTMENT 1 shall provide funding to the DEPARTMENT 2 for services under this Agreement through the date of termination.  Upon termination, the parties hereto agree that all reports and supporting documentation required for services rendered shall be provided by the DEPARTMENT 2 to DEPARTMENT 1.  ',
            index: 40
        },
        textToComment: 'sixty (60) days',
        commentContent: '90-day notice is essential to ensure vulnerable clinical clients have sufficient runway for uninterrupted medication management and shelter placement.',
        author: AUTHOR_1
    },
    {
        type: 'replace',
        target: { exactText: 'F.\tMODIFICATIONS', index: 42 },
        modified: 'F.\tMODIFICATIONS AND WRITTEN AMENDMENTS',
        author: AUTHOR_1
    },
    {
        type: 'replace',
        target: { exactText: 'G.\tGENERAL PROVISIONS AND CONDITIONS', index: 46 },
        modified: 'G.\tGENERAL PROVISIONS, STATUTORY ASSURANCES, AND CONDITIONS',
        author: AUTHOR_1
    },

    // --- 3. Attachment 3 Subbullets ---
    {
        type: 'replace',
        target: {
            exactText: 'Shall comply with the Health Insurance Portability and Accountability Act (HIPAA) of 1996, 42 U.S.C. 1320d et seq., which governs the protection of individually identifiable health information.',
            index: 134
        },
        modified: 'Shall comply with the Health Insurance Portability and Accountability Act (HIPAA) of 1996, 42 U.S.C. 1320d et seq., which governs the protection of individually identifiable health information.\n  * Execution of standard Business Associate Agreement (BAA) within 15 calendar days of award.\n  * Mandatory AES-256 encryption for all Electronic Protected Health Information (ePHI) at rest and TLS 1.3 in transit.\n  * Immediate notification of any security incident or breach within 24 hours to BCHD Privacy Officer.',
        author: AUTHOR_1
    },
    {
        type: 'replace',
        target: {
            exactText: 'Shall comply with all applicable nondiscrimination statutes, including but not limited to:',
            index: 137
        },
        modified: 'Shall comply with all applicable nondiscrimination statutes, including but not limited to:\n  * Title VI Language Access Plan (LAP) providing qualified medical interpreter services at no cost.\n  * Section 504 auxiliary aids and physical accessibility for persons with disabilities.\n  * Vital document translations in top non-English languages spoken by eligible Baltimore residents.',
        author: AUTHOR_1
    },
    {
        type: 'replace',
        target: {
            exactText: 'Shall comply with all applicable audit requirements of the Office of Management and Budget (OMB), including but not limited to OMB Circular A-133.  ',
            index: 144
        },
        modified: 'Shall comply with all applicable audit requirements of the Office of Management and Budget (OMB), including but not limited to OMB Circular A-133.  \n  * Annual Single Audit submission if federal expenditure exceeds $750,000 threshold under 2 CFR 200 Subpart F.\n  * Submission of Federal Audit Clearinghouse (FAC) reporting package within 30 days of audit report receipt.\n  * Corrective Action Plan (CAP) implementation timeline not to exceed 60 calendar days for any audit findings.',
        author: AUTHOR_1
    },

    // --- 4. Attachment 4: Brand New Document Appended with Tables, Lists, and Paragraphs ---
    {
        type: 'replace',
        target: { exactText: 'Date', index: 172 },
        modified: ATTACHMENT_4_BODY,
        structuredContent: true,
        author: AUTHOR_1
    }
];

export const AUTHOR_1_OPERATIONS = AUTHOR_1_PASS_1_OPERATIONS;

// ----------------------------------------------------------------------------
// Pass 1b: Author 1 (BCHD Lead Agency Counsel) Revision Alteration
// Demonstrates existingRevisions: 'merge-same-author' where an author alters their
// own earlier redlines on Section C (Term: 36m -> 24m) and Section D.3 (Payment: 45d -> 30d).
// ----------------------------------------------------------------------------
export const AUTHOR_1_ALTERATION_OPERATIONS = [
    {
        type: 'replace',
        target: {
            exactText: 'The term of this Agreement shall be thirty-six (36) months beginning on July 1, 2026 and ending on June 30, 2029, unless terminated earlier in accordance with this Agreement.'
        },
        modified: 'The term of this Agreement shall be twenty-four (24) months beginning on July 1, 2026 and ending on June 30, 2028, unless terminated earlier in accordance with this Agreement.',
        author: AUTHOR_1
    },
    {
        type: 'replace',
        target: {
            exactText: '3.\tDEPARTMENT 1 shall process journal entries and disburse funds to DEPARTMENT 2 within forty-five (45) calendar days of receipt of verified invoices and required supporting documentation.'
        },
        modified: '3.\tDEPARTMENT 1 shall process journal entries and disburse funds to DEPARTMENT 2 within thirty (30) calendar days of receipt of verified invoices and required supporting documentation.',
        author: AUTHOR_1
    }
];

// ----------------------------------------------------------------------------
// Pass 2: Author 2 (MOHS Counterparty Counsel)
// Weaves the two counsels together:
// - MOHS places counter-comments on BCHD clauses.
// - MOHS counters on the SAME provision (Section E Termination) with accept-all-first.
// - MOHS redlines Section G, Notices, Signatures, Attachments 1 & 2.
// ----------------------------------------------------------------------------
export const AUTHOR_2_OPERATIONS = [
    // --- 1. Counter-Comments on Author 1's Key Clauses ---
    {
        type: 'comment',
        target: {
            exactText: '3.\tDEPARTMENT 1 shall process journal entries and disburse funds to DEPARTMENT 2 within thirty (30) calendar days of receipt of verified invoices and required supporting documentation.'
        },
        textToComment: 'thirty (30) calendar days',
        commentContent: 'Counter-proposal: Request 15 business days instead of 30 calendar days. Community outreach subcontractors have bi-weekly payroll requirements and cannot absorb 30-day municipal disbursement lag.',
        author: AUTHOR_2
    },
    {
        type: 'comment',
        target: {
            exactText: 'The term of this Agreement shall be twenty-four (24) months beginning on July 1, 2026 and ending on June 30, 2028, unless terminated earlier in accordance with this Agreement.'
        },
        textToComment: 'twenty-four (24) months',
        commentContent: 'Agreed on 24-month base term, but MOHS requests inclusion of two optional one-year renewal terms subject to mutual written agreement and Board of Estimates approval.',
        author: AUTHOR_2
    },
    {
        type: 'comment',
        target: {
            exactText: 'Either party may terminate this Agreement without cause by giving to the other party written notification thereof at least ninety (90) calendar days prior to termination.  DEPARTMENT 1 shall provide funding to the DEPARTMENT 2 for services under this Agreement through the date of termination.  Upon termination, the parties hereto agree that all reports, clinical care transition summaries, and supporting documentation required for services rendered shall be provided by the DEPARTMENT 2 to DEPARTMENT 1 within fifteen (15) business days.'
        },
        textToComment: 'ninety (90) calendar days',
        commentContent: 'Counter-proposal: 90 days creates extended fiscal liability if federal award amounts are adjusted mid-year. MOHS proposes 45 calendar days with mutual wind-down cost reimbursement.',
        author: AUTHOR_2
    },

    // --- 2. Counter-Redline on the SAME Provision (Section E Termination) ---
    // MOHS counters on Section E by accepting BCHD's working text and proposing 45 days + wind-down costs
    {
        type: 'replace',
        target: {
            exactText: 'Either party may terminate this Agreement without cause by giving to the other party written notification thereof at least ninety (90) calendar days prior to termination.  DEPARTMENT 1 shall provide funding to the DEPARTMENT 2 for services under this Agreement through the date of termination.  Upon termination, the parties hereto agree that all reports, clinical care transition summaries, and supporting documentation required for services rendered shall be provided by the DEPARTMENT 2 to DEPARTMENT 1 within fifteen (15) business days.'
        },
        modified: 'Either party may terminate this Agreement without cause by giving to the other party written notification thereof at least forty-five (45) calendar days prior to termination, subject to mutual reimbursement of uncancelable wind-down costs.  DEPARTMENT 1 shall provide funding to the DEPARTMENT 2 for services under this Agreement through the date of termination.  Upon termination, the parties hereto agree that all reports, clinical care transition summaries, and supporting documentation required for services rendered shall be provided by the DEPARTMENT 2 to DEPARTMENT 1 within fifteen (15) business days.',
        author: AUTHOR_2,
        existingRevisions: 'accept-all-first'
    },

    // --- 3. Counterparty Redlines on General Provisions & Dispute Resolution ---
    {
        type: 'replace',
        target: {
            exactText: '1.\tDEPARTMENT 2 shall comply with all federal, state and local laws, ordinances, rules, regulations and federal or state grant requirements related to the funding under this Agreement.',
            index: 54
        },
        modified: '1.\tDEPARTMENT 2 shall comply with all federal, state and local laws, ordinances, rules, regulations and federal or state grant requirements related to the funding under this Agreement, provided that in the event of an alleged compliance default, DEPARTMENT 1 shall provide written notice and a thirty (30) calendar day opportunity to cure prior to initiating any funding suspension.',
        author: AUTHOR_2
    },
    {
        type: 'replace',
        target: {
            exactText: '2.\tAny funds advanced to DEPARTMENT 2 prior to the execution of this Agreement are subject to the terms and conditions of this Agreement.',
            index: 56
        },
        modified: '2.\tAny allowable pre-award grant funds advanced or expended by DEPARTMENT 2 for authorized startup costs prior to the execution of this Agreement are subject to the terms and conditions of this Agreement and reimbursable under the approved grant budget.',
        author: AUTHOR_2
    },
    {
        type: 'replace',
        target: {
            exactText: ' \t5.\tThis Agreement constitutes the entire, full and final understanding between the parties hereto and neither party shall be bound by any representations, statements, promises or agreements not expressly set forth herein. ',
            index: 73
        },
        modified: ' \t5.\tThis Agreement constitutes the entire, full and final understanding between the parties hereto and neither party shall be bound by any representations, statements, promises or agreements not expressly set forth herein. The parties agree to conduct an annual joint performance review in March of each contract year to assess program milestones and budgetary alignment.',
        author: AUTHOR_2
    },

    // --- 4. Counterparty Redlines: Notices Block for Department 2 ---
    {
        type: 'replace',
        target: { exactText: '_______________________________', occurrence: 5 },
        modified: "Mayor's Office of Homeless Services",
        author: AUTHOR_2
    },
    {
        type: 'replace',
        target: { exactText: '_______________________________', occurrence: 5 },
        modified: 'Attention: Executive Director & General Counsel',
        author: AUTHOR_2
    },
    {
        type: 'replace',
        target: { exactText: '_______________________________', occurrence: 5 },
        modified: '7 E. Redwood Street, 9th Floor, Baltimore, MD 21202',
        author: AUTHOR_2
    },
    {
        type: 'replace',
        target: { exactText: '_______________________________', occurrence: 5 },
        modified: 'Email: mohs.legal@baltimorecity.gov | Fax: (410) 396-8182',
        author: AUTHOR_2
    },

    // --- 5. Counterparty Redlines: Signatory Block ---
    {
        type: 'replace',
        target: { exactText: '\t\t\t\t\t\t\tSIGNATORY', occurrence: 2 },
        modified: '\t\t\t\t\t\t\tErnestina Del-Sarto, Executive Director',
        author: AUTHOR_2
    },
    {
        type: 'replace',
        target: { exactText: '\t\t\t\t\t\t\tDEPARTMENT 2', index: 106 },
        modified: "\t\t\t\t\t\t\tMayor's Office of Homeless Services (MOHS)",
        author: AUTHOR_2
    },

    // --- 6. Counterparty Redlines: Attachment 1 Scope and Budget ---
    {
        type: 'replace',
        target: { exactText: 'Program Scope:   ' },
        modified: 'Program Scope: MOHS shall oversee four (4) specialized mobile outreach units providing trauma-informed street outreach, coordinated entry intake, temporary emergency lodging placement, and clinical care navigation for unsheltered individuals across Baltimore City priority encampments.',
        author: AUTHOR_2
    },
    {
        type: 'replace',
        target: { exactText: 'Budget:' },
        modified: 'Budget: Total Allocation: $2,450,000 (Personnel & Outreach Staff: $1,420,000; Emergency Lodging Subsidies & Respite Beds: $680,000; Transportation & Medical Supplies: $150,000; Administrative Overhead: $200,000).',
        author: AUTHOR_2
    },

    // --- 7. Counterparty Redlines: Attachment 2 Grant Information ---
    {
        type: 'replace',
        target: { exactText: 'Grant Name:  ' },
        modified: 'Grant Name: SAMHSA Projects for Assistance in Transition from Homelessness (PATH) Collaborative Grant',
        author: AUTHOR_2
    },
    {
        type: 'replace',
        target: { exactText: 'CFDA Number [if Federal grant]:' },
        modified: 'CFDA Number [if Federal grant]: CFDA 93.150 (Comprehensive Mental Health and Substance Abuse Services)',
        author: AUTHOR_2
    },
    {
        type: 'replace',
        target: { exactText: 'Grant Number: ' },
        modified: 'Grant Number: 1-H79-SM084920-01-BCHD/MOHS',
        author: AUTHOR_2
    }
];

// ----------------------------------------------------------------------------
// Pass 3: Author 1 (BCHD Lead Agency Counsel) Final Compromise on Section E
// BCHD responds back on the SAME provision (Section E Termination) to reach
// final compromise: 60 calendar days with 45-day emergency shelter placement protocol.
// ----------------------------------------------------------------------------
export const AUTHOR_1_FINAL_OPERATIONS = [
    {
        type: 'comment',
        target: {
            exactText: 'Either party may terminate this Agreement without cause by giving to the other party written notification thereof at least forty-five (45) calendar days prior to termination, subject to mutual reimbursement of uncancelable wind-down costs.  DEPARTMENT 1 shall provide funding to the DEPARTMENT 2 for services under this Agreement through the date of termination.  Upon termination, the parties hereto agree that all reports, clinical care transition summaries, and supporting documentation required for services rendered shall be provided by the DEPARTMENT 2 to DEPARTMENT 1 within fifteen (15) business days.'
        },
        textToComment: 'forty-five (45) calendar days',
        commentContent: 'BCHD compromise proposal: We can accept a 60-day notice with 45-day clinical transition safeguard and wind-down cost coverage.',
        author: AUTHOR_1
    },
    {
        type: 'replace',
        target: {
            exactText: 'Either party may terminate this Agreement without cause by giving to the other party written notification thereof at least forty-five (45) calendar days prior to termination, subject to mutual reimbursement of uncancelable wind-down costs.  DEPARTMENT 1 shall provide funding to the DEPARTMENT 2 for services under this Agreement through the date of termination.  Upon termination, the parties hereto agree that all reports, clinical care transition summaries, and supporting documentation required for services rendered shall be provided by the DEPARTMENT 2 to DEPARTMENT 1 within fifteen (15) business days.'
        },
        modified: 'Either party may terminate this Agreement without cause by giving to the other party written notification thereof at least sixty (60) calendar days prior to termination, with forty-five (45) days for emergency shelter placements and mutual reimbursement of uncancelable wind-down costs.  DEPARTMENT 1 shall provide funding to the DEPARTMENT 2 for services under this Agreement through the date of termination.  Upon termination, the parties hereto agree that all reports, clinical care transition summaries, and supporting documentation required for services rendered shall be provided by the DEPARTMENT 2 to DEPARTMENT 1 within fifteen (15) business days.',
        author: AUTHOR_1,
        existingRevisions: 'accept-all-first'
    }
];

// ----------------------------------------------------------------------------
// Test 1: Preflight Verification on Author 1 Operations
// ----------------------------------------------------------------------------
async function testPreflight(docXml) {
    const preflight = preflightOperations(docXml, AUTHOR_1_PASS_1_OPERATIONS, AUTHOR_1);
    assert.equal(preflight.status, 'ok', 'Author 1 preflight must pass with status ok');
    assert.equal(preflight.issues?.length ?? 0, 0, 'Author 1 preflight must report zero issues');
}

// ----------------------------------------------------------------------------
// Test 2: Engine Sequential Multi-Author Execution
// Tests:
// - Pass 1: BCHD initial draft with clean headers (NO bullet lists).
// - Pass 1b: BCHD alteration (same-author merge on Term and Payment).
// - Multi-user conflict protection: MOHS refused without accept-all-first.
// - Pass 2: MOHS counterparty redlines on Section E (same provision) + others.
// - Pass 3: BCHD final compromise on Section E (same provision).
// ----------------------------------------------------------------------------
async function testEngineMultiAuthor(docXml) {
    // Pass 1: Author 1 (BCHD Lead Agency Counsel) Initial Draft
    const pass1 = await applyOperationsToDocumentXml(docXml, AUTHOR_1_PASS_1_OPERATIONS, AUTHOR_1);
    assert.equal(pass1.rolledBack, undefined, 'Pass 1 must not be rolled back');
    assert.equal(pass1.hasChanges, true, 'Pass 1 must produce changes');
    for (const r of pass1.results) {
        assert.equal(r.status, 'applied', `Pass 1 operation ${r.index} must be applied`);
    }

    // Structural checks for Pass 1:
    // Section headers A-G must NOT be bullet lists (no numPr on header paragraphs)
    const pass1PlainText = ingestWordOoxmlToPlainText(pass1.documentXml);
    assert.ok(pass1PlainText.includes('PURPOSE AND INTERAGENCY ALIGNMENT'), 'Header A must have new title');
    assert.ok(pass1PlainText.includes('SCOPE OF COORDINATED SERVICES AND RESPONSIBILITIES'), 'Header B must have new title');
    assert.ok(pass1PlainText.includes('TERM AND PERFORMANCE DURATION'), 'Header C must have new title');
    assert.ok(pass1PlainText.includes('COMPENSATION, FISCAL ACCOUNTING, AND INVOICING'), 'Header D must have new title');
    assert.ok(pass1PlainText.includes('EARLY TERMINATION PROTOCOLS AND TRANSITION'), 'Header E must have new title');
    assert.ok(pass1PlainText.includes('MODIFICATIONS AND WRITTEN AMENDMENTS'), 'Header F must have new title');
    assert.ok(pass1PlainText.includes('GENERAL PROVISIONS, STATUTORY ASSURANCES, AND CONDITIONS'), 'Header G must have new title');

    // Verify Header A paragraph is not a bullet list (numId must be 0 or absent, never positive)
    const alignIdx = pass1.documentXml.indexOf('INTERAGENCY ALIGNMENT');
    const headerAParaStart = pass1.documentXml.lastIndexOf('<w:p', alignIdx);
    const headerAParaEnd = pass1.documentXml.indexOf('</w:p>', alignIdx) + 6;
    const headerAPara = pass1.documentXml.slice(headerAParaStart, headerAParaEnd);
    assert.ok(!headerAPara.match(/<w:numId\s+w:val="[1-9]\d*"/), 'Header A must NOT have positive numId (must remain clean header, not bullet list)');

    // Attachment 3 nested subbullets: verify w:ilvl w:val="1"
    const subBulletMatches = pass1.documentXml.match(/<w:ilvl\s+w:val="1"\s*\/>/g) || [];
    assert.ok(subBulletMatches.length >= 6, `Expected at least 6 nested level-1 subbullets in Attachment 3, found ${subBulletMatches.length}`);

    // Attachment 4: verify newly appended document sections present
    assert.ok(pass1.documentXml.includes('ATTACHMENT 4 – JOINT AGENCY PUBLIC ANNOUNCEMENT'), 'Attachment 4 title must be present');
    assert.ok(pass1.documentXml.includes('PROTOCOL FOR PUBLIC STATEMENTS AND MEDIA INQUIRIES'), 'Attachment 4 protocols must be present');
    assert.ok(pass1.documentXml.includes('Joint Clearance Protocol:'), 'Attachment 4 policy list must be present');
    assert.ok(pass1.documentXml.includes('Phase 1: Operational Launch'), 'Attachment 4 milestone content must be present');
    const sourceTableCount = (docXml.match(/<w:tbl(?:\s|>)/g) || []).length;
    const pass1TableCount = (pass1.documentXml.match(/<w:tbl(?:\s|>)/g) || []).length;
    assert.equal(pass1TableCount, sourceTableCount + 2, 'Attachment 4 must add two real Word tables');
    assert.ok(pass1.documentXml.includes('<w:pStyle w:val="Heading1"'), 'Attachment 4 title must be a Word heading');
    assert.ok(pass1.documentXml.includes('<w:pStyle w:val="Heading2"'), 'Attachment 4 section labels must be Word headings');
    assert.ok(!pass1.documentXml.includes('| Agency / Role |'), 'Attachment 4 table syntax must not remain literal pipe text');

    // Pass 1b: BCHD alters its proposals (testing same-author merge on Section C Term & Section D.3 Disbursement)
    const pass1b = await applyOperationsToDocumentXml(pass1.documentXml, AUTHOR_1_ALTERATION_OPERATIONS, AUTHOR_1);
    assert.equal(pass1b.rolledBack, undefined, 'Pass 1b must not be rolled back');
    assert.equal(pass1b.hasChanges, true, 'Pass 1b must produce changes');
    for (const r of pass1b.results) {
        assert.equal(r.status, 'applied', `Pass 1b operation ${r.index} must be applied via same-author merge`);
    }
    assert.ok(pass1b.documentXml.includes('twenty-four (24) months'), 'Pass 1b merged Section C term to 24 months');
    assert.ok(pass1b.documentXml.includes('thirty (30) calendar days'), 'Pass 1b merged Section D.3 payment to 30 days');

    // Multi-user conflict protection check:
    // If MOHS attempts to redline Section E WITHOUT existingRevisions: 'accept-all-first', it must fail with EXISTING_REVISIONS
    const conflictAttempt = await applyOperationsToDocumentXml(pass1b.documentXml, [
        {
            type: 'replace',
            target: {
                exactText: 'Either party may terminate this Agreement without cause by giving to the other party written notification thereof at least ninety (90) calendar days prior to termination.  DEPARTMENT 1 shall provide funding to the DEPARTMENT 2 for services under this Agreement through the date of termination.  Upon termination, the parties hereto agree that all reports, clinical care transition summaries, and supporting documentation required for services rendered shall be provided by the DEPARTMENT 2 to DEPARTMENT 1 within fifteen (15) business days.'
            },
            modified: 'Conflicting modification attempt',
            author: AUTHOR_2
        }
    ], AUTHOR_2);
    assert.equal(conflictAttempt.results[0]?.status, 'error', 'Conflicting multi-user modification without policy must fail');
    assert.equal(conflictAttempt.results[0]?.error?.code, 'EXISTING_REVISIONS', 'Conflict error must be EXISTING_REVISIONS');

    // Pass 2: Author 2 (MOHS Counterparty Counsel) with accept-all-first on Section E
    const pass2 = await applyOperationsToDocumentXml(pass1b.documentXml, AUTHOR_2_OPERATIONS, AUTHOR_2);
    assert.equal(pass2.rolledBack, undefined, 'Pass 2 must not be rolled back');
    assert.equal(pass2.hasChanges, true, 'Pass 2 must produce changes');
    for (const r of pass2.results) {
        assert.equal(r.status, 'applied', `Pass 2 operation ${r.index} must be applied`);
    }

    // Pass 3: Author 1 (BCHD Lead Agency Counsel) counters back on the SAME provision (Section E) to reach compromise
    const pass3 = await applyOperationsToDocumentXml(pass2.documentXml, AUTHOR_1_FINAL_OPERATIONS, AUTHOR_1);
    assert.equal(pass3.rolledBack, undefined, 'Pass 3 must not be rolled back');
    assert.equal(pass3.hasChanges, true, 'Pass 3 must produce changes');
    for (const r of pass3.results) {
        assert.equal(r.status, 'applied', `Pass 3 operation ${r.index} must be applied`);
    }

    // Multi-Author Revisions Verification on Final Negotiated Document
    const authorMatches = [...pass3.documentXml.matchAll(/w:author="([^"]+)"/g)].map(m => m[1]);
    const authorCounts = {};
    for (const a of authorMatches) authorCounts[a] = (authorCounts[a] || 0) + 1;

    assert.ok(authorCounts[AUTHOR_1] >= 90, `Author 1 must have at least 90 revision tags, found ${authorCounts[AUTHOR_1]}`);
    assert.ok(authorCounts[AUTHOR_2] >= 25, `Author 2 must have at least 25 revision tags, found ${authorCounts[AUTHOR_2]}`);

    return { pass1, pass1b, pass2, pass3 };
}

// ----------------------------------------------------------------------------
// Test 3: Selective Author Accept and Reject Verification
// Verifies that revisions can be accepted/rejected by a specific author
// while preserving the other author's revisions intact.
// ----------------------------------------------------------------------------
async function testSelectiveRevisions(multiAuthorDocXml) {
    // 1. Accept only Author 2 (Counterparty)
    const acceptedAuth2 = acceptTrackedChangesInOoxml(multiAuthorDocXml, { author: AUTHOR_2 });
    assert.equal(acceptedAuth2.hasChanges, true);
    const remainingAfterAccept2 = [...acceptedAuth2.oxml.matchAll(/w:author="([^"]+)"/g)].map(m => m[1]);
    assert.ok(!remainingAfterAccept2.includes(AUTHOR_2), 'Author 2 revisions must be completely consumed upon accept');
    assert.ok(remainingAfterAccept2.includes(AUTHOR_1), 'Author 1 revisions must remain intact after accepting Author 2');

    // 2. Reject only Author 2 (Counterparty)
    const rejectedAuth2 = rejectTrackedChangesInOoxml(multiAuthorDocXml, { author: AUTHOR_2 });
    assert.equal(rejectedAuth2.hasChanges, true);
    const remainingAfterReject2 = [...rejectedAuth2.oxml.matchAll(/w:author="([^"]+)"/g)].map(m => m[1]);
    assert.ok(!remainingAfterReject2.includes(AUTHOR_2), 'Author 2 revisions must be completely removed upon reject');
    assert.ok(remainingAfterReject2.includes(AUTHOR_1), 'Author 1 revisions must remain intact after rejecting Author 2');

    // 3. Reject only Author 1 (Lead Agency)
    const rejectedAuth1 = rejectTrackedChangesInOoxml(multiAuthorDocXml, { author: AUTHOR_1 });
    assert.equal(rejectedAuth1.hasChanges, true);
    const remainingAfterReject1 = [...rejectedAuth1.oxml.matchAll(/w:author="([^"]+)"/g)].map(m => m[1]);
    assert.ok(!remainingAfterReject1.includes(AUTHOR_1), 'Author 1 revisions must be completely removed upon reject');
    assert.ok(remainingAfterReject1.includes(AUTHOR_2), 'Author 2 revisions must remain intact after rejecting Author 1');
}

// ----------------------------------------------------------------------------
// Test 4: Global Accept All & Reject All Differential Parity
// ----------------------------------------------------------------------------
async function testGlobalAcceptRejectAll(multiAuthorDocXml, originalDocXml) {
    // Accept All
    const acceptedAll = acceptTrackedChangesInOoxml(multiAuthorDocXml, { allAuthors: true });
    assert.equal(acceptedAll.hasChanges, true);
    assert.equal((acceptedAll.oxml.match(/<w:ins\b/g) || []).length, 0, 'No w:ins must remain in accepted document');
    assert.equal((acceptedAll.oxml.match(/<w:del\b/g) || []).length, 0, 'No w:del must remain in accepted document');

    const acceptedText = ingestWordOoxmlToPlainText(acceptedAll.oxml);
    assert.ok(acceptedText.includes('Baltimore City Health Department'), 'Accepted text must include BCHD');
    assert.ok(acceptedText.includes("Mayor's Office of Homeless Services"), 'Accepted text must include MOHS');
    assert.ok(acceptedText.includes('TWO MILLION FOUR HUNDRED FIFTY THOUSAND DOLLARS'), 'Accepted text must include $2.45M');
    assert.ok(acceptedText.includes('ATTACHMENT 4 – JOINT AGENCY PUBLIC ANNOUNCEMENT'), 'Accepted text must include Attachment 4');
    assert.ok(acceptedText.includes('thirty (30) calendar day opportunity to cure'), 'Accepted text must include 30-day cure period from Author 2');
    assert.ok(acceptedText.includes('sixty (60) calendar days prior to termination'), 'Accepted text must include final 60-day compromise');

    // Reject All
    const rejectedAll = rejectTrackedChangesInOoxml(multiAuthorDocXml, { allAuthors: true });
    assert.equal(rejectedAll.hasChanges, true);
    assert.equal((rejectedAll.oxml.match(/<w:ins\b/g) || []).length, 0, 'No w:ins must remain in rejected document');
    assert.equal((rejectedAll.oxml.match(/<w:del\b/g) || []).length, 0, 'No w:del must remain in rejected document');

    const rejectedText = ingestWordOoxmlToPlainText(rejectedAll.oxml);
    const originalText = ingestWordOoxmlToPlainText(originalDocXml);

    assert.ok(!rejectedText.includes('ATTACHMENT 4 – JOINT AGENCY PUBLIC ANNOUNCEMENT'), 'Rejected text must not include Attachment 4');
    assert.ok(rejectedText.includes('SPECIFY PURPOSE'), 'Rejected text must restore original SPECIFY PURPOSE placeholders');
    assert.ok(rejectedText.includes('NAME OF DEPARTMENT'), 'Rejected text must restore original NAME OF DEPARTMENT placeholders');
    assert.equal(rejectedText.includes('SPECIFY PURPOSE'), originalText.includes('SPECIFY PURPOSE'), 'Rejected text must mirror original placeholders');
}

// ----------------------------------------------------------------------------
// Test 5: Complete Package Pipeline via openDocx
// Executes Pass 1 -> Pass 1b (same-author merge) -> Pass 2 -> Pass 3.
// ----------------------------------------------------------------------------
async function testPackagePipeline(sourceBuffer) {
    const originalEntries = unzipEntries(sourceBuffer);

    // Pass 1 with openDocx (BCHD Initial Draft)
    const doc1 = openDocx(sourceBuffer);
    const res1 = await doc1.applyOperations(AUTHOR_1_PASS_1_OPERATIONS, { author: AUTHOR_1, atomic: true });
    assert.equal(res1.written, true, 'Pass 1 package must write successfully');
    assert.equal(res1.hasChanges, true, 'Pass 1 package must report changes');
    assert.ok(res1.artifactsChanged.includes('word/document.xml'));
    assert.ok(res1.artifactsChanged.includes('word/comments.xml'));
    assert.ok(res1.artifactsChanged.includes('word/numbering.xml'));

    // Pass 1b with openDocx (BCHD Alteration via same-author merge)
    const buf1 = res1.toBuffer();
    const doc1b = openDocx(buf1);
    const res1b = await doc1b.applyOperations(AUTHOR_1_ALTERATION_OPERATIONS, { author: AUTHOR_1, atomic: true });
    assert.equal(res1b.written, true, 'Pass 1b package must write successfully');
    assert.equal(res1b.hasChanges, true, 'Pass 1b package must report changes');

    // Pass 2 with openDocx (MOHS Counterparty proposals)
    const buf1b = res1b.toBuffer();
    const doc2 = openDocx(buf1b);
    const res2 = await doc2.applyOperations(AUTHOR_2_OPERATIONS, { author: AUTHOR_2, atomic: true });
    assert.equal(res2.written, true, 'Pass 2 package must write successfully');
    assert.equal(res2.hasChanges, true, 'Pass 2 package must report changes');

    // Pass 3 with openDocx (BCHD Final Compromise on Section E)
    const buf2 = res2.toBuffer();
    const doc3 = openDocx(buf2);
    const res3 = await doc3.applyOperations(AUTHOR_1_FINAL_OPERATIONS, { author: AUTHOR_1, atomic: true });
    assert.equal(res3.written, true, 'Pass 3 package must write successfully');
    assert.equal(res3.hasChanges, true, 'Pass 3 package must report changes');

    // Package Inspection
    const inspection = doc3.inspect();
    assert.ok(inspection.paragraphs.length >= 200, `Expected at least 200 paragraphs, found ${inspection.paragraphs.length}`);
    assert.equal(inspection.comments.length, 6, `Expected 6 comments total across both authors, found ${inspection.comments.length}`);

    const commentAuthors = [...new Set(inspection.comments.map(c => c.author))];
    assert.ok(commentAuthors.includes(AUTHOR_1), 'Comment authors must include Author 1');
    assert.ok(commentAuthors.includes(AUTHOR_2), 'Comment authors must include Author 2');

    // Structural OPC Package Validation
    const outputBuffer = res3.toBuffer();
    const outputEntries = unzipEntries(outputBuffer);
    await validateDocxPackage(new MemoryZip(outputEntries));

    // Untouched parts preservation
    const untouchedParts = [
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

    return buf1b;
}

// ----------------------------------------------------------------------------
// Test 6: Atomic Rollback Guarantee on Multi-Author Package
// ----------------------------------------------------------------------------
async function testAtomicRollback(pass1Buffer) {
    const doc = openDocx(pass1Buffer);

    const failingAuthor2Ops = [
        ...AUTHOR_2_OPERATIONS.slice(0, 4),
        {
            type: 'comment',
            target: { exactText: 'Non-existent target paragraph for intentional atomic failure' },
            textToComment: 'anchor',
            commentContent: 'Must trigger atomic rollback',
            author: AUTHOR_2
        },
        ...AUTHOR_2_OPERATIONS.slice(4)
    ];

    const result = await doc.applyOperations(failingAuthor2Ops, {
        author: AUTHOR_2,
        atomic: true
    });

    assert.equal(result.written, false, 'Failed batch must not write');
    assert.equal(result.rolledBack, true, 'Failed batch must roll back');
    assert.equal(result.status, 'error');
    assert.deepEqual(result.toBuffer(), pass1Buffer, 'Rolled back buffer must match pre-batch buffer byte-for-byte');
}

// ----------------------------------------------------------------------------
// Main Execution
// ----------------------------------------------------------------------------
async function main() {
    console.log('Loading SuperDoc New Interagency Agreement fixture...');
    const sourceBuffer = await loadSourceDocxBuffer();
    const sourceZip = unzipEntries(sourceBuffer);
    const docXml = sourceZip.get('word/document.xml').toString('utf8');

    console.log('  1. Running preflight on Author 1 operations...');
    await testPreflight(docXml);

    console.log('  2. Running engine multi-round negotiation (Pass 1 -> 1b -> Conflict Check -> Pass 2 -> Pass 3)...');
    const { pass3 } = await testEngineMultiAuthor(docXml);

    console.log('  3. Running selective author accept and reject verification...');
    await testSelectiveRevisions(pass3.documentXml);

    console.log('  4. Running global accept all & reject all differential parity...');
    await testGlobalAcceptRejectAll(pass3.documentXml, docXml);

    console.log('  5. Running complete package pipeline via openDocx with untouched-part preservation...');
    const pass1Buffer = await testPackagePipeline(sourceBuffer);

    console.log('  6. Running atomic rollback guarantee on multi-author package...');
    await testAtomicRollback(pass1Buffer);

    console.log('PASS: interagency_agreement_multi_author_tests.mjs - all 6 multi-author tests passed.');
}

await main();
