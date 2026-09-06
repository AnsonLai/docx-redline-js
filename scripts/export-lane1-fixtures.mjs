import '../tests/setup-xml-provider.mjs';

import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'fs';
import { dirname, join, resolve } from 'path';
import { fileURLToPath } from 'url';

import {
    acceptTrackedChangesInOoxml,
    rejectTrackedChangesInOoxml
} from '../index.js';
import { openDocx } from '../node/index.js';
import { buildZip } from './lib/minimal-zip.mjs';
import { unzipEntries } from './lib/zip-reader.mjs';
import {
    CORPUS_ID as SMOKING_CORPUS_ID,
    LEGAL_MARKUP_OPERATIONS as SMOKING_LEGAL_OPERATIONS,
    loadSourceDocxBuffer as loadSmokingDocxBuffer
} from '../tests/smoking_cessation_mixed_batch_tests.mjs';
import {
    CORPUS_ID as INTERAGENCY_CORPUS_ID,
    AUTHOR_1 as INTERAGENCY_AUTHOR_1,
    AUTHOR_2 as INTERAGENCY_AUTHOR_2,
    AUTHOR_1_PASS_1_OPERATIONS as INTERAGENCY_AUTHOR_1_PASS_1_OPS,
    AUTHOR_1_ALTERATION_OPERATIONS as INTERAGENCY_AUTHOR_1_ALT_OPS,
    AUTHOR_2_OPERATIONS as INTERAGENCY_AUTHOR_2_OPS,
    AUTHOR_1_FINAL_OPERATIONS as INTERAGENCY_AUTHOR_1_FINAL_OPS,
    loadSourceDocxBuffer as loadInteragencyDocxBuffer
} from '../tests/interagency_agreement_multi_author_tests.mjs';

const scriptDir = dirname(fileURLToPath(import.meta.url));
const repoRoot = dirname(scriptDir);

function argumentPath(name, fallback) {
    const index = process.argv.indexOf(name);
    if (index < 0) return fallback;
    if (!process.argv[index + 1]) throw new Error(`${name} requires a path`);
    return resolve(process.cwd(), process.argv[index + 1]);
}

const outputDir = argumentPath('--output-dir', join(repoRoot, 'tmp', 'lane1-docx'));
mkdirSync(outputDir, { recursive: true });

function packageWithPartXml(sourceEntries, partName, partXml) {
    return buildZip(Array.from(sourceEntries, ([name, data]) => ({
        name,
        data: name === partName ? Buffer.from(partXml, 'utf8') : data
    })));
}

const lane1Cases = [];

// ============================================================================
// Case 1: Smoking Cessation Services Agreement (NHS / Pharmacy Consortium)
// ============================================================================
console.log('Exporting Lane 1 Case: Smoking Cessation Contract Review...');
const smokingSourceBuffer = await loadSmokingDocxBuffer();
const smokingDoc = openDocx(smokingSourceBuffer);
const smokingAuthor = 'Senior Legal Counsel (NHS/Pharmacy Consortium)';
const smokingResult = await smokingDoc.applyOperations(SMOKING_LEGAL_OPERATIONS, {
    author: smokingAuthor,
    atomic: true,
    validate: true
});

if (!smokingResult.written) {
    throw new Error(`Failed to apply smoking cessation legal markup: ${smokingResult.error?.message}`);
}

const smokingTrackedBuffer = smokingResult.toBuffer();
const smokingTrackedEntries = unzipEntries(smokingTrackedBuffer);
const smokingTrackedXml = smokingTrackedEntries.get('word/document.xml').toString('utf8');

const smokingAccepted = acceptTrackedChangesInOoxml(smokingTrackedXml, { allAuthors: true });
const smokingAcceptedBuffer = packageWithPartXml(smokingTrackedEntries, 'word/document.xml', smokingAccepted.oxml);

const smokingRejected = rejectTrackedChangesInOoxml(smokingTrackedXml, { allAuthors: true });
const smokingRejectedBuffer = packageWithPartXml(smokingTrackedEntries, 'word/document.xml', smokingRejected.oxml);

const smokingCaseName = 'smoking-cessation-contract-review';
writeFileSync(join(outputDir, `${smokingCaseName}.source.docx`), smokingSourceBuffer);
writeFileSync(join(outputDir, `${smokingCaseName}.docx`), smokingTrackedBuffer);
writeFileSync(join(outputDir, `${smokingCaseName}.accepted.docx`), smokingAcceptedBuffer);
writeFileSync(join(outputDir, `${smokingCaseName}.rejected.docx`), smokingRejectedBuffer);
writeFileSync(join(outputDir, `${smokingCaseName}.document.xml`), smokingTrackedXml);

const smokingComments = SMOKING_LEGAL_OPERATIONS
    .filter(op => op.type === 'comment')
    .map((op, idx) => ({
        id: idx + 1,
        author: op.author || smokingAuthor,
        target: typeof op.target === 'string' ? op.target : op.target.exactText,
        anchor: op.textToComment,
        content: op.commentContent
    }));

const smokingOperationsList = SMOKING_LEGAL_OPERATIONS.map((op, idx) => ({
    index: idx + 1,
    type: op.type,
    target: typeof op.target === 'string' ? op.target : op.target.exactText,
    change: op.type === 'comment'
        ? `[Comment on "${op.textToComment}"] ${op.commentContent}`
        : op.type === 'replace'
            ? op.modified
            : `Format: ${JSON.stringify(op.format)}`,
    author: op.author || smokingAuthor,
    status: 'applied'
}));

const smokingMeta = {
    identity: `lane1:${smokingCaseName}`,
    name: smokingCaseName,
    title: 'Smoking Cessation Services Agreement (NHS / Pharmacy Consortium)',
    category: 'legal',
    sourceDocument: 'Smoking_Cessation_Contract_for_2020_-_2021_website.docx',
    sourceId: SMOKING_CORPUS_ID,
    sourceWords: 10770,
    sourceParagraphs: 533,
    author: smokingAuthor,
    description: 'Comprehensive 28-operation real-world legal markup of Southampton City Council NHS smoking cessation contract: recitals, multi-paragraph clauses, liability & indemnity caps, GDPR statutory updates, termination notice floors, Schedule 1 clinical protocol CO ppm monitoring, Schedule 1 Table 3 dates, and Schedule 1 Table 5 fee tariffs.',
    operationsCount: SMOKING_LEGAL_OPERATIONS.length,
    breakdown: {
        redlines: SMOKING_LEGAL_OPERATIONS.filter(o => o.type === 'replace').length,
        comments: SMOKING_LEGAL_OPERATIONS.filter(o => o.type === 'comment').length,
        formatting: SMOKING_LEGAL_OPERATIONS.filter(o => o.type === 'format').length
    },
    comments: smokingComments,
    operations: smokingOperationsList,
    revisions: {
        insertions: (smokingTrackedXml.match(/<w:ins\b/g) || []).length,
        deletions: (smokingTrackedXml.match(/<w:del\b/g) || []).length,
        formatting: (smokingTrackedXml.match(/<w:(?:rPrChange|pPrChange)\b/g) || []).length,
        comments: (smokingTrackedXml.match(/<w:commentRangeStart\b/g) || []).length
    }
};

writeFileSync(join(outputDir, `${smokingCaseName}.meta.json`), JSON.stringify(smokingMeta, null, 2));
lane1Cases.push(smokingMeta);

// ============================================================================
// Case 2: County Advisory Board Bylaws (Symbol Bullets & Governance)
// ============================================================================
const bylawsDocId = 'd27fe5513ca474bcd6c02136d2e4b8179203f5855c4333af81bc964755c7abe4';
const bylawsPath = join(repoRoot, 'tmp', 'superdoc-corpus', `${bylawsDocId}.docx`);

if (existsSync(bylawsPath)) {
    console.log('Exporting Lane 1 Case: County Advisory Board Bylaws...');
    const bylawsSourceBuffer = readFileSync(bylawsPath);
    const bylawsDoc = openDocx(bylawsSourceBuffer);
    const bylawsAuthor = 'BylawsCommittee';
    const bylawsOps = [
        {
            type: 'replace',
            target: { exactText: 'One (1) representative from the Board of Supervisors;' },
            modified: 'Two (2) voting representatives from the County Board of Supervisors;',
            author: bylawsAuthor
        },
        {
            type: 'comment',
            target: { exactText: 'Five (5) representatives from the incorporated cities/town;' },
            textToComment: 'incorporated cities/town',
            commentContent: 'Confirm municipal apportionment following 2024 census.',
            author: bylawsAuthor
        },
        {
            type: 'replace',
            target: { exactText: 'Three (3) representatives from the community at large.' },
            modified: 'Three (3) representatives from the community at large, including at least one youth advocate.',
            author: bylawsAuthor
        }
    ];

    const bylawsResult = await bylawsDoc.applyOperations(bylawsOps, {
        author: bylawsAuthor,
        atomic: true,
        validate: true
    });

    if (bylawsResult.written) {
        const bylawsTrackedBuffer = bylawsResult.toBuffer();
        const bylawsTrackedEntries = unzipEntries(bylawsTrackedBuffer);
        const bylawsTrackedXml = bylawsTrackedEntries.get('word/document.xml').toString('utf8');

        const bylawsAccepted = acceptTrackedChangesInOoxml(bylawsTrackedXml, { allAuthors: true });
        const bylawsAcceptedBuffer = packageWithPartXml(bylawsTrackedEntries, 'word/document.xml', bylawsAccepted.oxml);

        const bylawsRejected = rejectTrackedChangesInOoxml(bylawsTrackedXml, { allAuthors: true });
        const bylawsRejectedBuffer = packageWithPartXml(bylawsTrackedEntries, 'word/document.xml', bylawsRejected.oxml);

        const bylawsCaseName = 'bylaws-symbol-bullets';
        writeFileSync(join(outputDir, `${bylawsCaseName}.source.docx`), bylawsSourceBuffer);
        writeFileSync(join(outputDir, `${bylawsCaseName}.docx`), bylawsTrackedBuffer);
        writeFileSync(join(outputDir, `${bylawsCaseName}.accepted.docx`), bylawsAcceptedBuffer);
        writeFileSync(join(outputDir, `${bylawsCaseName}.rejected.docx`), bylawsRejectedBuffer);
        writeFileSync(join(outputDir, `${bylawsCaseName}.document.xml`), bylawsTrackedXml);

        const bylawsComments = bylawsOps
            .filter(op => op.type === 'comment')
            .map((op, idx) => ({
                id: idx + 1,
                author: op.author || bylawsAuthor,
                target: op.target.exactText,
                anchor: op.textToComment,
                content: op.commentContent
            }));

        const bylawsOperationsList = bylawsOps.map((op, idx) => ({
            index: idx + 1,
            type: op.type,
            target: op.target.exactText,
            change: op.type === 'comment'
                ? `[Comment on "${op.textToComment}"] ${op.commentContent}`
                : op.modified,
            author: op.author || bylawsAuthor,
            status: 'applied'
        }));

        const bylawsMeta = {
            identity: `lane1:${bylawsCaseName}`,
            name: bylawsCaseName,
            title: 'County Advisory Board Bylaws (Governance & Symbol Bullets)',
            category: 'administrative',
            sourceDocument: `${bylawsDocId}.docx`,
            sourceId: bylawsDocId,
            author: bylawsAuthor,
            description: 'Governance updates to Advisory Board membership: supervisor apportionment, census-based municipal representation comments, and youth advocacy inclusion.',
            operationsCount: bylawsOps.length,
            breakdown: {
                redlines: bylawsOps.filter(o => o.type === 'replace').length,
                comments: bylawsOps.filter(o => o.type === 'comment').length,
                formatting: 0
            },
            comments: bylawsComments,
            operations: bylawsOperationsList,
            revisions: {
                insertions: (bylawsTrackedXml.match(/<w:ins\b/g) || []).length,
                deletions: (bylawsTrackedXml.match(/<w:del\b/g) || []).length,
                formatting: (bylawsTrackedXml.match(/<w:(?:rPrChange|pPrChange)\b/g) || []).length,
                comments: (bylawsTrackedXml.match(/<w:commentRangeStart\b/g) || []).length
            }
        };

        writeFileSync(join(outputDir, `${bylawsCaseName}.meta.json`), JSON.stringify(bylawsMeta, null, 2));
        lane1Cases.push(bylawsMeta);
    }
}

// ============================================================================
// Case 3: New Interagency Agreement (Multi-Author Negotiation: BCHD vs. MOHS)
// ============================================================================
console.log('Exporting Lane 1 Case: New Interagency Agreement Multi-Author Negotiation...');
const interagencySourceBuffer = await loadInteragencyDocxBuffer();

// Pass 1: BCHD Lead Agency Counsel Initial Draft (clean headers, no bullets)
const interagencyDoc1 = openDocx(interagencySourceBuffer);
const interagencyRes1 = await interagencyDoc1.applyOperations(INTERAGENCY_AUTHOR_1_PASS_1_OPS, {
    author: INTERAGENCY_AUTHOR_1,
    atomic: true,
    validate: true
});
if (!interagencyRes1.written) {
    throw new Error(`Failed to apply interagency Pass 1 legal markup: ${interagencyRes1.error?.message}`);
}

// Pass 1b: BCHD Revision Alteration via Same-Author Merge (Term 36m -> 24m, Payment 45d -> 30d)
const interagencyDoc1b = openDocx(interagencyRes1.toBuffer());
const interagencyRes1b = await interagencyDoc1b.applyOperations(INTERAGENCY_AUTHOR_1_ALT_OPS, {
    author: INTERAGENCY_AUTHOR_1,
    atomic: true,
    validate: true
});
if (!interagencyRes1b.written) {
    throw new Error(`Failed to apply interagency Pass 1b alteration markup: ${interagencyRes1b.error?.message}`);
}

// Pass 2: MOHS Counterparty Counsel Review (counters on Section E, adds cure period, pre-award, Attachments 1 & 2)
const interagencyDoc2 = openDocx(interagencyRes1b.toBuffer());
const interagencyRes2 = await interagencyDoc2.applyOperations(INTERAGENCY_AUTHOR_2_OPS, {
    author: INTERAGENCY_AUTHOR_2,
    atomic: true,
    validate: true
});
if (!interagencyRes2.written) {
    throw new Error(`Failed to apply interagency Pass 2 counterparty markup: ${interagencyRes2.error?.message}`);
}

// Pass 3: BCHD Lead Agency Counsel Compromise on Section E (counters back on SAME provision: 60d + safeguard)
const interagencyDoc3 = openDocx(interagencyRes2.toBuffer());
const interagencyRes3 = await interagencyDoc3.applyOperations(INTERAGENCY_AUTHOR_1_FINAL_OPS, {
    author: INTERAGENCY_AUTHOR_1,
    atomic: true,
    validate: true
});
if (!interagencyRes3.written) {
    throw new Error(`Failed to apply interagency Pass 3 compromise markup: ${interagencyRes3.error?.message}`);
}

const interagencyTrackedBuffer = interagencyRes3.toBuffer();
const interagencyTrackedEntries = unzipEntries(interagencyTrackedBuffer);
const interagencyTrackedXml = interagencyTrackedEntries.get('word/document.xml').toString('utf8');

const interagencyAccepted = acceptTrackedChangesInOoxml(interagencyTrackedXml, { allAuthors: true });
const interagencyAcceptedBuffer = packageWithPartXml(interagencyTrackedEntries, 'word/document.xml', interagencyAccepted.oxml);

const interagencyRejected = rejectTrackedChangesInOoxml(interagencyTrackedXml, { allAuthors: true });
const interagencyRejectedBuffer = packageWithPartXml(interagencyTrackedEntries, 'word/document.xml', interagencyRejected.oxml);

const interagencyCaseName = 'interagency-agreement-multi-author';
writeFileSync(join(outputDir, `${interagencyCaseName}.source.docx`), interagencySourceBuffer);
writeFileSync(join(outputDir, `${interagencyCaseName}.docx`), interagencyTrackedBuffer);
writeFileSync(join(outputDir, `${interagencyCaseName}.accepted.docx`), interagencyAcceptedBuffer);
writeFileSync(join(outputDir, `${interagencyCaseName}.rejected.docx`), interagencyRejectedBuffer);
writeFileSync(join(outputDir, `${interagencyCaseName}.document.xml`), interagencyTrackedXml);

const allInteragencyOps = [
    ...INTERAGENCY_AUTHOR_1_PASS_1_OPS.map(op => ({ ...op, author: op.author || INTERAGENCY_AUTHOR_1 })),
    ...INTERAGENCY_AUTHOR_1_ALT_OPS.map(op => ({ ...op, author: op.author || INTERAGENCY_AUTHOR_1 })),
    ...INTERAGENCY_AUTHOR_2_OPS.map(op => ({ ...op, author: op.author || INTERAGENCY_AUTHOR_2 })),
    ...INTERAGENCY_AUTHOR_1_FINAL_OPS.map(op => ({ ...op, author: op.author || INTERAGENCY_AUTHOR_1 }))
];

const interagencyComments = allInteragencyOps
    .filter(op => op.type === 'comment')
    .map((op, idx) => ({
        id: idx + 1,
        author: op.author,
        target: typeof op.target === 'string' ? op.target : op.target.exactText,
        anchor: op.textToComment,
        content: op.commentContent
    }));

const interagencyOperationsList = allInteragencyOps.map((op, idx) => ({
    index: idx + 1,
    type: op.type,
    target: typeof op.target === 'string' ? op.target : op.target.exactText,
    change: op.type === 'comment'
        ? `[Comment on "${op.textToComment}"] ${op.commentContent}`
        : op.modified,
    author: op.author,
    status: 'applied'
}));

const interagencyMeta = {
    identity: `lane1:${interagencyCaseName}`,
    name: interagencyCaseName,
    title: 'New Interagency Agreement (Multi-Author Negotiation: BCHD vs. MOHS)',
    category: 'legal',
    sourceDocument: 'New Interagency Agreement.docx',
    sourceId: INTERAGENCY_CORPUS_ID,
    sourceWords: 1726,
    sourceParagraphs: 173,
    authors: [INTERAGENCY_AUTHOR_1, INTERAGENCY_AUTHOR_2],
    author: `${INTERAGENCY_AUTHOR_1} & ${INTERAGENCY_AUTHOR_2}`,
    description: 'Four-round woven negotiation on Baltimore City interagency healthcare & homelessness accord: (1) BCHD Lead Agency Counsel establishes clean legal section headers A-G (not bullets), programmatic terms ($2.45M), Attachment 3 sub-bullets, and Attachment 4 joint public announcement with tables & lists; (2) BCHD alters its initial redlines in a separate step (Term 36m -> 24m, Payment 45d -> 30d) via same-author merge; (3) MOHS Counterparty Counsel counters on the same termination provision (45 days + wind-down) and adds 30-day cure period, pre-award costs, annual review, and counter-comments; (4) BCHD responds back on the same termination provision reaching compromise (60 days + shelter safeguard).',
    operationsCount: allInteragencyOps.length,
    breakdown: {
        redlines: allInteragencyOps.filter(o => o.type === 'replace').length,
        comments: allInteragencyOps.filter(o => o.type === 'comment').length,
        formatting: 0
    },
    comments: interagencyComments,
    operations: interagencyOperationsList,
    revisions: {
        insertions: (interagencyTrackedXml.match(/<w:ins\b/g) || []).length,
        deletions: (interagencyTrackedXml.match(/<w:del\b/g) || []).length,
        formatting: (interagencyTrackedXml.match(/<w:(?:rPrChange|pPrChange)\b/g) || []).length,
        comments: (interagencyTrackedXml.match(/<w:commentRangeStart\b/g) || []).length
    }
};

writeFileSync(join(outputDir, `${interagencyCaseName}.meta.json`), JSON.stringify(interagencyMeta, null, 2));
lane1Cases.push(interagencyMeta);

const manifest = {
    generatedAt: new Date().toISOString(),
    casesCount: lane1Cases.length,
    cases: lane1Cases
};

writeFileSync(join(outputDir, 'manifest.json'), JSON.stringify(manifest, null, 2));
console.log(`PASS: Exported ${lane1Cases.length} Lane 1 compound visual fixtures to ${outputDir}`);
