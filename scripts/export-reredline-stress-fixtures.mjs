import { mkdirSync, writeFileSync } from 'fs';
import { resolve, join } from 'path';
import { fileURLToPath } from 'url';

import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { configureXmlProvider } from '../adapters/xml-adapter.js';
import { configureLogger } from '../adapters/logger.js';
import { validateRedlineOoxml } from '../core/redline-validation.js';
import { RevisionIdAllocator } from '../core/types.js';
import { applyRedlineToOxml, reconcileMarkdownTableOoxml } from '../index.js';
import { ingestWordOoxmlToPlainText } from '../pipeline/ingestion-export.js';
import { buildMinimalDocx } from './lib/minimal-zip.mjs';

configureXmlProvider({ DOMParser, XMLSerializer });
configureLogger(console, { level: 'silent' });

const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const serializer = new XMLSerializer();

const escapeXml = value => String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');

const run = (text, properties = '') => `<w:r>${properties ? `<w:rPr>${properties}</w:rPr>` : ''}<w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r>`;

function paragraph(text, { heading = false, pageBreakBefore = false, numId = null } = {}) {
    const pPr = [
        pageBreakBefore ? '<w:pageBreakBefore/>' : '',
        numId == null ? '' : `<w:numPr><w:ilvl w:val="0"/><w:numId w:val="${numId}"/></w:numPr>`,
        `<w:spacing w:after="${heading ? 120 : 100}" w:line="${heading ? 280 : 300}" w:lineRule="auto"/>`
    ].join('');
    const rPr = heading
        ? '<w:b/><w:color w:val="2E74B5"/><w:sz w:val="28"/><w:szCs w:val="28"/>'
        : '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/>';
    return `<w:p xmlns:w="${NS_W}"><w:pPr>${pPr}</w:pPr>${run(text, rPr)}</w:p>`;
}

function titleParagraph(title, subtitle) {
    return [
        `<w:p><w:pPr><w:spacing w:after="80"/></w:pPr>${run(title, '<w:b/><w:color w:val="1F4D78"/><w:sz w:val="36"/><w:szCs w:val="36"/>')}</w:p>`,
        `<w:p><w:pPr><w:spacing w:after="240"/></w:pPr>${run(subtitle, '<w:i/><w:color w:val="666666"/><w:sz w:val="20"/><w:szCs w:val="20"/>')}</w:p>`
    ].join('');
}

function markdownTable(headers, rows) {
    return [
        `| ${headers.join(' | ')} |`,
        `| ${headers.map(() => '---').join(' | ')} |`,
        ...rows.map(row => `| ${row.join(' | ')} |`)
    ].join('\n');
}

function tableXml(headers, rows) {
    const widths = [3000, 3000, 3360];
    const rowXml = (values, isHeader) => `<w:tr>${values.map((value, index) => `
      <w:tc>
        <w:tcPr><w:tcW w:w="${widths[index]}" w:type="dxa"/><w:vAlign w:val="center"/>${isHeader ? '<w:shd w:fill="E8EEF5"/>' : ''}</w:tcPr>
        <w:p><w:pPr><w:spacing w:after="40"/></w:pPr>${run(value, `${isHeader ? '<w:b/>' : ''}<w:sz w:val="20"/><w:szCs w:val="20"/>`)}</w:p>
      </w:tc>`).join('')}</w:tr>`;
    return `<w:tbl xmlns:w="${NS_W}">
      <w:tblPr>
        <w:tblW w:w="9360" w:type="dxa"/><w:tblInd w:w="120" w:type="dxa"/><w:tblLayout w:type="fixed"/>
        <w:tblBorders><w:top w:val="single" w:sz="4" w:color="AAB7C4"/><w:left w:val="single" w:sz="4" w:color="AAB7C4"/><w:bottom w:val="single" w:sz="4" w:color="AAB7C4"/><w:right w:val="single" w:sz="4" w:color="AAB7C4"/><w:insideH w:val="single" w:sz="4" w:color="D5DCE3"/><w:insideV w:val="single" w:sz="4" w:color="D5DCE3"/></w:tblBorders>
        <w:tblCellMar><w:top w:w="80" w:type="dxa"/><w:start w:w="120" w:type="dxa"/><w:bottom w:w="80" w:type="dxa"/><w:end w:w="120" w:type="dxa"/></w:tblCellMar>
      </w:tblPr>
      <w:tblGrid>${widths.map(width => `<w:gridCol w:w="${width}"/>`).join('')}</w:tblGrid>
      ${rowXml(headers, true)}${rows.map(row => rowXml(row, false)).join('')}
    </w:tbl>`;
}

const NUMBERING_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="${NS_W}">
  <w:abstractNum w:abstractNumId="10"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/><w:lvlJc w:val="left"/><w:pPr><w:tabs><w:tab w:val="num" w:pos="540"/></w:tabs><w:ind w:left="540" w:hanging="270"/></w:pPr></w:lvl></w:abstractNum>
  <w:abstractNum w:abstractNumId="11"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:tabs><w:tab w:val="num" w:pos="540"/></w:tabs><w:ind w:left="540" w:hanging="270"/></w:pPr></w:lvl></w:abstractNum>
  <w:num w:numId="10"><w:abstractNumId w:val="10"/></w:num>
  <w:num w:numId="11"><w:abstractNumId w:val="11"/></w:num>
</w:numbering>`;

const SCENARIOS = [
    {
        name: 'mixed-policy-review',
        title: 'Mixed Policy Review Stress Fixture',
        paragraphs: [
            ['The review board meets every month to assess open actions.', 'The review board meets every quarter to assess open actions.', 'The oversight committee meets every quarter to assess unresolved actions.'],
            ['All submissions require a manager signature.', 'All submissions require a manager signature and a dated approval note.', 'All submissions require a director signature and a dated approval note.'],
            ['The legacy exception remains available during the pilot period.', 'The legacy exception remains available.', 'The temporary exception remains available through year-end.'],
            ['Critical controls must be documented before launch.', 'Critical **controls** must be documented before launch.', 'Critical **safeguards** must be documented before production launch.'],
            ['The coordinator records decisions in the central register.', 'The coordinator records final decisions in the central register within two business days.', 'The governance lead records final decisions in the central register within one business day.']
        ],
        lists: [
            ['Confirm the meeting agenda.', 'Confirm the revised meeting agenda.', 'Confirm the final meeting agenda.'],
            ['Collect stakeholder comments.', 'Collect written stakeholder comments.', 'Collect written stakeholder approvals.'],
            ['Escalate unresolved risks.', 'Escalate unresolved operational risks.', 'Escalate unresolved operational and legal risks.'],
            ['Archive the superseded draft.', 'Archive the superseded working draft.', 'Archive the superseded approved draft.'],
            ['Record the decision owner.', 'Record the **decision owner**.', 'Record the **accountable decision owner**.'],
            ['Publish the implementation notice.', 'Publish the implementation notice within five days.', 'Publish the implementation notice within three days.']
        ],
        tables: [
            {
                headers: ['Workstream', 'Owner', 'Status'],
                source: [['Budget', 'Finance', 'Pending'], ['Policy', 'Legal', 'Draft'], ['Training', 'People', 'Planned']],
                round1: [['Budget', 'Treasury', 'Approved'], ['Policy', 'Legal', 'In review'], ['Training', 'People', 'Planned']],
                round2: [['Budget', 'Treasury', 'Complete'], ['Policy', 'Compliance', 'Approved'], ['Training', 'People', 'Scheduled']]
            },
            {
                headers: ['Risk', 'Rating', 'Response'],
                source: [['Access', 'Medium', 'Monitor'], ['Continuity', 'Low', 'Review']],
                round1: [['Access', 'High', 'Mitigate'], ['Privacy', 'Medium', 'Assess'], ['Continuity', 'Low', 'Review']],
                round2: [['Access', 'Medium', 'Mitigate'], ['Privacy', 'Low', 'Monitor'], ['Continuity', 'Low', 'Close']]
            },
            {
                headers: ['Milestone', 'Date', 'Lead'],
                source: [['Design', 'April 10', 'Avery'], ['Pilot', 'May 15', 'Blair'], ['Launch', 'June 20', 'Casey']],
                round1: [['Design', 'April 17', 'Avery'], ['Launch', 'June 27', 'Casey']],
                round2: [['Design', 'April 17', 'Avery'], ['Readiness', 'June 10', 'Blair'], ['Launch', 'July 4', 'Casey']]
            }
        ]
    },
    {
        name: 'mixed-contract-operations',
        title: 'Mixed Contract Operations Stress Fixture',
        paragraphs: [
            ['The supplier will retain audit records for three years.', 'The supplier will retain audit records for five years.', 'The contractor will retain complete audit records for seven years.'],
            ['Notices may be delivered by ordinary mail.', 'Notices may be delivered by registered mail or secure email.', 'Notices must be delivered by secure email with receipt confirmation.'],
            ['The annual renewal is automatic unless cancelled.', 'The annual renewal requires written confirmation.', 'Each renewal requires written confirmation from both parties.'],
            ['Service credits are calculated monthly.', 'Service **credits** are calculated monthly.', 'Service **adjustments** are calculated quarterly.'],
            ['The customer may request one compliance report.', 'The customer may request two compliance reports each year.', 'The customer may request quarterly compliance reports.']
        ],
        lists: [
            ['Verify insurance certificates.', 'Verify current insurance certificates.', 'Verify current insurance certificates and endorsements.'],
            ['Review subcontractor access.', 'Review approved subcontractor access.', 'Review and recertify approved subcontractor access.'],
            ['Log all security incidents.', 'Log all material security incidents.', 'Log all material security incidents within four hours.'],
            ['Preserve delivery receipts.', 'Preserve signed delivery receipts.', 'Preserve signed electronic delivery receipts.'],
            ['Confirm the remediation owner.', 'Confirm the **remediation owner**.', 'Confirm the **executive remediation owner**.'],
            ['Close completed obligations.', 'Close completed obligations after evidence review.', 'Close completed obligations after independent evidence review.']
        ],
        tables: [
            {
                headers: ['Clause', 'Position', 'Owner'],
                source: [['Liability', 'Open', 'Legal'], ['Security', 'Draft', 'Risk'], ['Privacy', 'Open', 'Privacy']],
                round1: [['Liability', 'Capped', 'Legal'], ['Security', 'Approved', 'Risk'], ['Privacy', 'Open', 'Privacy']],
                round2: [['Liability', 'Revised cap', 'Legal'], ['Security', 'Approved', 'Security'], ['Privacy', 'Approved', 'Privacy']]
            },
            {
                headers: ['Deliverable', 'Due', 'State'],
                source: [['Report', 'Day 10', 'Planned'], ['Certificate', 'Day 20', 'Planned']],
                round1: [['Report', 'Day 7', 'In progress'], ['Evidence pack', 'Day 14', 'Planned'], ['Certificate', 'Day 20', 'Planned']],
                round2: [['Report', 'Day 5', 'Complete'], ['Evidence pack', 'Day 12', 'In progress'], ['Certificate', 'Day 18', 'Planned']]
            },
            {
                headers: ['Region', 'Threshold', 'Reviewer'],
                source: [['Canada', '$50,000', 'Morgan'], ['United States', '$75,000', 'Riley'], ['Europe', 'EUR 60,000', 'Taylor']],
                round1: [['Canada', '$60,000', 'Morgan'], ['Europe', 'EUR 70,000', 'Taylor']],
                round2: [['Canada', '$65,000', 'Morgan'], ['United Kingdom', 'GBP 55,000', 'Jordan'], ['Europe', 'EUR 70,000', 'Taylor']]
            }
        ]
    }
];

function firstElement(xml, localName) {
    const document = new DOMParser().parseFromString(xml, 'application/xml');
    const element = document.getElementsByTagNameNS(NS_W, localName)[0];
    if (!element) throw new Error(`Expected w:${localName} in generated OOXML`);
    return serializer.serializeToString(element);
}

async function redlineParagraph(sourceXml, original, modified, author, allocator, existingRevisions = 'reject-input') {
    const result = await applyRedlineToOxml(sourceXml, original, modified, {
        generateRedlines: true,
        author,
        existingRevisions,
        _revisionIdAllocator: allocator
    });
    if (!result.hasChanges || result.status === 'error') {
        throw new Error(`Paragraph redline failed: ${result.error?.message || result.status || 'no change'}`);
    }
    return firstElement(result.oxml, 'p');
}

async function redlineTable(sourceXml, sourceRows, targetRows, headers, author, allocator, existingRevisions = 'reject-input') {
    const modifiedMarkdown = markdownTable(headers, targetRows);
    const result = await reconcileMarkdownTableOoxml(sourceXml, sourceRows[0][0], modifiedMarkdown, {
        generateRedlines: true,
        author,
        existingRevisions,
        _revisionIdAllocator: allocator
    });
    if (!result.hasChanges || result.status === 'error') {
        throw new Error(`Table redline failed: ${result.error?.message || result.status || 'no change'}`);
    }
    return firstElement(result.oxml, 'tbl');
}

function documentXml(scenario, stage, paragraphBlocks, listBlocks, tableBlocks) {
    const stageLabel = stage === 'source'
        ? 'Clean source document'
        : stage === 'round1'
            ? 'First review round - several paragraph, list, formatting, and table changes'
            : 'Second review round - prior revisions accepted per block, then re-redlined';
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${NS_W}"><w:body>
  ${titleParagraph(scenario.title, stageLabel)}
  ${paragraph('Paragraph revisions', { heading: true })}
  ${paragraphBlocks.join('\n  ')}
  ${paragraph('Bullet and numbered-list revisions', { heading: true, pageBreakBefore: true })}
  ${listBlocks.join('\n  ')}
  ${paragraph('Table revisions', { heading: true, pageBreakBefore: true })}
  ${tableBlocks.join(`\n  ${paragraph('', {})}\n  `)}
  <w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>
</w:body></w:document>`;
}

function revisionCounts(xml) {
    const doc = new DOMParser().parseFromString(xml, 'application/xml');
    return Object.fromEntries(['ins', 'del', 'rPrChange', 'pPrChange'].map(name => [
        name,
        doc.getElementsByTagNameNS(NS_W, name).length
    ]));
}

async function buildScenario(scenario) {
    const round1Allocator = new RevisionIdAllocator(1000);
    const round2Allocator = new RevisionIdAllocator(5000);
    const sourceParagraphs = scenario.paragraphs.map(([source]) => paragraph(source));
    const sourceLists = scenario.lists.map(([source], index) => paragraph(source, { numId: index < 4 ? 10 : 11 }));
    const sourceTables = scenario.tables.map(table => tableXml(table.headers, table.source));

    const round1Paragraphs = [];
    const round2Paragraphs = [];
    for (let index = 0; index < scenario.paragraphs.length; index += 1) {
        const [source, first, second] = scenario.paragraphs[index];
        const firstXml = await redlineParagraph(sourceParagraphs[index], source, first, 'Round One Reviewer', round1Allocator);
        round1Paragraphs.push(firstXml);
        round2Paragraphs.push(await redlineParagraph(firstXml, first.replace(/[*+]/g, ''), second, 'Round Two Reviewer', round2Allocator, 'accept-all-first'));
    }

    const round1Lists = [];
    const round2Lists = [];
    for (let index = 0; index < scenario.lists.length; index += 1) {
        const [source, first, second] = scenario.lists[index];
        const firstXml = await redlineParagraph(sourceLists[index], source, first, 'Round One Reviewer', round1Allocator);
        round1Lists.push(firstXml);
        round2Lists.push(await redlineParagraph(firstXml, first.replace(/[*+]/g, ''), second, 'Round Two Reviewer', round2Allocator, 'accept-all-first'));
    }

    const round1Tables = [];
    const round2Tables = [];
    for (let index = 0; index < scenario.tables.length; index += 1) {
        const table = scenario.tables[index];
        const firstXml = await redlineTable(sourceTables[index], table.source, table.round1, table.headers, 'Round One Reviewer', round1Allocator);
        round1Tables.push(firstXml);
        round2Tables.push(await redlineTable(firstXml, table.round1, table.round2, table.headers, 'Round Two Reviewer', round2Allocator, 'accept-all-first'));
    }

    return {
        source: documentXml(scenario, 'source', sourceParagraphs, sourceLists, sourceTables),
        round1: documentXml(scenario, 'round1', round1Paragraphs, round1Lists, round1Tables),
        rerelined: documentXml(scenario, 'rerelined', round2Paragraphs, round2Lists, round2Tables)
    };
}

export async function generateReredlineStressFixtures(outputDir) {
    const resolvedOutputDir = resolve(outputDir);
    mkdirSync(resolvedOutputDir, { recursive: true });
    const manifest = { generatedBy: 'scripts/export-reredline-stress-fixtures.mjs', scenarios: [] };

    for (const scenario of SCENARIOS) {
        const stages = await buildScenario(scenario);
        const entry = { name: scenario.name, stages: {} };
        for (const [stage, xml] of Object.entries(stages)) {
            const validation = validateRedlineOoxml(xml);
            if (!validation.valid) {
                throw new Error(`${scenario.name}/${stage} failed validation: ${JSON.stringify(validation.issues)}`);
            }
            const baseName = `${scenario.name}-${stage}`;
            writeFileSync(join(resolvedOutputDir, `${baseName}.docx`), buildMinimalDocx(xml, { numberingXml: NUMBERING_XML }));
            writeFileSync(join(resolvedOutputDir, `${baseName}.document.xml`), xml, 'utf8');
            entry.stages[stage] = {
                docx: `${baseName}.docx`,
                documentXml: `${baseName}.document.xml`,
                revisions: revisionCounts(xml),
                visibleText: ingestWordOoxmlToPlainText(xml)
            };
        }
        manifest.scenarios.push(entry);
    }

    writeFileSync(join(resolvedOutputDir, 'README.md'), `# Re-redlining stress fixtures

Generated by \`node scripts/export-reredline-stress-fixtures.mjs\`.

Each scenario contains five independently redlined prose paragraphs, six real
Word list paragraphs (four bullet and two numbered), and three tables covering
multi-cell updates, row insertion, and row deletion. The stages are:

- \`source\`: clean input with no tracked changes.
- \`round1\`: a heavy first review by \`Round One Reviewer\`.
- \`rerelined\`: each revised block is normalized with
  \`existingRevisions: "accept-all-first"\`, then changed again by
  \`Round Two Reviewer\` and assembled into one heavily redlined document.

The adjacent \`.document.xml\` files support direct validator/XSD inspection;
\`manifest.json\` records visible text and revision counts for every stage.
`, 'utf8');
    writeFileSync(join(resolvedOutputDir, 'manifest.json'), `${JSON.stringify(manifest, null, 2)}\n`, 'utf8');
    return manifest;
}

const isMain = process.argv[1] && resolve(process.argv[1]) === resolve(fileURLToPath(import.meta.url));
if (isMain) {
    const outputIndex = process.argv.indexOf('--output-dir');
    const outputDir = outputIndex >= 0 ? process.argv[outputIndex + 1] : join(process.cwd(), 'tests', 'fixtures', 'reredline-stress');
    if (!outputDir) throw new Error('--output-dir requires a path');
    const manifest = await generateReredlineStressFixtures(outputDir);
    console.log(`Wrote ${manifest.scenarios.length * 3} DOCX fixtures to ${resolve(outputDir)}`);
}
