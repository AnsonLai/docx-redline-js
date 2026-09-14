import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { buildZip } from '../scripts/lib/minimal-zip.mjs';
import { executeCli } from '../node/cli.js';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const contentTypes = `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`;
const rels = `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`;

const p = (id, text, props = '') => `<w:p w:paraId="${id}">${props}<w:r><w:t>${text}</w:t></w:r></w:p>`;
const delP = (id, text, markId, delId, author = 'Reviewer A') =>
  `<w:p w:paraId="${id}"><w:pPr><w:rPr><w:del w:id="${markId}" w:author="${author}" w:date="2026-02-01T00:00:00Z"/></w:rPr></w:pPr>` +
  `<w:del w:id="${delId}" w:author="${author}" w:date="2026-02-01T00:00:00Z"><w:r><w:delText>${text}</w:delText></w:r></w:del></w:p>`;

// Masked contract fixture: "Acme Cloud Services Agreement"
const documentXml = `<w:document xmlns:w="${W}"><w:body>
  ${p('00000001', '3. DATA PROTECTION AND CONFIDENTIALITY', '<w:pPr><w:pStyle w:val="Heading1"/></w:pPr>')}
  ${p('00000002', '3.1 Security Standards', '<w:pPr><w:pStyle w:val="Heading2"/></w:pPr>')}
  ${p('00000003', 'Acme shall maintain reasonable administrative, technical, and physical safeguards.')}
  ${delP('00000004', '3.2 Prohibited Data Types', '101', '102')}
  ${delP('00000005', 'Customer will not submit or process payment card information, health records subject to HIPAA, or social security numbers.', '103', '104')}
  ${p('00000006', '4. FEES AND INVOICING', '<w:pPr><w:pStyle w:val="Heading1"/></w:pPr>')}
  ${p('00000007', '4.1 Payment Terms', '<w:pPr><w:pStyle w:val="Heading2"/></w:pPr>')}
  ${p('00000008', 'Fees are due thirty (30) days from the invoice date.')}
  <w:sectPr/>
</w:body></w:document>`;

const fixture = buildZip([
    { name: '[Content_Types].xml', data: contentTypes },
    { name: 'word/document.xml', data: documentXml },
    { name: 'word/_rels/document.xml.rels', data: rels }
]);

const directory = await mkdtemp(path.join(tmpdir(), 'docx-deleted-clause-'));

try {
    const inputPath = path.join(directory, 'contract.docx');
    await writeFile(inputPath, fixture);

    // --- Scenario A: Agent searches for deleted clause in default accepted view ---
    // The clause was deleted, so default accepted view returns 0 matches.
    const searchHeadingAccepted = await executeCli(['extract', inputPath, '--search', '3.2', '--around', '2']);
    assert.equal(searchHeadingAccepted.status, 'ok');
    assert.equal(searchHeadingAccepted.selection.totalMatches, 0);
    assert.equal(searchHeadingAccepted.paragraphs.length, 0);
    // Tweak 1 verification: smart hint guides agent to --view rejected
    assert.ok(searchHeadingAccepted.selection.hint, 'Expected selection.hint to guide agent to rejected view');
    assert.match(searchHeadingAccepted.selection.hint, /0 matches in accepted view, but 1 match found in --view rejected/i);
    assert.match(searchHeadingAccepted.selection.hint, /re-run with --view rejected/i);

    // Also search for deleted body text substring
    const searchBodyAccepted = await executeCli(['extract', inputPath, '--search', 'payment card information', '--around', '1']);
    assert.equal(searchBodyAccepted.selection.totalMatches, 0);
    assert.ok(searchBodyAccepted.selection.hint);
    assert.match(searchBodyAccepted.selection.hint, /0 matches in accepted view, but 1 match found in --view rejected/i);

    // --- Scenario B: Agent follows hint and searches with --view rejected ---
    const searchRejected = await executeCli(['extract', inputPath, '--search', '3.2', '--view', 'rejected', '--around', '2']);
    assert.equal(searchRejected.status, 'ok');
    assert.equal(searchRejected.selection.totalMatches, 1);
    assert.equal(searchRejected.paragraphs.length, 5); // paragraphs 2, 3, 4 (match), 5, 6

    const matchPara = searchRejected.paragraphs.find(para => para.selectionRole === 'match');
    assert.ok(matchPara);
    assert.equal(matchPara.index, 4);
    assert.equal(matchPara.paragraphId, '00000004');
    assert.equal(matchPara.exactText, '3.2 Prohibited Data Types');
    assert.equal(matchPara.revisionView, 'rejected');

    const bodyPara = searchRejected.paragraphs.find(para => para.index === 5);
    assert.ok(bodyPara);
    assert.equal(bodyPara.paragraphId, '00000005');
    assert.equal(bodyPara.exactText, 'Customer will not submit or process payment card information, health records subject to HIPAA, or social security numbers.');

    // --- Scenario C: One-line verbatim and localized restore shortcuts ---
    const verbatimOutput = path.join(directory, 'verbatim.docx');
    const verbatimResult = await executeCli([
        'apply', inputPath,
        '--restore',
        '--target-id', matchPara.paragraphId,
        '--profile', 'agent',
        '--output', verbatimOutput
    ]);
    assert.equal(verbatimResult.status, 'ok', JSON.stringify(verbatimResult));
    assert.equal(verbatimResult.completion, true);
    assert.equal((await executeCli(['extract', verbatimOutput, '--search', '3.2'])).selection.totalMatches, 1);

    const localizedRestoreOutput = path.join(directory, 'localized-restore.docx');
    const localizedRestore = await executeCli([
        'apply', inputPath,
        '--restore',
        '--target-id', bodyPara.paragraphId,
        '--find', 'HIPAA',
        '--replace', 'applicable health-information law',
        '--profile', 'agent',
        '--output', localizedRestoreOutput
    ]);
    assert.equal(localizedRestore.status, 'ok', JSON.stringify(localizedRestore));
    assert.equal(localizedRestore.completion, true);
    assert.equal(localizedRestore.results[0].change.verification.acceptedViewMatchesCompiledText, true);
    const restoredBody = await executeCli(['extract', localizedRestoreOutput, '--search', 'applicable health-information law']);
    assert.equal(restoredBody.selection.totalMatches, 1);

    const weakRestore = await executeCli(['apply', inputPath, '--restore', '--target-ref', '4']);
    assert.equal(weakRestore.status, 'error');
    assert.equal(weakRestore.error.code, 'INVALID_OPERATION');

    // --- Scenario D: Restore the deleted clause with wordsmithing ---
    const operations = [
        {
            type: 'restore',
            target: {
                exactText: matchPara.exactText,
                paragraphId: matchPara.paragraphId,
                revisionView: 'rejected'
            },
            modified: '3.2 Prohibited Data Types'
        },
        {
            type: 'restore',
            target: {
                exactText: bodyPara.exactText,
                paragraphId: bodyPara.paragraphId,
                revisionView: 'rejected'
            },
            modified: 'Customer will not submit payment card information or health records without prior written agreement.'
        }
    ];

    const opsPath = path.join(directory, 'operations.json');
    const outputPath = path.join(directory, 'reviewed.docx');
    await writeFile(opsPath, JSON.stringify(operations, null, 2));

    const applyResult = await executeCli([
        'apply', inputPath,
        '--operations', opsPath,
        '--profile', 'agent',
        '--existing-revisions', 'slice-cross-author',
        '--output', outputPath
    ]);

    assert.equal(applyResult.status, 'ok');
    assert.equal(applyResult.completion, true);
    assert.equal(applyResult.written, true);
    assert.equal(applyResult.results.length, 2);
    assert.equal(applyResult.results[0].status, 'applied');
    assert.equal(applyResult.results[1].status, 'applied');
    assert.equal(applyResult.validation.generatedIssues.errors, 0);

    // --- Scenario E: Verify accepted view now contains the restored clause ---
    const verifyAccepted = await executeCli(['extract', outputPath, '--search', '3.2', '--around', '1']);
    assert.equal(verifyAccepted.status, 'ok');
    assert.equal(verifyAccepted.selection.totalMatches, 1);
    const restoredPara = verifyAccepted.paragraphs.find(para => para.selectionRole === 'match');
    assert.ok(restoredPara);
    assert.equal(restoredPara.exactText, '3.2 Prohibited Data Types');

    // --- Scenario F: Tweak 3 verification (Documentation links in CLI help) ---
    const applyHelp = await executeCli(['apply', '--help']);
    assert.equal(applyHelp.status, 'ok');
    assert.ok(Array.isArray(applyHelp.documentation));
    assert(applyHelp.documentation.length >= 4);
    // Ensure all documentation links are absolute repository URLs, not bare relative filenames
    for (const docUrl of applyHelp.documentation) {
        assert.match(docUrl, /^https:\/\/github\.com\/AnsonLai\/docx-redline-js\/blob\/main\//);
    }

    const globalHelp = await executeCli(['--help']);
    for (const docUrl of globalHelp.documentation) {
        assert.match(docUrl, /^https:\/\/github\.com\/AnsonLai\/docx-redline-js\/blob\/main\//);
    }
} finally {
    await rm(directory, { recursive: true, force: true });
}

console.log('agent CLI deleted clause discovery tests passed');
