import { createHash } from 'crypto';
import { mkdirSync, writeFileSync } from 'fs';
import { join, resolve } from 'path';

import { configureXmlProvider } from '../adapters/xml-adapter.js';
import { validateRedlineOoxml } from '../core/redline-validation.js';
import { preprocessMarkdown } from '../pipeline/markdown-processor.js';
import {
  applyOperationToDocumentXml,
  applyOperationsToDocumentXml
} from '../services/standalone-operation-runner.js';
import { buildMinimalDocx, buildMinimalDocxEntries } from './lib/minimal-zip.mjs';
import { unzipEntries } from './lib/zip-reader.mjs';
import {
  acceptTrackedChangesInOoxml,
  rejectTrackedChangesInOoxml
} from '../services/revision-comment-management.js';
import { WORD_TASK_CASES } from '../tests/fixtures/word-task-cases.mjs';

const { DOMParser, XMLSerializer } = await import('@xmldom/xmldom');
configureXmlProvider({ DOMParser, XMLSerializer });

const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const outputArgIndex = process.argv.indexOf('--output-dir');
const requestedOutputDir = outputArgIndex >= 0 ? process.argv[outputArgIndex + 1] : null;
if (outputArgIndex >= 0 && !requestedOutputDir) throw new Error('--output-dir requires a path');
const outputDir = requestedOutputDir
  ? resolve(process.cwd(), requestedOutputDir)
  : join(process.cwd(), 'tmp', 'validation-docx');
mkdirSync(outputDir, { recursive: true });

const escapeXmlText = text => String(text)
  .replace(/&/g, '&amp;')
  .replace(/</g, '&lt;')
  .replace(/>/g, '&gt;');

const baseDocument = text => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${NS_W}">
  <w:body>
    ${String(text).split(/\r?\n/).map(paragraphText =>
      `<w:p><w:r><w:t xml:space="preserve">${escapeXmlText(paragraphText)}</w:t></w:r></w:p>`
    ).join('\n    ')}
    <w:sectPr/>
  </w:body>
</w:document>`;

const cases = WORD_TASK_CASES;

let failures = 0;

for (const testCase of cases) {
  const sourceDocumentXml = testCase.sourceDocumentXml || baseDocument(testCase.sourceText || testCase.original);
  const operationOptions = { generateRedlines: true, ...(testCase.operationOptions || {}) };
  const result = Array.isArray(testCase.batchOperations)
    ? await applyOperationsToDocumentXml(
      sourceDocumentXml,
      testCase.batchOperations,
      'Validation',
      null,
      operationOptions
    )
    : await applyOperationToDocumentXml(
      sourceDocumentXml,
      testCase.operation || { type: 'redline', target: testCase.original, modified: testCase.modified },
      'Validation',
      null,
      operationOptions
    );

  if (testCase.expectAtomicRollback) {
    if (
      result?.hasChanges ||
      result?.rolledBack !== true ||
      result?.error?.code !== 'BATCH_OPERATION_FAILED' ||
      result?.documentXml !== sourceDocumentXml
    ) {
      console.error(`FAIL ${testCase.name}: expected an atomic batch rollback`);
      failures++;
      continue;
    }
  } else if ((!result?.hasChanges && !testCase.expectNoOp) || result?.status === 'error') {
    console.error(`FAIL ${testCase.name}: redline did not apply (status=${result?.status}, error=${result?.error?.message})`);
    failures++;
    continue;
  }
  if (testCase.expectNoOp && (result?.hasChanges || result?.status !== 'no-op' || result?.documentXml !== sourceDocumentXml)) {
    console.error(`FAIL ${testCase.name}: expected a byte-identical no-op preserving prior revisions`);
    failures++;
    continue;
  }

  if (Number.isInteger(testCase.maxRevisionId)) {
    const resultDoc = new DOMParser().parseFromString(result.documentXml, 'application/xml');
    const revisionNames = ['ins', 'del', 'moveFrom', 'moveTo', 'rPrChange', 'pPrChange', 'cellIns', 'cellDel'];
    const revisionIds = revisionNames
      .flatMap(name => Array.from(resultDoc.getElementsByTagNameNS(NS_W, name)))
      .map(node => Number.parseInt(node.getAttribute('w:id') || node.getAttribute('id') || '', 10))
      .filter(Number.isFinite);
    if (revisionIds.length === 0 || revisionIds.some(id => id > testCase.maxRevisionId)) {
      console.error(`FAIL ${testCase.name}: revision IDs exceeded ${testCase.maxRevisionId}: ${revisionIds.join(', ')}`);
      failures++;
      continue;
    }
  }

  if (testCase.requiredElements) {
    const resultDoc = new DOMParser().parseFromString(result.documentXml, 'application/xml');
    let missingRequiredElement = false;
    for (const [localName, minimumCount] of Object.entries(testCase.requiredElements)) {
      const actualCount = resultDoc.getElementsByTagNameNS(NS_W, localName).length;
      if (actualCount < minimumCount) {
        console.error(`FAIL ${testCase.name}: expected at least ${minimumCount} w:${localName} element(s), found ${actualCount}`);
        failures++;
        missingRequiredElement = true;
      }
    }
    if (missingRequiredElement) continue;
  }

  if (testCase.requiredNumberingFormats) {
    const numberingXml = testCase.packageParts?.numberingXml || result.numberingXml || '';
    for (const format of testCase.requiredNumberingFormats) {
      if (!numberingXml.includes(`<w:numFmt w:val="${format}"`)) {
        console.error(`FAIL ${testCase.name}: required numbering format ${format} is missing`);
        failures++;
        continue;
      }
    }
  }

  if (testCase.requiredElementParents || testCase.requiredElementText) {
    const resultDoc = new DOMParser().parseFromString(result.documentXml, 'application/xml');
    let structuralRequirementFailed = false;

    for (const [localName, parentLocalName] of Object.entries(testCase.requiredElementParents || {})) {
      const nodes = Array.from(resultDoc.getElementsByTagNameNS(NS_W, localName));
      const invalidNodes = nodes.filter(node => node.parentNode?.namespaceURI !== NS_W || node.parentNode?.localName !== parentLocalName);
      if (invalidNodes.length > 0) {
        console.error(`FAIL ${testCase.name}: ${invalidNodes.length} w:${localName} element(s) were not direct children of w:${parentLocalName}`);
        failures++;
        structuralRequirementFailed = true;
      }
    }

    for (const [localName, expectedTexts] of Object.entries(testCase.requiredElementText || {})) {
      const actualTexts = Array.from(resultDoc.getElementsByTagNameNS(NS_W, localName), node => node.textContent || '');
      if (JSON.stringify(actualTexts) !== JSON.stringify(expectedTexts)) {
        console.error(`FAIL ${testCase.name}: w:${localName} text mismatch; expected ${JSON.stringify(expectedTexts)}, found ${JSON.stringify(actualTexts)}`);
        failures++;
        structuralRequirementFailed = true;
      }
    }

    if (structuralRequirementFailed) continue;
  }

  const validation = validateRedlineOoxml(result.documentXml);
  const validationErrors = validation.issues.filter(issue => issue.severity === 'error');
  if (validationErrors.length > 0) {
    console.error(`FAIL ${testCase.name}: validateRedlineOoxml reported ${validationErrors.map(issue => issue.code).join(', ')}`);
    failures++;
    continue;
  }

  writeFileSync(join(outputDir, `${testCase.name}.document.xml`), result.documentXml, 'utf8');
  if (result.numberingXml) {
    writeFileSync(join(outputDir, `${testCase.name}.numbering.xml`), result.numberingXml, 'utf8');
  }

  const packageParts = {
    numberingXml: result.numberingXml || null,
    ...(testCase.packageParts || {})
  };
  let packageEntries;
  let docx;
  try {
    packageEntries = buildMinimalDocxEntries(result.documentXml, packageParts);
    docx = buildMinimalDocx(result.documentXml, packageParts);
  } catch (error) {
    console.error(`FAIL ${testCase.name}: package validation failed: ${error.message}`);
    failures++;
    continue;
  }

  const unpacked = unzipEntries(docx);
  const untouchedPartSha256 = {};
  for (const entry of packageEntries.filter(item => /^word\/(?:comments|footnotes|endnotes|header[0-9]+|footer[0-9]+)\.xml$/.test(item.name))) {
    const expectedBytes = Buffer.isBuffer(entry.data) ? entry.data : Buffer.from(entry.data, 'utf8');
    const actualBytes = unpacked.get(entry.name);
    if (!actualBytes?.equals(expectedBytes)) {
      console.error(`FAIL ${testCase.name}: packaged ${entry.name} was not byte-identical to the configured source part`);
      failures++;
      docx = null;
      break;
    }
    untouchedPartSha256[entry.name] = createHash('sha256').update(actualBytes).digest('hex');
  }
  if (!docx) continue;
  writeFileSync(join(outputDir, `${testCase.name}.docx`), docx);

  const sourceDocx = buildMinimalDocx(sourceDocumentXml, testCase.packageParts || {});
  const acceptedXml = acceptTrackedChangesInOoxml(result.documentXml, { allAuthors: true }).oxml;
  const rejectedXml = rejectTrackedChangesInOoxml(result.documentXml, { allAuthors: true }).oxml;
  const acceptedDocx = buildMinimalDocx(acceptedXml, packageParts);
  const rejectedDocx = buildMinimalDocx(rejectedXml, packageParts);
  writeFileSync(join(outputDir, `${testCase.name}.source.docx`), sourceDocx);
  writeFileSync(join(outputDir, `${testCase.name}.accepted.docx`), acceptedDocx);
  writeFileSync(join(outputDir, `${testCase.name}.rejected.docx`), rejectedDocx);

  // Expected text is derived from edit *intent*, not from this library's
  // accept/reject transforms, so external consumers (Word COM, LibreOffice)
  // act as independent oracles.
  const expected = {
    name: testCase.name,
    category: testCase.category,
    task: testCase.task,
    coverageMetadata: testCase.coverageMetadata,
    textFidelity: testCase.textFidelity || 'exact',
    assertionMode: testCase.assertionMode || 'exact',
    expectedAcceptedText: testCase.expectedAcceptedText ?? preprocessMarkdown(testCase.modified).cleanText,
    expectedRejectedText: testCase.expectedRejectedText ?? testCase.original,
    ...(testCase.assertionMode === 'contains' ? {
      expectedAcceptedContains: testCase.expectedAcceptedContains || [],
      expectedAcceptedAbsent: testCase.expectedAcceptedAbsent || [],
      expectedRejectedContains: testCase.expectedRejectedContains || [],
      expectedRejectedAbsent: testCase.expectedRejectedAbsent || []
    } : {}),
    sourceText: testCase.sourceText || testCase.original,
    modifiedText: preprocessMarkdown(testCase.modified).cleanText,
    requiredNumberingFormats: testCase.requiredNumberingFormats || [],
    untouchedPartSha256
  };
  writeFileSync(join(outputDir, `${testCase.name}.expected.json`), `${JSON.stringify(expected, null, 2)}\n`, 'utf8');

  console.log(`wrote ${testCase.name}: source, tracked, accepted, rejected, XML, and expectations`);
}

writeFileSync(join(outputDir, 'suite.json'), `${JSON.stringify({
  name: 'English legal and administrative Word differential suite',
  cases: cases.map(testCase => testCase.name)
}, null, 2)}\n`, 'utf8');

writeFileSync(join(outputDir, 'README.md'), `# Validation Fixtures

Generated by \`node scripts/export-validation-fixtures.mjs\`.

Each case produces:

- \`<name>.document.xml\` — the generated \`word/document.xml\` payload (for
  XSD validation and manual inspection).
- \`<name>.docx\` — a minimal package assembled by release tooling only (the
  published library still has no zip dependency).
- \`<name>.source.docx\`, \`<name>.accepted.docx\`, and
  \`<name>.rejected.docx\` — comparison states for the local HTML dashboard.
- \`<name>.expected.json\` — the accept-all / reject-all plain-text outcomes
  derived from edit intent, used by external-consumer differential checks.

Validation entry points:

- Word (differential accept/reject): \`npm run smoke:word:diff\`
- LibreOffice parse check: \`soffice --headless --convert-to pdf *.docx\`
- Schema check: \`xmllint --noout --schema wml.xsd *.document.xml\`
  (transitional schemas from ECMA-376 Part 4; see docs/VALIDATION.md)
`, 'utf8');

if (failures > 0) {
  console.error(`\n${failures} fixture case(s) failed.`);
  process.exit(1);
}

console.log(`Wrote validation fixtures to ${outputDir}`);
