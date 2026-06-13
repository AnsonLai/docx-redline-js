import { mkdirSync, writeFileSync } from 'fs';
import { join } from 'path';

import { configureXmlProvider } from '../adapters/xml-adapter.js';
import { applyOperationToDocumentXml } from '../services/standalone-operation-runner.js';

const { DOMParser, XMLSerializer } = await import('@xmldom/xmldom');
configureXmlProvider({ DOMParser, XMLSerializer });

const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const outputDir = join(process.cwd(), 'tmp', 'validation-docx');
mkdirSync(outputDir, { recursive: true });

const baseDocument = text => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${NS_W}">
  <w:body>
    <w:p><w:r><w:t>${text}</w:t></w:r></w:p>
    <w:sectPr/>
  </w:body>
</w:document>`;

const cases = [
  {
    name: 'simple-redline',
    documentXml: baseDocument('The old sentence.'),
    operation: { type: 'redline', target: 'The old sentence.', modified: 'The new sentence.' }
  },
  {
    name: 'paragraph-insert',
    documentXml: baseDocument('one'),
    operation: { type: 'redline', target: 'one', modified: 'one\ntwo' }
  },
  {
    name: 'format-only',
    documentXml: baseDocument('Make word bold'),
    operation: { type: 'redline', target: 'Make word bold', modified: 'Make **word** bold' }
  }
];

for (const testCase of cases) {
  const result = await applyOperationToDocumentXml(
    testCase.documentXml,
    testCase.operation,
    'Validation',
    null,
    { generateRedlines: true }
  );
  writeFileSync(join(outputDir, `${testCase.name}.document.xml`), result.documentXml, 'utf8');
  if (result.numberingXml) {
    writeFileSync(join(outputDir, `${testCase.name}.numbering.xml`), result.numberingXml, 'utf8');
  }
}

writeFileSync(join(outputDir, 'README.md'), `# Validation Fixtures

This folder contains generated OOXML parts for release-time validation.

The script writes document XML rather than complete .docx packages because this
package intentionally does not add a zip dependency.

To manually inspect these fixtures:

1. Copy a generated *.document.xml file into a minimal .docx package as word/document.xml.
2. Include any matching *.numbering.xml as word/numbering.xml.
3. Open with Word, LibreOffice, or another OOXML consumer.
`, 'utf8');

console.log(`Wrote validation fixtures to ${outputDir}`);
