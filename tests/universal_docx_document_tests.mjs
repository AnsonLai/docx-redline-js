import assert from 'node:assert/strict';
import { openDocx, computePackageRevisionToken } from '../document/docx-document.js';
import { zipDocx } from '../document/zip-archive.js';

const textEncoder = new TextEncoder();
const textDecoder = new TextDecoder();

// Helper to build a minimal valid docx package as pure Uint8Array
function createMinimalDocxUint8Array(paragraphText = 'Hello World') {
    const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
  <w:body>
    <w:p w14:paraId="00000001">
      <w:r><w:t>${paragraphText}</w:t></w:r>
    </w:p>
  </w:body>
</w:document>`;

    const contentTypesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`;

    const relsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

    const entries = new Map([
        ['[Content_Types].xml', textEncoder.encode(contentTypesXml)],
        ['_rels/.rels', textEncoder.encode(relsXml)],
        ['word/document.xml', textEncoder.encode(documentXml)]
    ]);

    return zipDocx(entries);
}

// 1. Open with Uint8Array, inspect, and get revision token
{
    const pkgBytes = createMinimalDocxUint8Array('Universal testing');
    assert.ok(pkgBytes instanceof Uint8Array);

    const doc = openDocx(pkgBytes);
    const token = doc.getRevisionToken();
    assert.equal(token.scope, 'package');
    assert.equal(token.algorithm, 'sha256');
    assert.equal(typeof token.value, 'string');
    assert.equal(token.value.length, 64);

    const inspection = doc.inspect();
    assert.equal(inspection.status, 'ok');
    assert.equal(inspection.paragraphs.length, 1);
    assert.equal(inspection.paragraphs[0].exactText, 'Universal testing');
    assert.equal(inspection.comments.length, 0);

    const directToken = computePackageRevisionToken(pkgBytes);
    assert.equal(directToken.value, token.value);
}

// 2. Apply operations and serialize with toUint8Array()
{
    const pkgBytes = createMinimalDocxUint8Array('Original content');
    const doc = openDocx(pkgBytes);

    const operations = [
        {
            type: 'replace',
            target: { text: 'Original content' },
            modified: 'Modified content'
        }
    ];

    const result = await doc.applyOperations(operations, {
        author: 'UniversalBot'
    });

    assert.equal(result.status, 'ok');
    assert.equal(result.written, true);
    assert.ok(result.uint8Array instanceof Uint8Array);
    assert.ok(result.toUint8Array() instanceof Uint8Array);

    // Verify document was updated in-memory
    const updatedBytes = doc.toUint8Array();
    assert.ok(updatedBytes instanceof Uint8Array);

    // Reopen updated document and verify text changed
    const reopenedDoc = openDocx(updatedBytes);
    const reopenedInspection = reopenedDoc.inspect();
    assert.equal(reopenedInspection.paragraphs[0].exactText, 'Modified content');
    assert.deepEqual(reopenedInspection.paragraphs[0].revisionAuthors, ['UniversalBot']);

    const reopenedText = textDecoder.decode(reopenedDoc.entries.get('word/document.xml'));
    assert.ok(reopenedText.includes('<w:ins'));
    assert.ok(reopenedText.includes('<w:del'));
    assert.ok(reopenedText.includes('Modified'));
}

console.log('PASS: universal docx document test suite');
