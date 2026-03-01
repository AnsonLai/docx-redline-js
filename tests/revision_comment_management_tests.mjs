import './setup-xml-provider.mjs';

import assert from 'assert';
import {
    acceptTrackedChangesInOoxml,
    rejectTrackedChangesInOoxml,
    deleteCommentsByAuthorInOoxml
} from '../index.js';

const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

function parseXmlStrict(xmlText, label) {
    const parser = new DOMParser();
    const xmlDoc = parser.parseFromString(xmlText, 'application/xml');
    const parseError = xmlDoc.getElementsByTagName('parsererror')[0];
    if (parseError) {
        throw new Error(`[XML parse error] ${label}: ${parseError.textContent || 'Unknown'}`);
    }
    return xmlDoc;
}

function buildRevisionDocXml() {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${NS_W}">
  <w:body>
    <w:p>
      <w:r><w:t>Start </w:t></w:r>
      <w:ins w:id="101" w:author="Alice" w:date="2026-01-01T00:00:00Z"><w:r><w:t>A</w:t></w:r></w:ins>
      <w:ins w:id="102" w:author="Bob" w:date="2026-01-01T00:00:00Z"><w:r><w:t>B</w:t></w:r></w:ins>
      <w:del w:id="201" w:author="Alice" w:date="2026-01-01T00:00:00Z"><w:r><w:delText>xa</w:delText></w:r></w:del>
      <w:del w:id="202" w:author="Bob" w:date="2026-01-01T00:00:00Z"><w:r><w:delText>xb</w:delText></w:r></w:del>
    </w:p>
    <w:tbl>
      <w:tr>
        <w:trPr><w:del w:id="301" w:author="Alice" w:date="2026-01-01T00:00:00Z"/></w:trPr>
        <w:tc><w:p><w:r><w:t>DeleteRow</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:trPr><w:ins w:id="302" w:author="Alice" w:date="2026-01-01T00:00:00Z"/></w:trPr>
        <w:tc><w:p><w:r><w:t>KeepRow</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>
    <w:sectPr/>
  </w:body>
</w:document>`;
}

function buildCommentPackageXml() {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
  <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document xmlns:w="${NS_W}">
        <w:body>
          <w:p>
            <w:commentRangeStart w:id="1"/>
            <w:r><w:t>Alice target</w:t></w:r>
            <w:commentRangeEnd w:id="1"/>
            <w:r><w:rPr></w:rPr><w:commentReference w:id="1"/></w:r>
          </w:p>
          <w:p>
            <w:commentRangeStart w:id="2"/>
            <w:r><w:t>Bob target</w:t></w:r>
            <w:commentRangeEnd w:id="2"/>
            <w:r><w:rPr></w:rPr><w:commentReference w:id="2"/></w:r>
          </w:p>
          <w:sectPr/>
        </w:body>
      </w:document>
    </pkg:xmlData>
  </pkg:part>
  <pkg:part pkg:name="/word/comments.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml">
    <pkg:xmlData>
      <w:comments xmlns:w="${NS_W}">
        <w:comment w:id="1" w:author="Alice"><w:p><w:r><w:t>alice comment</w:t></w:r></w:p></w:comment>
        <w:comment w:id="2" w:author="Bob"><w:p><w:r><w:t>bob comment</w:t></w:r></w:p></w:comment>
      </w:comments>
    </pkg:xmlData>
  </pkg:part>
</pkg:package>`;
}

async function testAcceptTrackedChangesBySpecificAuthor() {
    const source = buildRevisionDocXml();
    const result = acceptTrackedChangesInOoxml(source, { author: 'Alice' });

    assert.strictEqual(result.hasChanges, true, 'accept-by-author should report changes');
    assert.ok(result.acceptedCount >= 4, 'accept-by-author should count accepted markers');
    assert.ok(result.oxml.includes('<w:t>A</w:t>'), 'Alice insertion text should be kept');
    assert.ok(result.oxml.includes('<w:ins w:id="102"'), 'Bob insertion should remain tracked');
    assert.ok(result.oxml.includes('<w:del w:id="202"'), 'Bob deletion should remain tracked');
    assert.ok(!result.oxml.includes('<w:ins w:id="101"'), 'Alice insertion wrapper should be removed');
    assert.ok(!result.oxml.includes('<w:del w:id="201"'), 'Alice deletion wrapper should be removed');
    assert.ok(!result.oxml.includes('DeleteRow'), 'Alice deleted table row should be removed');
    assert.ok(result.oxml.includes('KeepRow'), 'Alice inserted table row should remain');
    assert.ok(!result.oxml.includes('<w:trPr><w:ins w:id="302"'), 'Alice row insertion marker should be removed');
}

async function testAcceptTrackedChangesAllAuthors() {
    const source = buildRevisionDocXml();
    const result = acceptTrackedChangesInOoxml(source, { allAuthors: true });

    assert.strictEqual(result.hasChanges, true, 'accept-all should report changes');
    assert.ok(result.acceptedCount >= 6, 'accept-all should count all accepted markers');
    assert.ok(!result.oxml.includes('<w:ins'), 'accept-all should remove all insertion wrappers');
    assert.ok(!result.oxml.includes('<w:del'), 'accept-all should remove all deletion wrappers');
}

async function testRejectTrackedChangesBySpecificAuthor() {
    const source = buildRevisionDocXml();
    const result = rejectTrackedChangesInOoxml(source, { author: 'Alice' });

    assert.strictEqual(result.hasChanges, true, 'reject-by-author should report changes');
    assert.ok(result.rejectedCount >= 4, 'reject-by-author should count rejected markers');
    assert.ok(!result.oxml.includes('<w:ins w:id="101"'), 'Alice insertion should be removed');
    assert.ok(result.oxml.includes('<w:ins w:id="102"'), 'Bob insertion should remain tracked');
    assert.ok(result.oxml.includes('<w:del w:id="202"'), 'Bob deletion should remain tracked');
    assert.ok(result.oxml.includes('<w:t>xa</w:t>'), 'Alice deletion text should be restored into normal text');
    assert.ok(result.oxml.includes('DeleteRow'), 'Alice deleted table row should remain');
    assert.ok(result.oxml.includes('<w:trPr/>') || result.oxml.includes('<w:trPr></w:trPr>'), 'Alice row deletion marker should be removed');
    assert.ok(!result.oxml.includes('KeepRow'), 'Alice inserted table row should be removed when rejected');
}

async function testRejectTrackedChangesAllAuthors() {
    const source = buildRevisionDocXml();
    const result = rejectTrackedChangesInOoxml(source, { allAuthors: true });

    assert.strictEqual(result.hasChanges, true, 'reject-all should report changes');
    assert.ok(result.rejectedCount >= 6, 'reject-all should count all rejected markers');
    assert.ok(!result.oxml.includes('<w:ins'), 'reject-all should remove all insertion wrappers');
    assert.ok(!result.oxml.includes('<w:del'), 'reject-all should remove all deletion wrappers');
    assert.ok(result.oxml.includes('<w:t>xa</w:t>'), 'reject-all should restore deletion text xa');
    assert.ok(result.oxml.includes('<w:t>xb</w:t>'), 'reject-all should restore deletion text xb');
    assert.ok(!result.oxml.includes('KeepRow'), 'reject-all should remove inserted table rows');
    assert.ok(result.oxml.includes('DeleteRow'), 'reject-all should keep deleted table rows');
}

async function testDeleteCommentsBySpecificAuthor() {
    const source = buildCommentPackageXml();
    const result = deleteCommentsByAuthorInOoxml(source, { author: 'Alice' });
    const doc = parseXmlStrict(result.oxml, 'delete comments by author output');

    assert.strictEqual(result.hasChanges, true, 'delete-comments-by-author should report changes');
    assert.strictEqual(result.commentsRemoved, 1, 'delete-comments-by-author should remove one comment node');
    assert.ok(result.referencesRemoved >= 3, 'delete-comments-by-author should remove related markers/references');

    const comments = Array.from(doc.getElementsByTagNameNS(NS_W, 'comment'));
    const markersStart = Array.from(doc.getElementsByTagNameNS(NS_W, 'commentRangeStart'));
    const markersEnd = Array.from(doc.getElementsByTagNameNS(NS_W, 'commentRangeEnd'));
    const refs = Array.from(doc.getElementsByTagNameNS(NS_W, 'commentReference'));

    assert.strictEqual(comments.length, 1, 'only one comment should remain');
    assert.strictEqual(comments[0].getAttribute('w:id'), '2', 'Bob comment should remain');
    assert.strictEqual(markersStart.some(node => (node.getAttribute('w:id') || node.getAttribute('id')) === '1'), false, 'Alice start marker should be removed');
    assert.strictEqual(markersEnd.some(node => (node.getAttribute('w:id') || node.getAttribute('id')) === '1'), false, 'Alice end marker should be removed');
    assert.strictEqual(refs.some(node => (node.getAttribute('w:id') || node.getAttribute('id')) === '1'), false, 'Alice comment reference should be removed');
    assert.strictEqual(markersStart.some(node => (node.getAttribute('w:id') || node.getAttribute('id')) === '2'), true, 'Bob start marker should remain');
}

async function testDeleteCommentsAllAuthors() {
    const source = buildCommentPackageXml();
    const result = deleteCommentsByAuthorInOoxml(source, { allAuthors: true });
    const doc = parseXmlStrict(result.oxml, 'delete comments all authors output');

    assert.strictEqual(result.hasChanges, true, 'delete-comments-all should report changes');
    assert.strictEqual(result.commentsRemoved, 2, 'delete-comments-all should remove all comments');
    assert.strictEqual(doc.getElementsByTagNameNS(NS_W, 'comment').length, 0, 'no comments should remain');
    assert.strictEqual(doc.getElementsByTagNameNS(NS_W, 'commentRangeStart').length, 0, 'no comment starts should remain');
    assert.strictEqual(doc.getElementsByTagNameNS(NS_W, 'commentRangeEnd').length, 0, 'no comment ends should remain');
    assert.strictEqual(doc.getElementsByTagNameNS(NS_W, 'commentReference').length, 0, 'no comment references should remain');
}

async function run() {
    await testAcceptTrackedChangesBySpecificAuthor();
    await testAcceptTrackedChangesAllAuthors();
    await testRejectTrackedChangesBySpecificAuthor();
    await testRejectTrackedChangesAllAuthors();
    await testDeleteCommentsBySpecificAuthor();
    await testDeleteCommentsAllAuthors();
    console.log('PASS: revision + comment management tests');
}

run().catch(err => {
    console.error('FAIL:', err.message);
    process.exit(1);
});
