import assert from 'assert/strict';

import { buildMinimalDocx, buildMinimalDocxEntries } from '../scripts/lib/minimal-zip.mjs';
import { unzipEntries } from '../scripts/lib/zip-reader.mjs';
import { createCommentsPart, createHeaderFooterPart, createNotesPart } from './fixtures/word-package-parts.mjs';

const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${NS_W}" xmlns:r="${NS_R}"><w:body><w:p>
<w:commentRangeStart w:id="4"/><w:r><w:t>Reviewed text</w:t></w:r><w:commentRangeEnd w:id="4"/>
<w:r><w:commentReference w:id="4"/><w:footnoteReference w:id="2"/><w:endnoteReference w:id="3"/></w:r>
<w:hyperlink r:id="rIdExternal"><w:r><w:t>Policy</w:t></w:r></w:hyperlink>
</w:p><w:sectPr><w:headerReference w:type="default" r:id="rIdHeader1"/><w:footerReference w:type="default" r:id="rIdFooter1"/></w:sectPr></w:body></w:document>`;

const commentsXml = `<w:comments xmlns:w="${NS_W}"><w:comment w:id="4" w:author="Reviewer"><w:p><w:r><w:t>Comment</w:t></w:r></w:p></w:comment></w:comments>`;
const footnotesXml = `<w:footnotes xmlns:w="${NS_W}"><w:footnote w:type="separator" w:id="-1"><w:p/></w:footnote><w:footnote w:type="continuationSeparator" w:id="0"><w:p/></w:footnote><w:footnote w:id="2"><w:p><w:r><w:t>Footnote</w:t></w:r></w:p></w:footnote></w:footnotes>`;
const endnotesXml = `<w:endnotes xmlns:w="${NS_W}"><w:endnote w:type="separator" w:id="-1"><w:p/></w:endnote><w:endnote w:type="continuationSeparator" w:id="0"><w:p/></w:endnote><w:endnote w:id="3"><w:p><w:r><w:t>Endnote</w:t></w:r></w:p></w:endnote></w:endnotes>`;
const headerXml = `<w:hdr xmlns:w="${NS_W}"><w:p><w:r><w:t>Header</w:t></w:r></w:p></w:hdr>`;
const footerXml = `<w:ftr xmlns:w="${NS_W}"><w:p><w:r><w:t>Footer</w:t></w:r></w:p></w:ftr>`;

const parts = {
    commentsXml,
    footnotesXml,
    endnotesXml,
    headers: [{ relationshipId: 'rIdHeader1', partName: 'header1.xml', xml: headerXml }],
    footers: [{ relationshipId: 'rIdFooter1', partName: 'footer1.xml', xml: footerXml }],
    externalHyperlinks: [{ relationshipId: 'rIdExternal', target: 'https://example.com/policy?version=1&lang=en' }]
};

function testComprehensivePackageIsDeterministicAndByteExact() {
    const expectedEntries = buildMinimalDocxEntries(documentXml, parts);
    const first = buildMinimalDocx(documentXml, parts);
    const second = buildMinimalDocx(documentXml, parts);
    assert.ok(first.equals(second), 'package output must be deterministic');

    const unpacked = unzipEntries(first);
    assert.deepEqual([...unpacked.keys()], expectedEntries.map(entry => entry.name));
    for (const entry of expectedEntries) {
        const expected = Buffer.isBuffer(entry.data) ? entry.data : Buffer.from(entry.data, 'utf8');
        assert.ok(unpacked.get(entry.name)?.equals(expected), `${entry.name} must remain byte-identical`);
    }

    const contentTypes = unpacked.get('[Content_Types].xml').toString('utf8');
    for (const partName of ['comments.xml', 'footnotes.xml', 'endnotes.xml', 'header1.xml', 'footer1.xml']) {
        assert.match(contentTypes, new RegExp(`/word/${partName.replace('.', '\\.')}"`));
    }

    const relationships = unpacked.get('word/_rels/document.xml.rels').toString('utf8');
    for (const id of ['rIdComments1', 'rIdFootnotes1', 'rIdEndnotes1', 'rIdHeader1', 'rIdFooter1', 'rIdExternal']) {
        assert.match(relationships, new RegExp(`Id="${id}"`));
    }
    assert.match(relationships, /TargetMode="External"/);
    assert.match(relationships, /version=1&amp;lang=en/);
}

function testInvalidPackageGraphsFailBeforeZipEmission() {
    assert.throws(
        () => buildMinimalDocx(documentXml, { ...parts, commentsXml: null }),
        /comment anchors require word\/comments\.xml/
    );
    assert.throws(
        () => buildMinimalDocx(documentXml, { ...parts, commentsXml: commentsXml.replace('w:id="4"', 'w:id="9"') }),
        /does not define referenced comment ID 4/
    );
    assert.throws(
        () => buildMinimalDocx(documentXml, { ...parts, footnotesXml: footnotesXml.replace('w:id="-1"', 'w:id="-2"') }),
        /separator ID -1/
    );
    assert.throws(
        () => buildMinimalDocx(documentXml, { ...parts, endnotesXml: endnotesXml.replace('w:id="3"', 'w:id="8"') }),
        /does not define referenced endnote ID 3/
    );
    assert.throws(
        () => buildMinimalDocx(documentXml, { ...parts, headers: [] }),
        /headerReference rIdHeader1/
    );
    assert.throws(
        () => buildMinimalDocx(documentXml, {
            ...parts,
            externalHyperlinks: [{ relationshipId: 'rIdHeader1', target: 'https://example.com' }]
        }),
        /relationship IDs must be unique/
    );
    assert.throws(
        () => buildMinimalDocx(documentXml, {
            ...parts,
            externalHyperlinks: [{ relationshipId: 'rIdExternal', target: 'file:///private/document' }]
        }),
        /unsupported external hyperlink target/
    );
}

function testReusablePartConstructors() {
    assert.match(createCommentsPart([{ id: 0, author: 'A & B', text: '<review>' }]), /A &amp; B/);
    assert.match(createCommentsPart([{ id: 0, author: 'A & B', text: '<review>' }]), /&lt;review&gt;/);
    assert.match(createNotesPart('footnote', [{ id: 1, text: 'Note' }]), /w:id="-1"/);
    assert.match(createNotesPart('endnote', [{ id: 2, text: 'Note' }]), /w:endnoteRef/);
    assert.match(createHeaderFooterPart('header', 'Agency & Office'), /Agency &amp; Office/);
    assert.throws(() => createNotesPart('margin', [{ id: 1, text: 'No' }]), /note kind/);
    assert.throws(() => createHeaderFooterPart('sidebar', 'No'), /part kind/);
}

testComprehensivePackageIsDeterministicAndByteExact();
testInvalidPackageGraphsFailBeforeZipEmission();
testReusablePartConstructors();

console.log('PASS: richer minimal DOCX package validation tests');
