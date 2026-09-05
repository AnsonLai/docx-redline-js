import assert from 'node:assert/strict';
import './setup-xml-provider.mjs';
import { applyRedlineToOxml, ingestWordOoxmlToPlainText } from '../index.js';
import { validateRedlineOoxml } from '../core/redline-validation.js';
import {
    acceptTrackedChangesInOoxml,
    rejectTrackedChangesInOoxml
} from '../services/revision-comment-management.js';
import { parseOoxmlSafe, createSerializer } from '../adapters/xml-adapter.js';
import { applySurgicalMode } from '../engine/surgical-mode.js';
import {
    applyOperationsToDocumentXml,
    applyOperationToDocumentXml
} from '../services/standalone-operation-runner.js';
import { validateDocumentOperation } from '../services/document-operation-contract.js';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

function getElements(root, tagName) {
    return Array.from(root.getElementsByTagNameNS(W, tagName));
}

// Test 1: Normal to bold run - formatting affinity (left vs right vs none)
{
    const oxml = `<w:p xmlns:w="${W}"><w:r><w:t xml:space="preserve">Normal </w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>Bold</w:t></w:r></w:p>`;
    const origText = 'Normal Bold';
    const modText = 'Normal Middle Bold';
    const serializer = createSerializer();

    // 1a: formatting: 'left' -> should NOT inherit bold
    {
        const { doc } = parseOoxmlSafe(oxml);
        const result = applySurgicalMode(doc, origText, modText, serializer, 'Alice', [], true, null, {}, {
            insertionAffinity: { formatting: 'left' }
        });
        assert.equal(result.hasChanges, true);
        const { doc: resDoc } = parseOoxmlSafe(result.oxml);
        const inss = getElements(resDoc, 'ins');
        assert.equal(inss.length, 1);
        const insRun = inss[0].firstChild;
        const b = insRun.getElementsByTagNameNS(W, 'b');
        assert.equal(b.length, 0, 'formatting: left should not inherit bold from right run');

        // Verify lifecycle equivalence
        const accepted = ingestWordOoxmlToPlainText(acceptTrackedChangesInOoxml(result.oxml, { allAuthors: true }).oxml);
        assert.equal(accepted.trim(), 'Normal Middle Bold');
        const rejected = ingestWordOoxmlToPlainText(rejectTrackedChangesInOoxml(result.oxml, { allAuthors: true }).oxml);
        assert.equal(rejected.trim(), 'Normal Bold');
    }

    // 1b: formatting: 'right' -> SHOULD inherit bold
    {
        const { doc } = parseOoxmlSafe(oxml);
        const result = applySurgicalMode(doc, origText, modText, serializer, 'Alice', [], true, null, {}, {
            insertionAffinity: { formatting: 'right' }
        });
        assert.equal(result.hasChanges, true);
        const { doc: resDoc } = parseOoxmlSafe(result.oxml);
        const inss = getElements(resDoc, 'ins');
        assert.equal(inss.length, 1);
        const insRun = inss[0].firstChild;
        const b = insRun.getElementsByTagNameNS(W, 'b');
        assert.equal(b.length, 1, 'formatting: right should inherit bold from right run');
    }

    // 1c: formatting: 'none' -> should NOT have any formatting
    {
        const { doc } = parseOoxmlSafe(oxml);
        const result = applySurgicalMode(doc, origText, modText, serializer, 'Alice', [], true, null, {}, {
            insertionAffinity: { formatting: 'none' }
        });
        assert.equal(result.hasChanges, true);
        const { doc: resDoc } = parseOoxmlSafe(result.oxml);
        const inss = getElements(resDoc, 'ins');
        assert.equal(inss.length, 1);
        const insRun = inss[0].firstChild;
        const b = insRun.getElementsByTagNameNS(W, 'b');
        assert.equal(b.length, 0, 'formatting: none should not inherit bold');
    }
}

// Test 2: Paragraph start/end with no carrier on one side
{
    const oxml = `<w:p xmlns:w="${W}"><w:r><w:rPr><w:b/></w:rPr><w:t>Text</w:t></w:r></w:p>`;
    const serializer = createSerializer();

    // 2a: Insertion at start with formatting: 'left' (no carrier on left)
    {
        const { doc } = parseOoxmlSafe(oxml);
        const result = applySurgicalMode(doc, 'Text', 'Start Text', serializer, 'Alice', [], true, null, {}, {
            insertionAffinity: { formatting: 'left' }
        });
        assert.equal(result.hasChanges, true);
        const { doc: resDoc } = parseOoxmlSafe(result.oxml);
        const inss = getElements(resDoc, 'ins');
        assert.equal(inss.length, 1);
        const insRun = inss[0].firstChild;
        const b = insRun.getElementsByTagNameNS(W, 'b');
        assert.equal(b.length, 0, 'Start insertion with formatting: left has no left carrier, so no bold');
    }

    // 2b: Insertion at start with formatting: 'right' (has right carrier with bold)
    {
        const { doc } = parseOoxmlSafe(oxml);
        const result = applySurgicalMode(doc, 'Text', 'Start Text', serializer, 'Alice', [], true, null, {}, {
            insertionAffinity: { formatting: 'right' }
        });
        assert.equal(result.hasChanges, true);
        const { doc: resDoc } = parseOoxmlSafe(result.oxml);
        const inss = getElements(resDoc, 'ins');
        assert.equal(inss.length, 1);
        const insRun = inss[0].firstChild;
        const b = insRun.getElementsByTagNameNS(W, 'b');
        assert.equal(b.length, 1, 'Start insertion with formatting: right inherits bold from right carrier');
    }

    // 2c: Insertion at end with formatting: 'right' (no carrier on right)
    {
        const { doc } = parseOoxmlSafe(oxml);
        const result = applySurgicalMode(doc, 'Text', 'Text End', serializer, 'Alice', [], true, null, {}, {
            insertionAffinity: { formatting: 'right' }
        });
        assert.equal(result.hasChanges, true);
        const { doc: resDoc } = parseOoxmlSafe(result.oxml);
        const inss = getElements(resDoc, 'ins');
        assert.equal(inss.length, 1);
        const insRun = inss[0].firstChild;
        const b = insRun.getElementsByTagNameNS(W, 'b');
        assert.equal(b.length, 0, 'End insertion with formatting: right has no right carrier, so no bold');
    }
}

// Test 3: Hyperlink inside vs outside and illegal placement
{
    const oxml = `<w:p xmlns:w="${W}"><w:r><w:t xml:space="preserve">Before </w:t></w:r><w:hyperlink r:id="rId1" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:r><w:t>Link</w:t></w:r></w:hyperlink><w:r><w:t xml:space="preserve"> After</w:t></w:r></w:p>`;
    const serializer = createSerializer();

    // 3a: Insert at start of hyperlink with hyperlink: 'inside'
    {
        const { doc } = parseOoxmlSafe(oxml);
        const result = applySurgicalMode(doc, 'Before Link After', 'Before InStart Link After', serializer, 'Alice', [], true, null, {}, {
            insertionAffinity: { hyperlink: 'inside' }
        });
        assert.equal(result.hasChanges, true);
        const { doc: resDoc } = parseOoxmlSafe(result.oxml);
        const inss = getElements(resDoc, 'ins');
        assert.equal(inss.length, 1);
        assert.equal(inss[0].parentNode.localName, 'hyperlink', 'Inserted ins must be child of hyperlink');
    }

    // 3b: Insert at start of hyperlink with hyperlink: 'outside'
    {
        const { doc } = parseOoxmlSafe(oxml);
        const result = applySurgicalMode(doc, 'Before Link After', 'Before OutStart Link After', serializer, 'Alice', [], true, null, {}, {
            insertionAffinity: { hyperlink: 'outside' }
        });
        assert.equal(result.hasChanges, true);
        const { doc: resDoc } = parseOoxmlSafe(result.oxml);
        const inss = getElements(resDoc, 'ins');
        assert.equal(inss.length, 1);
        assert.equal(inss[0].parentNode.localName, 'p', 'Inserted ins must be child of paragraph (outside hyperlink)');
        assert.equal(inss[0].nextSibling.localName, 'hyperlink', 'Inserted ins must precede hyperlink');
    }

    // 3c: Insert at end of hyperlink with hyperlink: 'inside'
    {
        const oxmlEnd = `<w:p xmlns:w="${W}"><w:r><w:t xml:space="preserve">Before </w:t></w:r><w:hyperlink r:id="rId1" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:r><w:t xml:space="preserve">Link </w:t></w:r></w:hyperlink><w:r><w:t>After</w:t></w:r></w:p>`;
        const { doc } = parseOoxmlSafe(oxmlEnd);
        const result = applySurgicalMode(doc, 'Before Link After', 'Before Link InEnd After', serializer, 'Alice', [], true, null, {}, {
            insertionAffinity: { hyperlink: 'inside' }
        });
        assert.equal(result.hasChanges, true);
        const { doc: resDoc } = parseOoxmlSafe(result.oxml);
        const inss = getElements(resDoc, 'ins');
        assert.equal(inss.length, 1);
        assert.equal(inss[0].parentNode.localName, 'hyperlink', 'Inserted ins must be child of hyperlink');
    }

    // 3d: Insert at end of hyperlink with hyperlink: 'outside'
    {
        const oxmlEnd = `<w:p xmlns:w="${W}"><w:r><w:t xml:space="preserve">Before </w:t></w:r><w:hyperlink r:id="rId1" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:r><w:t xml:space="preserve">Link </w:t></w:r></w:hyperlink><w:r><w:t>After</w:t></w:r></w:p>`;
        const { doc } = parseOoxmlSafe(oxmlEnd);
        const result = applySurgicalMode(doc, 'Before Link After', 'Before Link OutEnd After', serializer, 'Alice', [], true, null, {}, {
            insertionAffinity: { hyperlink: 'outside' }
        });
        assert.equal(result.hasChanges, true);
        const { doc: resDoc } = parseOoxmlSafe(result.oxml);
        const inss = getElements(resDoc, 'ins');
        assert.equal(inss.length, 1);
        assert.equal(inss[0].parentNode.localName, 'p', 'Inserted ins must be child of paragraph (outside hyperlink)');
        assert.equal(inss[0].previousSibling.localName, 'hyperlink', 'Inserted ins must follow hyperlink');
    }

    // 3e: Illegal outside placement from strictly interior position
    {
        const oxmlInterior = `<w:p xmlns:w="${W}"><w:r><w:t xml:space="preserve">Before </w:t></w:r><w:hyperlink r:id="rId1" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:r><w:t>First Second</w:t></w:r></w:hyperlink><w:r><w:t xml:space="preserve"> After</w:t></w:r></w:p>`;
        const { doc } = parseOoxmlSafe(oxmlInterior);
        const result = applySurgicalMode(doc, 'Before First Second After', 'Before First Inside Second After', serializer, 'Alice', [], true, null, {}, {
            insertionAffinity: { hyperlink: 'outside' }
        });
        assert.equal(result.hasChanges, false);
        assert.equal(result.status, 'error');
        assert.equal(result.error?.code, 'UNSUPPORTED_INSERTION_AFFINITY');
    }

    // 3f: Illegal inside placement when no hyperlink exists at boundary
    {
        const plainOxml = `<w:p xmlns:w="${W}"><w:r><w:t>Hello world</w:t></w:r></w:p>`;
        const { doc } = parseOoxmlSafe(plainOxml);
        const result = applySurgicalMode(doc, 'Hello world', 'Hello brave world', serializer, 'Alice', [], true, null, {}, {
            insertionAffinity: { hyperlink: 'inside' }
        });
        assert.equal(result.hasChanges, false);
        assert.equal(result.status, 'error');
        assert.equal(result.error?.code, 'UNSUPPORTED_INSERTION_AFFINITY');
    }
}

// Test 4: Bookmark range (inside vs outside)
{
    const oxml = `<w:p xmlns:w="${W}"><w:bookmarkStart w:id="0" w:name="_bm1"/><w:r><w:t>Bookmarked</w:t></w:r><w:bookmarkEnd w:id="0"/></w:p>`;
    const serializer = createSerializer();

    // 4a: Start of bookmark with bookmark: 'inside'
    {
        const { doc } = parseOoxmlSafe(oxml);
        const result = applySurgicalMode(doc, 'Bookmarked', 'Inside Bookmarked', serializer, 'Alice', [], true, null, {}, {
            insertionAffinity: { bookmark: 'inside' }
        });
        assert.equal(result.hasChanges, true);
        const { doc: resDoc } = parseOoxmlSafe(result.oxml);
        const inss = getElements(resDoc, 'ins');
        assert.equal(inss.length, 1);
        assert.equal(inss[0].previousSibling.localName, 'bookmarkStart', 'Ins must be placed after bookmarkStart');
    }

    // 4b: Start of bookmark with bookmark: 'outside'
    {
        const { doc } = parseOoxmlSafe(oxml);
        const result = applySurgicalMode(doc, 'Bookmarked', 'Outside Bookmarked', serializer, 'Alice', [], true, null, {}, {
            insertionAffinity: { bookmark: 'outside' }
        });
        assert.equal(result.hasChanges, true);
        const { doc: resDoc } = parseOoxmlSafe(result.oxml);
        const inss = getElements(resDoc, 'ins');
        assert.equal(inss.length, 1);
        assert.equal(inss[0].nextSibling.localName, 'bookmarkStart', 'Ins must be placed before bookmarkStart');
    }

    // 4c: End of bookmark with bookmark: 'inside'
    {
        const { doc } = parseOoxmlSafe(oxml);
        const result = applySurgicalMode(doc, 'Bookmarked', 'Bookmarked Inside', serializer, 'Alice', [], true, null, {}, {
            insertionAffinity: { bookmark: 'inside' }
        });
        assert.equal(result.hasChanges, true);
        const { doc: resDoc } = parseOoxmlSafe(result.oxml);
        const inss = getElements(resDoc, 'ins');
        assert.equal(inss.length, 1);
        assert.equal(inss[0].nextSibling.localName, 'bookmarkEnd', 'Ins must be placed before bookmarkEnd');
    }

    // 4d: End of bookmark with bookmark: 'outside'
    {
        const { doc } = parseOoxmlSafe(oxml);
        const result = applySurgicalMode(doc, 'Bookmarked', 'Bookmarked Outside', serializer, 'Alice', [], true, null, {}, {
            insertionAffinity: { bookmark: 'outside' }
        });
        assert.equal(result.hasChanges, true);
        const { doc: resDoc } = parseOoxmlSafe(result.oxml);
        const inss = getElements(resDoc, 'ins');
        assert.equal(inss.length, 1);
        assert.equal(inss[0].previousSibling.localName, 'bookmarkEnd', 'Ins must be placed after bookmarkEnd');
    }
}

// Test 5: Comment range (inside vs outside)
{
    const oxml = `<w:p xmlns:w="${W}"><w:commentRangeStart w:id="1"/><w:r><w:t>Commented</w:t></w:r><w:commentRangeEnd w:id="1"/></w:p>`;
    const serializer = createSerializer();

    // 5a: Start of comment with comment: 'inside'
    {
        const { doc } = parseOoxmlSafe(oxml);
        const result = applySurgicalMode(doc, 'Commented', 'Inside Commented', serializer, 'Alice', [], true, null, {}, {
            insertionAffinity: { comment: 'inside' }
        });
        assert.equal(result.hasChanges, true);
        const { doc: resDoc } = parseOoxmlSafe(result.oxml);
        const inss = getElements(resDoc, 'ins');
        assert.equal(inss.length, 1);
        assert.equal(inss[0].previousSibling.localName, 'commentRangeStart', 'Ins must follow commentRangeStart');
    }

    // 5b: Start of comment with comment: 'outside'
    {
        const { doc } = parseOoxmlSafe(oxml);
        const result = applySurgicalMode(doc, 'Commented', 'Outside Commented', serializer, 'Alice', [], true, null, {}, {
            insertionAffinity: { comment: 'outside' }
        });
        assert.equal(result.hasChanges, true);
        const { doc: resDoc } = parseOoxmlSafe(result.oxml);
        const inss = getElements(resDoc, 'ins');
        assert.equal(inss.length, 1);
        assert.equal(inss[0].nextSibling.localName, 'commentRangeStart', 'Ins must precede commentRangeStart');
    }

    // 5c: End of comment with comment: 'inside'
    {
        const { doc } = parseOoxmlSafe(oxml);
        const result = applySurgicalMode(doc, 'Commented', 'Commented Inside', serializer, 'Alice', [], true, null, {}, {
            insertionAffinity: { comment: 'inside' }
        });
        assert.equal(result.hasChanges, true);
        const { doc: resDoc } = parseOoxmlSafe(result.oxml);
        const inss = getElements(resDoc, 'ins');
        assert.equal(inss.length, 1);
        assert.equal(inss[0].nextSibling.localName, 'commentRangeEnd', 'Ins must precede commentRangeEnd');
    }

    // 5d: End of comment with comment: 'outside'
    {
        const { doc } = parseOoxmlSafe(oxml);
        const result = applySurgicalMode(doc, 'Commented', 'Commented Outside', serializer, 'Alice', [], true, null, {}, {
            insertionAffinity: { comment: 'outside' }
        });
        assert.equal(result.hasChanges, true);
        const { doc: resDoc } = parseOoxmlSafe(result.oxml);
        const inss = getElements(resDoc, 'ins');
        assert.equal(inss.length, 1);
        assert.equal(inss[0].previousSibling.localName, 'commentRangeEnd', 'Ins must follow commentRangeEnd');
    }
}

// Test 6: Field sequence preservation
{
    const oxml = `<w:p xmlns:w="${W}"><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> DATE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>2026-09-05</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>`;
    const serializer = createSerializer();

    const { doc } = parseOoxmlSafe(oxml);
    const result = applySurgicalMode(doc, '2026-09-05', '2026-09-05 (Today)', serializer, 'Alice', [], true, null, {}, {});
    assert.equal(result.hasChanges, true);

    const { doc: resDoc } = parseOoxmlSafe(result.oxml);
    const fldChars = getElements(resDoc, 'fldChar');
    assert.equal(fldChars.length, 3, 'All fldChar elements must be preserved');
    assert.equal(fldChars[0].getAttribute('w:fldCharType') || fldChars[0].getAttributeNS(W, 'fldCharType'), 'begin');
    assert.equal(fldChars[1].getAttribute('w:fldCharType') || fldChars[1].getAttributeNS(W, 'fldCharType'), 'separate');
    assert.equal(fldChars[2].getAttribute('w:fldCharType') || fldChars[2].getAttributeNS(W, 'fldCharType'), 'end');
}

// Test 7: SDT content preservation
{
    const oxml = `<w:p xmlns:w="${W}"><w:sdt><w:sdtContent><w:r><w:t>Structured text</w:t></w:r></w:sdtContent></w:sdt></w:p>`;
    const serializer = createSerializer();

    const { doc } = parseOoxmlSafe(oxml);
    const result = applySurgicalMode(doc, 'Structured text', 'Structured edited text', serializer, 'Alice', [], true, null, {}, {});
    assert.equal(result.hasChanges, true);

    const { doc: resDoc } = parseOoxmlSafe(result.oxml);
    const sdts = getElements(resDoc, 'sdt');
    assert.equal(sdts.length, 1, 'SDT wrapper must be preserved');
    const sdtContents = getElements(sdts[0], 'sdtContent');
    assert.equal(sdtContents.length, 1);
    const inss = getElements(sdtContents[0], 'ins');
    assert.equal(inss.length, 1, 'Inserted run must be inside sdtContent');
}

// Test 8: Same-author vs different-author insertion coalescing
{
    const oxml = `<w:p xmlns:w="${W}"><w:ins w:id="1" w:author="Alice" w:date="2026-09-05T12:00:00Z"><w:r><w:t>Existing</w:t></w:r></w:ins></w:p>`;
    const serializer = createSerializer();

    // 8a: Same author ('Alice') with revision: 'coalesce_same_author' -> coalesces into existing w:ins
    {
        const { doc } = parseOoxmlSafe(oxml);
        const result = applySurgicalMode(doc, 'Existing', 'Existing more', serializer, 'Alice', [], true, null, {}, {
            insertionAffinity: { revision: 'coalesce_same_author' }
        });
        assert.equal(result.hasChanges, true);
        const { doc: resDoc } = parseOoxmlSafe(result.oxml);
        const inss = getElements(resDoc, 'ins');
        assert.equal(inss.length, 1, 'Same author coalescing must reuse the single existing w:ins');
        const runs = getElements(inss[0], 'r');
        assert.equal(runs.length, 2, 'Existing w:ins should now contain both runs');
    }

    // 8b: Different author ('Bob') with revision: 'coalesce_same_author' -> separate w:ins
    {
        const { doc } = parseOoxmlSafe(oxml);
        const result = applySurgicalMode(doc, 'Existing', 'Existing other', serializer, 'Bob', [], true, null, {}, {
            insertionAffinity: { revision: 'coalesce_same_author' }
        });
        assert.equal(result.hasChanges, true);
        const { doc: resDoc } = parseOoxmlSafe(result.oxml);
        const inss = getElements(resDoc, 'ins');
        assert.equal(inss.length, 2, 'Different authors must produce separate w:ins elements');
        const authors = inss.map(i => i.getAttribute('w:author') || i.getAttributeNS(W, 'author'));
        assert.ok(authors.includes('Alice'));
        assert.ok(authors.includes('Bob'));
    }
}

// Test 9: Document Operation Contract and Standalone Runner
{
    // 9a: Contract validation passes for valid insertionAffinity
    const validOp = {
        type: 'redline',
        target: 'Sample paragraph',
        modified: 'Sample updated paragraph',
        insertionAffinity: {
            formatting: 'left',
            hyperlink: 'inside',
            revision: 'coalesce_same_author',
            bookmark: 'outside',
            comment: 'inside'
        }
    };
    assert.equal(validateDocumentOperation(validOp).valid, true);

    // 9b: Contract validation rejects invalid insertionAffinity properties
    assert.equal(validateDocumentOperation({
        ...validOp,
        insertionAffinity: { formatting: 'invalid' }
    }).valid, false);

    assert.equal(validateDocumentOperation({
        ...validOp,
        insertionAffinity: { hyperlink: 'invalid' }
    }).valid, false);

    assert.equal(validateDocumentOperation({
        ...validOp,
        insertionAffinity: { revision: 'invalid' }
    }).valid, false);

    assert.equal(validateDocumentOperation({
        ...validOp,
        insertionAffinity: 'not-an-object'
    }).valid, false);

    // 9c: Standalone runner passes insertionAffinity through
    const docXml = `<w:document xmlns:w="${W}"><w:body><w:p><w:r><w:t xml:space="preserve">Normal </w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>Bold</w:t></w:r></w:p></w:body></w:document>`;
    const batchResult = await applyOperationsToDocumentXml(docXml, [
        {
            type: 'redline',
            target: 'Normal Bold',
            modified: 'Normal Middle Bold',
            insertionAffinity: { formatting: 'left' }
        }
    ], 'Alice');

    assert.equal(batchResult.hasChanges, true);
    assert.equal(batchResult.results[0].status, 'applied');
}

console.log('All WP-08 insertion affinity tests passed successfully!');
