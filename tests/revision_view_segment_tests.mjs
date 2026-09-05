import assert from 'node:assert/strict';
import './setup-xml-provider.mjs';
import {
    extractParagraphRevisionSegments,
    extractCanonicalParagraphText,
    inspectDocumentParts
} from '../index.js';
import { parseOoxmlSafe } from '../adapters/xml-adapter.js';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

function parseParagraph(xmlString) {
    const wrapped = `<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>${xmlString}</w:body></w:document>`;
    const { doc, error } = parseOoxmlSafe(wrapped, 'application/xml');
    if (error || !doc) throw new Error(`Parse failed: ${error?.message}`);
    return doc.getElementsByTagNameNS(W, 'p')[0];
}

// Test 1: Empty and null inputs
{
    assert.deepEqual(extractParagraphRevisionSegments(null), []);
    assert.deepEqual(extractParagraphRevisionSegments(undefined), []);

    const emptyP = parseParagraph('<w:p/>');
    assert.deepEqual(extractParagraphRevisionSegments(emptyP), []);

    const pPrOnly = parseParagraph('<w:p><w:pPr><w:pStyle w:val="Normal"/></w:pPr></w:p>');
    assert.deepEqual(extractParagraphRevisionSegments(pPrOnly), []);
}

// Test 2: Baseline paragraph with plain text, tabs, breaks, and hyphens
{
    const xml = `<w:p>
        <w:r><w:t>Hello</w:t><w:tab/><w:t>world</w:t><w:br/><w:noBreakHyphen/><w:softHyphen/><w:t>!</w:t></w:r>
    </w:p>`;
    const p = parseParagraph(xml);
    const segments = extractParagraphRevisionSegments(p);

    assert.equal(segments.length, 1);
    assert.equal(segments[0].kind, 'baseline');
    assert.equal(segments[0].text, 'Hello\tworld\n\u2011\u00ad!');
    assert.equal(segments[0].acceptedStart, 0);
    assert.equal(segments[0].rejectedStart, 0);
    assert.equal(segments[0].author, undefined);
    assert.equal(segments[0].revisionId, undefined);

    // Parity with extractCanonicalParagraphText
    assert.equal(extractCanonicalParagraphText(p, { revisionView: 'accepted' }), segments[0].text);
    assert.equal(extractCanonicalParagraphText(p, { revisionView: 'rejected' }), segments[0].text);
}

// Test 3: Replacement exposes deletion and insertion as separate segments
{
    const xml = `<w:p>
        <w:r><w:t xml:space="preserve">The </w:t></w:r>
        <w:del w:id="1" w:author="EditorA">
            <w:r><w:delText>old</w:delText></w:r>
        </w:del>
        <w:ins w:id="2" w:author="EditorB">
            <w:r><w:t>new</w:t></w:r>
        </w:ins>
        <w:r><w:t xml:space="preserve"> clause</w:t></w:r>
    </w:p>`;
    const p = parseParagraph(xml);
    const segments = extractParagraphRevisionSegments(p);

    assert.equal(segments.length, 4);

    // Segment 0: "The "
    assert.equal(segments[0].text, 'The ');
    assert.equal(segments[0].kind, 'baseline');
    assert.equal(segments[0].acceptedStart, 0);
    assert.equal(segments[0].rejectedStart, 0);

    // Segment 1: "old" (deletion)
    assert.equal(segments[1].text, 'old');
    assert.equal(segments[1].kind, 'deletion');
    assert.equal(segments[1].author, 'EditorA');
    assert.equal(segments[1].revisionId, '1');
    assert.equal(segments[1].acceptedStart, null);
    assert.equal(segments[1].rejectedStart, 4);

    // Segment 2: "new" (insertion)
    assert.equal(segments[2].text, 'new');
    assert.equal(segments[2].kind, 'insertion');
    assert.equal(segments[2].author, 'EditorB');
    assert.equal(segments[2].revisionId, '2');
    assert.equal(segments[2].acceptedStart, 4);
    assert.equal(segments[2].rejectedStart, null);

    // Segment 3: " clause"
    assert.equal(segments[3].text, ' clause');
    assert.equal(segments[3].kind, 'baseline');
    assert.equal(segments[3].acceptedStart, 7);
    assert.equal(segments[3].rejectedStart, 7);

    // Required assertions: Concatenation matches extractCanonicalParagraphText
    const acceptedText = segments.filter(s => s.acceptedStart !== null).map(s => s.text).join('');
    const rejectedText = segments.filter(s => s.rejectedStart !== null).map(s => s.text).join('');
    assert.equal(acceptedText, 'The new clause');
    assert.equal(rejectedText, 'The old clause');
    assert.equal(acceptedText, extractCanonicalParagraphText(p, { revisionView: 'accepted' }));
    assert.equal(rejectedText, extractCanonicalParagraphText(p, { revisionView: 'rejected' }));
}

// Test 4: Emoji UTF-16 code units count
{
    const xml = `<w:p>
        <w:r><w:t>Hello 🚀 world</w:t></w:r>
        <w:ins w:id="5" w:author="Alice">
            <w:r><w:t> 🎉 party</w:t></w:r>
        </w:ins>
    </w:p>`;
    const p = parseParagraph(xml);
    const segments = extractParagraphRevisionSegments(p);

    assert.equal(segments.length, 2);
    // Rocket emoji is 2 UTF-16 code units: length of "Hello 🚀 world" is 6 + 2 + 6 = 14
    assert.equal(segments[0].text.length, 14);
    assert.equal(segments[0].acceptedStart, 0);
    assert.equal(segments[0].rejectedStart, 0);

    // " 🎉 party" has party popper (2 code units) -> length 1 + 2 + 6 = 9
    assert.equal(segments[1].text.length, 9);
    assert.equal(segments[1].acceptedStart, 14);
    assert.equal(segments[1].rejectedStart, null);
}

// Test 5: Entity decoding does not alter DOM offsets
{
    const xml = `<w:p>
        <w:r><w:t>Fish &amp; Chips</w:t></w:r>
        <w:ins w:id="7" w:author="Chef">
            <w:r><w:t> &lt;tasty&gt;</w:t></w:r>
        </w:ins>
    </w:p>`;
    const p = parseParagraph(xml);
    const segments = extractParagraphRevisionSegments(p);

    assert.equal(segments[0].text, 'Fish & Chips');
    assert.equal(segments[0].text.length, 12);
    assert.equal(segments[0].acceptedStart, 0);

    assert.equal(segments[1].text, ' <tasty>');
    assert.equal(segments[1].acceptedStart, 12);
}

// Test 6: Move revisions (moveFrom and moveTo) map consistently
{
    const xml = `<w:p>
        <w:r><w:t>Start </w:t></w:r>
        <w:moveFrom w:id="10" w:author="Mover">
            <w:r><w:delText>source clause </w:delText></w:r>
        </w:moveFrom>
        <w:moveTo w:id="11" w:author="Mover">
            <w:r><w:t>destination clause </w:t></w:r>
        </w:moveTo>
        <w:r><w:t>End</w:t></w:r>
    </w:p>`;
    const p = parseParagraph(xml);
    const segments = extractParagraphRevisionSegments(p);

    assert.equal(segments.length, 4);

    // moveFrom: invisible in accepted, visible in rejected
    assert.equal(segments[1].kind, 'move_from');
    assert.equal(segments[1].author, 'Mover');
    assert.equal(segments[1].revisionId, '10');
    assert.equal(segments[1].acceptedStart, null);
    assert.equal(segments[1].rejectedStart, 6);

    // moveTo: visible in accepted, invisible in rejected
    assert.equal(segments[2].kind, 'move_to');
    assert.equal(segments[2].author, 'Mover');
    assert.equal(segments[2].revisionId, '11');
    assert.equal(segments[2].acceptedStart, 6);
    assert.equal(segments[2].rejectedStart, null);

    const acceptedText = segments.filter(s => s.acceptedStart !== null).map(s => s.text).join('');
    const rejectedText = segments.filter(s => s.rejectedStart !== null).map(s => s.text).join('');
    assert.equal(acceptedText, 'Start destination clause End');
    assert.equal(rejectedText, 'Start source clause End');
    assert.equal(acceptedText, extractCanonicalParagraphText(p, { revisionView: 'accepted' }));
    assert.equal(rejectedText, extractCanonicalParagraphText(p, { revisionView: 'rejected' }));
}

// Test 7: Multi-run merging within same container vs distinct containers
{
    // Consecutive runs inside same insertion wrapper merge
    const xmlSameIns = `<w:p>
        <w:ins w:id="20" w:author="Author1">
            <w:r><w:t>Part A </w:t></w:r>
            <w:r><w:rPr><w:b/></w:rPr><w:t>Part B</w:t></w:r>
        </w:ins>
    </w:p>`;
    const pSameIns = parseParagraph(xmlSameIns);
    const segsSame = extractParagraphRevisionSegments(pSameIns);
    assert.equal(segsSame.length, 1);
    assert.equal(segsSame[0].text, 'Part A Part B');
    assert.equal(segsSame[0].kind, 'insertion');

    // Consecutive insertions with different revision IDs do NOT merge
    const xmlDiffIns = `<w:p>
        <w:ins w:id="21" w:author="Author1">
            <w:r><w:t>Part A </w:t></w:r>
        </w:ins>
        <w:ins w:id="22" w:author="Author1">
            <w:r><w:t>Part B</w:t></w:r>
        </w:ins>
    </w:p>`;
    const pDiffIns = parseParagraph(xmlDiffIns);
    const segsDiff = extractParagraphRevisionSegments(pDiffIns);
    assert.equal(segsDiff.length, 2);
    assert.equal(segsDiff[0].revisionId, '21');
    assert.equal(segsDiff[1].revisionId, '22');

    // Hyperlink container boundary prevents merging with outside paragraph runs
    const xmlHyperlink = `<w:p>
        <w:r><w:t>Before </w:t></w:r>
        <w:hyperlink r:id="rId9"><w:r><w:t>link</w:t></w:r></w:hyperlink>
        <w:r><w:t> After</w:t></w:r>
    </w:p>`;
    const pHyperlink = parseParagraph(xmlHyperlink);
    const segsHyperlink = extractParagraphRevisionSegments(pHyperlink);
    assert.equal(segsHyperlink.length, 3);
    assert.equal(segsHyperlink[0].text, 'Before ');
    assert.equal(segsHyperlink[1].text, 'link');
    assert.equal(segsHyperlink[2].text, ' After');

    // Option mergeRuns: false keeps separate runs unmerged
    const segsNoMerge = extractParagraphRevisionSegments(pSameIns, { mergeRuns: false });
    assert.equal(segsNoMerge.length, 2);
    assert.equal(segsNoMerge[0].text, 'Part A ');
    assert.equal(segsNoMerge[1].text, 'Part B');
}

// Test 8: inspectDocumentParts returns segments for each paragraph
{
    const documentXml = `<w:document xmlns:w="${W}"><w:body>
        <w:p><w:r><w:t>First para</w:t></w:r></w:p>
        <w:p>
            <w:r><w:t>Second </w:t></w:r>
            <w:ins w:id="30" w:author="Reviewer"><w:r><w:t>revised </w:t></w:r></w:ins>
            <w:r><w:t>para</w:t></w:r>
        </w:p>
    </w:body></w:document>`;

    const inspection = inspectDocumentParts({ documentXml });
    assert.equal(inspection.status, 'ok');
    assert.equal(inspection.paragraphs.length, 2);

    assert(Array.isArray(inspection.paragraphs[0].segments));
    assert.equal(inspection.paragraphs[0].segments.length, 1);
    assert.equal(inspection.paragraphs[0].segments[0].text, 'First para');

    assert.equal(inspection.paragraphs[1].segments.length, 3);
    assert.equal(inspection.paragraphs[1].segments[0].text, 'Second ');
    assert.equal(inspection.paragraphs[1].segments[1].kind, 'insertion');
    assert.equal(inspection.paragraphs[1].segments[1].author, 'Reviewer');
    assert.equal(inspection.paragraphs[1].segments[1].revisionId, '30');
    assert.equal(inspection.paragraphs[1].segments[2].text, 'para');
}

console.log('revision view segment tests passed');
