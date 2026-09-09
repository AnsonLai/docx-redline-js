import assert from 'node:assert/strict';

import './setup-xml-provider.mjs';
import {
    acceptTrackedChangesInOoxml,
    rejectTrackedChangesInOoxml,
    validateRedlineOoxml
} from '../index.js';
import { parseOoxmlSafe } from '../adapters/xml-adapter.js';
import { extractCanonicalParagraphText } from '../core/paragraph-text.js';
import {
    applyOperationToDocumentXml,
    applyOperationsToDocumentXml
} from '../services/standalone-operation-runner.js';
import { openDocx } from '../node/index.js';
import { buildZip } from '../scripts/lib/minimal-zip.mjs';
import { unzipEntries } from '../scripts/lib/zip-reader.mjs';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const A = 'Reviewer A';
const B = 'Reviewer B';
const DATE = '2026-09-09T12:00:00Z';

function documentXml(body) {
    return `<w:document xmlns:w="${W}" xmlns:w14="${W14}"><w:body>${body}<w:sectPr/></w:body></w:document>`;
}

function docxPackage(xml) {
    const contentTypes = '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        + '<Default Extension="xml" ContentType="application/xml"/>'
        + '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        + '</Types>';
    return buildZip([
        { name: '[Content_Types].xml', data: contentTypes },
        { name: 'word/document.xml', data: xml }
    ]);
}

function deletedParagraph(text, paraId = 'DEAD0001', contentId = 11, markId = 10) {
    return `<w:p w14:paraId="${paraId}">`
        + `<w:pPr><w:rPr><w:del w:id="${markId}" w:author="${A}" w:date="${DATE}"/></w:rPr></w:pPr>`
        + `<w:del w:id="${contentId}" w:author="${A}" w:date="${DATE}"><w:r><w:rPr><w:b/>`
        + `<w:rPrChange w:id="${contentId + 1}" w:author="${A}" w:date="${DATE}"><w:rPr/></w:rPrChange></w:rPr>`
        + `<w:delText>${text}</w:delText></w:r></w:del></w:p>`;
}

function deletedParagraphWithContent(content, paraId = 'DEAD0001', contentId = 11, markId = 10, anchors = '') {
    return `<w:p w14:paraId="${paraId}"><w:pPr><w:rPr>`
        + `<w:del w:id="${markId}" w:author="${A}" w:date="${DATE}"/>`
        + `</w:rPr></w:pPr>${anchors}<w:del w:id="${contentId}" w:author="${A}" w:date="${DATE}">`
        + `${content}</w:del></w:p>`;
}

function textVector(xml) {
    return paragraphs(xml).map(paragraph => extractCanonicalParagraphText(paragraph));
}

function parse(xml) {
    const parsed = parseOoxmlSafe(xml, 'application/xml');
    assert(parsed.doc, parsed.error?.message);
    return parsed.doc;
}

function paragraphs(xml) {
    return Array.from(parse(xml).getElementsByTagNameNS(W, 'p'));
}

function revisionIds(xml) {
    const doc = parse(xml);
    return ['ins', 'del', 'rPrChange', 'pPrChange'].flatMap(name => (
        Array.from(doc.getElementsByTagNameNS(W, name))
            .map(node => node.getAttribute('w:id'))
            .filter(Boolean)
    ));
}

console.log('--- WP09c-e validation, identity, and rejected-view insertion ---');

// A dirty baseline is allowed when the requested restoration adds no new error.
{
    const dirty = `<w:p><w:ins w:id="90" w:author="Legacy" w:date="${DATE}"><w:r><w:t>One</w:t></w:r></w:ins>`
        + `<w:ins w:id="90" w:author="Legacy" w:date="${DATE}"><w:r><w:t>Two</w:t></w:r></w:ins></w:p>`;
    const input = documentXml(dirty + deletedParagraph('Archived clause.') + '<w:p><w:r><w:t>Following text.</w:t></w:r></w:p>');
    assert.equal(validateRedlineOoxml(input).valid, false, 'fixture must carry a pre-existing validation error');
    const result = await applyOperationToDocumentXml(input, {
        type: 'restore',
        target: { paragraphId: 'DEAD0001' },
        modified: 'Restated clause.',
        author: B
    }, B, null, { strictTargets: true });
    assert.equal(result.status, 'ok', JSON.stringify({ error: result.error, results: result.results }));
    assert.equal(result.hasChanges, true);

    const facade = await openDocx(docxPackage(input)).applyOperations([{
        type: 'restore',
        target: { paragraphId: 'DEAD0001' },
        modified: 'Restated clause.',
        author: B
    }], { author: B, atomic: true, strictTargets: true });
    assert.equal(facade.written, true, JSON.stringify(facade.error));
    assert(facade.validation.originalIssues.some(issue => issue.code === 'DUPLICATE_REVISION_ID'));
    assert.deepEqual(facade.validation.generatedIssues, []);
}

// Rejected-view insertion splits Reviewer A's deletion around a top-level Reviewer B insertion.
{
    const sourceText = 'The account holder must pay each invoice.';
    const input = documentXml(deletedParagraph(sourceText) + '<w:p><w:r><w:t>Following retained text.</w:t></w:r></w:p>');
    const operation = {
        type: 'insert',
        target: { paragraphId: 'DEAD0001', revisionView: 'rejected' },
        anchor: { exactText: 'must pay', occurrence: 1, offset: 5 },
        modified: '[clarification] ',
        author: B,
        existingRevisions: 'slice-cross-author'
    };
    const result = await applyOperationToDocumentXml(input, operation, B, null, { strictTargets: true });
    assert.equal(result.status, 'ok', JSON.stringify(result.error));
    assert.equal(result.operationType, 'rejected-insert');
    const paragraph = paragraphs(result.documentXml)[0];
    const order = Array.from(paragraph.childNodes)
        .filter(node => node.nodeType === 1)
        .map(node => node.localName);
    assert.deepEqual(order, ['pPr', 'del', 'ins', 'del']);
    const topLevel = Array.from(paragraph.childNodes).filter(node => ['del', 'ins'].includes(node.localName));
    assert.deepEqual(topLevel.map(node => node.getAttribute('w:author')), [A, B, A]);
    assert.equal(extractCanonicalParagraphText(paragraph, { revisionView: 'rejected' }), sourceText);
    assert(extractCanonicalParagraphText(paragraph, { revisionView: 'accepted' }).includes('[clarification] '));
    const ids = revisionIds(result.documentXml);
    assert.equal(new Set(ids).size, ids.length, 'split carriers and cloned run history need distinct IDs');
    assert.deepEqual(validateRedlineOoxml(result.documentXml).issues.filter(issue => issue.severity === 'error'), []);

    const rejectB = rejectTrackedChangesInOoxml(result.documentXml, { author: B });
    assert(!rejectB.oxml.includes('[clarification]'));
    assert(rejectB.oxml.includes(`w:author="${A}"`));
    const rejectA = rejectTrackedChangesInOoxml(result.documentXml, { author: A });
    assert(extractCanonicalParagraphText(paragraphs(rejectA.oxml)[0]).includes('must [clarification] pay'));
    const rejectAll = rejectTrackedChangesInOoxml(result.documentXml, { allAuthors: true });
    assert.equal(extractCanonicalParagraphText(paragraphs(rejectAll.oxml)[0]), sourceText);
    assert(!rejectAll.oxml.includes('[clarification]'));
    const acceptAll = acceptTrackedChangesInOoxml(result.documentXml, { allAuthors: true });
    assert(acceptAll.oxml.includes('[clarification]'));
    assert(!acceptAll.oxml.includes('account holder'));
    const acceptA = acceptTrackedChangesInOoxml(result.documentXml, { author: A });
    assert(acceptA.oxml.includes('[clarification]'));
    assert(!acceptA.oxml.includes('account holder'));
    const acceptB = acceptTrackedChangesInOoxml(result.documentXml, { author: B });
    assert(extractCanonicalParagraphText(paragraphs(acceptB.oxml)[0]).includes('[clarification]'));
    assert(!extractCanonicalParagraphText(paragraphs(acceptB.oxml)[0]).includes('account holder'));

    const missingPolicy = await applyOperationToDocumentXml(input, {
        ...operation,
        existingRevisions: undefined
    }, B, null, { strictTargets: true });
    assert.equal(missingPolicy.status, 'error');
    assert.equal(missingPolicy.error.code, 'INVALID_OPERATION');

    const missingAnchor = await applyOperationToDocumentXml(input, {
        ...operation,
        anchor: undefined
    }, B, null, { strictTargets: true });
    assert.equal(missingAnchor.status, 'error');
    assert.equal(missingAnchor.error.code, 'INVALID_OPERATION');

    const facade = await openDocx(docxPackage(input)).applyOperations([operation], {
        author: B,
        atomic: true,
        strictTargets: true
    });
    assert.equal(facade.written, true, JSON.stringify(facade.error));
    assert.equal(facade.results[0].status, 'applied');
    assert.deepEqual(facade.validation.generatedIssues, []);
    const facadeDocument = unzipEntries(facade.toBuffer()).get('word/document.xml').toString();
    const facadeOrder = Array.from(paragraphs(facadeDocument)[0].childNodes)
        .filter(node => node.nodeType === 1)
        .map(node => node.localName);
    assert.deepEqual(facadeOrder, ['pPr', 'del', 'ins', 'del']);
}

// Oracle A boundary matrix: carrier start/end, run boundary, and formatted interior.
{
    const content = '<w:r><w:rPr><w:b/></w:rPr><w:delText xml:space="preserve">Alpha </w:delText></w:r>'
        + '<w:r><w:rPr><w:i/></w:rPr><w:delText>beta</w:delText></w:r>'
        + '<w:r><w:rPr><w:u w:val="single"/></w:rPr><w:delText xml:space="preserve"> gamma</w:delText></w:r>';
    const cases = [
        { label: 'start', anchor: { exactText: 'Alpha beta gamma', occurrence: 1, offset: 0 }, order: ['pPr', 'ins', 'del'] },
        { label: 'end', anchor: { exactText: 'Alpha beta gamma', occurrence: 1, offset: 16 }, order: ['pPr', 'del', 'ins'] },
        { label: 'run boundary', anchor: { exactText: 'Alpha beta', occurrence: 1, offset: 6 }, order: ['pPr', 'del', 'ins', 'del'] },
        { label: 'formatted interior', anchor: { exactText: 'beta', occurrence: 1, offset: 2 }, order: ['pPr', 'del', 'ins', 'del'] }
    ];
    for (const testCase of cases) {
        const input = documentXml(
            deletedParagraphWithContent(content)
            + '<w:p><w:r><w:t>Following retained text.</w:t></w:r></w:p>'
        );
        const result = await applyOperationToDocumentXml(input, {
            type: 'insert',
            target: { paragraphId: 'DEAD0001', revisionView: 'rejected' },
            anchor: testCase.anchor,
            modified: '[B]',
            author: B,
            existingRevisions: 'slice-cross-author'
        }, B, null, { strictTargets: true });
        assert.equal(result.status, 'ok', `${testCase.label}: ${JSON.stringify(result.error)}`);
        const paragraph = paragraphs(result.documentXml)[0];
        assert.deepEqual(
            Array.from(paragraph.childNodes).filter(node => node.nodeType === 1).map(node => node.localName),
            testCase.order,
            testCase.label
        );
        assert.equal(extractCanonicalParagraphText(paragraph, { revisionView: 'rejected' }), 'Alpha beta gamma');
        assert.equal(new Set(revisionIds(result.documentXml)).size, revisionIds(result.documentXml).length);
        assert.deepEqual(validateRedlineOoxml(result.documentXml).issues.filter(issue => issue.severity === 'error'), []);
        assert(!paragraph.getElementsByTagNameNS(W, 'pPr')[0].getElementsByTagNameNS(W, 'ins').length,
            `${testCase.label}: run insertion must not insert the paragraph mark`);
    }

    const controlsInput = documentXml(
        deletedParagraphWithContent(
            '<w:r><w:delText>A</w:delText><w:tab/><w:delText>B</w:delText><w:br/><w:delText>C</w:delText></w:r>'
        )
        + '<w:p><w:r><w:t>Following retained text.</w:t></w:r></w:p>'
    );
    const controls = await applyOperationToDocumentXml(controlsInput, {
        type: 'insert',
        target: { paragraphId: 'DEAD0001', revisionView: 'rejected' },
        anchor: { exactText: 'A\tB\nC', occurrence: 1, offset: 2 },
        modified: '[B]',
        author: B,
        existingRevisions: 'slice-cross-author'
    }, B, null, { strictTargets: true });
    assert.equal(controls.status, 'ok', JSON.stringify(controls.error));
    const controlsParagraph = paragraphs(controls.documentXml)[0];
    assert.equal(extractCanonicalParagraphText(controlsParagraph, { revisionView: 'rejected' }), 'A\tB\nC');
    assert.equal(controlsParagraph.getElementsByTagNameNS(W, 'tab').length, 1);
    assert.equal(controlsParagraph.getElementsByTagNameNS(W, 'br').length, 1);
}

// Repeated anchors require an occurrence, and structural split boundaries fail closed.
{
    const repeatedInput = documentXml(
        deletedParagraph('term and term and term')
        + '<w:p><w:r><w:t>Following retained text.</w:t></w:r></w:p>'
    );
    const baseOperation = {
        type: 'insert',
        target: { paragraphId: 'DEAD0001', revisionView: 'rejected' },
        modified: '[B]',
        author: B,
        existingRevisions: 'slice-cross-author'
    };
    const ambiguous = await applyOperationToDocumentXml(repeatedInput, {
        ...baseOperation,
        anchor: { exactText: 'term', offset: 2 }
    }, B, null, { strictTargets: true });
    assert.equal(ambiguous.status, 'error');
    assert.equal(ambiguous.error.code, 'AMBIGUOUS_ANCHOR');
    assert.equal(ambiguous.documentXml, repeatedInput);

    const selected = await applyOperationToDocumentXml(repeatedInput, {
        ...baseOperation,
        anchor: { exactText: 'term', occurrence: 2, offset: 2 }
    }, B, null, { strictTargets: true });
    assert.equal(selected.status, 'ok', JSON.stringify(selected.error));
    assert.equal(
        extractCanonicalParagraphText(paragraphs(rejectTrackedChangesInOoxml(selected.documentXml, { author: A }).oxml)[0]),
        'term and te[B]rm and term'
    );

    const unsafeFixtures = [
        deletedParagraphWithContent(
            '<w:r><w:delText>Commented text</w:delText></w:r>',
            'DEAD0001',
            41,
            40,
            '<w:commentRangeStart w:id="7"/>'
        ),
        deletedParagraphWithContent(
            '<w:hyperlink r:id="rId7" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            + '<w:r><w:delText>Linked text</w:delText></w:r></w:hyperlink>',
            'DEAD0001',
            51,
            50
        ),
        deletedParagraphWithContent(
            '<w:r><w:fldChar w:fldCharType="begin"/><w:delText>Field text</w:delText></w:r>',
            'DEAD0001',
            61,
            60
        )
    ];
    for (const fixture of unsafeFixtures) {
        const input = documentXml(fixture + '<w:p><w:r><w:t>Following retained text.</w:t></w:r></w:p>');
        const targetText = extractCanonicalParagraphText(paragraphs(input)[0], { revisionView: 'rejected' });
        const refused = await applyOperationToDocumentXml(input, {
            ...baseOperation,
            anchor: { exactText: targetText, occurrence: 1, offset: 1 }
        }, B, null, { strictTargets: true });
        assert.equal(refused.status, 'error');
        assert.equal(refused.error.code, 'UNSAFE_REVISION_BOUNDARY');
        assert.equal(refused.documentXml, input);
    }
}

// Combined Word-oracle document: inline insertion remains stable beside a
// post-source two-paragraph restoration block.
{
    const input = documentXml(
        deletedParagraph('The account holder must pay.', 'DEAD0001', 11, 10)
        + deletedParagraph('Archived heading', 'DEAD0002', 21, 20)
        + deletedParagraph('Archived body text.', 'DEAD0003', 31, 30)
        + '<w:p w14:paraId="NEXT0001"><w:r><w:t>Following retained text.</w:t></w:r></w:p>'
    );
    const result = await applyOperationsToDocumentXml(input, [
        {
            type: 'insert',
            target: { paragraphId: 'DEAD0001', revisionView: 'rejected' },
            anchor: { exactText: 'must pay', occurrence: 1, offset: 5 },
            modified: '[B] ',
            author: B,
            existingRevisions: 'slice-cross-author'
        },
        {
            type: 'restore',
            target: { paragraphId: 'DEAD0002', revisionView: 'rejected' },
            targetEnd: { paragraphId: 'DEAD0003', revisionView: 'rejected' },
            modified: ['Restored heading', 'Restored body text.'],
            author: B
        }
    ], B, null, { strictTargets: true, atomic: true });
    assert.equal(result.status, 'ok', JSON.stringify({ error: result.error, results: result.results }));
    assert.deepEqual(textVector(result.documentXml), [
        '[B] ', '', '', 'Restored heading', 'Restored body text.', 'Following retained text.'
    ]);
    assert.deepEqual(textVector(rejectTrackedChangesInOoxml(result.documentXml, { author: B }).oxml), [
        '', '', '', 'Following retained text.'
    ]);
    assert.deepEqual(textVector(acceptTrackedChangesInOoxml(result.documentXml, { author: B }).oxml), [
        '[B] ', '', '', 'Restored heading', 'Restored body text.', 'Following retained text.'
    ]);
    assert.deepEqual(textVector(rejectTrackedChangesInOoxml(result.documentXml, { author: A }).oxml), [
        'The account holder must [B] pay.',
        'Archived heading',
        'Archived body text.',
        'Restored heading',
        'Restored body text.',
        'Following retained text.'
    ]);
    assert.deepEqual(textVector(rejectTrackedChangesInOoxml(result.documentXml, { allAuthors: true }).oxml), [
        'The account holder must pay.', 'Archived heading', 'Archived body text.', 'Following retained text.'
    ]);
    const acceptA = textVector(acceptTrackedChangesInOoxml(result.documentXml, { author: A }).oxml);
    const acceptAll = textVector(acceptTrackedChangesInOoxml(result.documentXml, { allAuthors: true }).oxml);
    assert.deepEqual(acceptA, ['[B] Restored heading', 'Restored body text.', 'Following retained text.']);
    assert.deepEqual(acceptAll, ['[B] Restored heading', 'Restored body text.', 'Following retained text.']);
    assert.equal(new Set(revisionIds(result.documentXml)).size, revisionIds(result.documentXml).length);
    assert.deepEqual(validateRedlineOoxml(result.documentXml).issues.filter(issue => issue.severity === 'error'), []);
}

// Paragraph expansion may inherit effective formatting, but never revision-history nodes or IDs.
{
    const input = documentXml(
        `<w:p w14:paraId="BASE0001"><w:pPr><w:spacing w:after="120"/>`
        + `<w:pPrChange w:id="201" w:author="${A}" w:date="${DATE}"><w:pPr/></w:pPrChange></w:pPr>`
        + `<w:r><w:rPr><w:i/><w:rPrChange w:id="202" w:author="${A}" w:date="${DATE}"><w:rPr/></w:rPrChange></w:rPr>`
        + '<w:t>Anchor paragraph.</w:t></w:r></w:p>'
        + '<w:p><w:r><w:t>Following retained text.</w:t></w:r></w:p>'
    );
    const result = await applyOperationToDocumentXml(input, {
        type: 'redline',
        target: { paragraphId: 'BASE0001' },
        modified: 'Anchor paragraph.\nExpanded paragraph.',
        author: B,
        existingRevisions: 'slice-cross-author'
    }, B, null, { strictTargets: true });
    assert.notEqual(result.status, 'error', JSON.stringify(result.error));
    assert.equal(result.hasChanges, true);
    const ids = revisionIds(result.documentXml);
    assert.equal(new Set(ids).size, ids.length);
    assert.deepEqual(validateRedlineOoxml(result.documentXml).issues.filter(issue => issue.severity === 'error'), []);
    const expanded = paragraphs(result.documentXml).find(paragraph => (
        extractCanonicalParagraphText(paragraph).includes('Expanded paragraph.')
    ));
    assert(expanded);
    assert.equal(expanded.getElementsByTagNameNS(W, 'pPrChange').length, 0);
    assert.equal(expanded.getElementsByTagNameNS(W, 'rPrChange').length, 0);
}

console.log('PASS: WP09c-e validation, identity, and rejected-view insertion tests');
