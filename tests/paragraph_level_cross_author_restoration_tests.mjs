import assert from 'assert/strict';

import './setup-xml-provider.mjs';
import {
    acceptTrackedChangesInOoxml,
    rejectTrackedChangesInOoxml,
    validateRedlineOoxml
} from '../index.js';
import { parseOoxmlSafe } from '../adapters/xml-adapter.js';
import { extractCanonicalParagraphText } from '../core/paragraph-text.js';
import {
    createParagraphFingerprint,
    getDocumentParagraphNodes
} from '../core/paragraph-targeting.js';
import {
    applyOperationToDocumentXml,
    applyOperationsToDocumentXml,
    preflightOperations
} from '../services/standalone-operation-runner.js';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const DATE = '2026-09-08T18:00:00Z';
const A = 'Reviewer A';
const B = 'Reviewer B';
const C = 'Reviewer C';

function escapeXml(text) {
    return String(text).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function run(text, rPr = '', deleted = false) {
    const tag = deleted ? 'w:delText' : 'w:t';
    const preserve = /^\s|\s$/.test(text) ? ' xml:space="preserve"' : '';
    return `<w:r>${rPr}<${tag}${preserve}>${escapeXml(text)}</${tag}></w:r>`;
}

function normalParagraph(text, id) {
    return `<w:p w14:paraId="${id}">${run(text)}</w:p>`;
}

function deletedParagraph({
    text,
    id,
    revisionBase,
    author = A,
    content = null,
    pPr = '',
    anchors = ''
}) {
    return `<w:p w14:paraId="${id}" w14:textId="11111111" w:rsidR="00112233">`
        + `<w:pPr>${pPr}<w:rPr><w:del w:id="${revisionBase}" w:author="${author}" w:date="${DATE}"/></w:rPr></w:pPr>`
        + `${anchors}<w:del w:id="${revisionBase + 1}" w:author="${author}" w:date="${DATE}">${content || run(text, '', true)}</w:del>`
        + '</w:p>';
}

function doc(body) {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="${W}" xmlns:w14="${W14}"><w:body>${body}</w:body></w:document>`;
}

function parse(xml) {
    const parsed = parseOoxmlSafe(xml, 'application/xml');
    assert.equal(parsed.error, null);
    return parsed.doc;
}

function paragraphTexts(xml) {
    return getDocumentParagraphNodes(parse(xml)).map(paragraph => extractCanonicalParagraphText(paragraph));
}

function paragraphs(xml) {
    return getDocumentParagraphNodes(parse(xml));
}

function author(node) {
    return node.getAttribute('w:author') || node.getAttribute('author');
}

function paraId(node) {
    return node.getAttribute('w14:paraId') || node.getAttributeNS(W14, 'paraId');
}

function count(xml, localName) {
    return parse(xml).getElementsByTagNameNS(W, localName).length;
}

function baseDocument(target = deletedParagraph({ text: 'Restored clause.', id: 'DEAD0001', revisionBase: 1 })) {
    return doc(
        normalParagraph('Before.', 'BEFORE01')
        + target
        + normalParagraph('After.', 'AFTER001')
    );
}

async function restore(input, modified = 'Restored clause.', overrides = {}) {
    return applyOperationToDocumentXml(input, {
        type: 'restore',
        target: { paragraphId: 'DEAD0001' },
        modified,
        author: B,
        ...overrides
    }, B, null, { strictTargets: true });
}

// Plain restoration, metadata, receipts, paragraph order, and six lifecycle outcomes.
{
    const input = baseDocument();
    const originalSourceXml = new XMLSerializer().serializeToString(paragraphs(input)[1]);
    const preflight = preflightOperations(input, [{
        type: 'restore', target: { paragraphId: 'DEAD0001' }, modified: 'Restored clause.', author: B
    }], B, { strictTargets: true });
    assert.equal(preflight.valid, true);
    assert.equal(preflight.results[0].operationType, 'restore');

    const result = await restore(input);
    assert.equal(result.status, 'ok', JSON.stringify(result.error));
    assert.equal(result.hasChanges, true);
    assert.deepEqual(paragraphTexts(result.documentXml), ['Before.', 'Restored clause.', '', 'After.']);
    assert.equal(validateRedlineOoxml(result.documentXml).valid, true);

    const nodes = paragraphs(result.documentXml);
    const restored = nodes[1];
    const source = nodes[2];
    assert.notEqual(paraId(restored), paraId(source));
    assert.equal(paraId(source), 'DEAD0001');
    assert(!restored.hasAttribute('w14:textId'));
    assert(!restored.hasAttribute('w:rsidR'));
    assert.equal(author(restored.getElementsByTagNameNS(W, 'ins')[0]), B);
    assert.equal(author(restored.getElementsByTagNameNS(W, 'ins')[1]), B);
    assert.equal(author(source.getElementsByTagNameNS(W, 'del')[0]), A);
    assert.equal(new XMLSerializer().serializeToString(source), originalSourceXml);
    assert.equal(result.receipt.revisionItems.length, 2);
    assert(result.receipt.revisionItems.every(item => item.kind === 'ins'));

    const resultDoc = parse(result.documentXml);
    const allIds = ['ins', 'del'].flatMap(localName => (
        Array.from(resultDoc.getElementsByTagNameNS(W, localName)).map(node => node.getAttribute('w:id'))
    ));
    assert.equal(new Set(allIds).size, allIds.length);

    const acceptAll = acceptTrackedChangesInOoxml(result.documentXml, { allAuthors: true });
    assert.deepEqual(paragraphTexts(acceptAll.oxml), ['Before.', 'Restored clause.', 'After.']);

    const rejectB = rejectTrackedChangesInOoxml(result.documentXml, { author: B });
    assert.deepEqual(paragraphTexts(rejectB.oxml), ['Before.', '', 'After.']);
    assert.equal(count(rejectB.oxml, 'del'), 2);

    const rejectA = rejectTrackedChangesInOoxml(result.documentXml, { author: A });
    assert.deepEqual(paragraphTexts(rejectA.oxml), ['Before.', 'Restored clause.', 'Restored clause.', 'After.']);

    const acceptA = acceptTrackedChangesInOoxml(result.documentXml, { author: A });
    assert.deepEqual(paragraphTexts(acceptA.oxml), ['Before.', 'Restored clause.', 'After.']);
    assert.equal(count(acceptA.oxml, 'ins'), 2);

    const acceptB = acceptTrackedChangesInOoxml(result.documentXml, { author: B });
    assert.deepEqual(paragraphTexts(acceptB.oxml), ['Before.', 'Restored clause.', '', 'After.']);
    assert.equal(count(acceptB.oxml, 'del'), 2);

    const rejectAll = rejectTrackedChangesInOoxml(result.documentXml, { allAuthors: true });
    assert.deepEqual(paragraphTexts(rejectAll.oxml), ['Before.', 'Restored clause.', 'After.']);

    const repeated = await restore(result.documentXml);
    assert.equal(repeated.status, 'no-op');
    assert.equal(repeated.hasChanges, false);
    assert.equal(repeated.documentXml, result.documentXml);
}

// Adjusted restoration preserves source run formatting where text survives.
{
    const target = deletedParagraph({
        text: 'Bold plain.',
        id: 'DEAD0001',
        revisionBase: 10,
        content: run('Bold', '<w:rPr><w:b/></w:rPr>', true) + run(' plain.', '', true)
    });
    const result = await restore(baseDocument(target), 'Bold revised.');
    assert.equal(result.status, 'ok', JSON.stringify(result.error));
    assert.deepEqual(paragraphTexts(result.documentXml), ['Before.', 'Bold revised.', '', 'After.']);
    const restored = paragraphs(result.documentXml)[1];
    assert(restored.getElementsByTagNameNS(W, 'b').length >= 1);
}

// Mixed bold/italic/underline source runs survive a verbatim restoration.
{
    const target = deletedParagraph({
        text: 'Bold italic underlined.',
        id: 'DEAD0001',
        revisionBase: 14,
        content: run('Bold ', '<w:rPr><w:b/></w:rPr>', true)
            + run('italic ', '<w:rPr><w:i/></w:rPr>', true)
            + run('underlined.', '<w:rPr><w:u w:val="single"/></w:rPr>', true)
    });
    const result = await restore(baseDocument(target), 'Bold italic underlined.');
    assert.equal(result.status, 'ok');
    const restored = paragraphs(result.documentXml)[1];
    assert.equal(restored.getElementsByTagNameNS(W, 'b').length, 1);
    assert.equal(restored.getElementsByTagNameNS(W, 'i').length, 1);
    assert.equal(restored.getElementsByTagNameNS(W, 'u').length, 1);
}

// Paragraph properties use the allowlist; revision/change/identity metadata is not cloned.
{
    const pPr = '<w:pStyle w:val="Clause"/><w:numPr><w:ilvl w:val="1"/><w:numId w:val="9"/></w:numPr>'
        + '<w:spacing w:after="120"/><w:pPrChange w:id="80" w:author="Reviewer A"><w:pPr/></w:pPrChange>';
    const result = await restore(baseDocument(deletedParagraph({
        text: 'Listed clause.', id: 'DEAD0001', revisionBase: 20, pPr
    })), 'Listed clause.');
    assert.equal(result.status, 'ok');
    const restored = paragraphs(result.documentXml)[1];
    assert.equal(restored.getElementsByTagNameNS(W, 'pStyle')[0].getAttribute('w:val'), 'Clause');
    assert.equal(restored.getElementsByTagNameNS(W, 'numId')[0].getAttribute('w:val'), '9');
    assert.equal(restored.getElementsByTagNameNS(W, 'pPrChange').length, 0);
    assert.equal(restored.getElementsByTagNameNS(W, 'del').length, 0);
}

// Bookmarks/comments remain only on A's paragraph and are reported as dropped from B's clone.
{
    const anchors = '<w:bookmarkStart w:id="4" w:name="ClauseAnchor"/><w:commentRangeStart w:id="7"/>';
    const target = deletedParagraph({ text: 'Anchored.', id: 'DEAD0001', revisionBase: 30, anchors })
        .replace('</w:p>', '<w:bookmarkEnd w:id="4"/><w:commentRangeEnd w:id="7"/><w:r><w:commentReference w:id="7"/></w:r></w:p>');
    const input = baseDocument(target);
    const result = await restore(input, 'Anchored.');
    assert.equal(result.status, 'ok');
    assert.equal(count(result.documentXml, 'bookmarkStart'), count(input, 'bookmarkStart'));
    assert.equal(count(result.documentXml, 'commentRangeStart'), count(input, 'commentRangeStart'));
    const restored = paragraphs(result.documentXml)[1];
    assert.equal(restored.getElementsByTagNameNS(W, 'bookmarkStart').length, 0);
    assert.equal(restored.getElementsByTagNameNS(W, 'commentReference').length, 0);
    assert(result.receipt.warnings.includes('RESTORATION_DROPPED_BOOKMARK:ClauseAnchor'));
    assert(result.receipt.warnings.includes('RESTORATION_DROPPED_COMMENT:7'));
}

// Multi-paragraph range produces one contiguous inserted block before all A paragraphs.
{
    const input = doc(
        normalParagraph('Before.', 'BEFORE01')
        + deletedParagraph({ text: 'First.', id: 'DEAD0001', revisionBase: 40 })
        + deletedParagraph({ text: 'Second.', id: 'DEAD0002', revisionBase: 42 })
        + normalParagraph('After.', 'AFTER001')
    );
    const result = await applyOperationToDocumentXml(input, {
        type: 'restore',
        target: { index: 2 },
        targetEndRef: 3,
        modified: ['First restored.', 'Second restored.'],
        author: B
    }, B, null, { strictTargets: true });
    assert.equal(result.status, 'ok');
    assert.deepEqual(paragraphTexts(result.documentXml), [
        'Before.', 'First restored.', 'Second restored.', '', '', 'After.'
    ]);
    const nodes = paragraphs(result.documentXml);
    assert.equal(paraId(nodes[3]), 'DEAD0001');
    assert.equal(paraId(nodes[4]), 'DEAD0002');
    assert.equal(result.receipt.revisionItems.length, 4);

    const descriptorResult = await applyOperationToDocumentXml(input, {
        type: 'restore',
        target: { paragraphId: 'DEAD0001' },
        targetEnd: { paragraphId: 'DEAD0002' },
        modified: ['First restored.', 'Second restored.'],
        author: B
    }, B, null, { strictTargets: true });
    assert.equal(descriptorResult.status, 'ok');
    assert.deepEqual(paragraphTexts(descriptorResult.documentXml), [
        'Before.', 'First restored.', 'Second restored.', '', '', 'After.'
    ]);
}

// Independent restorations remain independently adjacent and preserve source order.
{
    const input = doc(
        deletedParagraph({ text: 'First.', id: 'DEAD0001', revisionBase: 44 })
        + deletedParagraph({ text: 'Second.', id: 'DEAD0002', revisionBase: 46 })
        + normalParagraph('After.', 'AFTER001')
    );
    const result = await applyOperationsToDocumentXml(input, [
        { type: 'restore', target: { paragraphId: 'DEAD0001' }, modified: 'First.', author: B },
        { type: 'restore', target: { paragraphId: 'DEAD0002' }, modified: 'Second.', author: B }
    ], B, null, { strictTargets: true, atomic: true });
    assert.equal(result.status, 'ok');
    assert.deepEqual(paragraphTexts(result.documentXml), ['First.', '', 'Second.', '', 'After.']);
    assert.equal(result.receipts.length, 2);
    assert(result.receipts.every(receipt => receipt.revisionItems.length === 2));
}

// Reapplying with changed text replaces the existing same-author restoration instead of duplicating it.
{
    const first = await restore(baseDocument());
    const changed = await restore(first.documentXml, 'Adjusted restoration.');
    assert.equal(changed.status, 'ok', JSON.stringify(changed.error));
    assert.deepEqual(paragraphTexts(changed.documentXml), ['Before.', 'Adjusted restoration.', '', 'After.']);
    assert.equal(paragraphs(changed.documentXml).filter(paragraph => (
        extractCanonicalParagraphText(paragraph) === 'Adjusted restoration.'
    )).length, 1);
}

// Strict fingerprint targeting works from the accepted view; rejected-view mutation stays read-only.
{
    const input = baseDocument();
    const target = paragraphs(input)[1];
    const fingerprint = createParagraphFingerprint(target, { text: '', index: 2, revisionView: 'accepted' });
    const fingerprintResult = await applyOperationToDocumentXml(input, {
        type: 'restore', target: { fingerprint }, modified: 'Restored clause.', author: B
    }, B, null, { strictTargets: true });
    assert.equal(fingerprintResult.status, 'ok', JSON.stringify(fingerprintResult.error));

    const rejectedView = await applyOperationToDocumentXml(input, {
        type: 'restore',
        target: { exactText: 'Restored clause.', revisionView: 'rejected' },
        modified: 'Restored clause.',
        author: B
    }, B, null, { strictTargets: true });
    assert.equal(rejectedView.status, 'error');
    assert.equal(rejectedView.error.code, 'UNSUPPORTED_REVISION_VIEW_MUTATION');
    assert.equal(rejectedView.documentXml, input);

    const rejectedInspection = preflightOperations(input, [{
        type: 'restore',
        target: { exactText: 'Restored clause.', revisionView: 'rejected' },
        modified: 'Restored clause.',
        author: B
    }], B, { strictTargets: true });
    assert.equal(rejectedInspection.valid, true);
    assert.equal(rejectedInspection.results[0].resolvedTarget.revisionView, 'rejected');
}

// Non-triggering taxonomy rows retain their established ordinary editing paths.
{
    const sameAuthorInput = baseDocument(deletedParagraph({
        text: 'Original.', id: 'DEAD0001', revisionBase: 54, author: B
    }));
    const sameAuthor = await applyOperationToDocumentXml(sameAuthorInput, {
        type: 'redline',
        target: { paragraphId: 'DEAD0001' },
        modified: 'Same-author revision.',
        author: B,
        existingRevisions: 'merge-same-author'
    }, B, null, { strictTargets: true });
    assert.equal(sameAuthor.status, 'ok');
    assert.deepEqual(paragraphTexts(sameAuthor.documentXml), ['Before.', 'Same-author revision.', 'After.']);

    const noMarkInput = doc(
        normalParagraph('Before.', 'BEFORE01')
        + `<w:p w14:paraId="DEAD0001"><w:del w:id="58" w:author="${A}" w:date="${DATE}">${run('Original.', '', true)}</w:del></w:p>`
        + normalParagraph('After.', 'AFTER001')
    );
    const noMark = await applyOperationToDocumentXml(noMarkInput, {
        type: 'redline',
        target: { paragraphId: 'DEAD0001' },
        modified: 'Ordinary insertion.',
        author: B,
        existingRevisions: 'slice-cross-author'
    }, B, null, { strictTargets: true });
    assert.equal(noMark.status, 'ok');
    assert.deepEqual(paragraphTexts(noMark.documentXml), ['Before.', 'Ordinary insertion.', 'After.']);

    const partialInput = baseDocument(deletedParagraph({
        text: 'Removed.', id: 'DEAD0001', revisionBase: 59
    }).replace('</w:p>', `${run('Surviving.')}</w:p>`));
    const partial = await applyOperationToDocumentXml(partialInput, {
        type: 'redline',
        target: { paragraphId: 'DEAD0001' },
        modified: 'Surviving, revised.',
        author: B,
        existingRevisions: 'slice-cross-author'
    }, B, null, { strictTargets: true });
    assert.equal(partial.status, 'ok');
    assert.deepEqual(paragraphTexts(partial.documentXml), ['Before.', 'Surviving, revised.', 'After.']);
}

// Structural and taxonomy refusals remain byte-exact under atomic execution.
{
    const cases = [
        {
            name: 'same author',
            input: baseDocument(deletedParagraph({ text: 'Restored clause.', id: 'DEAD0001', revisionBase: 60, author: B })),
            code: 'RESTORATION_STATE_REQUIRED'
        },
        {
            name: 'surviving content',
            input: baseDocument(deletedParagraph({ text: 'Removed.', id: 'DEAD0001', revisionBase: 62 })
                .replace('</w:p>', `${run('Survives.')}</w:p>`)),
            code: 'RESTORATION_STATE_REQUIRED'
        },
        {
            name: 'content deletion without paragraph mark deletion',
            input: doc(
                normalParagraph('Before.', 'BEFORE01')
                + `<w:p w14:paraId="DEAD0001"><w:del w:id="63" w:author="${A}" w:date="${DATE}">${run('Restored clause.', '', true)}</w:del></w:p>`
                + normalParagraph('After.', 'AFTER001')
            ),
            code: 'RESTORATION_STATE_REQUIRED'
        },
        {
            name: 'section break',
            input: baseDocument(deletedParagraph({
                text: 'Restored clause.', id: 'DEAD0001', revisionBase: 64, pPr: '<w:sectPr><w:pgSz w:w="12240"/></w:sectPr>'
            })),
            code: 'SECTION_BREAK_PARAGRAPH'
        },
        {
            name: 'move source',
            input: baseDocument(deletedParagraph({ text: 'Restored clause.', id: 'DEAD0001', revisionBase: 66 })
                .replace(
                    '<w:del w:id="67" w:author="Reviewer A" w:date="2026-09-08T18:00:00Z">',
                    '<w:moveFrom w:id="67" w:author="Reviewer A" w:date="2026-09-08T18:00:00Z">'
                ).replace('</w:del></w:p>', '</w:moveFrom></w:p>')),
            code: 'UNSUPPORTED_MOVE_REVISION'
        },
        {
            name: 'source inside move-from range',
            input: doc(
                `<w:p w14:paraId="BEFORE01">${run('Before.')}<w:moveFromRangeStart w:id="91" w:author="${A}"/></w:p>`
                + deletedParagraph({ text: 'Restored clause.', id: 'DEAD0001', revisionBase: 72 })
                + `<w:p w14:paraId="AFTER001"><w:moveFromRangeEnd w:id="91"/>${run('After.')}</w:p>`
            ),
            code: 'UNSUPPORTED_MOVE_REVISION'
        },
        {
            name: 'final paragraph',
            input: doc(normalParagraph('Before.', 'BEFORE01') + deletedParagraph({
                text: 'Restored clause.', id: 'DEAD0001', revisionBase: 68
            })),
            code: 'UNSAFE_PARAGRAPH_PLACEMENT'
        },
        {
            name: 'paragraph immediately before body section properties',
            input: doc(
                normalParagraph('Before.', 'BEFORE01')
                + deletedParagraph({ text: 'Restored clause.', id: 'DEAD0001', revisionBase: 70 })
                + '<w:sectPr><w:pgSz w:w="12240"/></w:sectPr>'
            ),
            code: 'UNSAFE_PARAGRAPH_PLACEMENT'
        }
    ];
    for (const fixture of cases) {
        const result = await applyOperationsToDocumentXml(fixture.input, [{
            type: 'restore', target: { paragraphId: 'DEAD0001' }, modified: 'Restored clause.', author: B
        }], B, null, { strictTargets: true, atomic: true });
        assert.equal(result.status, 'error', fixture.name);
        assert.equal(result.results[0].error.code, fixture.code, fixture.name);
        assert.equal(result.documentXml, fixture.input, fixture.name);
        assert.equal(result.rolledBack, true, fixture.name);
    }
}

// Restore is explicitly tracked and range cardinality is validated before mutation.
{
    const input = baseDocument();
    const direct = await applyOperationToDocumentXml(input, {
        type: 'restore',
        target: { paragraphId: 'DEAD0001' },
        modified: 'Restored clause.',
        author: B,
        generateRedlines: false
    }, B, null, { strictTargets: true });
    assert.equal(direct.status, 'error');
    assert.equal(direct.error.code, 'INVALID_OPERATION');
    assert.equal(direct.documentXml, input);

    const mismatch = await applyOperationsToDocumentXml(input, [{
        type: 'restore',
        target: { paragraphId: 'DEAD0001' },
        modified: ['One.', 'Two.'],
        author: B
    }], B, null, { strictTargets: true, atomic: true });
    assert.equal(mismatch.status, 'error');
    assert.equal(mismatch.results[0].error.code, 'RESTORATION_COUNT_MISMATCH');
    assert.equal(mismatch.documentXml, input);
}

// Progressive batches commit valid restoration receipts while preserving an independently failed operation.
{
    const input = baseDocument();
    const result = await applyOperationsToDocumentXml(input, [
        { type: 'restore', target: { paragraphId: 'DEAD0001' }, modified: 'Restored clause.', author: B },
        { type: 'restore', target: { paragraphId: 'MISSING1' }, modified: 'Missing.', author: B }
    ], B, null, { strictTargets: true, atomic: false, continueOnError: true });
    assert.equal(result.status, 'partial');
    assert.equal(result.results[0].status, 'applied');
    assert.equal(result.results[0].receipt.finalDisposition, 'applied');
    assert.equal(result.results[0].receipt.revisionItems.length, 2);
    assert.equal(result.results[1].status, 'error');
    assert.deepEqual(paragraphTexts(result.documentXml), ['Before.', 'Restored clause.', '', 'After.']);
}

// Table-cell restoration is supported when a same-cell successor exists; deleted rows fail closed.
{
    const target = deletedParagraph({ text: 'Cell clause.', id: 'DEAD0001', revisionBase: 80 });
    const table = `<w:tbl><w:tr><w:tc>${target}${normalParagraph('Cell successor.', 'CELLNEXT')}</w:tc></w:tr></w:tbl>`;
    const supported = await restore(doc(table), 'Cell clause.');
    assert.equal(supported.status, 'ok');
    assert.deepEqual(paragraphTexts(supported.documentXml), ['Cell clause.', '', 'Cell successor.']);

    const deletedRow = `<w:tbl><w:tr><w:trPr><w:del w:id="90" w:author="${A}" w:date="${DATE}"/></w:trPr><w:tc>${target}${normalParagraph('Cell successor.', 'CELLNEXT')}</w:tc></w:tr></w:tbl>`;
    const refused = await restore(doc(deletedRow), 'Cell clause.');
    assert.equal(refused.status, 'error');
    assert.equal(refused.error.code, 'UNSAFE_DELETED_TABLE_ROW');
}

// A third author can continue editing B's restored paragraph through ordinary slicing.
{
    const restored = await restore(baseDocument());
    const restoredId = paraId(paragraphs(restored.documentXml)[1]);
    const followUp = await applyOperationToDocumentXml(restored.documentXml, {
        type: 'redline',
        target: { paragraphId: restoredId },
        modified: 'Restored clause, revised again.',
        author: C,
        existingRevisions: 'slice-cross-author'
    }, C, null, { strictTargets: true });
    assert.equal(followUp.status, 'ok');
    assert.equal(paragraphTexts(followUp.documentXml)[1], 'Restored clause, revised again.');
    assert.equal(validateRedlineOoxml(followUp.documentXml).valid, true);
}

console.log('PASS: paragraph-level cross-author restoration tests');
