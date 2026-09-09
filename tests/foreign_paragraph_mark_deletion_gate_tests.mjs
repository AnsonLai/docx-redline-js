import assert from 'assert/strict';

import './setup-xml-provider.mjs';
import {
    acceptTrackedChangesInOoxml,
    applyRedlineToOxml,
    ingestWordOoxmlToPlainText,
    rejectTrackedChangesInOoxml,
    validateRedlineOoxml
} from '../index.js';
import { parseOoxmlSafe } from '../adapters/xml-adapter.js';
import { inspectForeignDeletedParagraphTarget } from '../core/paragraph-revision-safety.js';
import { applyOperationsToDocumentXml } from '../services/standalone-operation-runner.js';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const DATE = '2026-09-08T18:00:00Z';
const OWNER = 'Reviewer A';
const EDITOR = 'Reviewer B';
const RESTORED = 'Restored clause.';

function deletion(id, text = RESTORED, author = OWNER) {
    return `<w:del w:id="${id}" w:author="${author}" w:date="${DATE}"><w:r><w:delText>${text}</w:delText></w:r></w:del>`;
}

function paragraphMarkDeletion(id, author = OWNER) {
    return `<w:pPr><w:rPr><w:del w:id="${id}" w:author="${author}" w:date="${DATE}"/></w:rPr></w:pPr>`;
}

function paragraph(content, id = 'DEAD0001') {
    return `<w:p xmlns:w="${W}" xmlns:w14="${W14}" w14:paraId="${id}">${content}</w:p>`;
}

function documentXml(content) {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="${W}" xmlns:w14="${W14}"><w:body>${content}</w:body></w:document>`;
}

function parseParagraph(xml) {
    const parsed = parseOoxmlSafe(xml, 'application/xml');
    assert.equal(parsed.error, null);
    return parsed.doc.getElementsByTagNameNS(W, 'p')[0];
}

const deletedParagraph = paragraph(`${paragraphMarkDeletion(1)}${deletion(2)}`);

// Low-level paragraph mutation refuses the unsafe same-paragraph resurrection.
{
    const result = await applyRedlineToOxml(deletedParagraph, '', RESTORED, {
        author: EDITOR,
        existingRevisions: 'slice-cross-author'
    });
    assert.equal(result.status, 'error');
    assert.equal(result.error.code, 'FOREIGN_PARAGRAPH_MARK_DELETION');
    assert.equal(result.error.ownerAuthor, OWNER);
    assert.equal(result.hasChanges, false);
    assert.equal(result.oxml, deletedParagraph);

    const directEdit = await applyRedlineToOxml(deletedParagraph, '', RESTORED, {
        author: EDITOR,
        generateRedlines: false
    });
    assert.equal(directEdit.status, 'error');
    assert.equal(directEdit.error.code, 'FOREIGN_PARAGRAPH_MARK_DELETION');
    assert.equal(directEdit.oxml, deletedParagraph);
}

// The document runner surfaces the operation error and rolls atomic work back byte-for-byte.
{
    const original = documentXml(`${deletedParagraph}<w:p w14:paraId="NEXT0001"><w:r><w:t>Successor.</w:t></w:r></w:p>`);
    const result = await applyOperationsToDocumentXml(original, [{
        type: 'redline',
        target: { paragraphId: 'DEAD0001' },
        modified: RESTORED
    }], EDITOR, null, {
        atomic: true,
        strictTargets: true,
        existingRevisions: 'slice-cross-author'
    });
    assert.equal(result.status, 'error');
    assert.equal(result.error.code, 'BATCH_OPERATION_FAILED');
    assert.equal(result.results[0].error.code, 'FOREIGN_PARAGRAPH_MARK_DELETION');
    assert.equal(result.results[0].error.ownerAuthor, OWNER);
    assert.equal(result.rolledBack, true);
    assert.equal(result.hasChanges, false);
    assert.equal(result.documentXml, original);
}

// Validation warns about an already-authored unsafe shape without invalidating OOXML.
const unsafeShape = paragraph(
    `${paragraphMarkDeletion(10)}${deletion(11)}`
    + `<w:ins w:id="12" w:author="${EDITOR}" w:date="${DATE}"><w:r><w:t>${RESTORED}</w:t></w:r></w:ins>`
);
{
    const validation = validateRedlineOoxml(unsafeShape);
    const issue = validation.issues.find(item => item.code === 'FOREIGN_PARAGRAPH_MARK_DELETION');
    assert.equal(validation.valid, true);
    assert.equal(issue?.severity, 'warning');
    assert.match(issue?.message || '', /Reviewer A/);
}

// Lifecycle evidence: accepting all changes to a terminal unsafe paragraph loses
// Reviewer B's apparent restoration; rejecting A exposes both copies.
{
    const accepted = acceptTrackedChangesInOoxml(unsafeShape, { allAuthors: true });
    assert.equal(ingestWordOoxmlToPlainText(accepted.oxml), '');

    const rejectedA = rejectTrackedChangesInOoxml(unsafeShape, { author: OWNER });
    assert.equal(ingestWordOoxmlToPlainText(rejectedA.oxml), `${RESTORED}${RESTORED}`);
}

// Trigger taxonomy exclusions: surviving content, no paragraph-mark deletion,
// and a mark deletion owned by the current author are not resurrection state.
{
    const pendingMerge = paragraph(
        `${paragraphMarkDeletion(20)}<w:r><w:t>Surviving.</w:t></w:r>${deletion(21, ' Removed.')}`,
        'LEGAL001'
    );
    assert.equal(inspectForeignDeletedParagraphTarget(parseParagraph(pendingMerge), EDITOR).matches, false);
    assert(!validateRedlineOoxml(
        pendingMerge.replace('</w:p>', `<w:ins w:id="22" w:author="${EDITOR}" w:date="${DATE}"><w:r><w:t> Pending.</w:t></w:r></w:ins></w:p>`)
    ).issues.some(issue => issue.code === 'FOREIGN_PARAGRAPH_MARK_DELETION'));

    const contentOnlyDeletion = paragraph(deletion(30), 'CONTENT1');
    assert.equal(inspectForeignDeletedParagraphTarget(parseParagraph(contentOnlyDeletion), EDITOR).matches, false);

    const sameAuthor = paragraph(`${paragraphMarkDeletion(40, EDITOR)}${deletion(41, RESTORED, EDITOR)}`, 'SAME0001');
    assert.equal(inspectForeignDeletedParagraphTarget(parseParagraph(sameAuthor), EDITOR).matches, false);

    const anchoredDeletion = paragraph(
        `${paragraphMarkDeletion(50)}<w:commentRangeStart w:id="7"/>${deletion(51)}`
        + '<w:commentRangeEnd w:id="7"/><w:r><w:commentReference w:id="7"/></w:r>',
        'ANCHOR01'
    );
    assert.equal(inspectForeignDeletedParagraphTarget(parseParagraph(anchoredDeletion), EDITOR).matches, true);
}

console.log('PASS: foreign paragraph-mark deletion gate tests');
