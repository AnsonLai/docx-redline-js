import './setup-xml-provider.mjs';

// A tracked (or untracked) text edit must keep every footnote and endnote reference it does not explicitly
// delete, in Accept All and in Reject All. Regression for GitHub issues #6, #7 and #8: the reconstruction route
// re-placed a reference only when all text before it, or all text after it, was unchanged, so the word before each
// reference was deleted together with the reference and reinserted without it. Also covers the
// ACCEPT_NON_TEXT_CONTENT safety net in the revision fidelity verifier.
import assert from 'assert/strict';
import { readFileSync } from 'node:fs';
import { acceptTrackedChangesInOoxml, inspectDocumentParts, openDocx, rejectTrackedChangesInOoxml, validateRedlineOoxml } from '../index.js';
import { applyOperationsToDocumentXml } from '../services/standalone-operation-runner.js';
import { verifyNonTextContent } from '../services/revision-fidelity-verifier.js';
import { parseXml } from '../adapters/xml-adapter.js';
import { unzipDocx, zipDocx } from '../document/zip-archive.js';

const NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const W = `xmlns:w="${NS}"`;
const run = text => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const fn = id => `<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="${id}"/></w:r>`;
const en = id => `<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:endnoteReference w:id="${id}"/></w:r>`;
const docXml = body => `<w:document ${W}><w:body>${body}<w:sectPr/></w:body></w:document>`;
const strong = q => ({ index: q.index, exactText: q.exactText, paragraphId: q.paragraphId, fingerprint: q.fingerprint });

/** Reference ids in document order, each with the visible text that precedes it in its paragraph. */
function references(xml) {
    const doc = parseXml(xml);
    const out = [];
    for (const paragraph of Array.from(doc.getElementsByTagNameNS(NS, 'p'))) {
        let text = '';
        const walk = node => {
            for (const child of Array.from(node.childNodes)) {
                if (child.nodeType !== 1) continue;
                if (child.localName === 't') text += child.textContent;
                else if (child.localName === 'footnoteReference' || child.localName === 'endnoteReference') {
                    out.push(`${child.localName === 'footnoteReference' ? 'fn' : 'en'}${child.getAttribute('w:id')}@${text}`);
                } else walk(child);
            }
        };
        walk(paragraph);
    }
    return out;
}

async function edit(body, modify, options = {}) {
    const documentXml = docXml(body);
    const target = inspectDocumentParts({ documentXml }).paragraphs[0];
    const result = await applyOperationsToDocumentXml(documentXml, [{ type: 'redline', target: strong(target), modified: modify(target.exactText) }],
        'Editor', {}, { atomic: true, strictTargets: true, ...options });
    assert.equal(result.status, 'ok', JSON.stringify(result.results?.find(item => item.error)?.error ?? result.error));
    assert.deepEqual(validateRedlineOoxml(result.documentXml).issues.filter(issue => issue.severity === 'error'), []);
    const accepted = acceptTrackedChangesInOoxml(result.documentXml, { allAuthors: true }).oxml;
    const rejected = rejectTrackedChangesInOoxml(result.documentXml, { allAuthors: true }).oxml;
    return { source: documentXml, xml: result.documentXml, accepted, rejected, sourceText: target.exactText };
}
const acceptedText = xml => inspectDocumentParts({ documentXml: xml }).paragraphs.map(q => q.exactText);

const A = 'The buyer pays the deposit within ten days.';
const B = 'The seller holds it in escrow.';

// ---------------------------------------------------------------------------------------------------------
// 1. The three reported shapes
// ---------------------------------------------------------------------------------------------------------
{
    // #6: the reference is the paragraph's last child; the edit replaces the first words.
    const view = await edit(`<w:p>${run(A)}${fn(1)}</w:p>`, text => text.replace('The buyer', 'The purchaser'));
    assert.deepEqual(references(view.accepted), ['fn1@The purchaser pays the deposit within ten days.'], '#6 Accept All keeps the reference after the sentence');
    assert.deepEqual(references(view.rejected), references(view.source), '#6 Reject All');
    assert.doesNotMatch(view.xml, /<w:del\b(?:(?!<\/w:del>).)*footnoteReference/s, '#6 the reference is not deleted');
    assert.equal((view.xml.match(/days\./g) || []).length, 1, '#6 "days." is not deleted and reinserted');
}
{
    // #7: a bookmark around the first sentence, then a reference, then a second sentence that is edited.
    const view = await edit(`<w:p><w:bookmarkStart w:id="10" w:name="clause"/>${run(A)}<w:bookmarkEnd w:id="10"/>${fn(1)}${run(' ')}${run(B)}</w:p>`,
        text => text.replace(B, 'The seller holds it in a separate escrow account.'));
    assert.deepEqual(references(view.accepted), [`fn1@${A}`], '#7 Accept All keeps the reference');
    assert.deepEqual(references(view.rejected), references(view.source), '#7 Reject All');
    const firstSentence = view.xml.slice(view.xml.indexOf('<w:bookmarkStart'), view.xml.indexOf('<w:bookmarkEnd'));
    assert.doesNotMatch(firstSentence, /<w:(ins|del)\b/, '#7 the untouched first sentence carries no revisions');
    assert.ok(view.xml.indexOf('<w:bookmarkEnd') < view.xml.indexOf('<w:footnoteReference'), '#7 the bookmark still ends before the reference');
    assert.deepEqual(acceptedText(view.accepted), [`${A} The seller holds it in a separate escrow account.`]);
}
{
    // #8: five sentences each followed by a reference; only the first sentence is edited.
    const S = ['One pays.', 'Two holds.', 'Three signs.', 'Four files.', 'Five reports.'];
    const view = await edit(`<w:p>${S.map((s, i) => `${run(s)}${fn(i + 1)}${run(' ')}`).join('')}${run('The donor keeps a record.')}</w:p>`,
        text => text.replace('One pays.', 'One pays promptly.'));
    assert.deepEqual(references(view.accepted).map(ref => ref.split('@')[0]), ['fn1', 'fn2', 'fn3', 'fn4', 'fn5'], '#8 every reference survives Accept All');
    assert.equal(references(view.accepted)[0], 'fn1@One pays promptly.', '#8 the edited sentence keeps its reference');
    assert.deepEqual(references(view.rejected), references(view.source), '#8 Reject All');
    assert.equal((view.xml.match(/<w:del\b/g) || []).length, 1, '#8 only the edited word is deleted');
}

// ---------------------------------------------------------------------------------------------------------
// 2. More shapes: endnotes, mid-sentence references, rewritten text around a reference, formatting
// ---------------------------------------------------------------------------------------------------------
{
    const end = await edit(`<w:p>${run(A)}${en(2)}</w:p>`, text => text.replace('The buyer', 'The purchaser'));
    assert.deepEqual(references(end.accepted), ['en2@The purchaser pays the deposit within ten days.'], 'endnote references survive');

    const middle = await edit(`<w:p>${run('The deposit')}${fn(3)}${run(' is due within ten days.')}</w:p>`, text => text.replace('ten days', 'fourteen days'));
    assert.deepEqual(references(middle.accepted), ['fn3@The deposit'], 'a mid-sentence reference stays after its word');

    // The word before the reference is replaced: the reference follows the replacement.
    const replaced = await edit(`<w:p>${run('Review period expires on Friday')}${fn(4)}${run('.')}</w:p>`, () => 'Review period expires on Monday afternoon.');
    assert.deepEqual(references(replaced.accepted), ['fn4@Review period expires on Monday afternoon'], 'the reference follows the replaced phrase');
    for (const inserted of Array.from(parseXml(replaced.xml).getElementsByTagNameNS(NS, 'ins'))) {
        assert.equal(inserted.getElementsByTagNameNS(NS, 'vertAlign').length, 0, 'inserted text does not inherit the reference\'s superscript');
    }

    // Text inserted right after a reference (explicit token placement) keeps normal formatting.
    const afterRef = await edit(`<w:p>${run('Due on Friday')}${fn(5)}${run('.')}</w:p>`, () => 'Due on Friday{{__FN_5__}} and Monday.');
    assert.deepEqual(references(afterRef.accepted), ['fn5@Due on Friday']);
    for (const inserted of Array.from(parseXml(afterRef.xml).getElementsByTagNameNS(NS, 'ins'))) {
        assert.equal(inserted.getElementsByTagNameNS(NS, 'vertAlign').length, 0, 'text after a reference is not superscript');
    }

    // A full rewrite of the sentence still keeps the reference.
    const rewritten = await edit(`<w:p>${run(A)}${fn(6)}</w:p>`, () => 'Payment is due on signing.');
    assert.deepEqual(references(rewritten.accepted).map(ref => ref.split('@')[0]), ['fn6'], 'a rewritten sentence keeps its reference');
    assert.deepEqual(references(rewritten.rejected), references(rewritten.source));
}

// ---------------------------------------------------------------------------------------------------------
// 3. Other routes: untracked edits, localized replacements, openDocx
// ---------------------------------------------------------------------------------------------------------
{
    const untracked = await edit(`<w:p>${run(A)}${fn(1)}</w:p>`, text => text.replace('The buyer', 'The purchaser'), { generateRedlines: false });
    assert.deepEqual(references(untracked.xml), ['fn1@The purchaser pays the deposit within ten days.'], 'untracked edits keep references');

    const documentXml = docXml(`<w:p>${run(A)}${fn(1)}</w:p>`);
    const target = inspectDocumentParts({ documentXml }).paragraphs[0];
    const localized = await applyOperationsToDocumentXml(documentXml, [{ type: 'redline', target: strong(target), replacements: [{ find: 'buyer', replace: 'purchaser' }] }],
        'Editor', {}, { atomic: true, strictTargets: true });
    assert.equal(localized.status, 'ok', JSON.stringify(localized.results?.find(item => item.error)?.error));
    assert.deepEqual(references(acceptTrackedChangesInOoxml(localized.documentXml, { allAuthors: true }).oxml), ['fn1@The purchaser pays the deposit within ten days.']);

    const entries = new Map(unzipDocx(new Uint8Array(readFileSync(new URL('./fixtures/sample_doc_test.docx', import.meta.url)))));
    entries.set('word/document.xml', new TextEncoder().encode(docXml(`<w:p>${run(A)}${fn(1)}</w:p>`)));
    const doc = openDocx(zipDocx(entries));
    const result = await doc.applyOperations([{ type: 'redline', target: strong(doc.inspect().paragraphs[0]), modified: A.replace('The buyer', 'The purchaser') }], { author: 'Editor' });
    assert.equal(result.status, 'ok', JSON.stringify(result.error ?? result.results));
    const written = new TextDecoder().decode(openDocx(result.toUint8Array()).entries.get('word/document.xml'));
    assert.deepEqual(references(acceptTrackedChangesInOoxml(written, { allAuthors: true }).oxml).map(ref => ref.split('@')[0]), ['fn1'], 'openDocx keeps the reference');
}

// ---------------------------------------------------------------------------------------------------------
// 4. Explicit deletion may remove references
// ---------------------------------------------------------------------------------------------------------
{
    const documentXml = docXml(`<w:p>${run('Keep this.')}</w:p><w:p>${run(A)}${fn(1)}</w:p>`);
    const target = inspectDocumentParts({ documentXml }).paragraphs[1];
    const deleted = await applyOperationsToDocumentXml(documentXml, [{ type: 'redline', target: strong(target), modified: '' }], 'Editor', {}, { atomic: true, strictTargets: true });
    assert.equal(deleted.status, 'ok', 'deleting the paragraph is not refused');
    assert.deepEqual(references(acceptTrackedChangesInOoxml(deleted.documentXml, { allAuthors: true }).oxml), []);
    assert.deepEqual(references(rejectTrackedChangesInOoxml(deleted.documentXml, { allAuthors: true }).oxml), ['fn1@' + A]);
}

// ---------------------------------------------------------------------------------------------------------
// 5. The ACCEPT_NON_TEXT_CONTENT safety net
// ---------------------------------------------------------------------------------------------------------
{
    const before = parseXml(docXml(`<w:p>${run(A)}${fn(1)}${en(2)}</w:p>`));
    // Lost references (a tracked deletion of the reference runs) are violations; ids are compared.
    const lost = parseXml(docXml(`<w:p>${run(A)}<w:del w:id="9" w:author="X" w:date="2026-01-01T00:00:00Z">${fn(1)}</w:del></w:p>`));
    const check = verifyNonTextContent({ before, after: lost });
    assert.equal(check.ok, false);
    assert.deepEqual(check.violations.map(v => [v.invariant, v.content]), [['ACCEPT_NON_TEXT_CONTENT', 'footnoteReference:1'], ['ACCEPT_NON_TEXT_CONTENT', 'endnoteReference:2']]);
    assert.equal(check.violations[0].source.sourceParagraphIndex, 1);
    // A renumbered reference is a different reference.
    assert.equal(verifyNonTextContent({ before, after: parseXml(docXml(`<w:p>${run(A)}${fn(7)}${en(2)}</w:p>`)) }).ok, false);
    // Moving or keeping references is fine; an inserted reference is fine; references already deleted in the source do not count.
    assert.equal(verifyNonTextContent({ before, after: parseXml(docXml(`<w:p>${en(2)}${run(A)}<w:ins w:id="9" w:author="X" w:date="2026-01-01T00:00:00Z">${fn(1)}</w:ins></w:p>`)) }).ok, true);
    const alreadyDeleted = parseXml(docXml(`<w:p>${run(A)}<w:del w:id="9" w:author="X" w:date="2026-01-01T00:00:00Z">${fn(1)}</w:del></w:p>`));
    assert.equal(verifyNonTextContent({ before: alreadyDeleted, after: alreadyDeleted }).ok, true);
    // Drawings and embedded objects are compared by count.
    const drawing = '<w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"/></w:drawing></w:r>';
    assert.equal(verifyNonTextContent({ before: parseXml(docXml(`<w:p>${run(A)}${drawing}</w:p>`)), after: parseXml(docXml(`<w:p>${run(A)}</w:p>`)) }).ok, false);
}

console.log('note_reference_preservation_tests: all passed');
