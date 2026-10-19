import './setup-xml-provider.mjs';

// A caller's structuredContent: false keeps inline Markdown literal in paragraphs that a
// redline adds (before, after or around the target, list siblings, ranges), not only in
// the rewritten target. Regression: the plain-adjacency insertion route built each new
// paragraph through an engine call that never received plainTextContent.
import assert from 'assert/strict';
import { readFileSync } from 'node:fs';
import { acceptTrackedChangesInOoxml, inspectDocumentParts, openDocx, rejectTrackedChangesInOoxml } from '../index.js';
import { applyOperationsToDocumentXml } from '../services/standalone-operation-runner.js';
import { unzipDocx, zipDocx } from '../document/zip-archive.js';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const numbering = `<w:numbering ${W}>`
    + '<w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>'
    + '<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%2)"/></w:lvl></w:abstractNum>'
    + '<w:abstractNum w:abstractNumId="2"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/></w:lvl></w:abstractNum>'
    + '<w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="2"/></w:num></w:numbering>';
const p = text => `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
const item = (text, level = 0, numId = 1) => `<w:p><w:pPr><w:numPr><w:ilvl w:val="${level}"/><w:numId w:val="${numId}"/></w:numPr></w:pPr><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
const docXml = body => `<w:document ${W}><w:body>${body}<w:sectPr/></w:body></w:document>`;
const strong = q => ({ index: q.index, exactText: q.exactText, paragraphId: q.paragraphId, fingerprint: q.fingerprint });
const inspect = xml => inspectDocumentParts({ documentXml: xml, numberingXml: numbering }).paragraphs;
const MARKERS = 'Use __bold__, _italic_, **strong**, ~~gone~~ and ++under++';

async function apply(body, buildOperation, options = {}) {
    const documentXml = docXml(body);
    const paragraphs = inspect(documentXml);
    const operation = { type: 'redline', ...buildOperation(paragraphs) };
    const result = await applyOperationsToDocumentXml(documentXml, [operation], 'Tester', { numberingXml: numbering },
        { atomic: true, strictTargets: true, ...options });
    assert.equal(result.status, 'ok', JSON.stringify(result.error ?? result.results));
    const acceptedXml = acceptTrackedChangesInOoxml(result.documentXml, { allAuthors: true }).oxml;
    return {
        accepted: inspect(acceptedXml),
        acceptedXml,
        rejected: inspect(rejectTrackedChangesInOoxml(result.documentXml, { allAuthors: true }).oxml).map(q => q.exactText),
        source: paragraphs.map(q => q.exactText)
    };
}
const texts = view => view.accepted.map(q => q.exactText);
const noEmphasisRuns = xml => assert.doesNotMatch(xml, /<w:(b|i|strike|u)\b(?![^>]*w:val="(0|false|none)")/, 'no emphasis formatting');

// ---------------------------------------------------------------------------------------------------------
// 1. Plain paragraphs added around the target
// ---------------------------------------------------------------------------------------------------------
for (const [name, modified, expected] of [
    ['after', `Plain intro.\n${MARKERS}`, ['Plain intro.', MARKERS]],
    ['before', `${MARKERS}\nPlain intro.`, [MARKERS, 'Plain intro.']],
    ['before and after', `New __a__\nPlain intro.\nNew _b_`, ['New __a__', 'Plain intro.', 'New _b_']],
    ['several after', `Plain intro.\nOne __x__\nTwo _y_\nThree **z**`, ['Plain intro.', 'One __x__', 'Two _y_', 'Three **z**']],
    ['signature line', `Plain intro.\nSigned: ${'_'.repeat(20)} Date: ${'_'.repeat(10)}`, ['Plain intro.', `Signed: ${'_'.repeat(20)} Date: ${'_'.repeat(10)}`]],
    ['changed target', `Plain intro, changed.\n${MARKERS}`, ['Plain intro, changed.', MARKERS]],
    ['heading-like text', 'Plain intro.\n# Not a heading', ['Plain intro.', '# Not a heading']],
]) {
    const view = await apply(p('Plain intro.'), ps => ({ target: strong(ps[0]), modified, structuredContent: false }));
    assert.deepEqual(texts(view), expected, `${name}: Accept All`);
    assert.deepEqual(view.rejected, view.source, `${name}: Reject All restores the source`);
    noEmphasisRuns(view.acceptedXml);
    assert.doesNotMatch(view.acceptedXml, /w:pStyle w:val="Heading/, `${name}: no heading style`);
}

// The batch option applies when the operation does not say; the operation wins over the batch.
{
    const batch = await apply(p('Plain intro.'), ps => ({ target: strong(ps[0]), modified: `Plain intro.\n${MARKERS}` }), { structuredContent: false });
    assert.deepEqual(texts(batch), ['Plain intro.', MARKERS]);
    const override = await apply(p('Plain intro.'), ps => ({ target: strong(ps[0]), modified: 'Plain intro.\nUse __bold__ here', structuredContent: true }), { structuredContent: false });
    assert.deepEqual(texts(override), ['Plain intro.', 'Use bold here'], 'structuredContent: true still formats inserted paragraphs');
    assert.match(override.acceptedXml, /<w:b\/>|<w:b w:val="(1|true)"\/>/);
}

// replace, untracked edits and openDocx follow the same contract.
{
    const replaced = await apply(p('Plain intro.'), ps => ({ type: 'replace', target: strong(ps[0]), modified: `Plain intro.\n${MARKERS}`, structuredContent: false }));
    assert.deepEqual(texts(replaced), ['Plain intro.', MARKERS]);
    const untracked = await apply(p('Plain intro.'), ps => ({ target: strong(ps[0]), modified: `Plain intro.\n${MARKERS}`, structuredContent: false }), { generateRedlines: false });
    assert.deepEqual(texts(untracked), ['Plain intro.', MARKERS]);
    assert.deepEqual(untracked.rejected, ['Plain intro.', MARKERS], 'untracked: nothing to reject');

    const entries = new Map(unzipDocx(new Uint8Array(readFileSync(new URL('./fixtures/sample_doc_test.docx', import.meta.url)))));
    entries.set('word/document.xml', new TextEncoder().encode(docXml(p('Plain intro.'))));
    const doc = openDocx(zipDocx(entries));
    const result = await doc.applyOperations([{ type: 'redline', target: strong(doc.inspect().paragraphs[0]), modified: `Plain intro.\n${MARKERS}`, structuredContent: false }], { author: 'Tester' });
    assert.equal(result.status, 'ok', JSON.stringify(result.error ?? result.results));
    const written = new TextDecoder().decode(openDocx(result.toUint8Array()).entries.get('word/document.xml'));
    assert.deepEqual(inspect(acceptTrackedChangesInOoxml(written, { allAuthors: true }).oxml).map(q => q.exactText), ['Plain intro.', MARKERS]);
}

// ---------------------------------------------------------------------------------------------------------
// 2. Ranges
// ---------------------------------------------------------------------------------------------------------
{
    const added = await apply(p('One.') + p('Two.'), ps => ({ target: strong(ps[0]), targetEnd: strong(ps[1]), modified: `One.\nTwo.\n${MARKERS}`, structuredContent: false }));
    assert.deepEqual(texts(added), ['One.', 'Two.', MARKERS]);
    assert.deepEqual(added.rejected, ['One.', 'Two.']);
    const rewritten = await apply(p('One.') + p('Two.'), ps => ({ target: strong(ps[0]), targetEnd: strong(ps[1]), modified: 'One __a__.\nTwo _b_.', structuredContent: false }));
    assert.deepEqual(texts(rewritten), ['One __a__.', 'Two _b_.']);
}

// ---------------------------------------------------------------------------------------------------------
// 3. List siblings keep the list and stay literal (decimal level 0, level 1, bullets)
// ---------------------------------------------------------------------------------------------------------
for (const [name, body, targetIndex, modified, expectedTexts, level, numId] of [
    ['decimal level 0', item('First item') + item('Second item'), 0, 'First item\nNew __x__ item', ['First item', 'New __x__ item', 'Second item'], 0, 1],
    ['decimal level 1', item('Top') + item('Sub one', 1), 1, 'Sub one\nSub __two__', ['Top', 'Sub one', 'Sub __two__'], 1, 1],
    ['bullets', item('Apples', 0, 2), 0, 'Apples\nPears _and_ plums', ['Apples', 'Pears _and_ plums'], 0, 2],
]) {
    const view = await apply(body, ps => ({ target: strong(ps[targetIndex]), modified, structuredContent: false }));
    assert.deepEqual(texts(view), expectedTexts, `${name}: Accept All`);
    assert.deepEqual(view.rejected, view.source, `${name}: Reject All`);
    const inserted = view.accepted.find(q => q.exactText === expectedTexts.find(t => /_/.test(t)));
    assert.ok(inserted?.list, `${name}: the new paragraph is a list item`);
    assert.equal(inserted.list.level, level, `${name}: list level`);
    assert.match(view.acceptedXml, new RegExp(`<w:numId w:val="${numId}"/>`), `${name}: same list`);
}

// ---------------------------------------------------------------------------------------------------------
// 4. Paragraph-mark revisions follow Word's model, in both structuredContent modes.
//    'after' = Enter at the end of the target: the target's mark is the inserted break and the last new
//    paragraph inherits the target's original mark. At the end of the body, the final paragraph mark
//    must never be tracked (Word cannot reject it cleanly). 'before': each new paragraph's own mark.
// ---------------------------------------------------------------------------------------------------------
{
    const { parseXml } = await import('../adapters/xml-adapter.js');
    const NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
    // Per body paragraph: the tracked revision on its mark ('ins' | 'del' | null) and that marker's author.
    const marks = xml => Array.from(parseXml(xml).getElementsByTagNameNS(NS, 'body')[0].childNodes)
        .filter(node => node.nodeType === 1 && node.localName === 'p')
        .map(paragraph => {
            const pPr = Array.from(paragraph.childNodes).find(node => node.localName === 'pPr');
            const rPr = Array.from(pPr?.childNodes || []).find(node => node.localName === 'rPr');
            const marker = Array.from(rPr?.childNodes || []).find(node => node.localName === 'ins' || node.localName === 'del');
            return marker ? `${marker.localName}:${marker.getAttribute('w:author')}` : null;
        });
    const run = async (body, build, sc, extra = {}) => {
        const documentXml = docXml(body);
        const result = await applyOperationsToDocumentXml(documentXml, [{ type: 'redline', structuredContent: sc, ...build(inspect(documentXml)) }],
            'Tester', { numberingXml: numbering }, { atomic: true, strictTargets: true, ...extra });
        assert.equal(result.status, 'ok', JSON.stringify(result.error ?? result.results));
        return result.documentXml;
    };
    for (const sc of [false, true]) {
        const label = `structuredContent ${sc}`;
        // Body end: the reported Word-lane failure.
        const end = await run(p('First.') + p('Last paragraph.'), ps => ({ target: strong(ps[1]), modified: 'Last paragraph.\nAppended.' }), sc);
        assert.deepEqual(marks(end), [null, 'ins:Tester', null], `${label}: target mark is the inserted break; the final mark is untracked`);
        // Several appended paragraphs: every new mark but the last is tracked.
        const several = await run(p('First.') + p('Last.'), ps => ({ target: strong(ps[1]), modified: 'Last.\nOne.\nTwo.' }), sc);
        assert.deepEqual(marks(several), [null, 'ins:Tester', 'ins:Tester', null], `${label}: several appended`);
        assert.deepEqual(inspect(rejectTrackedChangesInOoxml(several, { allAuthors: true }).oxml).map(q => q.exactText), ['First.', 'Last.']);
        // Mid-document append keeps the same shape.
        const middle = await run(p('Target.') + p('After.'), ps => ({ target: strong(ps[0]), modified: 'Target.\nNew.' }), sc);
        assert.deepEqual(marks(middle), ['ins:Tester', null, null], `${label}: mid-document append`);
    }
    // Plain-text specifics: insert before keeps each new paragraph's own mark; a target whose mark was
    // already inserted by another reviewer hands that identity to the last new paragraph.
    const before = await run(p('First.') + p('Last.'), ps => ({ target: strong(ps[1]), modified: 'New one.\nNew two.\nLast.' }), false);
    assert.deepEqual(marks(before), [null, 'ins:Tester', 'ins:Tester', null], 'before: new paragraphs track their own marks');
    const seeded = p('First.') + '<w:p><w:pPr><w:rPr><w:ins w:id="7" w:author="Other" w:date="2026-01-01T00:00:00Z"/></w:rPr></w:pPr>'
        + '<w:ins w:id="8" w:author="Other" w:date="2026-01-01T00:00:00Z"><w:r><w:t>Seeded.</w:t></w:r></w:ins></w:p>' + p('Tail.');
    const handedOver = await run(seeded, ps => ({ target: strong(ps[1]), modified: 'Seeded.\nAppended.' }), false,
        { existingRevisions: 'slice-cross-author' });
    assert.deepEqual(marks(handedOver), [null, 'ins:Tester', 'ins:Other', null], 'the inherited mark keeps the other reviewer\'s identity');
    assert.match(handedOver, /w:id="7" w:author="Other"/, 'inherited marker keeps its id');
    assert.deepEqual(inspect(rejectTrackedChangesInOoxml(handedOver, { allAuthors: true }).oxml).map(q => q.exactText), ['First.', 'Tail.']);
    assert.deepEqual(inspect(acceptTrackedChangesInOoxml(handedOver, { allAuthors: true }).oxml).map(q => q.exactText), ['First.', 'Seeded.', 'Appended.', 'Tail.']);
}

console.log('plain_text_paragraph_insertion_tests: all passed');
