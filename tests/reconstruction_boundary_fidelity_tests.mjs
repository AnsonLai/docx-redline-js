import './setup-xml-provider.mjs';

import assert from 'node:assert/strict';

import { openDocx } from '../index.js';
import { unzipDocx, zipDocx } from '../document/zip-archive.js';

// Regressions for GitHub issues #3 (hyperlink absorbs adjacent plain
// punctuation) and #4 (Reject All reorders text around a manual line break).

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const encoder = new TextEncoder();
const decoder = new TextDecoder();

function fixture(paragraphBody) {
    return zipDocx(new Map([
        ['[Content_Types].xml', encoder.encode(`<Types xmlns="${CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`)],
        ['_rels/.rels', encoder.encode(`<Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`)],
        ['word/document.xml', encoder.encode(`<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body><w:p>${paragraphBody}</w:p><w:sectPr/></w:body></w:document>`)],
        ['word/_rels/document.xml.rels', encoder.encode(`<Relationships xmlns="${REL}"><Relationship Id="rIdLink" Type="${R}/hyperlink" Target="https://example.org/" TargetMode="External"/></Relationships>`)]
    ]));
}

const paragraphText = doc => doc.inspect().paragraphs.map(paragraph => paragraph.exactText).join('\n');

function hyperlinkTexts(bytes) {
    const xml = decoder.decode(unzipDocx(bytes).get('word/document.xml'));
    return [...xml.matchAll(/<w:hyperlink\b([^>]*)>([\s\S]*?)<\/w:hyperlink>/g)].map(match => ({
        rId: /r:id="([^"]*)"/.exec(match[1])?.[1] ?? null,
        text: [...match[2].matchAll(/<w:t\b[^>]*>([\s\S]*?)<\/w:t>/g)].map(text => text[1]).join('')
    }));
}

async function applyAndResolve({ label, source, expectedSource, edit, expectedAccepted }) {
    const doc = openDocx(source);
    const inspected = doc.inspect().paragraphs[0];
    assert.equal(paragraphText(doc), expectedSource, `${label}: source text`);

    const applied = await doc.applyOperations([{
        type: 'replace',
        target: { paragraphId: inspected.paragraphId, fingerprint: inspected.fingerprint },
        ...edit,
        insertionAffinity: { hyperlink: 'preserve', formatting: 'left' }
    }], { author: 'Boundary fidelity', atomic: true });
    assert.equal(applied.status, 'ok', `${label}: apply ${JSON.stringify(applied.error || applied.results)}`);

    const tracked = doc.toUint8Array();
    const resolved = {};
    for (const [state, resolution, expectedText] of [
        ['accepted', 'accept', expectedAccepted],
        ['rejected', 'reject', expectedSource]
    ]) {
        const resolvedDoc = openDocx(tracked);
        const result = await resolvedDoc.resolveRevisions(resolution, { allAuthors: true });
        assert.equal(result.status, 'ok', `${label}: ${state} resolution`);
        assert.equal(paragraphText(resolvedDoc), expectedText, `${label}: ${state} text`);
        resolved[state] = resolvedDoc.toUint8Array();
    }
    return resolved;
}

const hyperlinkSource = '<w:hyperlink r:id="rIdLink"><w:r><w:t>example.org</w:t></w:r></w:hyperlink><w:r><w:t>.</w:t></w:r>';
for (const [mode, edit] of [
    ['localized replacement', { replacements: [{ find: 'example.org', replace: 'example.net' }] }],
    ['full modified text', { modified: 'example.net.' }]
]) {
    const label = `hyperlink boundary (${mode})`;
    const resolved = await applyAndResolve({
        label,
        source: fixture(hyperlinkSource),
        expectedSource: 'example.org.',
        edit,
        expectedAccepted: 'example.net.'
    });
    assert.deepEqual(hyperlinkTexts(resolved.accepted), [{ rId: 'rIdLink', text: 'example.net' }], `${label}: accepted link`);
    assert.deepEqual(hyperlinkTexts(resolved.rejected), [{ rId: 'rIdLink', text: 'example.org' }], `${label}: rejected link`);
}

// A replacement wholly outside the link must not pull link text into the
// revision either.
{
    const label = 'hyperlink boundary (trailing plain text edit)';
    const resolved = await applyAndResolve({
        label,
        source: fixture('<w:hyperlink r:id="rIdLink"><w:r><w:t>example.org</w:t></w:r></w:hyperlink><w:r><w:t>.foo</w:t></w:r>'),
        expectedSource: 'example.org.foo',
        edit: { modified: 'example.org.bar' },
        expectedAccepted: 'example.org.bar'
    });
    assert.deepEqual(hyperlinkTexts(resolved.accepted), [{ rId: 'rIdLink', text: 'example.org' }], `${label}: accepted link`);
}

const lineBreakSource = '<w:r><w:tab/><w:t>tabbed</w:t><w:br/><w:t xml:space="preserve">Line with </w:t></w:r>';
for (const [mode, edit] of [
    ['localized replacement', { replacements: [{ find: 'tabbed', replace: 'aligned' }] }],
    ['full modified text', { modified: '\taligned\nLine with ' }]
]) {
    await applyAndResolve({
        label: `manual line break (${mode})`,
        source: fixture(lineBreakSource),
        expectedSource: '\ttabbed\nLine with ',
        edit,
        expectedAccepted: '\taligned\nLine with '
    });
}

// Text changed on both sides of the break keeps each change on its own side.
await applyAndResolve({
    label: 'manual line break (edits on both sides)',
    source: fixture('<w:r><w:t>before</w:t><w:br/><w:t>after</w:t></w:r>'),
    expectedSource: 'before\nafter',
    edit: { modified: 'earlier\nlater' },
    expectedAccepted: 'earlier\nlater'
});

console.log('PASS: reconstruction boundary fidelity tests');
