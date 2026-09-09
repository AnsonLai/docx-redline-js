import assert from 'assert/strict';

import './setup-xml-provider.mjs';
import {
    acceptTrackedChangesInOoxml,
    applyRedlineToOxml,
    rejectTrackedChangesInOoxml,
    validateRedlineOoxml
} from '../index.js';
import { parseOoxmlSafe } from '../adapters/xml-adapter.js';
import { extractCanonicalParagraphText } from '../core/paragraph-text.js';
import { applyOperationsToDocumentXml } from '../services/standalone-operation-runner.js';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const DATE = '2026-09-08T16:00:00Z';

function escapeXml(text) {
    return String(text)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;');
}

function run(text, rPr = '') {
    const preserve = /^\s|\s$/.test(text) ? ' xml:space="preserve"' : '';
    return `<w:r>${rPr}<w:t${preserve}>${escapeXml(text)}</w:t></w:r>`;
}

function insertion(id, author, content) {
    return `<w:ins w:id="${id}" w:author="${author}" w:date="${DATE}">${content}</w:ins>`;
}

function paragraph(content) {
    return `<w:p xmlns:w="${W}" xmlns:r="${R}">${content}</w:p>`;
}

function parsed(xml) {
    const result = parseOoxmlSafe(xml, 'application/xml');
    assert.equal(result.error, null);
    return result.doc;
}

function acceptedText(xml) {
    const doc = parsed(xml);
    return extractCanonicalParagraphText(doc.getElementsByTagNameNS(W, 'p')[0]);
}

function authorOf(node) {
    return node.getAttributeNS(W, 'author') || node.getAttribute('w:author');
}

function assertValid(xml, label) {
    const validation = validateRedlineOoxml(xml);
    assert.equal(validation.valid, true, `${label}: ${JSON.stringify(validation.issues)}`);
    const emptyRevisions = Array.from(parsed(xml).getElementsByTagNameNS(W, 'ins'))
        .filter(node => !(node.textContent || '').length && node.getElementsByTagNameNS(W, 'tab').length === 0);
    assert.equal(emptyRevisions.length, 0, `${label}: empty insertion wrapper`);
}

async function assertInsertionRoundTrip({ name, source, original, position, payload, runner = false }) {
    const modified = original.slice(0, position) + payload + original.slice(position);
    const result = await applyRedlineToOxml(source, original, modified, {
        author: 'Current Reviewer',
        existingRevisions: 'slice-cross-author',
        pairReplacements: true,
        structuredContent: false
    });
    assert.equal(result.status, 'ok', `${name}: ${JSON.stringify(result.error)}`);
    assert.equal(result.hasChanges, true, `${name}: expected a change`);
    assert.equal(acceptedText(result.oxml), modified, `${name}: generated accepted view`);
    assertValid(result.oxml, name);

    const currentInsertions = Array.from(parsed(result.oxml).getElementsByTagNameNS(W, 'ins'))
        .filter(node => authorOf(node) === 'Current Reviewer');
    assert(currentInsertions.length >= 1, `${name}: current-author insertion missing`);

    const acceptAll = acceptTrackedChangesInOoxml(result.oxml, { allAuthors: true });
    assert.equal(acceptedText(acceptAll.oxml), modified, `${name}: Accept All`);
    const rejectCurrent = rejectTrackedChangesInOoxml(result.oxml, { author: 'Current Reviewer' });
    assert.equal(acceptedText(rejectCurrent.oxml), original, `${name}: Reject Current Reviewer`);

    if (runner) {
        const documentXml = `<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>${source}<w:sectPr/></w:body></w:document>`;
        const batch = await applyOperationsToDocumentXml(documentXml, [{
            type: 'redline',
            target: { exactText: original },
            modified,
            author: 'Current Reviewer',
            existingRevisions: 'slice-cross-author',
            structuredContent: false
        }], 'Current Reviewer', null, { atomic: true, strictTargets: true });
        assert.equal(batch.status, 'ok', `${name} runner: ${JSON.stringify(batch.error)}`);
        assert.equal(batch.results[0]?.status, 'applied', `${name} runner operation`);
        assert.equal(acceptedText(batch.documentXml), modified, `${name} runner accepted view`);
    }
}

const payloads = ['X', ' ', '  ', '\u00a0', '—', '😀', '\t', ' & < > ', '👩‍⚖️', '§ 2.1 '];
const carrierCases = [
    { name: 'single-run', parts: ['Agreement text'] },
    { name: 'repeated-phrase', parts: ['Policy applies. Policy applies.'] },
    { name: 'multi-run', parts: ['First ', 'second ', 'third'] },
    { name: 'formatted-runs', parts: ['Bold ', 'plain ', 'italic'], formats: ['<w:rPr><w:b/></w:rPr>', '', '<w:rPr><w:i/></w:rPr>'] },
    { name: 'unicode', parts: ['世界 ', '😀 ', 'café'] },
    { name: 'combining', parts: ['Cafe\u0301 ', 're\u0301sume\u0301 ', 'term'] },
    { name: 'existing-whitespace', parts: ['Alpha  ', 'beta\u00a0', 'gamma'] },
    { name: 'numeric-repeat', parts: ['12345 ', '12345 ', '12345'] }
];

let executed = 0;
for (const [caseIndex, testCase] of carrierCases.entries()) {
    const carrierText = testCase.parts.join('');
    const original = `Baseline ${carrierText}`;
    const carrierRuns = testCase.parts.map((part, index) => run(part, testCase.formats?.[index] || '')).join('');
    const source = paragraph(run('Baseline ') + insertion(10 + caseIndex, 'Prior Reviewer', carrierRuns));
    const runBoundaries = testCase.parts.slice(0, -1).map((_, index) => {
        return 'Baseline '.length + testCase.parts.slice(0, index + 1).join('').length;
    });
    const positions = [...new Set([
        'Baseline '.length,
        'Baseline '.length + 1,
        ...runBoundaries,
        'Baseline '.length + Math.floor(carrierText.length / 2),
        original.length - 1,
        original.length
    ])].filter(position => position >= 'Baseline '.length && position <= original.length);

    for (const [positionIndex, position] of positions.entries()) {
        await assertInsertionRoundTrip({
            name: `${testCase.name}@${position}`,
            source,
            original,
            position,
            payload: payloads[(caseIndex + positionIndex) % payloads.length],
            runner: executed < 12
        });
        executed++;
    }
}

// Every XML-/Unicode-/whitespace-sensitive payload gets a dedicated interior case.
for (const [index, payload] of payloads.entries()) {
    const source = paragraph(run('Base ') + insertion(400 + index, 'Prior Reviewer', run('left right')));
    await assertInsertionRoundTrip({
        name: `payload-${index}`,
        source,
        original: 'Base left right',
        position: 'Base left'.length,
        payload,
        runner: index < 3
    });
}

// One operation with insertions in three different structural containers.
{
    const source = paragraph(
        insertion(500, 'Prior Reviewer', run('Alpha '))
        + `<w:hyperlink r:id="rIdMulti">${run('Link ')}</w:hyperlink>`
        + insertion(501, 'Other Reviewer', run('Omega'))
    );
    const original = 'Alpha Link Omega';
    const modified = 'Al-FIRST-pha Li-SECOND-nk Ome-THIRD-ga';
    const result = await applyRedlineToOxml(source, original, modified, {
        author: 'Current Reviewer',
        existingRevisions: 'slice-cross-author',
        pairReplacements: true,
        structuredContent: false
    });
    assert.equal(result.status, 'ok', JSON.stringify(result.error));
    assert.equal(acceptedText(result.oxml), modified);
    assertValid(result.oxml, 'multi-container-insertions');
    assert.equal(parsed(result.oxml).getElementsByTagNameNS(W, 'hyperlink').length, 1);
}

// Consecutive rounds by the same reviewer and then a third reviewer.
{
    const source = paragraph(run('Base ') + insertion(600, 'Prior Reviewer', run('contract text')));
    const firstText = 'Base contract NEW text';
    const first = await applyRedlineToOxml(source, 'Base contract text', firstText, {
        author: 'Current Reviewer', existingRevisions: 'slice-cross-author', structuredContent: false
    });
    assert.equal(first.status, 'ok', JSON.stringify(first.error));

    const secondText = 'Base contract NEWER text';
    const second = await applyRedlineToOxml(first.oxml, firstText, secondText, {
        author: 'Current Reviewer', existingRevisions: 'slice-cross-author', structuredContent: false
    });
    assert.equal(second.status, 'ok', JSON.stringify(second.error));
    assert.equal(acceptedText(second.oxml), secondText);
    assertValid(second.oxml, 'same-reviewer-second-round');

    const thirdText = 'Base contract THIRD NEWER text';
    const third = await applyRedlineToOxml(second.oxml, secondText, thirdText, {
        author: 'Third Reviewer', existingRevisions: 'slice-cross-author', structuredContent: false
    });
    assert.equal(third.status, 'ok', JSON.stringify(third.error));
    assert.equal(acceptedText(third.oxml), thirdText);
    assertValid(third.oxml, 'third-reviewer-round');
}

// Nested structural containers may succeed exactly or fail closed, but may never throw or return corrupt success.
for (const [name, content, original, modified] of [
    [
        'hyperlink-inside-foreign-insertion',
        insertion(700, 'Prior Reviewer', run('Before ') + `<w:hyperlink r:id="rIdNested">${run('link')}</w:hyperlink>` + run(' after')),
        'Before link after',
        'Before linked link after'
    ],
    [
        'simple-field-beside-foreign-insertion',
        insertion(701, 'Prior Reviewer', run('Before ')) + `<w:fldSimple w:instr=" DATE ">${run('September 8')}</w:fldSimple>` + insertion(702, 'Prior Reviewer', run(' after')),
        'Before September 8 after',
        'Before current September 8 after'
    ]
]) {
    let result;
    await assert.doesNotReject(async () => {
        result = await applyRedlineToOxml(paragraph(content), original, modified, {
            author: 'Current Reviewer', existingRevisions: 'slice-cross-author', structuredContent: false
        });
    }, `${name}: public transform must not throw`);
    if (result.status === 'ok') {
        assert.equal(acceptedText(result.oxml), modified);
        assertValid(result.oxml, name);
    } else {
        assert.equal(result.hasChanges, false, `${name}: safe failure must be unapplied`);
        assert(['TARGET_NOT_FOUND', 'PATCH_ROUNDTRIP_MISMATCH', 'UNSAFE_REVISION_NESTING'].includes(result.error?.code),
            `${name}: unexpected error ${JSON.stringify(result.error)}`);
    }
}

// Hyperlink interiors and both sides of hyperlink boundaries with repeated text.
{
    const source = paragraph(
        insertion(100, 'Prior Reviewer', run('Before Policy '))
        + `<w:hyperlink r:id="rIdPolicy">${run('Policy link')}</w:hyperlink>`
        + insertion(101, 'Other Reviewer', run(' after Policy link.'))
    );
    const original = 'Before Policy Policy link after Policy link.';
    for (const [index, position] of [
        original.indexOf('Policy'),
        original.indexOf('Policy link') + 3,
        original.indexOf(' after'),
        original.lastIndexOf('Policy') + 'Policy'.length,
        original.length
    ].entries()) {
        await assertInsertionRoundTrip({
            name: `hyperlink-boundary-${index}`,
            source,
            original,
            position,
            payload: payloads[index],
            runner: index < 2
        });
        const result = await applyRedlineToOxml(source, original,
            original.slice(0, position) + payloads[index] + original.slice(position), {
                author: 'Current Reviewer', existingRevisions: 'slice-cross-author', structuredContent: false
            });
        const hyperlinks = parsed(result.oxml).getElementsByTagNameNS(W, 'hyperlink');
        assert.equal(hyperlinks.length, 1);
        assert.equal(hyperlinks[0].getAttribute('r:id'), 'rIdPolicy');
    }
}

// Bookmarks, comments, and a prior nested deletion remain structurally intact.
for (const [name, middle] of [
    ['bookmark', '<w:bookmarkStart w:id="8" w:name="Clause"/><w:bookmarkEnd w:id="8"/>'],
    ['comment', '<w:commentRangeStart w:id="9"/><w:commentRangeEnd w:id="9"/>'],
    ['nested-deletion', `<w:del w:id="202" w:author="Earlier Reviewer" w:date="${DATE}"><w:r><w:delText xml:space="preserve">hidden text </w:delText></w:r></w:del>`]
]) {
    const source = paragraph(insertion(200, 'Prior Reviewer', run('Alpha ') + middle + run('Omega')));
    await assertInsertionRoundTrip({
        name,
        source,
        original: 'Alpha Omega',
        position: 'Alpha '.length,
        payload: 'inserted ',
        runner: true
    });
}

// Mixed same-author and foreign carriers must never return schema-invalid nested insertions.
{
    const source = paragraph(
        insertion(300, 'Current Reviewer', run('Current text '))
        + insertion(301, 'Prior Reviewer', run('foreign text'))
    );
    const original = 'Current text foreign text';
    const modified = 'Current NEW text foreign text';
    const result = await applyRedlineToOxml(source, original, modified, {
        author: 'Current Reviewer',
        existingRevisions: 'slice-cross-author',
        structuredContent: false
    });
    if (result.status === 'ok') {
        assert.equal(acceptedText(result.oxml), modified);
        assertValid(result.oxml, 'mixed-current-and-foreign');
    } else {
        assert(['PATCH_ROUNDTRIP_MISMATCH', 'UNSAFE_REVISION_NESTING'].includes(result.error?.code),
            `unexpected safe failure: ${JSON.stringify(result.error)}`);
        assert.equal(result.hasChanges, false);
        assert.equal(result.oxml, source);
    }
}

assert(executed >= 40, `expected at least 40 generated carrier cases, got ${executed}`);
console.log(`PASS: ${executed + payloads.length + 14} cross-author insertion stress cases`);
