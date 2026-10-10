import assert from 'node:assert/strict';

import './setup-xml-provider.mjs';
import { parseOoxmlSafe, serializeXml } from '../adapters/xml-adapter.js';
import { RevisionIdAllocator, NS_W } from '../core/types.js';
import {
    clearParagraphMarkRevision,
    markParagraphDeleted,
    markParagraphInserted,
    paragraphMarkPPrXml,
    paragraphMarkRevisionXml,
    planParagraphMarks,
    restoreParagraphMarkInserted,
    setParagraphPropertiesMarkRevision
} from '../engine/paragraph-mark-revisions.js';
import { applyOperationsToDocumentXml } from '../services/standalone-operation-runner.js';
import { wordRevisionParagraphTexts } from '../core/word-revision-semantics.js';

const W = NS_W;
const parse = xml => {
    const parsed = parseOoxmlSafe(xml);
    assert.ok(parsed.doc, parsed.error?.message);
    return parsed.doc;
};
const paragraphsOf = doc => Array.from(doc.getElementsByTagNameNS('*', 'p'));
const markOf = (p, kind) => Array.from(p.getElementsByTagNameNS('*', kind))
    .filter(n => n.parentNode?.localName === 'rPr' && n.parentNode.parentNode?.localName === 'pPr' && n.parentNode.parentNode.parentNode === p);
const ctx = (allocator = new RevisionIdAllocator(100)) => ({ author: 'Unit Author', date: '2026-01-02T03:04:05Z', revisionIdAllocator: allocator });

// DOM: inserted and deleted marks, ids from the supplied allocator, pPr created first, rPr last.
{
    const doc = parse(`<w:document xmlns:w="${W}"><w:body><w:p><w:r><w:t>A</w:t></w:r></w:p><w:p><w:pPr><w:pStyle w:val="X"/></w:pPr></w:p></w:body></w:document>`);
    const [a, b] = paragraphsOf(doc);
    const allocator = new RevisionIdAllocator(100);
    const ins = markParagraphInserted(a, ctx(allocator));
    const del = markParagraphDeleted(b, ctx(allocator));
    assert.equal(ins.localName, 'ins');
    assert.equal(ins.getAttribute('w:id'), '100');
    assert.equal(ins.getAttribute('w:author'), 'Unit Author');
    assert.equal(ins.getAttribute('w:date'), '2026-01-02T03:04:05Z');
    assert.equal(del.getAttribute('w:id'), '101', 'ids come from the supplied allocator');
    assert.equal(a.firstChild.localName, 'pPr', 'pPr is the first child');
    assert.equal(markOf(a, 'ins').length, 1);
    assert.equal(markOf(b, 'del').length, 1);
    assert.equal(b.firstChild.lastChild.localName, 'rPr', 'rPr is last in pPr');

    // Replaces the opposite kind and never duplicates.
    markParagraphDeleted(a, ctx(allocator));
    assert.equal(markOf(a, 'ins').length, 0);
    assert.equal(markOf(a, 'del').length, 1);
    markParagraphDeleted(a, ctx(allocator));
    assert.equal(markOf(a, 'del').length, 1);

    assert.equal(clearParagraphMarkRevision(a), true);
    assert.equal(markOf(a, 'del').length, 0);
    assert.equal(a.getElementsByTagNameNS('*', 'rPr').length, 0, 'empty rPr removed');
    assert.equal(clearParagraphMarkRevision(a), false);
}

// DOM: explicit metadata skips allocation; no allocator falls back to the owner document.
{
    const doc = parse(`<w:document xmlns:w="${W}"><w:body><w:p/></w:body></w:document>`);
    const [p] = paragraphsOf(doc);
    const marker = markParagraphInserted(p, { metadata: { id: 9, author: 'Pre', date: '2026-02-02T00:00:00Z' } });
    assert.equal(marker.getAttribute('w:id'), '9');
    assert.equal(marker.getAttribute('w:author'), 'Pre');
    const other = markParagraphDeleted(p, { author: 'Doc Author' });
    assert.ok(Number.isInteger(Number(other.getAttribute('w:id'))));
}

// restoreParagraphMarkInserted keeps the captured identity and is idempotent.
{
    const doc = parse(`<w:document xmlns:w="${W}"><w:body><w:p/></w:body></w:document>`);
    const [p] = paragraphsOf(doc);
    const marker = restoreParagraphMarkInserted(p, { id: '55', author: 'Keeper', date: '2026-03-03T00:00:00Z' });
    assert.equal(marker.getAttribute('w:id'), '55');
    assert.equal(restoreParagraphMarkInserted(p, { id: '56', author: 'Keeper' }), null);
    assert.equal(markOf(p, 'ins').length, 1);
}

// Detached pPr (list generation parses a pPr string).
{
    const doc = parse(`<w:root xmlns:w="${W}"><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr></w:root>`);
    const pPr = doc.documentElement.firstChild;
    setParagraphPropertiesMarkRevision(pPr, 'del', ctx());
    const xml = serializeXml(pPr);
    assert.match(xml, /<w:rPr><w:del w:id="100" w:author="Unit Author" w:date="2026-01-02T03:04:05Z"\/><\/w:rPr>/);
    assert.throws(() => setParagraphPropertiesMarkRevision(pPr, 'bogus', ctx()), /Unsupported/);
}

// rPr placement follows CT_PPr order (before sectPr / pPrChange) and an existing rPr stays put.
{
    const doc = parse(`<w:root xmlns:w="${W}"><w:p><w:pPr><w:pStyle w:val="X"/><w:pPrChange w:id="1" w:author="A" w:date="2026-01-01T00:00:00Z"><w:pPr/></w:pPrChange></w:pPr></w:p><w:p><w:pPr><w:rPr><w:b/></w:rPr><w:sectPr/></w:pPr></w:p></w:root>`);
    const [first, second] = paragraphsOf(doc);
    markParagraphInserted(first, ctx());
    markParagraphDeleted(second, ctx());
    const names = p => Array.from(p.firstChild.childNodes).map(n => n.localName);
    assert.deepEqual(names(first), ['pStyle', 'rPr', 'pPrChange']);
    assert.deepEqual(names(second), ['rPr', 'sectPr']);
    assert.deepEqual(Array.from(second.firstChild.firstChild.childNodes).map(n => n.localName), ['b', 'del']);
}

// String builders share allocation and escaping with the DOM path.
{
    const allocator = new RevisionIdAllocator(7);
    const marker = paragraphMarkRevisionXml('ins', { ...ctx(allocator), author: 'A & "B"' });
    assert.equal(marker, '<w:ins w:id="7" w:author="A &amp; &quot;B&quot;" w:date="2026-01-02T03:04:05Z"/>');
    assert.equal(paragraphMarkPPrXml('del', ctx(allocator)),
        '<w:pPr><w:rPr><w:del w:id="8" w:author="Unit Author" w:date="2026-01-02T03:04:05Z"/></w:rPr></w:pPr>');
}

// planParagraphMarks invariants.
{
    const plan = planParagraphMarks({
        sourceParagraphs: ['a', 'b', 'c'],
        targetParagraphs: ['a2', 'x', 'y'],
        alignment: [
            { sourceIndex: 0, targetIndex: 0 },
            { sourceIndex: 1, targetIndex: null },
            { sourceIndex: 2, targetIndex: 1, propertyChange: true },
            { sourceIndex: null, targetIndex: 2 }
        ]
    });
    assert.deepEqual(plan.map(e => e.action), ['keep', 'deleteMark', 'propertyChange', 'insertMark']);
    assert.throws(() => planParagraphMarks({ sourceParagraphs: ['a', 'b'], targetParagraphs: [], alignment: [{ sourceIndex: 0 }] }), /drops 1 source/);
    assert.throws(() => planParagraphMarks({ sourceParagraphs: [], targetParagraphs: ['x'], alignment: [] }), /without a source/);
    assert.throws(() => planParagraphMarks({ sourceParagraphs: ['a'], targetParagraphs: [], alignment: [{ sourceIndex: 0 }, { sourceIndex: 0 }] }), /more than once/);
    assert.throws(() => planParagraphMarks({ sourceParagraphs: [], targetParagraphs: [], alignment: [{}] }), /neither/);
}

// The three list-generation rounds as fixtures: the generated marks must match the plan's counts.
{
    const item = text => `<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>${text ? `<w:r><w:t>${text}</w:t></w:r>` : ''}</w:p>`;
    const body = items => `<w:document xmlns:w="${W}"><w:body><w:p><w:r><w:t>Intro</w:t></w:r></w:p>${items.map(item).join('')}<w:sectPr/></w:body></w:document>`;
    const rounds = [
        { name: 'range collapse', items: ['Alpha', 'Beta', 'Gamma'], target: { index: 2, exactText: 'Alpha' }, targetEnd: { index: 4, exactText: 'Gamma' } },
        { name: 'empty item', items: ['Alpha', '', 'Gamma'], target: { index: 2, exactText: 'Alpha' }, targetEnd: { index: 4, exactText: 'Gamma' } },
        { name: 'all-empty range', items: ['', ''], target: { index: 2, exactText: '' }, targetEnd: { index: 3, exactText: '' } }
    ];
    const targetTexts = ['One', 'Two'];
    for (const round of rounds) {
        const sourceParagraphs = ['Intro', ...round.items];
        const targetParagraphs = ['Intro', ...targetTexts];
        const alignment = [{ sourceIndex: 0, targetIndex: 0 }];
        round.items.forEach((_, i) => alignment.push({ sourceIndex: i + 1, targetIndex: null }));
        targetTexts.forEach((_, i) => alignment.push({ sourceIndex: null, targetIndex: i + 1 }));
        const plan = planParagraphMarks({ sourceParagraphs, targetParagraphs, alignment });
        const expectedDeleted = plan.filter(e => e.action === 'deleteMark').length;
        const expectedInserted = plan.filter(e => e.action === 'insertMark').length;

        const result = await applyOperationsToDocumentXml(body(round.items), [{
            type: 'redline', target: round.target, targetEnd: round.targetEnd, modified: '- One\n- Two'
        }], 'Round Author', {}, { atomic: true, strictTargets: true, structuredContent: true, generateRedlines: true });
        assert.equal(result.status, 'ok', `${round.name}: ${JSON.stringify(result.error || result.results)}`);
        const doc = parse(result.documentXml);
        const paragraphs = paragraphsOf(doc);
        assert.equal(paragraphs.filter(p => markOf(p, 'del').length).length, expectedDeleted, `${round.name}: one deleted paragraph per deleted source paragraph`);
        assert.equal(paragraphs.filter(p => markOf(p, 'ins').length).length, expectedInserted, `${round.name}: every new paragraph carries an inserted mark`);
        assert.equal(paragraphs.filter(p => !markOf(p, 'ins').length).length, sourceParagraphs.length,
            `${round.name}: untracked source-side marks equal the source paragraph count`);
        assert.equal(wordRevisionParagraphTexts(doc, 'rejected').length, sourceParagraphs.length, `${round.name}: Reject All keeps the source paragraph count`);
    }
}

console.log('PASS: paragraph mark revisions');
