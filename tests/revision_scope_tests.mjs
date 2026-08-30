import './setup-xml-provider.mjs';

import assert from 'assert/strict';

import { applyRedlineToOxml } from '../index.js';
import { RevisionIdAllocator } from '../core/types.js';
import { applyOperationToDocumentXml } from '../services/standalone-operation-runner.js';
import { assertUniqueRevisionIds } from './helpers/ooxml-assertions.mjs';

const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const DATE = '2026-01-01T00:00:00Z';

function paragraph(content, extraNamespaces = '') {
    return `<w:p xmlns:w="${NS_W}" ${extraNamespaces}>${content}</w:p>`;
}

function documentXml(content) {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`
        + `<w:document xmlns:w="${NS_W}"><w:body>${content}<w:sectPr/></w:body></w:document>`;
}

function revisionIds(xml) {
    const doc = new DOMParser().parseFromString(xml, 'application/xml');
    const names = ['ins', 'del', 'moveFrom', 'moveTo', 'rPrChange', 'pPrChange', 'cellIns', 'cellDel'];
    return names.flatMap(name => Array.from(doc.getElementsByTagNameNS(NS_W, name)))
        .map(node => Number.parseInt(node.getAttribute('w:id') || node.getAttribute('id') || '', 10))
        .filter(Number.isFinite);
}

function revisionAuthors(xml) {
    const doc = new DOMParser().parseFromString(xml, 'application/xml');
    return ['ins', 'del', 'rPrChange', 'pPrChange']
        .flatMap(name => Array.from(doc.getElementsByTagNameNS(NS_W, name)))
        .map(node => node.getAttribute('w:author') || node.getAttribute('author'))
        .filter(Boolean);
}

async function testHostileRevisionIdIsClampedAndDoesNotPoisonNextDocument() {
    const hostile = paragraph(
        `<w:ins w:id="2147483000" w:author="Prior" w:date="${DATE}"><w:r><w:t>Old</w:t></w:r></w:ins>`
        + '<w:r><w:t xml:space="preserve"> clause</w:t></w:r>'
    );
    const hostileResult = await applyRedlineToOxml(hostile, 'Old clause', 'Updated clause', {
        author: 'ScopedA',
        existingRevisions: 'accept-all-first'
    });

    assert.equal(hostileResult.status, 'ok');
    const hostileIds = revisionIds(hostileResult.oxml);
    assert.ok(hostileIds.length >= 1);
    assert.ok(hostileIds.every(id => id >= 1000 && id < 10000), hostileIds.join(','));
    assertUniqueRevisionIds(hostileResult.oxml);

    const clean = paragraph('<w:r><w:t>Clean source.</w:t></w:r>');
    const cleanResult = await applyRedlineToOxml(clean, 'Clean source.', 'Clean result.', {
        author: 'ScopedB'
    });
    const cleanIds = revisionIds(cleanResult.oxml);
    assert.ok(cleanIds.length >= 1);
    assert.ok(cleanIds.every(id => id >= 1000 && id < 10000), cleanIds.join(','));
    assert.equal(Math.min(...cleanIds), 1000,
        'an unrelated clean document should receive a fresh low-range allocator');
    assertUniqueRevisionIds(cleanResult.oxml);
}

async function testBookmarkAndRelationshipIdsDoNotSeedRevisionIds() {
    const source = paragraph(
        '<w:bookmarkStart w:id="2147483000" w:name="largeBookmark"/>'
        + '<w:hyperlink r:id="2147482999"><w:r><w:t>Linked source.</w:t></w:r></w:hyperlink>'
        + '<w:bookmarkEnd w:id="2147483000"/>',
        `xmlns:r="${NS_R}"`
    );
    const result = await applyRedlineToOxml(source, 'Linked source.', 'Linked result.', {
        author: 'ScopedLink'
    });
    const ids = revisionIds(result.oxml);
    assert.ok(ids.length >= 1);
    assert.equal(Math.min(...ids), 1000);
    assert.ok(ids.every(id => id < 10000));
    assertUniqueRevisionIds(result.oxml);
}

function testAllocatorSeedsFromCommentsButNotOtherIdSpaces() {
    const source = `<w:comments xmlns:w="${NS_W}" xmlns:r="${NS_R}">`
        + '<w:comment w:id="1100" w:author="Reviewer"><w:p><w:r><w:t>Note</w:t></w:r></w:p></w:comment>'
        + '<w:bookmarkStart w:id="2000000000" w:name="ignored"/>'
        + '<w:hyperlink r:id="2000000001"/>'
        + '</w:comments>';
    const doc = new DOMParser().parseFromString(source, 'application/xml');
    const allocator = new RevisionIdAllocator();
    assert.equal(allocator.seed(doc), 1101);
    assert.equal(allocator.next(), 1101);
}

async function testStandaloneInvocationsAreDocumentScoped() {
    const hostile = documentXml(paragraph(
        `<w:ins w:id="2147483000" w:author="Prior" w:date="${DATE}"><w:r><w:t>Prior</w:t></w:r></w:ins>`
        + '<w:r><w:t xml:space="preserve"> text.</w:t></w:r>'
    ));
    const hostileResult = await applyOperationToDocumentXml(
        hostile,
        { type: 'replace', target: 'Prior text.', modified: 'Updated text.' },
        'StandaloneScopedA',
        null,
        { existingRevisions: 'accept-all-first' }
    );
    assert.ok(revisionIds(hostileResult.documentXml).every(id => id < 10000));
    assertUniqueRevisionIds(hostileResult.documentXml);

    const clean = documentXml('<w:p><w:r><w:t>Standalone clean.</w:t></w:r></w:p>');
    const cleanResult = await applyOperationToDocumentXml(
        clean,
        { type: 'replace', target: 'Standalone clean.', modified: 'Standalone result.' },
        'StandaloneScopedB'
    );
    const cleanIds = revisionIds(cleanResult.documentXml);
    assert.equal(Math.min(...cleanIds), 1000);
    assertUniqueRevisionIds(cleanResult.documentXml);
}

async function testInterleavedAuthorsStayPerCall() {
    const sourceA = paragraph('<w:r><w:t>Alpha item</w:t></w:r>');
    const sourceB = paragraph('<w:r><w:t>Beta item</w:t></w:r>');

    const [resultA, resultB] = await Promise.all([
        applyRedlineToOxml(sourceA, 'Alpha item', '- Alpha item\n- Alpha follow-up', {
            author: 'Concurrent Alpha'
        }),
        applyRedlineToOxml(sourceB, 'Beta item', '- Beta item\n- Beta follow-up', {
            author: 'Concurrent Beta'
        })
    ]);

    const authorsA = revisionAuthors(resultA.oxml);
    const authorsB = revisionAuthors(resultB.oxml);
    assert.ok(authorsA.length >= 1);
    assert.ok(authorsB.length >= 1);
    assert.ok(authorsA.every(author => author === 'Concurrent Alpha'), authorsA.join(','));
    assert.ok(authorsB.every(author => author === 'Concurrent Beta'), authorsB.join(','));
    assertUniqueRevisionIds(resultA.oxml);
    assertUniqueRevisionIds(resultB.oxml);
}

await testHostileRevisionIdIsClampedAndDoesNotPoisonNextDocument();
await testBookmarkAndRelationshipIdsDoNotSeedRevisionIds();
testAllocatorSeedsFromCommentsButNotOtherIdSpaces();
await testStandaloneInvocationsAreDocumentScoped();
await testInterleavedAuthorsStayPerCall();

console.log('PASS: document-scoped revision IDs and per-call authors');
