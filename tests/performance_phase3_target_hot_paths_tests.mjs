import './setup-xml-provider.mjs';

import assert from 'node:assert/strict';
import { DOMParser } from '@xmldom/xmldom';
import { RevisionIdAllocator } from '../core/types.js';
import {
    buildParagraphMetadataIndex,
    buildTargetReferenceSnapshot,
    createParagraphFingerprint,
    resolveTargetParagraph
} from '../core/paragraph-targeting.js';
import { DocumentOperationSession } from '../services/document-operation-session.js';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const paragraphs = Array.from({ length: 1000 }, (_, index) => {
    const text = index === 499 ? 'Duplicate target' : index === 799 ? 'Duplicate target' : `Paragraph ${index + 1}`;
    const tableWrapped = index === 799;
    const paragraph = `<w:p w14:paraId="${String(index + 1).padStart(8, '0')}"><w:r><w:t>${text}</w:t></w:r></w:p>`;
    return tableWrapped ? `<w:tbl><w:tr><w:tc>${paragraph}</w:tc></w:tr></w:tbl>` : paragraph;
}).join('');
const xml = `<w:document xmlns:w="${W}" xmlns:w14="${W14}"><w:body>${paragraphs}<w:p><w:ins w:id="9001" w:author="Prior" w:date="2026-01-01T00:00:00Z"><w:r><w:t>Revised</w:t></w:r></w:ins></w:p><w:sectPr/></w:body></w:document>`;

const doc = new DOMParser().parseFromString(xml, 'application/xml');
const index = buildParagraphMetadataIndex(doc);
assert.equal(index.entries.length, 1001);
assert.equal(index.byNormalizedText.get('Duplicate target').length, 2);
assert.equal(index.byId.get('00000800').index, 800);
assert.equal(index.entries[799].inTable, true);
assert.equal(index.entries[799].fingerprint, createParagraphFingerprint(index.entries[799].paragraph));

const snapshot = buildTargetReferenceSnapshot(doc, index);
assert.equal(snapshot.get(800).text, 'Duplicate target');
assert.equal(snapshot.get(800).inTable, true);

const descriptor = { text: 'Duplicate target', index: 800, inTable: true };
const cached = resolveTargetParagraph(doc, { targetDescriptor: descriptor, strictAmbiguity: true, paragraphMetadataIndex: index });
const uncached = resolveTargetParagraph(doc, { targetDescriptor: descriptor, strictAmbiguity: true });
assert.equal(cached.resolvedBy, uncached.resolvedBy);
assert.equal(cached.paragraph, uncached.paragraph);

const allocator = new RevisionIdAllocator();
const originalGetElements = doc.getElementsByTagName;
doc.getElementsByTagName = () => { throw new Error('seed must not allocate a universal element array'); };
assert.equal(allocator.seed(doc), 9002);
doc.getElementsByTagName = originalGetElements;

const session = new DocumentOperationSession(xml);
const firstIndex = session.getParagraphIndex();
assert.equal(session.getParagraphIndex(), firstIndex, 'session should reuse metadata until invalidated');
session.invalidateParagraphIndex();
assert.notEqual(session.getParagraphIndex(), firstIndex, 'invalidated session must not retain DOM-derived metadata');

console.log('PASS: Phase 3 target metadata and revision traversal hot paths');
