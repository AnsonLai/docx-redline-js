import './setup-xml-provider.mjs';

import assert from 'node:assert/strict';
import { parseOoxmlSafe } from '../adapters/xml-adapter.js';
import {
    createParagraphTextIndex,
    resolveTextInParagraphIndex
} from '../services/comment-locator.js';
import { injectCommentsIntoOoxml } from '../services/comment-engine.js';
import { inspectDocumentParts } from '../services/document-inspection.js';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

function paragraphFromRuns(runs) {
    const xml = `<w:document xmlns:w="${W}"><w:body><w:p>${runs}</w:p></w:body></w:document>`;
    const parsed = parseOoxmlSafe(xml, 'application/xml');
    assert.equal(parsed.error, null);
    return parsed.doc.getElementsByTagNameNS(W, 'p')[0];
}

function indexFor(runs) {
    return createParagraphTextIndex(paragraphFromRuns(runs));
}

const multiRun = indexFor(
    '<w:r><w:t>during the\u00a0</w:t></w:r>'
    + '<w:r><w:rPr><w:u w:val="single"/></w:rPr><w:t>Subscription Term</w:t></w:r>'
    + '<w:r><w:t>\u00a0only</w:t></w:r>'
);
const equivalent = resolveTextInParagraphIndex(multiRun, 'during the Subscription Term only');
assert.equal(equivalent.found, true);
assert.equal(equivalent.resolvedBy, 'space_equivalent_anchor');
assert.deepEqual([equivalent.start, equivalent.end], [0, multiRun.fullText.length]);
assert.notEqual(equivalent.startRun, equivalent.endRun);

const reverseEquivalent = resolveTextInParagraphIndex(
    indexFor('<w:r><w:t>Alpha Beta</w:t></w:r>'),
    'Alpha\u00a0Beta'
);
assert.equal(reverseEquivalent.found, true);
assert.equal(reverseEquivalent.resolvedBy, 'space_equivalent_anchor');

const exact = resolveTextInParagraphIndex(
    indexFor('<w:r><w:t>Prefix exact suffix</w:t></w:r>'),
    'exact'
);
assert.equal(exact.found, true);
assert.equal(exact.resolvedBy, 'exact_anchor');
assert.deepEqual([exact.start, exact.end, exact.startOffset, exact.endOffset], [7, 12, 7, 12]);

const ambiguousExact = resolveTextInParagraphIndex(
    indexFor('<w:r><w:t>term and term</w:t></w:r>'),
    'term'
);
assert.equal(ambiguousExact.found, false);
assert.equal(ambiguousExact.error.code, 'AMBIGUOUS_ANCHOR');
assert.deepEqual(ambiguousExact.error.candidates, [{ start: 0, end: 4 }, { start: 9, end: 13 }]);

const ambiguousEquivalent = resolveTextInParagraphIndex(
    indexFor('<w:r><w:t>A\u00a0B and A B</w:t></w:r>'),
    'A B'
);
assert.equal(ambiguousEquivalent.found, true,
    'an exact match must take priority over additional space-equivalent candidates');
assert.equal(ambiguousEquivalent.resolvedBy, 'exact_anchor');
const onlyEquivalentAmbiguous = resolveTextInParagraphIndex(
    indexFor('<w:r><w:t>A\u00a0B and A\u00a0B</w:t></w:r>'),
    'A B'
);
assert.equal(onlyEquivalentAmbiguous.found, false);
assert.equal(onlyEquivalentAmbiguous.error.code, 'AMBIGUOUS_ANCHOR');

const punctuationMismatch = resolveTextInParagraphIndex(
    indexFor('<w:r><w:t>Alpha Beta.</w:t></w:r>'),
    'Alpha Beta,'
);
assert.equal(punctuationMismatch.found, false);
assert.equal(punctuationMismatch.error.code, 'ANCHOR_NOT_FOUND');
const tabMismatch = resolveTextInParagraphIndex(
    indexFor('<w:r><w:t>Alpha\tBeta</w:t></w:r>'),
    'Alpha Beta'
);
assert.equal(tabMismatch.found, false, 'tabs must not be normalized to spaces');

const source = `<w:document xmlns:w="${W}"><w:body><w:p>`
    + '<w:r><w:t>during the\u00a0</w:t></w:r>'
    + '<w:r><w:rPr><w:u w:val="single"/></w:rPr><w:t>Subscription Term</w:t></w:r>'
    + '<w:r><w:t>\u00a0only</w:t></w:r>'
    + '</w:p></w:body></w:document>';
let allocated = 50;
const injected = injectCommentsIntoOoxml(source, [{
    paragraphIndex: 1,
    textToFind: 'during the Subscription Term only',
    commentContent: 'Review.'
}], { author: 'Agent', commentIdAllocator: () => allocated++ });
assert.equal(injected.status, undefined);
assert.equal(injected.commentsApplied, 1);
assert.equal(injected.resolvedAnchors[0].resolvedBy, 'space_equivalent_anchor');
assert.equal(allocated, 51);
assert.match(injected.oxml, /w:commentRangeStart w:id="50"/);
assert.match(injected.oxml, /w:commentRangeEnd w:id="50"/);
assert.match(injected.oxml, /<w:u w:val="single"/);

const subspanSource = `<w:document xmlns:w="${W}"><w:body><w:p><w:r>`
    + '<w:t>Prefix Alpha\u00a0Beta suffix</w:t>'
    + '</w:r></w:p></w:body></w:document>';
const subspan = injectCommentsIntoOoxml(subspanSource, [{
    paragraphIndex: 1,
    textToFind: 'Alpha Beta',
    commentContent: 'Mapped subspan.'
}], { author: 'Agent' });
assert.equal(subspan.commentsApplied, 1);
const inspectedSubspan = inspectDocumentParts({
    documentXml: subspan.oxml,
    commentsXml: subspan.commentsXml
});
assert.equal(inspectedSubspan.paragraphs[0].exactText, 'Prefix Alpha\u00a0Beta suffix');
assert.equal(inspectedSubspan.comments[0].anchoredText, 'Alpha\u00a0Beta',
    'mapped offsets must enclose the original NBSP text');

let failedAllocations = 0;
const failed = injectCommentsIntoOoxml(source, [{
    paragraphIndex: 1,
    textToFind: 'genuinely absent',
    commentContent: 'Must fail.'
}], { commentIdAllocator: () => { failedAllocations += 1; return 99; } });
assert.equal(failed.status, 'error');
assert.equal(failed.error.code, 'ANCHOR_NOT_FOUND');
assert.equal(failedAllocations, 0, 'failed resolution must not consume a comment ID');
assert.equal(failed.oxml, source);

console.log('comment anchor locator tests passed');
