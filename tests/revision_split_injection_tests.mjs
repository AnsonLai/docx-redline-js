import './setup-xml-provider.mjs';

import assert from 'assert/strict';
import { validateRedlineOoxml } from '../core/redline-validation.js';
import { applyHighlightToOoxml } from '../engine/formatting-removal.js';
import { injectCommentsIntoOoxml, resetRevisionIdCounter } from '../services/comment-engine.js';
import { elementsByLocalName, parseXml } from './helpers/ooxml-assertions.mjs';

const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const trackedProperties = '<w:rPr><w:b/><w:rPrChange w:id="68" w:author="Reviewer" w:date="2026-01-01T00:00:00Z"><w:rPr/></w:rPrChange></w:rPr>';

function paragraph(runXml) {
    return `<w:document xmlns:w="${NS_W}"><w:body><w:p>${runXml}</w:p></w:body></w:document>`;
}

function assertValidUniqueRevisionIds(xml, expectedRPrChangeCount) {
    const validation = validateRedlineOoxml(xml);
    assert.equal(validation.valid, true, JSON.stringify(validation.issues));

    const doc = parseXml(xml);
    const changes = elementsByLocalName(doc, 'rPrChange');
    const ids = changes.map(change => change.getAttribute('w:id') || change.getAttribute('id'));
    assert.equal(changes.length, expectedRPrChangeCount);
    assert.equal(new Set(ids).size, ids.length, 'split fragments must have unique revision ids');
}

function injectComment(oxml) {
    resetRevisionIdCounter(1000);
    return injectCommentsIntoOoxml(oxml, [{
        paragraphIndex: 1,
        textToFind: 'bravo charlie',
        commentContent: 'Review this phrase'
    }], { author: 'Commenter' });
}

const caseA = injectComment(paragraph(
    `<w:r>${trackedProperties}<w:t>Alpha bravo charlie delta</w:t></w:r>`
));
assert.equal(caseA.hasChanges, true);
assertValidUniqueRevisionIds(caseA.oxml, 3);

const caseB = injectComment(paragraph(
    `<w:ins w:id="67" w:author="Reviewer" w:date="2026-01-01T00:00:00Z"><w:r>${trackedProperties}<w:t>Alpha bravo charlie delta</w:t></w:r></w:ins>`
));
assert.equal(caseB.hasChanges, true);
assertValidUniqueRevisionIds(caseB.oxml, 3);

const caseC = injectComment(paragraph(
    '<w:ins w:id="67" w:author="Reviewer" w:date="2026-01-01T00:00:00Z"><w:r><w:rPr><w:b/></w:rPr><w:t>Alpha bravo charlie delta</w:t></w:r></w:ins>'
));
assert.equal(caseC.hasChanges, true);
assertValidUniqueRevisionIds(caseC.oxml, 0);

const highlighted = applyHighlightToOoxml(
    paragraph(`<w:r>${trackedProperties}<w:t>Alpha bravo charlie delta</w:t></w:r>`),
    'bravo charlie',
    'yellow'
);
assertValidUniqueRevisionIds(highlighted, 3);

const highlightedInsideInsertion = applyHighlightToOoxml(
    paragraph(`<w:ins w:id="67" w:author="Reviewer" w:date="2026-01-01T00:00:00Z"><w:r>${trackedProperties}<w:t>Alpha bravo charlie delta</w:t></w:r></w:ins>`),
    'bravo charlie',
    'yellow'
);
assertValidUniqueRevisionIds(highlightedInsideInsertion, 3);

const trackedHighlight = applyHighlightToOoxml(
    paragraph(`<w:r>${trackedProperties}<w:t>Alpha bravo charlie delta</w:t></w:r>`),
    'bravo charlie',
    'yellow',
    { generateRedlines: true, author: 'Highlighter' }
);
assertValidUniqueRevisionIds(trackedHighlight, 3);

console.log('revision_split_injection_tests.mjs ... PASS');
