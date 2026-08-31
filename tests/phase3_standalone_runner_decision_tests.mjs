import './setup-xml-provider.mjs';

import assert from 'assert/strict';

import { applyOperationToDocumentXml, applyOperationsToDocumentXml } from '../services/standalone-operation-runner.js';

const NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const documentXml = `<?xml version="1.0"?><w:document xmlns:w="${NS}"><w:body>
  <w:p><w:r><w:t>Review the filing deadline.</w:t></w:r></w:p>
  <w:p><w:r><w:t>Second operation target.</w:t></w:r></w:p>
  <w:sectPr/>
</w:body></w:document>`;

// Highlight is a supported full-document operation and must localize the
// requested text without changing the untargeted paragraph.
const highlighted = await applyOperationToDocumentXml(documentXml, {
    type: 'highlight',
    target: 'Review the filing deadline.',
    textToHighlight: 'filing deadline',
    color: 'yellow'
}, 'Phase 3', null, { generateRedlines: true });
assert.equal(highlighted.hasChanges, true);
assert.match(highlighted.documentXml, /<w:highlight w:val="yellow"/);
assert.match(highlighted.documentXml, /Second operation target\./);
assert.match(highlighted.documentXml, /<w:rPrChange/);

const missingHighlight = await applyOperationToDocumentXml(documentXml, {
    type: 'highlight',
    target: 'Review the filing deadline.',
    textToHighlight: 'absent words',
    color: 'green'
}, 'Phase 3');
assert.equal(missingHighlight.hasChanges, false);
assert.equal(missingHighlight.documentXml, documentXml);

// continueOnError:false stops at the failed operation. Atomic mode rolls back
// the earlier comment artifact and document mutation; non-atomic mode retains
// the earlier operation but still does not attempt the final replacement.
const operations = [
    {
        type: 'comment',
        target: 'Review the filing deadline.',
        textToComment: 'filing deadline',
        commentContent: 'Confirm this date.'
    },
    { type: 'replace', target: 'Missing target.', modified: 'Must fail.' },
    { type: 'replace', target: 'Second operation target.', modified: 'Should not run.' }
];
const atomic = await applyOperationsToDocumentXml(documentXml, operations, 'Phase 3', {}, {
    generateRedlines: false,
    continueOnError: false
});
assert.equal(atomic.rolledBack, true);
assert.equal(atomic.documentXml, documentXml);
assert.equal(atomic.commentsXml, null);
assert.deepEqual(atomic.executionOrder, [1, 2]);
assert.deepEqual(atomic.results.map(result => result.status), ['applied', 'error']);

const partialContext = {};
const partial = await applyOperationsToDocumentXml(documentXml, operations, 'Phase 3', partialContext, {
    generateRedlines: false,
    continueOnError: false,
    atomic: false
});
assert.equal(partial.hasChanges, true);
assert.ok(partial.commentsXml?.includes('Confirm this date.'));
assert.deepEqual(partial.executionOrder, [1, 2]);
assert.doesNotMatch(partial.documentXml, /Should not run/);
assert.ok(partialContext.targetRefSnapshot instanceof Map);

// Parse failure remains a stable structured error and produces no artifacts.
const malformed = await applyOperationsToDocumentXml('<broken', operations, 'Phase 3');
assert.equal(malformed.status, 'error');
assert.equal(malformed.hasChanges, false);
assert.deepEqual(malformed.results, []);
assert.deepEqual(malformed.numberingXmlParts, []);

console.log('PASS: Phase 3 standalone highlight, stop-on-error, and artifact rollback decisions');
