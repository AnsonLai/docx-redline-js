import assert from 'node:assert/strict';
import './setup-xml-provider.mjs';

import {
    applyRedlineToOxml,
    ingestWordOoxmlToPlainText,
    sanitizeAiResponse
} from '../index.js';
import { acceptTrackedChangesInOoxml } from '../services/revision-comment-management.js';

const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

function paragraph(text) {
    return `<w:p xmlns:w="${NS_W}"><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
}

const priorRevision = `<w:p xmlns:w="${NS_W}">`
    + '<w:r><w:t xml:space="preserve">A </w:t></w:r>'
    + '<w:del w:id="1" w:author="Prior" w:date="2026-01-01T00:00:00Z"><w:r><w:delText>old</w:delText></w:r></w:del>'
    + '<w:ins w:id="2" w:author="Prior" w:date="2026-01-01T00:00:00Z"><w:r><w:t>new</w:t></w:r></w:ins>'
    + '<w:r><w:t xml:space="preserve"> end</w:t></w:r>'
    + '</w:p>';

const preservedNoOp = await applyRedlineToOxml(priorRevision, 'A new end', 'A new end', {
    existingRevisions: 'accept-all-first'
});
assert.equal(preservedNoOp.hasChanges, false);
assert.equal(preservedNoOp.status, 'no-op');
assert.equal(preservedNoOp.oxml, priorRevision);
assert(preservedNoOp.oxml.includes('w:author="Prior"'));

const keptNormalized = await applyRedlineToOxml(priorRevision, 'A new end', 'A new end', {
    existingRevisions: 'accept-all-first-keep-normalized'
});
assert.equal(keptNormalized.hasChanges, true);
assert.equal(keptNormalized.status, 'ok');
assert(!keptNormalized.oxml.includes('w:author="Prior"'));
assert(keptNormalized.warnings?.some(warning => warning.includes('Existing revisions were accepted')));
assert.equal(keptNormalized.oxml, acceptTrackedChangesInOoxml(priorRevision, { allAuthors: true }).oxml);

const corruptionSamples = [
    'The rate is $X$ per unit as defined in Schedule A.',
    'Costs range from $ten thousand$ upward.',
    String.raw`Escape sequences such as \n must be preserved literally.`,
    String.raw`Escape sequences such as \r\n must be preserved literally.`,
    'Here is the text: this clause is part of the actual contract body.',
    'Here is the redline: this clause is also part of the actual contract body.'
];

for (const modified of corruptionSamples) {
    const result = await applyRedlineToOxml(paragraph('Original clause.'), 'Original clause.', modified);
    assert.equal(result.hasChanges, true, modified);
    const accepted = acceptTrackedChangesInOoxml(result.oxml, { allAuthors: true });
    assert.equal(ingestWordOoxmlToPlainText(accepted.oxml), modified, modified);
    assert(!result.warnings?.some(warning => warning.includes('Input was sanitized')));
}

const inlinePreface = 'Here is the text: this sentence must stay intact.';
assert.equal(sanitizeAiResponse(inlinePreface), inlinePreface);
assert.equal(sanitizeAiResponse(String.raw`The rate is $X$ and literal \n stays.`), String.raw`The rate is $X$ and literal \n stays.`);

const rawAssistantText = 'Here is the text:\nSanitized clause.';
const sanitized = await applyRedlineToOxml(paragraph('Original clause.'), 'Original clause.', rawAssistantText, {
    sanitizeInput: true
});
assert(sanitized.warnings?.includes('Input was sanitized; pass sanitizeInput: false to disable.'));
const acceptedSanitized = acceptTrackedChangesInOoxml(sanitized.oxml, { allAuthors: true });
assert.equal(ingestWordOoxmlToPlainText(acceptedSanitized.oxml), 'Sanitized clause.');

console.log('PASS: Phase 4 input mutation and existing-revision no-op policies');
