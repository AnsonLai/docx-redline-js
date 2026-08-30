import assert from 'node:assert/strict';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import * as api from '../index.js';

api.configureXmlProvider({ DOMParser, XMLSerializer });
const loggedWarnings = [];
const loggedErrors = [];
api.configureLogger({
    log() {},
    warn(...args) { loggedWarnings.push(args.map(String).join(' ')); },
    error(...args) { loggedErrors.push(args.map(String).join(' ')); }
});

const malformed = '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:r><w:t>hello</w:t></w:p>';
const inputs = [malformed, '', null, undefined, '<html><body/></html>'];
const commentsXml = '<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>';

const invoke = {
    applyFormattingRemovalToOoxml: value => api.applyFormattingRemovalToOoxml(value, 'hello', ['bold']),
    applyHighlightToOoxml: value => api.applyHighlightToOoxml(value, 'hello', 'yellow'),
    applyRedlineToOxml: value => api.applyRedlineToOxml(value, 'hello', 'goodbye'),
    applyRedlineToOxmlWithListFallback: value => api.applyRedlineToOxmlWithListFallback(value, 'hello', 'goodbye'),
    reconcileMarkdownTableOoxml: value => api.reconcileMarkdownTableOoxml(value, 'hello', '| A |\n|---|\n| B |'),
    ingestOoxml: value => api.ingestOoxml(value),
    ingestWordOoxmlToPlainText: value => api.ingestWordOoxmlToPlainText(value),
    ingestWordOoxmlToMarkdown: value => api.ingestWordOoxmlToMarkdown(value),
    ingestWordOoxmlToPlainTextResult: value => api.ingestWordOoxmlToPlainTextResult(value),
    ingestWordOoxmlToMarkdownResult: value => api.ingestWordOoxmlToMarkdownResult(value),
    injectCommentsIntoOoxml: value => api.injectCommentsIntoOoxml(value, [{ paragraphIndex: 1, textToFind: 'hello', commentContent: 'note' }]),
    injectCommentsIntoPackage: value => api.injectCommentsIntoPackage(value, commentsXml),
    acceptTrackedChangesInOoxml: value => api.acceptTrackedChangesInOoxml(value, { allAuthors: true }),
    rejectTrackedChangesInOoxml: value => api.rejectTrackedChangesInOoxml(value, { allAuthors: true }),
    deleteCommentsByAuthorInOoxml: value => api.deleteCommentsByAuthorInOoxml(value, { allAuthors: true }),
    extractParagraphIdFromOoxml: value => api.extractParagraphIdFromOoxml(value),
    extractReplacementNodesFromOoxml: value => api.extractReplacementNodesFromOoxml(value),
    parseOoxml: value => api.parseOoxml(value),
    parseOoxmlSafe: value => api.parseOoxmlSafe(value),
    validateRedlineOoxml: value => api.validateRedlineOoxml(value)
};

// Discover OOXML-consuming main-entry exports so a newly added API cannot
// silently evade this matrix. Output-only builders are explicitly excluded.
const outputOnlyOoxmlExports = new Set(['generateTableOoxml', 'serializeOoxml', 'serializeToOoxml']);
const discovered = Object.keys(api)
    .filter(name => /Ooxml|Oxml/.test(name) && typeof api[name] === 'function')
    .filter(name => !outputOnlyOoxmlExports.has(name));
const explicitlyNamedWithoutOoxml = new Set(['injectCommentsIntoPackage']);
const coveredDiscovered = Object.keys(invoke).filter(name => !explicitlyNamedWithoutOoxml.has(name));
assert.deepEqual(discovered.sort(), coveredDiscovered.sort(), 'OOXML export error-contract matrix is stale');

for (const [name, call] of Object.entries(invoke)) {
    for (const value of inputs) {
        await assert.doesNotReject(async () => call(value), `${name} threw for ${String(value)}`);
    }
}

for (const name of ['acceptTrackedChangesInOoxml', 'rejectTrackedChangesInOoxml', 'deleteCommentsByAuthorInOoxml']) {
    const result = await invoke[name](malformed);
    assert.equal(result.status, 'error', `${name} status`);
    assert.equal(result.error?.code, 'PARSE_ERROR', `${name} error code`);
    assert.equal(result.hasChanges, false, `${name} hasChanges`);
    assert.equal(result.oxml, malformed, `${name} preserves caller OOXML`);
}

for (const name of ['ingestWordOoxmlToPlainTextResult', 'ingestWordOoxmlToMarkdownResult']) {
    const result = await invoke[name](malformed);
    assert.deepEqual(
        { text: result.text, status: result.status, code: result.error?.code },
        { text: '', status: 'error', code: 'PARSE_ERROR' },
        `${name} distinguishes malformed input from empty content`
    );
}

loggedWarnings.length = 0;
loggedErrors.length = 0;
const recovered = api.parseOoxmlSafe('<a>&nosuch;</a>');
assert.equal(recovered.error, null);
assert(recovered.warnings.some(message => message.includes('entity not found')));
assert(loggedWarnings.some(message => message.includes('entity not found')));
api.parseOoxmlSafe(malformed);
assert(loggedErrors.some(message => message.includes('XML parse error')));

const documentXml = '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>First clause</w:t></w:r></w:p><w:p><w:r><w:t>Second clause</w:t></w:r></w:p></w:body></w:document>';
const matching = await api.applyRedlineToOxml(documentXml, ' First   clause \nSecond clause ', 'First clause\nRevised clause');
assert.notEqual(matching.error?.code, 'TARGET_NOT_FOUND', 'normalized multi-line target should match');

const missing = await api.applyRedlineToOxml(documentXml, 'First clause\nMissing clause', 'First clause\nRevised clause');
assert.equal(missing.status, 'error');
assert.equal(missing.error?.code, 'TARGET_NOT_FOUND');
assert.equal(missing.hasChanges, false);
assert.equal(missing.oxml, documentXml);

console.log(`PASS: unified error contract covers ${discovered.length} OOXML exports and ${inputs.length} malformed-input classes`);
