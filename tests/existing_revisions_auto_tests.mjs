import './setup-xml-provider.mjs';

// existingRevisions: 'auto' merges a target's own-author history and slices
// around another or unattributed reviewer's history, per target.
import assert from 'assert/strict';
import { acceptTrackedChangesInOoxml, rejectTrackedChangesInOoxml, inspectDocumentParts } from '../index.js';
import { applyOperationsToDocumentXml, preflightOperations } from '../services/standalone-operation-runner.js';
import { resolveExistingRevisionsPolicy } from '../core/word-xml.js';
import { parseXml } from '../adapters/xml-adapter.js';
import { executeCli } from '../node/cli.js';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const ins = (author, text) => `<w:ins w:id="${Math.floor(Math.random() * 1e6)}" w:author="${author}" w:date="2026-01-01T00:00:00Z"><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:ins>`;
const run = text => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const documentXml = `<w:document ${W}><w:body>`
    + `<w:p>${run('Foreign start ')}${ins('Reviewer', 'added ')}${run('end.')}</w:p>`
    + `<w:p>${run('Own start ')}${ins('Me', 'added ')}${run('end.')}</w:p>`
    + `<w:p>${run('Plain paragraph.')}</w:p>`
    + `<w:sectPr/></w:body></w:document>`;
const source = inspectDocumentParts({ documentXml }).paragraphs;
const strong = p => ({ index: p.index, exactText: p.exactText, paragraphId: p.paragraphId, fingerprint: p.fingerprint });
const texts = (xml, resolve) => inspectDocumentParts({ documentXml: resolve(xml, { allAuthors: true }).oxml }).paragraphs.map(p => p.exactText);

// Resolver.
{
    const doc = parseXml(documentXml);
    const [foreign, own, plain] = Array.from(doc.getElementsByTagNameNS('*', 'p'));
    assert.equal(resolveExistingRevisionsPolicy('auto', foreign, 'Me'), 'slice-cross-author');
    assert.equal(resolveExistingRevisionsPolicy('auto', own, ' me '), 'merge-same-author');
    assert.equal(resolveExistingRevisionsPolicy('auto', plain, 'Me'), 'merge-same-author');
    assert.equal(resolveExistingRevisionsPolicy('reject-input', foreign, 'Me'), 'reject-input');
}

// merge-same-author refuses the foreign target; auto slices around it.
{
    const operation = { type: 'redline', target: strong(source[0]), structuredContent: false, modified: 'Foreign start added end. More.' };
    const refused = await applyOperationsToDocumentXml(documentXml, [{ ...operation, existingRevisions: 'merge-same-author' }], 'Me', null, { atomic: true, strictTargets: true });
    assert.notEqual(refused.status, 'ok', 'merge-same-author refuses foreign history');

    const preflight = await preflightOperations(documentXml, [{ ...operation, existingRevisions: 'auto' }], 'Me');
    assert.equal(preflight.results[0].existingRevisions, 'slice-cross-author', 'preflight reports the resolved policy');
    assert.equal(preflight.results[0].error, undefined);

    const result = await applyOperationsToDocumentXml(documentXml, [{ ...operation, existingRevisions: 'auto' }], 'Me', null, { atomic: true, strictTargets: true });
    assert.equal(result.status, 'ok', JSON.stringify(result.error ?? result.results));
    assert.equal(texts(result.documentXml, acceptTrackedChangesInOoxml)[0], 'Foreign start added end. More.');
    assert.equal(texts(result.documentXml, rejectTrackedChangesInOoxml)[0], 'Foreign start end.', 'Reject All restores the reviewer baseline');
    assert.match(result.documentXml, /w:author="Reviewer"/, 'reviewer history is preserved');
}

// auto merges own history, and a batch-level auto applies per target.
{
    const result = await applyOperationsToDocumentXml(documentXml, [
        { type: 'redline', target: strong(source[0]), structuredContent: false, modified: 'Foreign start added end!' },
        { type: 'redline', target: strong(source[1]), structuredContent: false, modified: 'Own start added end!' },
        { type: 'redline', target: strong(source[2]), structuredContent: false, modified: 'Plain paragraph!' },
    ], 'Me', null, { atomic: true, strictTargets: true, existingRevisions: 'auto' });
    assert.equal(result.status, 'ok', JSON.stringify(result.error ?? result.results));
    assert.deepEqual(texts(result.documentXml, acceptTrackedChangesInOoxml), ['Foreign start added end!', 'Own start added end!', 'Plain paragraph!']);
    assert.deepEqual(texts(result.documentXml, rejectTrackedChangesInOoxml), ['Foreign start end.', 'Own start end.', 'Plain paragraph.']);
}

// The CLI advertises 'auto' as an additive capability without a contract bump.
{
    const version = await executeCli(['version']);
    assert.equal(version.contractVersion, 9);
    assert.ok(version.capabilities.includes('existing-revisions-auto-v1'));
}

console.log('existing_revisions_auto_tests: all passed');
