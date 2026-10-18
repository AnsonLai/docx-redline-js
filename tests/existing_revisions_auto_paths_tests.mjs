import './setup-xml-provider.mjs';

// existingRevisions: 'auto' through every entry point (resolver edge cases, the
// engine, the bare runner, preflight, openDocx and the CLI), and equivalence
// with the explicit policy it resolves to.
import assert from 'assert/strict';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { Readable } from 'node:stream';
import {
    acceptTrackedChangesInOoxml, applyRedlineToOxml, inspectDocumentParts, openDocx,
    rejectTrackedChangesInOoxml, resolveExistingRevisionsPolicy
} from '../index.js';
import { applyOperationsToDocumentXml, preflightOperations } from '../services/standalone-operation-runner.js';
import { EXISTING_REVISIONS_POLICIES, isExistingRevisionsPolicy } from '../services/document-operation-contract.js';
import { parseXml } from '../adapters/xml-adapter.js';
import { unzipDocx, zipDocx } from '../document/zip-archive.js';
import { executeCli } from '../node/cli.js';

const NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const W = `xmlns:w="${NS}"`;
const DATE = '2026-01-01T00:00:00Z';
let nextId = 100;
const ins = (author, text) => `<w:ins w:id="${nextId++}"${author == null ? '' : ` w:author="${author}"`} w:date="${DATE}"><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:ins>`;
const del = (author, text) => `<w:del w:id="${nextId++}" w:author="${author}" w:date="${DATE}"><w:r><w:delText xml:space="preserve">${text}</w:delText></w:r></w:del>`;
const run = text => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const para = content => `<w:p>${content}</w:p>`;
const docXml = body => `<w:document ${W}><w:body>${body}<w:sectPr/></w:body></w:document>`;
const strong = p => ({ index: p.index, exactText: p.exactText, paragraphId: p.paragraphId, fingerprint: p.fingerprint });
const view = (xml, resolve) => inspectDocumentParts({ documentXml: resolve(xml, { allAuthors: true }).oxml }).paragraphs.map(p => p.exactText);
const authorsOf = xml => [...new Set([...xml.matchAll(/w:author="([^"]*)"/g)].map(m => m[1]))].sort();
const firstParagraph = xml => parseXml(xml).getElementsByTagNameNS(NS, 'p')[0];

// ---------------------------------------------------------------------------------------------------------
// 1. Resolver edge cases
// ---------------------------------------------------------------------------------------------------------
{
    const resolve = (content, author = 'Me') => resolveExistingRevisionsPolicy('auto', firstParagraph(docXml(para(content))), author);
    assert.equal(resolve(run('plain')), 'merge-same-author', 'no tracked changes');
    assert.equal(resolve(ins('Me', 'x')), 'merge-same-author', 'only own insertions');
    assert.equal(resolve(ins('ME', 'x') + del('me', 'y')), 'merge-same-author', 'author match ignores case');
    assert.equal(resolve(ins('Me', 'x') + ins('Reviewer', 'y')), 'slice-cross-author', 'own plus foreign');
    assert.equal(resolve(del('Reviewer', 'gone')), 'slice-cross-author', 'foreign deletion only');
    assert.equal(resolve(ins(null, 'x')), 'slice-cross-author', 'unattributed changes count as another author');
    assert.equal(resolve(ins('Me', 'x') + ins(null, 'y')), 'slice-cross-author', 'own plus unattributed is not all own');
    const paragraphs = Array.from(parseXml(docXml(para(run('plain')) + para(ins('Me', 'x')) + para(ins('Reviewer', 'y')))).getElementsByTagNameNS(NS, 'p'));
    assert.equal(resolveExistingRevisionsPolicy('auto', paragraphs.slice(0, 2), 'Me'), 'merge-same-author', 'a range of own and plain paragraphs');
    assert.equal(resolveExistingRevisionsPolicy('auto', paragraphs, 'Me'), 'slice-cross-author', 'any foreign paragraph in a range');
    assert.equal(resolveExistingRevisionsPolicy('auto', [], 'Me'), 'merge-same-author', 'an empty range');
    assert.equal(resolve(ins('Me', 'x'), ''), 'slice-cross-author', 'an empty operation author owns nothing');
    assert.equal(resolveExistingRevisionsPolicy('auto', null, 'Me'), 'merge-same-author', 'no node');
    for (const policy of EXISTING_REVISIONS_POLICIES.filter(item => item !== 'auto')) {
        assert.equal(resolveExistingRevisionsPolicy(policy, firstParagraph(docXml(para(ins('Reviewer', 'x')))), 'Me'), policy, `${policy} passes through`);
    }
    assert.equal(resolveExistingRevisionsPolicy(undefined, null, 'Me'), undefined);
    assert.equal(resolveExistingRevisionsPolicy(null, null, 'Me'), null);
}

// ---------------------------------------------------------------------------------------------------------
// 2. Contract and schema
// ---------------------------------------------------------------------------------------------------------
{
    assert.equal(isExistingRevisionsPolicy('auto'), true);
    assert.equal(isExistingRevisionsPolicy('Auto'), false, 'policies are case-sensitive');
    assert.equal(isExistingRevisionsPolicy('automatic'), false);
    const schema = JSON.parse(await readFile(new URL('../docs/schemas/document-operations.schema.json', import.meta.url), 'utf8'));
    const schemaText = JSON.stringify(schema);
    assert.ok(schemaText.includes('"existingRevisions":{"enum":["auto",'), 'schema lists auto');

    const source = docXml(para(run('Plain.')));
    const target = strong(inspectDocumentParts({ documentXml: source }).paragraphs[0]);
    const bad = await applyOperationsToDocumentXml(source, [{ type: 'redline', target, modified: 'Plain!', existingRevisions: 'automatic' }], 'Me', null, { strictTargets: true });
    assert.notEqual(bad.status, 'ok', 'an unknown policy is still refused');
    const badBatch = await applyOperationsToDocumentXml(source, [{ type: 'redline', target, modified: 'Plain!' }], 'Me', null, { strictTargets: true, existingRevisions: 'AUTO' });
    assert.notEqual(badBatch.status, 'ok', 'an unknown batch policy is still refused');
}

// ---------------------------------------------------------------------------------------------------------
// 3. Engine: applyRedlineToOxml with 'auto' equals the resolved explicit policy
// ---------------------------------------------------------------------------------------------------------
{
    const paragraph = content => `<w:p ${W}>${content}</w:p>`;
    const cases = [
        { name: 'foreign', oxml: paragraph(run('Base ') + ins('Reviewer', 'added ') + run('end.')), original: 'Base added end.', modified: 'Base added end. More.', expected: 'slice-cross-author' },
        { name: 'own', oxml: paragraph(run('Base ') + ins('Me', 'added ') + run('end.')), original: 'Base added end.', modified: 'Base added end. More.', expected: 'merge-same-author' },
        { name: 'plain', oxml: paragraph(run('Base end.')), original: 'Base end.', modified: 'Base end. More.', expected: 'merge-same-author' }
    ];
    for (const item of cases) {
        const auto = await applyRedlineToOxml(item.oxml, item.original, item.modified, { author: 'Me', existingRevisions: 'auto' });
        const explicit = await applyRedlineToOxml(item.oxml, item.original, item.modified, { author: 'Me', existingRevisions: item.expected });
        assert.equal(auto.status, 'ok', `${item.name}: ${JSON.stringify(auto.error)}`);
        assert.equal(auto.status, explicit.status, item.name);
        assert.deepEqual(view(auto.oxml, acceptTrackedChangesInOoxml), view(explicit.oxml, acceptTrackedChangesInOoxml), `${item.name}: Accept All matches ${item.expected}`);
        assert.deepEqual(view(auto.oxml, rejectTrackedChangesInOoxml), view(explicit.oxml, rejectTrackedChangesInOoxml), `${item.name}: Reject All matches ${item.expected}`);
        assert.deepEqual(authorsOf(auto.oxml), authorsOf(explicit.oxml), `${item.name}: same authors as ${item.expected}`);
    }
}

// ---------------------------------------------------------------------------------------------------------
// 4. Runner: per-target resolution in one batch, ranges, tables, untracked edits, preflight
// ---------------------------------------------------------------------------------------------------------
const mixedXml = docXml(
    para(run('Foreign start ') + ins('Reviewer', 'added ') + run('end.'))
    + para(run('Foreign deletion ') + del('Reviewer', 'removed ') + run('kept.'))
    + para(run('Own start ') + ins('Me', 'added ') + run('end.'))
    + para(run('Plain paragraph.'))
    + `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol/></w:tblGrid><w:tr><w:tc>${para(run('Cell ') + ins('Reviewer', 'note ') + run('text.'))}</w:tc></w:tr></w:tbl>`
    + para(run('Range one.'))
    + para(run('Range ') + ins('Reviewer', 'two '))
);
const mixed = inspectDocumentParts({ documentXml: mixedXml }).paragraphs;
const byText = prefix => strong(mixed.find(p => p.exactText.startsWith(prefix)));
{
    const operations = [
        { type: 'redline', target: byText('Foreign start'), modified: 'Foreign start added end!' },
        { type: 'redline', target: byText('Foreign deletion'), modified: 'Foreign deletion kept!' },
        { type: 'redline', target: byText('Own start'), modified: 'Own start added end!' },
        { type: 'redline', target: byText('Plain'), modified: 'Plain paragraph!' },
        { type: 'redline', target: byText('Cell'), modified: 'Cell note text!' }
    ];
    const preflight = await preflightOperations(mixedXml, operations, 'Me', { existingRevisions: 'auto' });
    assert.deepEqual(preflight.results.map(item => item.existingRevisions),
        ['slice-cross-author', 'slice-cross-author', 'merge-same-author', 'merge-same-author', 'slice-cross-author'],
        'preflight resolves each target');
    assert.ok(preflight.results.every(item => !item.error), JSON.stringify(preflight.results.map(item => item.error)));

    const result = await applyOperationsToDocumentXml(mixedXml, operations, 'Me', null, { atomic: true, strictTargets: true, existingRevisions: 'auto' });
    assert.equal(result.status, 'ok', JSON.stringify(result.error ?? result.results));
    const accepted = view(result.documentXml, acceptTrackedChangesInOoxml);
    const rejected = view(result.documentXml, rejectTrackedChangesInOoxml);
    assert.deepEqual(accepted.slice(0, 5), ['Foreign start added end!', 'Foreign deletion kept!', 'Own start added end!', 'Plain paragraph!', 'Cell note text!']);
    assert.deepEqual(rejected.slice(0, 5), ['Foreign start end.', 'Foreign deletion removed kept.', 'Own start end.', 'Plain paragraph.', 'Cell text.'],
        'Reject All restores every reviewer baseline');
    assert.ok(authorsOf(result.documentXml).includes('Reviewer'), 'reviewer history is preserved');

    // The same batch with the explicit per-target policies gives the same views.
    const explicit = await applyOperationsToDocumentXml(mixedXml, operations.map((operation, index) => ({
        ...operation, existingRevisions: preflight.results[index].existingRevisions
    })), 'Me', null, { atomic: true, strictTargets: true });
    assert.equal(explicit.status, 'ok', JSON.stringify(explicit.error ?? explicit.results));
    assert.deepEqual(view(explicit.documentXml, acceptTrackedChangesInOoxml), accepted);
    assert.deepEqual(view(explicit.documentXml, rejectTrackedChangesInOoxml), rejected);
}
{
    // A range whose first paragraph is plain and whose second carries another reviewer's insertion:
    // auto resolves over the whole range, not just the start.
    const operation = { type: 'redline', target: byText('Range one'), targetEnd: byText('Range two'), modified: 'Range one!\nRange two!' };
    const preflight = await preflightOperations(mixedXml, [operation], 'Me', { existingRevisions: 'auto' });
    assert.equal(preflight.results[0].existingRevisions, 'slice-cross-author', 'preflight resolves over the range');
    const result = await applyOperationsToDocumentXml(mixedXml, [operation], 'Me', null, { atomic: true, strictTargets: true, existingRevisions: 'auto' });
    assert.equal(result.status, 'ok', JSON.stringify(result.error ?? result.results));
    assert.deepEqual(view(result.documentXml, acceptTrackedChangesInOoxml).slice(-2), ['Range one!', 'Range two!']);
    assert.deepEqual(view(result.documentXml, rejectTrackedChangesInOoxml).slice(-2), ['Range one.', 'Range ']);
}
{
    // Untracked edits under auto keep the foreign insertion and create no revision of their own.
    const result = await applyOperationsToDocumentXml(mixedXml, [
        { type: 'redline', target: byText('Foreign start'), modified: 'Foreign start added end!' },
        { type: 'redline', target: byText('Plain'), modified: 'Plain paragraph!' }
    ], 'Me', null, { atomic: true, strictTargets: true, existingRevisions: 'auto', generateRedlines: false });
    assert.equal(result.status, 'ok', JSON.stringify(result.error ?? result.results));
    assert.deepEqual(view(result.documentXml, acceptTrackedChangesInOoxml).slice(0, 4), ['Foreign start added end!', 'Foreign deletion kept.', 'Own start added end.', 'Plain paragraph!']);
    const dates = [...new Set([...result.documentXml.matchAll(/w:date="([^"]*)"/g)].map(m => m[1]))];
    assert.deepEqual(dates, [DATE], 'untracked edits add no revisions (every revision is a source revision)');
}
{
    // Rejected-view insertion keeps requiring the explicit slicing policy.
    const operation = {
        type: 'insert', target: { ...byText('Foreign deletion'), revisionView: 'rejected' },
        anchor: { exactText: 'removed', occurrence: 1, offset: 0 }, modified: 'now ', existingRevisions: 'auto'
    };
    const result = await applyOperationsToDocumentXml(mixedXml, [operation], 'Me', null, { strictTargets: true });
    assert.notEqual(result.status, 'ok');
    assert.match(JSON.stringify(result), /slice-cross-author/);
}

// ---------------------------------------------------------------------------------------------------------
// 5. openDocx and the CLI
// ---------------------------------------------------------------------------------------------------------
const encoder = new TextEncoder();
const template = unzipDocx(new Uint8Array(await readFile(new URL('./fixtures/sample_doc_test.docx', import.meta.url))));
const packageXml = docXml(para(run('Foreign start ') + ins('Reviewer', 'added ') + run('end.')) + para(run('Plain paragraph.')));
const entries = new Map(template);
entries.set('word/document.xml', encoder.encode(packageXml));
const packageBytes = zipDocx(entries);
const packageOperations = paragraphs => [
    { type: 'redline', target: strong(paragraphs[0]), modified: 'Foreign start added end!' },
    { type: 'redline', target: strong(paragraphs[1]), modified: 'Plain paragraph!' }
];
{
    const doc = openDocx(packageBytes);
    const result = await doc.applyOperations(packageOperations(doc.inspect().paragraphs), { author: 'Me', atomic: true, existingRevisions: 'auto' });
    assert.equal(result.status, 'ok', JSON.stringify(result.error ?? result.results));
    const reopened = openDocx(result.toUint8Array()).inspect().paragraphs;
    assert.deepEqual(reopened.map(p => p.exactText), ['Foreign start added end!', 'Plain paragraph!']);
    assert.deepEqual([...reopened[0].revisionAuthors].sort(), ['Me', 'Reviewer']);
}
{
    const directory = await mkdtemp(path.join(tmpdir(), 'docx-auto-policy-'));
    try {
        const input = path.join(directory, 'input.docx');
        const output = path.join(directory, 'output.docx');
        await writeFile(input, packageBytes);
        const inspected = openDocx(packageBytes).inspect().paragraphs;
        const result = await executeCli([
            'apply', input, '--operations', '-', '--profile', 'agent', '--atomic',
            '--existing-revisions', 'auto', '--author', 'Me', '--output', output
        ], { stdin: Readable.from([JSON.stringify(packageOperations(inspected))]) });
        assert.equal(result.status, 'ok', JSON.stringify(result));
        assert.equal(result.completion, true);
        assert.equal(result.effectiveOptions.existingRevisions, 'auto');
        const written = openDocx(new Uint8Array(await readFile(output))).inspect().paragraphs;
        assert.deepEqual(written.map(p => p.exactText), ['Foreign start added end!', 'Plain paragraph!']);

        const refused = await executeCli([
            'apply', input, '--operations', '-', '--existing-revisions', 'automatic', '--output', path.join(directory, 'refused.docx')
        ], { stdin: Readable.from([JSON.stringify(packageOperations(inspected))]) });
        assert.equal(refused.status, 'error');
        assert.equal(refused.error.code, 'INVALID_OPERATION');
    } finally {
        await rm(directory, { recursive: true, force: true });
    }
}

console.log('existing_revisions_auto_paths_tests: all passed');
