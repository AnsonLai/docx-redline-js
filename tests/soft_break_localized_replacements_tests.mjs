import assert from 'node:assert/strict';

import './setup-xml-provider.mjs';
import { inspectDocumentParts, acceptTrackedChangesInOoxml, rejectTrackedChangesInOoxml } from '../index.js';
import { applyOperationsToDocumentXml } from '../services/standalone-operation-runner.js';
import { parseOoxmlSafe } from '../adapters/xml-adapter.js';
import { readFileSync } from 'node:fs';
import { openDocx } from '../index.js';
import { unzipDocx, zipDocx } from '../document/zip-archive.js';
import { applyOperationToDocumentXml } from '../services/standalone-operation-runner.js';
import { DocumentOperationSession } from '../services/document-operation-session.js';
import { compileOperationBatch } from '../services/operation-batch-compiler.js';

// Regression: localized `replacements` were refused for every paragraph whose
// accepted-view exactText contains "\n" (a w:br soft line break).

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const documentXml = `<w:document ${W}><w:body><w:p>`
    + '<w:r><w:t>A. First recital.</w:t></w:r>'
    + '<w:r><w:br/><w:t>B. The Parties desire a potential business relationship.</w:t></w:r>'
    + '<w:r><w:br/><w:t>C. Third recital, first and last.</w:t></w:r>'
    + '</w:p><w:sectPr/></w:body></w:document>';
const paragraph = inspectDocumentParts({ documentXml }).paragraphs[0];
assert.ok(/\n/.test(paragraph.exactText), 'precondition: exactText has a line break');

const textOf = xml => {
    const parsed = parseOoxmlSafe(xml);
    assert.ok(parsed.doc, 'parseable');
    return Array.from(parsed.doc.getElementsByTagNameNS('*', 'r')).map(run => Array.from(run.childNodes).map(node => (
        node.localName === 't' || node.localName === 'delText' ? node.textContent : node.localName === 'br' ? '\n' : ''
    )).join('')).join('');
};
const countBr = xml => (xml.match(/<w:br\b/g) || []).length;
const apply = (replacements, extra = {}) => applyOperationsToDocumentXml(documentXml, [{
    type: 'redline',
    target: { index: 1, exactText: paragraph.exactText, paragraphId: paragraph.paragraphId, fingerprint: paragraph.fingerprint },
    replacements,
    ...extra
}], 'Tester', null, { atomic: true, generateRedlines: true });

// 1. Single-line find applies as a tracked change and preserves the breaks.
{
    const result = await apply([{ find: 'a potential business relationship', replace: 'project Titan' }]);
    assert.equal(result.status, 'ok', JSON.stringify(result.error || result.results));
    const xml = result.documentXml;
    assert.equal(countBr(xml), 2, 'both w:br preserved');
    assert.ok(/<w:ins\b/.test(xml) && /<w:del\b/.test(xml));
    assert.equal(textOf(acceptTrackedChangesInOoxml(xml, { allAuthors: true }).oxml),
        'A. First recital.\nB. The Parties desire project Titan.\nC. Third recital, first and last.');
    const rejected = rejectTrackedChangesInOoxml(xml, { allAuthors: true }).oxml;
    assert.equal(textOf(rejected), paragraph.exactText);
    assert.equal(countBr(rejected), 2);
}

// 2. Occurrence counting spans lines; ambiguity and explicit occurrence work.
{
    const ambiguous = await apply([{ find: 'recital', replace: 'item' }]);
    assert.equal(ambiguous.status, 'error');
    const second = await apply([{ find: 'recital', replace: 'item', occurrence: 2 }]);
    assert.equal(second.status, 'ok', JSON.stringify(second.error || second.results));
    assert.equal(textOf(acceptTrackedChangesInOoxml(second.documentXml, { allAuthors: true }).oxml),
        ['A. First recital.', 'B. The Parties desire a potential business relationship.', 'C. Third item, first and last.'].join('\n'));
    const result = await apply([
        { find: 'First recital', replace: 'Opening recital' },
        { find: 'Third recital', replace: 'Closing recital' }
    ]);
    assert.equal(result.status, 'ok', JSON.stringify(result.error || result.results));
    assert.equal(countBr(result.documentXml), 2);
    assert.equal(textOf(acceptTrackedChangesInOoxml(result.documentXml, { allAuthors: true }).oxml),
        'A. Opening recital.\nB. The Parties desire a potential business relationship.\nC. Closing recital, first and last.');
    assert.equal(textOf(rejectTrackedChangesInOoxml(result.documentXml, { allAuthors: true }).oxml), paragraph.exactText);
}

// 3. Replacement that introduces a line break stays refused.
{
    const result = await apply([{ find: 'First recital', replace: 'One\nTwo' }]);
    assert.equal(result.status, 'error');
}

// 4. A find spanning a line break is refused; accepting it would keep the w:br.
{
    const result = await apply([{ find: 'recital.\nB. The Parties', replace: 'recital. The Parties' }]);
    assert.equal(result.status, 'error');
    assert.match(JSON.stringify(result), /cannot span a line break/);
}

// 5. Weak target (index + exactText only, no paragraphId/fingerprint) behaves exactly like the strong target.
// Regression: the batch compiler left weak targets whose exactText has a line break "dynamic", skipped the
// replacement compilation, and applied an empty "modified": the whole paragraph was tracked as deleted with ok.
const weakTarget = { index: 1, exactText: paragraph.exactText };
const strongTarget = { ...weakTarget, paragraphId: paragraph.paragraphId, fingerprint: paragraph.fingerprint };
const expectedAccepted = 'A. First recital.\nB. The Parties desire project Titan.\nC. Third recital, first and last.';
const normalizeMarkup = xml => xml.replace(/w:id="\d+"/g, '').replace(/w:date="[^"]*"/g, '');
{
    const find = [{ find: 'a potential business relationship', replace: 'project Titan' }];
    const run = target => applyOperationsToDocumentXml(documentXml, [{ type: 'redline', target, replacements: find }],
        'Tester', null, { atomic: true, generateRedlines: true });
    const weak = await run(weakTarget);
    const strong = await run(strongTarget);
    for (const [name, result] of [['weak', weak], ['strong', strong]]) {
        assert.equal(result.status, 'ok', name + JSON.stringify(result.error || result.results));
        assert.equal(textOf(acceptTrackedChangesInOoxml(result.documentXml, { allAuthors: true }).oxml), expectedAccepted, name);
        assert.equal(countBr(result.documentXml), 2, name);
        assert.equal(result.results[0].change.verification.acceptedViewMatchesCompiledText, true, name);
    }
    assert.equal(normalizeMarkup(weak.documentXml), normalizeMarkup(strong.documentXml), 'weak and strong targets write identical markup');
    // The single-operation entry point takes the same path.
    const single = await applyOperationToDocumentXml(documentXml, { type: 'redline', target: weakTarget, replacements: find }, 'Tester', null, {});
    assert.equal(textOf(acceptTrackedChangesInOoxml(single.documentXml, { allAuthors: true }).oxml), expectedAccepted);

    // Packaged openDocx path.
    const entries = unzipDocx(readFileSync(new URL('./fixtures/agentic-lists/nested-lists-source.docx', import.meta.url)));
    const original = new TextDecoder().decode(entries.get('word/document.xml'));
    const sectPr = original.match(/<w:sectPr[\s\S]*<\/w:sectPr>\s*<\/w:body>/)[0];
    const head = original.slice(0, original.indexOf('<w:body>') + '<w:body>'.length);
    const bodyXml = documentXml.match(/<w:p>[\s\S]*<\/w:p>/)[0];
    entries.set('word/document.xml', new TextEncoder().encode(head + bodyXml + sectPr + '</w:document>'));
    const bytes = new Uint8Array(zipDocx(entries));
    const packagedParagraph = openDocx(bytes).inspect({}).paragraphs[0];
    const packaged = await openDocx(bytes).applyOperations([{
        type: 'redline', target: { index: 1, exactText: packagedParagraph.exactText }, replacements: find
    }], { author: 'Tester' });
    assert.equal(packaged.status, 'ok', JSON.stringify(packaged.error || packaged.results));
    const acceptedPackage = await openDocx(packaged.toUint8Array()).resolveRevisions('accept', { allAuthors: true });
    const acceptedXml = new TextDecoder().decode(unzipDocx(acceptedPackage.toUint8Array()).get('word/document.xml'));
    assert.equal(textOf(acceptedXml), expectedAccepted, 'packaged weak target keeps the paragraph');
}

// 6. Accept-view intent guard: a localized replacement whose result is not the compiled text is refused before
// commit, whatever path produced it (here: the pre-fix shape, an empty "modified" that tracks the paragraph deleted).
{
    const session = new DocumentOperationSession(documentXml, {});
    const source = { type: 'redline', target: strongTarget, replacements: [{ find: 'a potential business relationship', replace: 'project Titan' }] };
    const compilation = compileOperationBatch(session.document, [source], { strictTargets: true });
    session.sourceTargetRegistry = compilation.registry;
    const bad = { ...compilation.compiledOperations[0], modified: '' };
    assert.ok(bad._localizedReplacementCompilation, 'precondition: compiled localized replacement');
    const refused = await applyOperationToDocumentXml(documentXml, bad, 'Tester', null, { _documentOperationSession: session });
    assert.equal(refused.status, 'error');
    assert.equal(refused.error.code, 'PATCH_ROUNDTRIP_MISMATCH');
    assert.equal(refused.hasChanges, false);
    assert.equal(refused.documentXml, documentXml, 'nothing committed');
    assert.ok(!/<w:del\b/.test(session.serializeCurrent()), 'savepoint restored');
}

console.log('soft_break_localized_replacements_tests: all passed');
