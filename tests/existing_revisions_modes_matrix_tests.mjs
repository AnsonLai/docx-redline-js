import './setup-xml-provider.mjs';

import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';

import {
    applyRedlineToOxml,
    acceptTrackedChangesInOoxml,
    rejectTrackedChangesInOoxml,
    ingestWordOoxmlToPlainText,
    reconcileMarkdownTableOoxml
} from '../index.js';
import {
    applyOperationsToDocumentXml
} from '../services/standalone-operation-runner.js';
import { preflightOperations } from '../services/operation-preflight.js';
import { buildZip } from '../scripts/lib/minimal-zip.mjs';
import { executeCli } from '../node/cli.js';

const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

function createDocumentXml(bodyInner) {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${NS_W}">
  <w:body>
    ${bodyInner}
    <w:sectPr/>
  </w:body>
</w:document>`;
}

function makeRevisedParagraph({
    prefix = 'Clause 1: ',
    delText = 'Payment within 30 days.',
    insText = 'Payment within 45 days.',
    suffix = ' Late fees apply.',
    author = 'Alice',
    paraId = 'P1'
} = {}) {
    const authorAttr = author ? `w:author="${author}"` : '';
    return `<w:p xmlns:w="${NS_W}" w:paraId="${paraId}">`
        + `<w:r><w:t xml:space="preserve">${prefix}</w:t></w:r>`
        + `<w:del w:id="101" ${authorAttr} w:date="2026-01-01T00:00:00Z"><w:r><w:delText>${delText}</w:delText></w:r></w:del>`
        + `<w:ins w:id="102" ${authorAttr} w:date="2026-01-01T00:00:00Z"><w:r><w:t>${insText}</w:t></w:r></w:ins>`
        + `<w:r><w:t xml:space="preserve">${suffix}</w:t></w:r>`
        + `</w:p>`;
}

function makeCommentedRevisedParagraph({
    author = 'Alice',
    commentId = '1',
    paraId = 'P_COM'
} = {}) {
    return `<w:p xmlns:w="${NS_W}" w:paraId="${paraId}">`
        + `<w:r><w:t xml:space="preserve">Clause: </w:t></w:r>`
        + `<w:commentRangeStart w:id="${commentId}"/>`
        + `<w:del w:id="201" w:author="${author}" w:date="2026-01-01T00:00:00Z"><w:r><w:delText>net 30</w:delText></w:r></w:del>`
        + `<w:ins w:id="202" w:author="${author}" w:date="2026-01-01T00:00:00Z"><w:r><w:t>net 60</w:t></w:r></w:ins>`
        + `<w:commentRangeEnd w:id="${commentId}"/>`
        + `<w:r><w:commentReference w:id="${commentId}"/></w:r>`
        + `</w:p>`;
}

// ===========================================================================
// 1. applyRedlineToOxml: Matrix of all 4 existingRevisions modes
// ===========================================================================

console.log('Testing Section 1: applyRedlineToOxml policy matrix...');

// ---------------------------------------------------------------------------
// 1.1 Mode 'merge-same-author' (Explicit and Default)
// ---------------------------------------------------------------------------
{
    const sourceXml = makeRevisedParagraph({ author: 'Alice' });
    const acceptedOriginal = 'Clause 1: Payment within 45 days. Late fees apply.';
    const baselineOriginal = 'Clause 1: Payment within 30 days. Late fees apply.';

    // 1.1.a Default when option omitted should be 'merge-same-author'
    const defaultRes = await applyRedlineToOxml(
        sourceXml,
        acceptedOriginal,
        'Clause 1: Payment within 60 days. Late fees apply.',
        { author: 'Alice' }
    );
    assert.equal(defaultRes.status, 'ok');
    assert.equal(defaultRes.hasChanges, true);
    // Verified: re-diffed from baseline (30 days) to new (60 days)
    const acc = acceptTrackedChangesInOoxml(defaultRes.oxml, { author: 'Alice' });
    const rej = rejectTrackedChangesInOoxml(defaultRes.oxml, { author: 'Alice' });
    assert.equal(ingestWordOoxmlToPlainText(acc.oxml), 'Clause 1: Payment within 60 days. Late fees apply.');
    assert.equal(ingestWordOoxmlToPlainText(rej.oxml), baselineOriginal, 'Rejecting must restore original baseline, not intermediate');

    // 1.1.b Same-author complete revert back to baseline
    const revertRes = await applyRedlineToOxml(
        sourceXml,
        acceptedOriginal,
        baselineOriginal,
        { author: 'Alice', existingRevisions: 'merge-same-author' }
    );
    assert.equal(revertRes.status, 'ok');
    assert.equal(revertRes.hasChanges, true);
    assert.equal(ingestWordOoxmlToPlainText(revertRes.oxml), baselineOriginal);
    assert.ok(revertRes.warnings?.some(w => w.includes('Previous revisions by the same author were reverted to baseline')));

    // 1.1.c Different author fails closed with EXISTING_REVISIONS
    const diffAuthorRes = await applyRedlineToOxml(
        sourceXml,
        acceptedOriginal,
        'Clause 1: Payment within 60 days. Late fees apply.',
        { author: 'Bob', existingRevisions: 'merge-same-author' }
    );
    assert.equal(diffAuthorRes.status, 'error');
    assert.equal(diffAuthorRes.error?.code, 'EXISTING_REVISIONS');
    assert.equal(diffAuthorRes.hasChanges, false);
    assert.ok(diffAuthorRes.error?.message.includes('Alice'));

    // 1.1.d Multiple authors in target paragraph fails closed with EXISTING_REVISIONS
    const mixedAuthorsXml = `<w:p xmlns:w="${NS_W}">`
        + `<w:ins w:id="1" w:author="Alice" w:date="2026-01-01T00:00:00Z"><w:r><w:t>Alice addition. </w:t></w:r></w:ins>`
        + `<w:ins w:id="2" w:author="Bob" w:date="2026-01-01T00:00:00Z"><w:r><w:t>Bob addition.</w:t></w:r></w:ins>`
        + `</w:p>`;
    const mixedRes = await applyRedlineToOxml(
        mixedAuthorsXml,
        'Alice addition. Bob addition.',
        'Alice addition. Bob revision.',
        { author: 'Alice', existingRevisions: 'merge-same-author' }
    );
    assert.equal(mixedRes.status, 'error');
    assert.equal(mixedRes.error?.code, 'EXISTING_REVISIONS');
    assert.equal(mixedRes.hasChanges, false);

    // 1.1.e Unattributed revision fails closed with EXISTING_REVISIONS
    const unattributedXml = makeRevisedParagraph({ author: '' });
    const unattrRes = await applyRedlineToOxml(
        unattributedXml,
        acceptedOriginal,
        'New clause text.',
        { author: 'Alice', existingRevisions: 'merge-same-author' }
    );
    assert.equal(unattrRes.status, 'error');
    assert.equal(unattrRes.error?.code, 'EXISTING_REVISIONS');
    assert.ok(unattrRes.error?.message.includes('unattributed'));

    // 1.1.f Target paragraph with comments fails closed with COMMENTED_CONTENT_MERGE
    const commentedXml = makeCommentedRevisedParagraph({ author: 'Alice' });
    const comRes = await applyRedlineToOxml(
        commentedXml,
        'Clause: net 60',
        'Clause: net 90',
        { author: 'Alice', existingRevisions: 'merge-same-author' }
    );
    assert.equal(comRes.status, 'error');
    assert.equal(comRes.error?.code, 'COMMENTED_CONTENT_MERGE');
    assert.equal(comRes.hasChanges, false);

    // 1.1.g Direct edit without redlines (generateRedlines: false) by same author
    const directRes = await applyRedlineToOxml(
        sourceXml,
        acceptedOriginal,
        'Clause 1: Clean text without redlines.',
        { author: 'Alice', generateRedlines: false, existingRevisions: 'merge-same-author' }
    );
    assert.equal(directRes.status, 'ok');
    assert.equal(directRes.hasChanges, true);
    assert.ok(!directRes.oxml.includes('<w:ins'));
    assert.ok(!directRes.oxml.includes('<w:del'));
    assert.equal(ingestWordOoxmlToPlainText(directRes.oxml), 'Clause 1: Clean text without redlines.');
}

// ---------------------------------------------------------------------------
// 1.2 Mode 'accept-all-first'
// ---------------------------------------------------------------------------
{
    const sourceXml = makeRevisedParagraph({ author: 'Alice' });
    const acceptedOriginal = 'Clause 1: Payment within 45 days. Late fees apply.';

    // 1.2.a Different author succeeds under accept-all-first
    const diffAuthorRes = await applyRedlineToOxml(
        sourceXml,
        acceptedOriginal,
        'Clause 1: Payment within 90 days. Late fees apply.',
        { author: 'Bob', existingRevisions: 'accept-all-first' }
    );
    assert.equal(diffAuthorRes.status, 'ok');
    assert.equal(diffAuthorRes.hasChanges, true);
    assert.ok(!diffAuthorRes.oxml.includes('w:author="Alice"'), 'Alice prior revisions must be accepted/removed');
    assert.ok(diffAuthorRes.oxml.includes('w:author="Bob"'), 'Bob new revisions must be present');
    const acc = acceptTrackedChangesInOoxml(diffAuthorRes.oxml, { author: 'Bob' });
    const rej = rejectTrackedChangesInOoxml(diffAuthorRes.oxml, { author: 'Bob' });
    assert.equal(ingestWordOoxmlToPlainText(acc.oxml), 'Clause 1: Payment within 90 days. Late fees apply.');
    assert.equal(ingestWordOoxmlToPlainText(rej.oxml), acceptedOriginal, 'Rejecting Bob revision restores Alice accepted text');

    // 1.2.b Multiple authors normalized under accept-all-first
    const mixedAuthorsXml = `<w:p xmlns:w="${NS_W}">`
        + `<w:ins w:id="1" w:author="Alice" w:date="2026-01-01T00:00:00Z"><w:r><w:t>Alice addition. </w:t></w:r></w:ins>`
        + `<w:ins w:id="2" w:author="Bob" w:date="2026-01-01T00:00:00Z"><w:r><w:t>Bob addition.</w:t></w:r></w:ins>`
        + `</w:p>`;
    const charlieRes = await applyRedlineToOxml(
        mixedAuthorsXml,
        'Alice addition. Bob addition.',
        'Alice addition. Charlie revision.',
        { author: 'Charlie', existingRevisions: 'accept-all-first' }
    );
    assert.equal(charlieRes.status, 'ok');
    assert.equal(charlieRes.hasChanges, true);
    assert.ok(!charlieRes.oxml.includes('w:author="Alice"'), 'Alice revision author attribute must be removed');
    assert.ok(!charlieRes.oxml.includes('w:author="Bob"'), 'Bob revision author attribute must be removed');
    assert.ok(charlieRes.oxml.includes('w:author="Charlie"'), 'Charlie revision author attribute must be present');

    // 1.2.c No-op edit under accept-all-first preserves original OOXML
    const noOpRes = await applyRedlineToOxml(
        sourceXml,
        acceptedOriginal,
        acceptedOriginal,
        { author: 'Bob', existingRevisions: 'accept-all-first' }
    );
    assert.equal(noOpRes.status, 'no-op');
    assert.equal(noOpRes.hasChanges, false);
    assert.equal(noOpRes.oxml, sourceXml, 'Original revisions must be preserved on no-op under accept-all-first');
}

// ---------------------------------------------------------------------------
// 1.3 Mode 'accept-all-first-keep-normalized'
// ---------------------------------------------------------------------------
{
    const sourceXml = makeRevisedParagraph({ author: 'Alice' });
    const acceptedOriginal = 'Clause 1: Payment within 45 days. Late fees apply.';

    // 1.3.a Normal edit succeeds
    const editRes = await applyRedlineToOxml(
        sourceXml,
        acceptedOriginal,
        'Clause 1: Payment within 90 days. Late fees apply.',
        { author: 'Bob', existingRevisions: 'accept-all-first-keep-normalized' }
    );
    assert.equal(editRes.status, 'ok');
    assert.equal(editRes.hasChanges, true);

    // 1.3.b No-op edit returns normalized OOXML with hasChanges: true and warning
    const noOpRes = await applyRedlineToOxml(
        sourceXml,
        acceptedOriginal,
        acceptedOriginal,
        { author: 'Bob', existingRevisions: 'accept-all-first-keep-normalized' }
    );
    assert.equal(noOpRes.status, 'ok');
    assert.equal(noOpRes.hasChanges, true);
    assert.ok(!noOpRes.oxml.includes('<w:ins'));
    assert.ok(!noOpRes.oxml.includes('<w:del'));
    assert.ok(noOpRes.warnings?.some(w => w.includes('Existing revisions were accepted before redlining')));
    assert.equal(ingestWordOoxmlToPlainText(noOpRes.oxml), acceptedOriginal);
}

// ---------------------------------------------------------------------------
// 1.4 Mode 'reject-input'
// ---------------------------------------------------------------------------
{
    const sourceXml = makeRevisedParagraph({ author: 'Alice' });
    const acceptedOriginal = 'Clause 1: Payment within 45 days. Late fees apply.';

    // 1.4.a Same author fails with EXISTING_REVISIONS
    const sameRes = await applyRedlineToOxml(
        sourceXml,
        acceptedOriginal,
        'Clause 1: Payment within 60 days. Late fees apply.',
        { author: 'Alice', existingRevisions: 'reject-input' }
    );
    assert.equal(sameRes.status, 'error');
    assert.equal(sameRes.error?.code, 'EXISTING_REVISIONS');
    assert.equal(sameRes.hasChanges, false);
    assert.equal(sameRes.oxml, sourceXml);

    // 1.4.b Different author fails with EXISTING_REVISIONS
    const diffRes = await applyRedlineToOxml(
        sourceXml,
        acceptedOriginal,
        'Clause 1: Payment within 60 days. Late fees apply.',
        { author: 'Bob', existingRevisions: 'reject-input' }
    );
    assert.equal(diffRes.status, 'error');
    assert.equal(diffRes.error?.code, 'EXISTING_REVISIONS');
    assert.equal(diffRes.hasChanges, false);
    assert.equal(diffRes.oxml, sourceXml);
}

// ===========================================================================
// 2. Standalone Operations Runner & Preflight Matrix
// ===========================================================================

console.log('Testing Section 2: Document operations runner and preflight across all modes...');

{
    const p1 = makeRevisedParagraph({ author: 'Alice', paraId: 'P1', delText: 'old one', insText: 'new one', prefix: 'Para 1: ', suffix: '' });
    // P2 is insertion-only: under reject-input returns EXISTING_REVISIONS
    const p2 = `<w:p xmlns:w="${NS_W}" w:paraId="P2"><w:ins w:id="103" w:author="Alice" w:date="2026-01-01T00:00:00Z"><w:r><w:t>Para 2: new two</w:t></w:r></w:ins></w:p>`;
    // P2b has deletion: under reject-input returns UNSAFE_REVISION_NESTING
    const p2b = makeRevisedParagraph({ author: 'Alice', paraId: 'P2B', delText: 'old two b', insText: 'new two b', prefix: 'Para 2B: ', suffix: '' });
    const p3 = makeRevisedParagraph({ author: 'Bob', paraId: 'P3', delText: 'old three', insText: 'new three', prefix: 'Para 3: ', suffix: '' });
    const p4 = makeCommentedRevisedParagraph({ author: 'Alice', commentId: '10', paraId: 'P4' });

    const docXml = createDocumentXml(`${p1}\n${p2}\n${p2b}\n${p3}\n${p4}`);

    // Preflight tests:
    // Op 1: Alice on P1 under merge-same-author -> ready
    // Op 2: Alice on P2 (insertion-only) under reject-input -> error (EXISTING_REVISIONS)
    // Op 2b: Alice on P2B (with deletion) under reject-input -> error (UNSAFE_REVISION_NESTING)
    // Op 3: Alice on P3 (Bob's revision) under merge-same-author -> error (EXISTING_REVISIONS)
    // Op 4: Alice on P3 (Bob's revision) under accept-all-first -> ready
    // Op 5: Alice on P4 (commented) under merge-same-author -> error (COMMENTED_CONTENT_MERGE)

    const preflightOps = [
        { type: 'replace', target: 'Para 1: new one', modified: 'Para 1: final one', existingRevisions: 'merge-same-author' },
        { type: 'replace', target: 'Para 2: new two', modified: 'Para 2: final two', existingRevisions: 'reject-input' },
        { type: 'replace', target: 'Para 2B: new two b', modified: 'Para 2B: final two b', existingRevisions: 'reject-input' },
        { type: 'replace', target: 'Para 3: new three', modified: 'Para 3: final three', existingRevisions: 'merge-same-author' },
        { type: 'replace', target: 'Para 3: new three', modified: 'Para 3: final three', existingRevisions: 'accept-all-first' },
        { type: 'replace', target: 'Clause: net 60', modified: 'Clause: net 90', existingRevisions: 'merge-same-author' }
    ];

    const pfResult = preflightOperations(docXml, preflightOps, 'Alice', {
        _existingCommentDetails: { '10': { author: 'Reviewer', text: 'Important clause' } }
    });

    assert.equal(pfResult.results[0].status, 'ready', 'Op 1 (same author merge) must be ready');
    assert.equal(pfResult.results[1].status, 'error', 'Op 2 (reject-input insertion only) must error');
    assert.equal(pfResult.results[1].error?.code, 'EXISTING_REVISIONS');
    assert.equal(pfResult.results[2].status, 'error', 'Op 2b (reject-input deletion) must error');
    assert.equal(pfResult.results[2].error?.code, 'UNSAFE_REVISION_NESTING');
    assert.equal(pfResult.results[3].status, 'error', 'Op 3 (different author under merge-same-author) must error');
    assert.equal(pfResult.results[3].error?.code, 'EXISTING_REVISIONS');
    assert.equal(pfResult.results[4].status, 'ready', 'Op 4 (different author under accept-all-first) must be ready');
    assert.equal(pfResult.results[5].status, 'error', 'Op 5 (commented content merge) must error');
    assert.equal(pfResult.results[5].error?.code, 'COMMENTED_CONTENT_MERGE');

    // Runner tests: progressive execution (atomic: false)
    const runOps = [
        { type: 'replace', target: 'Para 1: new one', modified: 'Para 1: final one', existingRevisions: 'merge-same-author' },
        { type: 'replace', target: 'Para 3: new three', modified: 'Para 3: final three', existingRevisions: 'accept-all-first' }
    ];

    const runResult = await applyOperationsToDocumentXml(docXml, runOps, 'Alice', null, { atomic: false });
    assert.equal(runResult.status, 'ok');
    assert.equal(runResult.hasChanges, true);
    assert.equal(runResult.results[0].status, 'applied');
    assert.equal(runResult.results[1].status, 'applied');

    // Verify output document content using canonical plain text
    const plain = ingestWordOoxmlToPlainText(runResult.documentXml);
    assert.ok(plain.includes('Para 1: final one'));
    assert.ok(plain.includes('Para 3: final three'));
}

// ===========================================================================
// 3. Tables: reconcileMarkdownTableOoxml across modes
// ===========================================================================

console.log('Testing Section 3: Tables policy matrix...');

{
    const tableXml = `<w:tbl xmlns:w="${NS_W}">`
        + `<w:tr>`
        + `<w:tc><w:p><w:r><w:t>Header A</w:t></w:r></w:p></w:tc>`
        + `<w:tc><w:p><w:r><w:t>Header B</w:t></w:r></w:p></w:tc>`
        + `</w:tr>`
        + `<w:tr>`
        + `<w:tc><w:p>`
        + `<w:del w:id="301" w:author="Alice" w:date="2026-01-01T00:00:00Z"><w:r><w:delText>old cell</w:delText></w:r></w:del>`
        + `<w:ins w:id="302" w:author="Alice" w:date="2026-01-01T00:00:00Z"><w:r><w:t>new cell</w:t></w:r></w:ins>`
        + `</w:p></w:tc>`
        + `<w:tc><w:p><w:r><w:t>static cell</w:t></w:r></w:p></w:tc>`
        + `</w:tr>`
        + `</w:tbl>`;

    const origText = ingestWordOoxmlToPlainText(tableXml);
    const modMd = `| Header A | Header B |\n| --- | --- |\n| newer cell | static cell |`;

    // 3.1 Same author merge in table cell
    const tblSame = await reconcileMarkdownTableOoxml(tableXml, origText, modMd, {
        author: 'Alice',
        existingRevisions: 'merge-same-author'
    });
    assert.equal(tblSame.status, 'ok');
    assert.equal(tblSame.hasChanges, true);

    // 3.2 Different author under merge-same-author fails
    const tblDiff = await reconcileMarkdownTableOoxml(tableXml, origText, modMd, {
        author: 'Bob',
        existingRevisions: 'merge-same-author'
    });
    assert.equal(tblDiff.status, 'error');
    assert.equal(tblDiff.error?.code, 'EXISTING_REVISIONS');

    // 3.3 Different author under accept-all-first succeeds
    const tblAccept = await reconcileMarkdownTableOoxml(tableXml, origText, modMd, {
        author: 'Bob',
        existingRevisions: 'accept-all-first'
    });
    assert.equal(tblAccept.status, 'ok');
    assert.equal(tblAccept.hasChanges, true);

    // 3.4 Reject-input fails
    const tblReject = await reconcileMarkdownTableOoxml(tableXml, origText, modMd, {
        author: 'Alice',
        existingRevisions: 'reject-input'
    });
    assert.equal(tblReject.status, 'error');
    assert.equal(tblReject.error?.code, 'EXISTING_REVISIONS');
}

// ===========================================================================
// 4. CLI / Node Facade Integration with --existing-revisions
// ===========================================================================

console.log('Testing Section 4: CLI integration with --existing-revisions...');

{
    const contentTypes = `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">`
        + `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>`
        + `<Default Extension="xml" ContentType="application/xml"/>`
        + `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>`
        + `</Types>`;
    const rels = `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`;
    const docWithRevision = createDocumentXml(makeRevisedParagraph({
        author: 'Alice',
        paraId: 'P_CLI',
        prefix: 'Term: ',
        delText: '30 days',
        insText: '60 days',
        suffix: ''
    }));

    const pkgFixture = buildZip([
        { name: '[Content_Types].xml', data: contentTypes },
        { name: 'word/document.xml', data: docWithRevision },
        { name: 'word/_rels/document.xml.rels', data: rels }
    ]);

    const testDir = await mkdtemp(path.join(tmpdir(), 'docx-revisions-cli-'));
    try {
        const inputPath = path.join(testDir, 'revised.docx');
        await writeFile(inputPath, pkgFixture);

        // 4.1 CLI default (merge-same-author): Alice succeeds
        const aliceOut = path.join(testDir, 'alice_out.docx');
        const aliceRes = await executeCli([
            'apply', inputPath,
            '--target', 'Term: 60 days',
            '--modified', 'Term: 90 days',
            '--author', 'Alice',
            '--output', aliceOut
        ]);
        assert.equal(aliceRes.status, 'ok');
        assert.equal(aliceRes.written, true);

        // 4.2 CLI default (merge-same-author): Bob fails with EXISTING_REVISIONS
        const bobOut = path.join(testDir, 'bob_out.docx');
        const bobRes = await executeCli([
            'apply', inputPath,
            '--target', 'Term: 60 days',
            '--modified', 'Term: 120 days',
            '--author', 'Bob',
            '--output', bobOut
        ]);
        assert.equal(bobRes.status, 'error');
        assert.equal(bobRes.results[0].error?.code, 'EXISTING_REVISIONS');
        assert.equal(bobRes.written, false);

        // 4.3 CLI with --existing-revisions accept-all-first: Bob succeeds
        const bobAcceptOut = path.join(testDir, 'bob_accept_out.docx');
        const bobAcceptRes = await executeCli([
            'apply', inputPath,
            '--target', 'Term: 60 days',
            '--modified', 'Term: 120 days',
            '--author', 'Bob',
            '--existing-revisions', 'accept-all-first',
            '--output', bobAcceptOut
        ]);
        assert.equal(bobAcceptRes.status, 'ok');
        assert.equal(bobAcceptRes.written, true);

        // 4.4 CLI with --existing-revisions reject-input: Alice fails with EXISTING_REVISIONS
        const aliceRejectOut = path.join(testDir, 'alice_reject_out.docx');
        const aliceRejectRes = await executeCli([
            'apply', inputPath,
            '--target', 'Term: 60 days',
            '--modified', 'Term: 90 days',
            '--author', 'Alice',
            '--existing-revisions', 'reject-input',
            '--output', aliceRejectOut
        ]);
        assert.equal(aliceRejectRes.status, 'error');
        assert.equal(aliceRejectRes.results[0].error?.code, 'UNSAFE_REVISION_NESTING');
        assert.equal(aliceRejectRes.written, false);

        // 4.5 CLI preflight with --existing-revisions
        const pfAccept = await executeCli([
            'preflight', inputPath,
            '--target', 'Term: 60 days',
            '--modified', 'Term: 120 days',
            '--author', 'Bob',
            '--existing-revisions', 'accept-all-first'
        ]);
        assert.equal(pfAccept.valid, true);

        const pfReject = await executeCli([
            'preflight', inputPath,
            '--target', 'Term: 60 days',
            '--modified', 'Term: 120 days',
            '--author', 'Bob',
            '--existing-revisions', 'reject-input'
        ]);
        assert.equal(pfReject.valid, false);
        assert.equal(pfReject.results[0].error?.code, 'UNSAFE_REVISION_NESTING');

    } finally {
        await rm(testDir, { recursive: true, force: true });
    }
}

console.log('All existingRevisions mode matrix tests passed successfully!');
