import assert from 'assert/strict';

import './setup-xml-provider.mjs';
import {
    acceptTrackedChangesInOoxml,
    applyRedlineToOxml,
    rejectTrackedChangesInOoxml,
    validateRedlineOoxml
} from '../index.js';
import { parseOoxmlSafe } from '../adapters/xml-adapter.js';
import { extractCanonicalParagraphText } from '../core/paragraph-text.js';
import { applyOperationsToDocumentXml } from '../services/standalone-operation-runner.js';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const DATE = '2026-09-08T18:00:00Z';
const AUTHOR = 'Boundary Reviewer';
const nbsp = '\u00a0';

function escapeXml(text) {
    return String(text)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;');
}

function run(text, rPr = '') {
    const preserve = /^\s|\s$/.test(text) ? ' xml:space="preserve"' : '';
    return `<w:r>${rPr}<w:t${preserve}>${escapeXml(text)}</w:t></w:r>`;
}

function hyperlink(id, text) {
    return `<w:hyperlink r:id="${id}">${run(text, '<w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr>')}</w:hyperlink>`;
}

function insertion(id, author, content) {
    return `<w:ins w:id="${id}" w:author="${author}" w:date="${DATE}">${content}</w:ins>`;
}

function paragraph(content) {
    return `<w:p xmlns:w="${W}" xmlns:r="${R}">${content}</w:p>`;
}

function parse(xml) {
    const parsed = parseOoxmlSafe(xml, 'application/xml');
    assert.equal(parsed.error, null);
    return parsed.doc;
}

function acceptedText(xml) {
    const paragraphNode = parse(xml).getElementsByTagNameNS(W, 'p')[0];
    return extractCanonicalParagraphText(paragraphNode);
}

function hyperlinkIds(xml) {
    return Array.from(parse(xml).getElementsByTagNameNS(W, 'hyperlink'))
        .map(node => node.getAttribute('r:id'));
}

const cases = [
    {
        name: 'replacement-inside-single-run',
        source: paragraph(run('Alpha old term omega')),
        original: 'Alpha old term omega',
        modified: 'Alpha replacement term omega'
    },
    {
        name: 'replacement-at-run-start',
        source: paragraph(run('Prefix') + run(' old suffix', '<w:rPr><w:b/></w:rPr>')),
        original: 'Prefix old suffix',
        modified: 'Prefix, revised suffix'
    },
    {
        name: 'replacement-at-run-end',
        source: paragraph(run('Prefix old ') + run('suffix', '<w:rPr><w:i/></w:rPr>')),
        original: 'Prefix old suffix',
        modified: 'Prefix revised—suffix'
    },
    {
        name: 'privacy-qualifier-after-hyperlink',
        source: paragraph(
            run('Privacy Policy located at')
            + run(nbsp)
            + hyperlink('rIdPrivacy', 'example.com/privacy')
            + run(`${nbsp}(the “`)
            + run('Privacy Policy', '<w:rPr><w:b/><w:u w:val="single"/></w:rPr>')
            + run('”).')
        ),
        original: `Privacy Policy located at${nbsp}example.com/privacy${nbsp}(the “Privacy Policy”).`,
        modified: `Privacy Policy located at${nbsp}example.com/privacy, as in effect on execution${nbsp}(the “Privacy Policy”).`,
        links: ['rIdPrivacy'],
        runner: true
    },
    {
        name: 'qualifier-before-hyperlink',
        source: paragraph(run('See the policy ') + hyperlink('rIdPolicy', 'here') + run(' for details.')),
        original: 'See the policy here for details.',
        modified: 'See the current policy: here for details.',
        links: ['rIdPolicy'],
        runner: true
    },
    {
        name: 'whole-spacer-run-between-hyperlinks',
        source: paragraph(hyperlink('rIdA', 'Alpha') + run(' ') + hyperlink('rIdB', 'Beta')),
        original: 'Alpha Beta',
        modified: 'Alpha and Beta',
        links: ['rIdA', 'rIdB'],
        runner: true
    },
    {
        name: 'formatted-run-after-hyperlink',
        source: paragraph(hyperlink('rIdClause', 'Section 2.8') + run(' applies', '<w:rPr><w:b/></w:rPr>')),
        original: 'Section 2.8 applies',
        modified: 'Section 2.8, as amended, applies',
        links: ['rIdClause']
    },
    {
        name: 'bookmark-adjacent-replacement',
        source: paragraph(run('Defined term ') + '<w:bookmarkStart w:id="4" w:name="Term"/>' + run('Value') + '<w:bookmarkEnd w:id="4"/>'),
        original: 'Defined term Value',
        modified: 'Defined term: Value'
    },
    {
        name: 'comment-boundary-adjacent-replacement',
        source: paragraph(run('Clause ') + '<w:commentRangeStart w:id="7"/>' + run('text') + '<w:commentRangeEnd w:id="7"/>' + run(' continues')),
        original: 'Clause text continues',
        modified: 'Clause—text continues'
    },
    {
        name: 'two-replacements-across-distinct-runs',
        source: paragraph(run('The old ') + hyperlink('rIdTerms', 'terms') + run(' apply on the old date.')),
        original: 'The old terms apply on the old date.',
        modified: 'The revised terms apply on the execution date.',
        links: ['rIdTerms'],
        runner: true
    },
    {
        name: 'replacement-after-foreign-insertion',
        source: paragraph(insertion(80, 'Prior Reviewer', run('Inserted clause')) + run(' applies.')),
        original: 'Inserted clause applies.',
        modified: 'Inserted clause, as revised, applies.'
    },
    {
        name: 'replacement-inside-foreign-insertion',
        source: paragraph(insertion(81, 'Prior Reviewer', run('Inserted old clause'))),
        original: 'Inserted old clause',
        modified: 'Inserted revised clause'
    }
];

for (const testCase of cases) {
    const result = await applyRedlineToOxml(testCase.source, testCase.original, testCase.modified, {
        author: AUTHOR,
        existingRevisions: 'slice-cross-author',
        pairReplacements: true,
        structuredContent: false
    });
    assert.equal(result.status, 'ok', `${testCase.name}: ${JSON.stringify(result.error)}`);
    assert.equal(result.hasChanges, true, `${testCase.name}: expected a change`);
    assert.equal(acceptedText(result.oxml), testCase.modified, `${testCase.name}: current view`);
    assert.deepEqual(hyperlinkIds(result.oxml), testCase.links || [], `${testCase.name}: hyperlinks`);

    const validation = validateRedlineOoxml(result.oxml);
    assert.equal(validation.valid, true, `${testCase.name}: ${JSON.stringify(validation.issues)}`);

    const accepted = acceptTrackedChangesInOoxml(result.oxml, { allAuthors: true });
    assert.equal(acceptedText(accepted.oxml), testCase.modified, `${testCase.name}: Accept All`);

    const rejected = rejectTrackedChangesInOoxml(result.oxml, { author: AUTHOR });
    assert.equal(acceptedText(rejected.oxml), testCase.original, `${testCase.name}: Reject Reviewer`);

    const revisions = [
        ...Array.from(parse(result.oxml).getElementsByTagNameNS(W, 'ins')),
        ...Array.from(parse(result.oxml).getElementsByTagNameNS(W, 'del'))
    ];
    const ids = revisions.map(node => node.getAttributeNS(W, 'id') || node.getAttribute('w:id'));
    assert.equal(new Set(ids).size, ids.length, `${testCase.name}: revision IDs must be unique`);

    if (testCase.runner) {
        const documentXml = `<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>${testCase.source}<w:sectPr/></w:body></w:document>`;
        const batch = await applyOperationsToDocumentXml(documentXml, [{
            type: 'redline',
            target: { exactText: testCase.original },
            modified: testCase.modified,
            author: AUTHOR,
            existingRevisions: 'slice-cross-author',
            structuredContent: false
        }], AUTHOR, null, { atomic: true, strictTargets: true });
        assert.equal(batch.status, 'ok', `${testCase.name} runner: ${JSON.stringify(batch.error)}`);
        assert.equal(batch.results[0]?.status, 'applied', `${testCase.name}: runner operation`);
        assert.equal(acceptedText(batch.documentXml), testCase.modified, `${testCase.name}: runner current view`);
        assert.deepEqual(hyperlinkIds(batch.documentXml), testCase.links || [], `${testCase.name}: runner hyperlinks`);
    }
}

console.log(`PASS: ${cases.length} cross-author replacement-anchor scenarios`);
