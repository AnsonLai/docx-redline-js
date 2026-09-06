import assert from 'node:assert/strict';
import { DOMParser as BaseDOMParser, XMLSerializer as BaseXMLSerializer } from '@xmldom/xmldom';
import { configureXmlProvider, parseOoxmlSafe } from '../adapters/xml-adapter.js';
import { buildTargetReferenceSnapshot, getParagraphText } from '../core/paragraph-targeting.js';
import { validateRedlineOoxml } from '../core/redline-validation.js';
import {
    acceptTrackedChangesInOoxml,
    rejectTrackedChangesInOoxml
} from '../services/revision-comment-management.js';
import {
    applyOperationToDocumentXml,
    applyOperationsToDocumentXml,
    orderOperationsForStableTargets
} from '../services/standalone-operation-runner.js';

const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const NS_W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const SOURCE = `<w:document xmlns:w="${NS_W}" xmlns:w14="${NS_W14}"><w:body>
<w:p w14:paraId="AA000001"><w:r><w:t>Opening clause for review.</w:t></w:r></w:p>
<w:p w14:paraId="AA000002"><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="7"/></w:numPr></w:pPr><w:r><w:t>First established list item</w:t></w:r></w:p>
<w:tbl><w:tr><w:tc><w:p w14:paraId="AA000003"><w:r><w:t>Old table cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>
<w:p w14:paraId="AA000004"><w:r><w:t>Closing clause remains visible.</w:t></w:r></w:p>
<w:sectPr/></w:body></w:document>`;

const counters = {
    exactSourceParses: 0,
    fullDocumentSerializations: 0,
    sessionParses: 0,
    sessionSerializations: 0
};

class CountingDOMParser {
    constructor(options) {
        this.parser = new BaseDOMParser(options);
    }

    parseFromString(source, contentType) {
        if (source === SOURCE) counters.exactSourceParses += 1;
        return this.parser.parseFromString(source, contentType);
    }
}

class CountingXMLSerializer {
    constructor() {
        this.serializer = new BaseXMLSerializer();
    }

    serializeToString(node) {
        const output = this.serializer.serializeToString(node);
        const rootName = String(node?.documentElement?.localName || node?.documentElement?.nodeName || '');
        const paragraphCount = node?.getElementsByTagNameNS?.(NS_W, 'p')?.length || 0;
        if ((rootName === 'document' || rootName === 'w:document') && paragraphCount >= 4) {
            counters.fullDocumentSerializations += 1;
        }
        return output;
    }
}

configureXmlProvider({ DOMParser: CountingDOMParser, XMLSerializer: CountingXMLSerializer });

const operations = [
    {
        type: 'comment',
        target: { paragraphId: 'AA000001', exactText: 'Opening clause for review.' },
        textToComment: 'Opening clause',
        commentContent: 'Confirm this opening.',
        author: 'Reviewer'
    },
    {
        type: 'replace',
        target: { paragraphId: 'AA000002', exactText: 'First established list item' },
        modified: 'First established list item\nA carefully added second list item for review',
        author: 'List Editor'
    },
    {
        type: 'replace',
        target: { paragraphId: 'AA000003', exactText: 'Old table cell' },
        modified: 'New table cell',
        author: 'Table Editor'
    },
    {
        type: 'highlight',
        target: { paragraphId: 'AA000004', exactText: 'Closing clause remains visible.' },
        textToHighlight: 'remains visible',
        color: 'yellow',
        author: 'Reviewer'
    }
];

const batch = await applyOperationsToDocumentXml(SOURCE, operations, 'Fallback', null, {
    generateRedlines: true,
    _sessionInstrumentation: {
        onDocumentParse: () => { counters.sessionParses += 1; },
        onDocumentSerialize: () => { counters.sessionSerializations += 1; }
    }
});

assert.equal(batch.hasChanges, true);
assert.equal(batch.rolledBack, undefined);
assert.equal(counters.exactSourceParses, 1, 'the complete source must be parsed once');
assert.equal(counters.fullDocumentSerializations, 1, 'the live document must be serialized once');
assert.equal(counters.sessionParses, 1);
assert.equal(counters.sessionSerializations, 1);
assert.deepEqual(batch.executionOrder, [1, 2, 3, 4]);
assert.deepEqual(batch.authorsUsed, ['Reviewer', 'List Editor', 'Table Editor']);
assert.equal(batch.results.every(result => result.status === 'applied'), true);
assert.match(batch.documentXml, /w:commentRangeStart/);
assert.match(batch.documentXml, /w:highlight[^>]+w:val="yellow"/);
assert.match(batch.documentXml, /<w:tbl>/);
assert.match(batch.documentXml, /<w:numId w:val="7"/);

function acceptedText(xml) {
    const accepted = acceptTrackedChangesInOoxml(xml, { allAuthors: true });
    assert.equal(accepted.error, undefined);
    const parsed = parseOoxmlSafe(accepted.oxml, 'application/xml');
    assert.equal(parsed.error, null);
    return Array.from(parsed.doc.getElementsByTagNameNS(NS_W, 'p')).map(getParagraphText);
}

function rejectedText(xml) {
    const rejected = rejectTrackedChangesInOoxml(xml, { allAuthors: true });
    assert.equal(rejected.error, undefined);
    const parsed = parseOoxmlSafe(rejected.oxml, 'application/xml');
    assert.equal(parsed.error, null);
    return Array.from(parsed.doc.getElementsByTagNameNS(NS_W, 'p')).map(getParagraphText);
}

const sourceDoc = parseOoxmlSafe(SOURCE, 'application/xml').doc;
const sequentialContext = { targetRefSnapshot: buildTargetReferenceSnapshot(sourceDoc) };
let sequentialXml = SOURCE;
let sequentialCommentsXml = null;
for (const operation of orderOperationsForStableTargets(operations)) {
    const result = await applyOperationToDocumentXml(
        sequentialXml,
        operation,
        'Fallback',
        sequentialContext,
        { generateRedlines: true }
    );
    assert.notEqual(result.status, 'error');
    sequentialXml = result.documentXml;
    if (result.commentsXml) sequentialCommentsXml = result.commentsXml;
}

const batchAcceptedText = acceptedText(batch.documentXml);
const batchRejectedText = rejectedText(batch.documentXml);
assert.deepEqual(batchAcceptedText, acceptedText(sequentialXml), 'accepted view must match legacy sequential execution');
assert.deepEqual(batchRejectedText, rejectedText(sequentialXml), 'rejected view must match legacy sequential execution');
assert.deepEqual(batchAcceptedText, [
    'Opening clause for review.',
    'First established list item',
    'A carefully added second list item for review',
    'New table cell',
    'Closing clause remains visible.'
]);
assert.deepEqual(batchRejectedText, [
    'Opening clause for review.',
    'First established list item',
    'Old table cell',
    'Closing clause remains visible.'
]);
assert.deepEqual(validateRedlineOoxml(batch.documentXml), { valid: true, issues: [] });
assert.equal((batch.documentXml.match(/<w:ins\b/g) || []).length, (sequentialXml.match(/<w:ins\b/g) || []).length);
assert.equal((batch.documentXml.match(/<w:del\b/g) || []).length, (sequentialXml.match(/<w:del\b/g) || []).length);
assert.equal((batch.documentXml.match(/w:commentRangeStart/g) || []).length, 1);
assert.equal((sequentialXml.match(/w:commentRangeStart/g) || []).length, 1);
assert.match(batch.commentsXml, /Confirm this opening\./);
assert.match(sequentialCommentsXml, /Confirm this opening\./);

const beforeNoOpSerializations = counters.sessionSerializations;
const noOp = await applyOperationsToDocumentXml(SOURCE, [{
    type: 'replace',
    target: { paragraphId: 'AA000004' },
    modified: 'Closing clause remains visible.'
}], 'No-op', null, {
    _sessionInstrumentation: {
        onDocumentSerialize: () => { counters.sessionSerializations += 1; }
    }
});
assert.equal(noOp.hasChanges, false);
assert.equal(noOp.documentXml, SOURCE, 'a no-op batch must preserve exact source bytes');
assert.equal(counters.sessionSerializations, beforeNoOpSerializations, 'a no-op batch must not serialize the document');

const beforeRollbackSerializations = counters.sessionSerializations;
const rollback = await applyOperationsToDocumentXml(SOURCE, [
    {
        type: 'replace',
        target: { paragraphId: 'AA000004' },
        modified: 'This edit must be rolled back.'
    },
    {
        type: 'replace',
        target: { exactText: 'Missing target' },
        modified: 'Never applied.'
    }
], 'Rollback', null, {
    atomic: true,
    _sessionInstrumentation: {
        onDocumentSerialize: () => { counters.sessionSerializations += 1; }
    }
});
assert.equal(rollback.rolledBack, true);
assert.equal(rollback.documentXml, SOURCE, 'atomic rollback must preserve exact source bytes');
assert.equal(counters.sessionSerializations, beforeRollbackSerializations, 'rolled-back batches must not serialize the document');

console.log('PASS: performance Phase 1 live session preserves redline semantics');
