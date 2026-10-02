import assert from 'node:assert/strict';

import './setup-xml-provider.mjs';
import { inspectDocumentParts, acceptTrackedChangesInOoxml, rejectTrackedChangesInOoxml } from '../index.js';
import { applyOperationsToDocumentXml } from '../services/standalone-operation-runner.js';
import { parseOoxmlSafe } from '../adapters/xml-adapter.js';

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

console.log('soft_break_localized_replacements_tests: all passed');
