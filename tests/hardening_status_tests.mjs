import './setup-xml-provider.mjs';

import assert from 'assert/strict';

import { applyRedlineToOxml } from '../index.js';
import { elementsByLocalName, parseXmlFragment } from './helpers/ooxml-assertions.mjs';

const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

function paragraph(inner) {
    return `<w:p xmlns:w="${NS_W}">${inner}</w:p>`;
}

async function testMalformedXmlReportsParseError() {
    const result = await applyRedlineToOxml('<w:p><w:r>', 'x', 'y', {
        author: 'Hardening'
    });

    assert.equal(result.hasChanges, false);
    assert.equal(result.status, 'error');
    assert.equal(result.error?.code, 'PARSE_ERROR');
}

async function testMissingTargetReportsTargetNotFound() {
    const source = paragraph('<w:r><w:t>Hello</w:t></w:r>');
    const result = await applyRedlineToOxml(source, 'Missing', 'Changed', {
        author: 'Hardening'
    });

    assert.equal(result.hasChanges, false);
    assert.equal(result.status, 'error');
    assert.equal(result.error?.code, 'TARGET_NOT_FOUND');
    assert.equal(result.oxml, source);
}

async function testNoOpStatus() {
    const source = paragraph('<w:r><w:t>Hello</w:t></w:r>');
    const result = await applyRedlineToOxml(source, 'Hello', 'Hello', {
        author: 'Hardening'
    });

    assert.equal(result.hasChanges, false);
    assert.equal(result.status, 'no-op');
}

async function testGeneratedRevisionIdsSeedAboveExistingIds() {
    const source = paragraph([
        '<w:ins w:id="5000" w:author="Prior" w:date="2026-01-01T00:00:00Z"><w:r><w:t>Old</w:t></w:r></w:ins>',
        '<w:r><w:t> text</w:t></w:r>'
    ].join(''));
    const result = await applyRedlineToOxml(source, 'Old text', 'New text', {
        author: 'Hardening',
        existingRevisions: 'accept-all-first'
    });

    assert.equal(result.hasChanges, true);
    assert.equal(result.status, 'ok');
    const doc = parseXmlFragment(result.oxml);
    const ids = elementsByLocalName(doc, 'ins')
        .concat(elementsByLocalName(doc, 'del'))
        .map(node => Number.parseInt(node.getAttribute('w:id') || node.getAttribute('id') || '', 10))
        .filter(Number.isFinite);
    assert.ok(ids.length > 0, 'Expected generated revision ids');
    assert.ok(ids.every(id => id > 5000), `Expected generated ids above 5000, got ${ids.join(', ')}`);
}

await testMalformedXmlReportsParseError();
await testMissingTargetReportsTargetNotFound();
await testNoOpStatus();
await testGeneratedRevisionIdsSeedAboveExistingIds();

console.log('PASS: hardening status tests');
