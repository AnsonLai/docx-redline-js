import assert from 'node:assert/strict';

import { configureXmlProvider } from '../adapters/xml-adapter.js';
import { openDocx } from '../index.js';
import { buildMatrixFixtures } from './helpers/reject-roundtrip-fixtures.mjs';
import { generateOperationsForFixture, makeRng } from './helpers/operation-generators.mjs';

/**
 * E4 (pre-0.9.0 plan): the list conversion paths must not make the XML parser emit diagnostics (xmldom hierarchy
 * errors, malformed-markup warnings) for the OOXML they generate, with either parser provider a consumer can end up
 * with. No jsdom is imported on purpose: the host has no DOMParser, so "fallback" is the library's bundled pure-JS
 * parser, and "xmldom" is the same package configured explicitly through configureXmlProvider.
 *
 * Capture is on console.warn / console.error (the adapter logger's channels); the library's own info logging on
 * console.log is not a diagnostic.
 */

const LIST_KINDS = new Set([
    'plain-to-list', 'paragraph-format-numbering', 'list-format-convert-plain', 'list-format-change-format',
    'list-format-join', 'list-format-redline-compose', 'list-sequence-redline'
]);
const FIXTURES = new Set(['nested-lists-source', 'typed-headers', 'adjacent-lists', 'empty-list-items']);

async function runProvider(label, configure) {
    const xmldom = await import('@xmldom/xmldom');
    if (configure === 'xmldom') configureXmlProvider({ DOMParser: xmldom.DOMParser, XMLSerializer: xmldom.XMLSerializer });
    else configureXmlProvider({ DOMParser: null, XMLSerializer: null });
    assert.equal(typeof globalThis.DOMParser, 'undefined', 'no host DOMParser leaks into the provider test');

    const captured = [];
    const originalWarn = console.warn;
    const originalError = console.error;
    const originalLog = console.log;
    console.warn = (...args) => captured.push(['warn', args.join(' ')]);
    console.error = (...args) => captured.push(['error', args.join(' ')]);
    console.log = () => {};
    let applied = 0;
    try {
        for (const fixture of buildMatrixFixtures().filter(item => FIXTURES.has(item.name))) {
            const inspection = openDocx(fixture.bytes).inspect({ existingRevisions: undefined });
            const rng = makeRng(20261011);
            // Generate plenty of cases, keep only the list-conversion kinds.
            const { cases } = generateOperationsForFixture(rng, inspection.paragraphs, 120);
            for (const generated of cases.filter(item => LIST_KINDS.has(item.kind))) {
                const operations = Array.isArray(generated.operation) ? generated.operation : [generated.operation];
                const result = await openDocx(fixture.bytes).applyOperations(operations, { author: 'Diagnostics Author', ...fixture.options });
                if (result.status !== 'error' && result.hasChanges) applied++;
            }
        }
    } finally {
        console.warn = originalWarn;
        console.error = originalError;
        console.log = originalLog;
    }
    assert.ok(applied >= 10, `${label}: exercised the list paths (${applied} applied)`);
    assert.deepEqual(captured, [], `${label}: parser diagnostics for generated list OOXML:\n${captured.map(([level, text]) => `${level}: ${text}`).join('\n')}`);
    return applied;
}

const fallback = await runProvider('fallback', 'fallback');
const xmldom = await runProvider('xmldom', 'xmldom');
console.log(`PASS: no parser diagnostics for generated list OOXML (fallback ${fallback} applied, xmldom ${xmldom} applied)`);
