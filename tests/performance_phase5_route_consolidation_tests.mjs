import './setup-xml-provider.mjs';

import assert from 'node:assert/strict';
import * as rootApi from '../index.js';
import { applyRedlineToOxml } from '../engine/oxml-engine.js';
import { ReconciliationPipeline } from '../pipeline/pipeline.js';
import { ingestOoxml } from '../pipeline/ingestion.js';
import { acceptTrackedChangesInOoxml, rejectTrackedChangesInOoxml } from '../services/revision-comment-management.js';
import { createRouteFrequencyCollector, RECONCILIATION_CAPABILITY_MATRIX } from '../engine/route-selection.js';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const source = `<w:p xmlns:w="${W}"><w:r><w:rPr><w:b/></w:rPr><w:t>Original clause</w:t></w:r></w:p>`;
const target = '1. First clause\n2. Second clause';
const collector = createRouteFrequencyCollector();
const direct = await applyRedlineToOxml(source, 'Original clause', target, {
    author:'Agent', _routeInstrumentation: collector
});
const legacy = await new ReconciliationPipeline({ author:'Agent', generateRedlines:true }).execute(source, target);
const legacyWrapped = rootApi.wrapInDocumentFragment(legacy.ooxml, {
    includeNumbering: legacy.includeNumbering,
    numberingXml: legacy.numberingXml
});

assert.equal(direct.hasChanges, true);
assert.equal(legacy.isValid, true);
assert.deepEqual(collector.snapshot(), { listDirect: 1 });
assert.ok(RECONCILIATION_CAPABILITY_MATRIX.listDirect.numbering);

const acceptedDirect = acceptTrackedChangesInOoxml(direct.oxml, { allAuthors:true });
const acceptedLegacy = acceptTrackedChangesInOoxml(legacyWrapped, { allAuthors:true });
const rejectedDirect = rejectTrackedChangesInOoxml(direct.oxml, { allAuthors:true });
const rejectedLegacy = rejectTrackedChangesInOoxml(legacyWrapped, { allAuthors:true });
assert.equal(ingestOoxml(acceptedDirect.oxml).acceptedText, ingestOoxml(acceptedLegacy.oxml).acceptedText);
assert.equal(ingestOoxml(rejectedDirect.oxml).acceptedText, ingestOoxml(rejectedLegacy.oxml).acceptedText);
assert.match(direct.oxml, /numbering\.xml/);
assert.match(direct.oxml, /w:numPr/);

const multiParagraph = `<w:document xmlns:w="${W}"><w:body><w:p><w:r><w:t>1. One</w:t></w:r></w:p><w:p><w:r><w:t>2. Two</w:t></w:r></w:p></w:body></w:document>`;
const compatibilityCollector = createRouteFrequencyCollector();
await applyRedlineToOxml(multiParagraph, '1. One\n2. Two', '1. One changed\n2. Two', {
    author:'Agent', _routeInstrumentation: compatibilityCollector
});
assert.deepEqual(compatibilityCollector.snapshot(), { listCompatibilityPipeline: 1 });

assert.equal(rootApi.ReconciliationPipeline, ReconciliationPipeline);
assert.equal(typeof rootApi.serializeToOoxml, 'function');
assert.equal(typeof rootApi.wrapInDocumentFragment, 'function');

console.log('PASS: Phase 5 route consolidation and compatibility window');
