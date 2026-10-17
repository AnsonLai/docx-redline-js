import './setup-xml-provider.mjs';

// compareDocumentInspections ignores host-regenerated identity (paraIds, rsids,
// revision ids/dates, run splits) and reports visible or review differences.
import assert from 'assert/strict';
import { compareDocumentInspections, inspectDocumentParts } from '../index.js';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"';
const doc = body => `<w:document ${W}><w:body>${body}<w:sectPr/></w:body></w:document>`;
const inspect = body => inspectDocumentParts({ documentXml: doc(body) });

const prepared = inspect('<w:p w14:paraId="1A2B3C4D"><w:r><w:t xml:space="preserve">Hello </w:t></w:r>'
    + '<w:ins w:id="1" w:author="Me" w:date="2026-01-01T00:00:00Z"><w:r><w:t>new world</w:t></w:r></w:ins></w:p>');
// Same content as Word writes it back: new paraId, rsids, revision id/date, and split runs.
const committed = inspect('<w:p w14:paraId="7F7F7F7F" w:rsidR="00AB12CD"><w:r w:rsidR="00AB12CD"><w:t xml:space="preserve">Hel</w:t></w:r><w:r><w:t xml:space="preserve">lo </w:t></w:r>'
    + '<w:ins w:id="99" w:author="Me" w:date="2026-10-17T08:00:00Z"><w:r><w:t xml:space="preserve">new </w:t></w:r></w:ins>'
    + '<w:ins w:id="100" w:author="Me" w:date="2026-10-17T08:00:00Z"><w:r><w:t>world</w:t></w:r></w:ins></w:p>');
assert.deepEqual(compareDocumentInspections(prepared, committed), { matches: true, differences: [] });

const untracked = inspect('<w:p><w:r><w:t>Hello new world</w:t></w:r></w:p>');
const lostReview = compareDocumentInspections(prepared, untracked);
assert.equal(lostReview.matches, false, 'losing the tracked insertion is a difference');
assert.ok(lostReview.differences.some(item => item.path === 'paragraphs[0].review'));

const extra = compareDocumentInspections(prepared, inspect('<w:p><w:r><w:t>Hello </w:t></w:r><w:ins w:id="1" w:author="Me" w:date="2026-01-01T00:00:00Z"><w:r><w:t>new world</w:t></w:r></w:ins></w:p><w:p/>'));
assert.equal(extra.matches, false);
assert.deepEqual(extra.differences[0], { path: 'paragraphs.length', expected: 1, actual: 2 });

console.log('inspection_comparison_tests: all passed');
