/**
 * Package part registry contract.
 *
 * Every registry entry marked `verified: true` must have been observed, with identical content type
 * and relationship type, in a document saved by real Word (tests/fixtures/word-authored/*.docx,
 * regenerate with the scripts/generate-word-*-fixture.ps1 scripts). A `verified` flag that no fixture
 * backs up fails here, so the flag cannot be claimed on memory alone.
 */
import assert from 'node:assert/strict';
import { readdirSync, readFileSync } from 'node:fs';
import { unzipEntries } from '../scripts/lib/zip-reader.mjs';
import { correctedContentTypeFor, findPartSpecByPartName, getPartSpec, listPartSpecs } from '../services/package-parts.js';

const fixtureDir = new URL('./fixtures/word-authored/', import.meta.url);
const fixtures = readdirSync(fixtureDir).filter(name => name.endsWith('.docx')).sort();
assert.ok(fixtures.length >= 2, 'expected the Word-authored fixtures to be present');

const attrs = tag => Object.fromEntries([...tag.matchAll(/([\w:]+)="([^"]*)"/g)].map(m => [m[1], m[2]]));
const observedContentTypes = new Map(); // spec id -> Set of content types
const observedRelTypes = new Map();     // spec id -> Set of relationship types
const note = (map, id, value) => { if (!map.has(id)) map.set(id, new Set()); map.get(id).add(value); };

for (const fixture of fixtures) {
    const entries = unzipEntries(readFileSync(new URL(fixture, fixtureDir)));
    const read = name => entries.get(name)?.toString('utf8') || '';

    for (const tag of read('[Content_Types].xml').match(/<Override\b[^>]*>/g) || []) {
        const { PartName, ContentType } = attrs(tag);
        const spec = findPartSpecByPartName(PartName);
        if (spec) note(observedContentTypes, spec.id, ContentType);
    }
    for (const tag of read('word/_rels/document.xml.rels').match(/<Relationship\b[^>]*>/g) || []) {
        const { Type, Target } = attrs(tag);
        const spec = findPartSpecByPartName(`word/${Target}`);
        if (spec) note(observedRelTypes, spec.id, Type);
    }
    for (const tag of read('_rels/.rels').match(/<Relationship\b[^>]*>/g) || []) {
        const { Type, Target } = attrs(tag);
        const spec = findPartSpecByPartName(Target);
        if (spec) note(observedRelTypes, spec.id, Type);
    }
}

for (const spec of listPartSpecs()) {
    const contentTypes = observedContentTypes.get(spec.id);
    const relTypes = observedRelTypes.get(spec.id);
    if (spec.verified) {
        assert.ok(contentTypes, `'${spec.id}' is marked verified but no Word fixture contains it`);
        assert.deepEqual([...contentTypes], [spec.contentType], `'${spec.id}' content type differs from what Word wrote`);
        assert.ok(relTypes, `'${spec.id}' is marked verified but no Word fixture has its relationship`);
        assert.deepEqual([...relTypes], [spec.relType], `'${spec.id}' relationship type differs from what Word wrote`);
    } else if (contentTypes) {
        // An unverified entry that a fixture now covers must either match or be promoted to verified.
        assert.deepEqual([...contentTypes], [spec.contentType], `'${spec.id}' is unverified and disagrees with Word; fix it`);
    }
    assert.ok(!spec.legacyContentTypes.includes(spec.contentType), `'${spec.id}' lists its own correct type as legacy`);
}

// Lookup behavior.
assert.equal(findPartSpecByPartName('/word/commentsExtended.xml').id, 'commentsExtended');
assert.equal(findPartSpecByPartName('word/COMMENTSEXTENDED.XML').id, 'commentsExtended');
assert.equal(findPartSpecByPartName('/word/header12.xml').id, 'header');
assert.equal(findPartSpecByPartName('/word/footer3.xml').id, 'footer');
assert.equal(findPartSpecByPartName('/word/headers.xml'), null);
assert.equal(findPartSpecByPartName('/word/theme/theme1.xml'), null);
assert.throws(() => getPartSpec('nope'), /Unknown package part/);

// Repair only rewrites values registered as known-wrong; it never "corrects" anything it has no evidence about.
const badExtended = 'application/vnd.ms-word.commentsExtended+xml';
assert.equal(correctedContentTypeFor('/word/commentsExtended.xml', badExtended), getPartSpec('commentsExtended').contentType);
assert.equal(correctedContentTypeFor('/word/commentsExtended.xml', getPartSpec('commentsExtended').contentType), null);
assert.equal(correctedContentTypeFor('/word/document.xml', 'application/vnd.ms-word.document.macroEnabled.main+xml'), null);
assert.equal(correctedContentTypeFor('/word/theme/theme1.xml', badExtended), null);

console.log('PASS: package part registry');
