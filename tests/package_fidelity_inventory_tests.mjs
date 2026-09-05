import assert from 'node:assert/strict';
import {
    inventoryPackage,
    comparePackageEntries,
    assertPackageFidelity,
    normalizeOpcEntryName
} from './helpers/package-fidelity.mjs';
import { zipDocx, unzipDocx } from '../node/zip-archive.js';

// 1. Path normalization
assert.equal(normalizeOpcEntryName('word/document.xml'), 'word/document.xml');
assert.equal(normalizeOpcEntryName('/word/document.xml'), 'word/document.xml');
assert.equal(normalizeOpcEntryName('word\\document.xml'), 'word/document.xml');
assert.equal(normalizeOpcEntryName('.\\word/document.xml'), 'word/document.xml');
assert.equal(normalizeOpcEntryName('./word//document.xml'), 'word/document.xml');
assert.throws(() => normalizeOpcEntryName(''), /Invalid empty ZIP entry name/);

// 2. Duplicate normalized path rejection
const duplicateEntries = new Map([
    ['word/document.xml', Buffer.from('<w:document/>')],
    ['word\\document.xml', Buffer.from('<w:document/>')]
]);
assert.throws(() => inventoryPackage(duplicateEntries), /Duplicate normalized ZIP entry path detected/);

// 3. Recompressing identical entries produces no content differences
const binaryMedia = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0xff, 0xaa]);
const originalEntries = new Map([
    ['[Content_Types].xml', Buffer.from('<Types/>')],
    ['_rels/.rels', Buffer.from('<Relationships/>')],
    ['word/document.xml', Buffer.from('<w:document><w:body><w:p/></w:body></w:document>')],
    ['word/media/image1.png', binaryMedia]
]);

const zipped = zipDocx(originalEntries);
const rezipped = zipDocx(unzipDocx(zipped));

const inventoryA = inventoryPackage(zipped);
const inventoryB = inventoryPackage(rezipped);

assert.equal(inventoryA.entries.size, 4);
assert.equal(inventoryB.entries.size, 4);

const diffRecompress = comparePackageEntries(inventoryA, inventoryB);
assert.equal(diffRecompress.identical, true);
assert.equal(diffRecompress.unexpectedChanged.length, 0);
assert.equal(diffRecompress.unexpectedAdded.length, 0);
assert.equal(diffRecompress.unexpectedRemoved.length, 0);
assert.equal(diffRecompress.unchanged.length, 4);
assert.doesNotThrow(() => assertPackageFidelity(inventoryA, inventoryB));

// 4. One-byte payload change is detected
const modifiedDocXml = Buffer.from('<w:document><w:body><w:p/></w:body></w:document>');
modifiedDocXml[modifiedDocXml.length - 1] = 0x20; // alter 1 byte
const alteredEntries = new Map(originalEntries);
alteredEntries.set('word/document.xml', modifiedDocXml);

const diffOneByte = comparePackageEntries(originalEntries, alteredEntries);
assert.equal(diffOneByte.identical, false);
assert.equal(diffOneByte.unexpectedChanged.length, 1);
assert.equal(diffOneByte.unexpectedChanged[0].name, 'word/document.xml');
assert.match(diffOneByte.report, /word\/document\.xml/);
assert.throws(() => assertPackageFidelity(originalEntries, alteredEntries), /PACKAGE_FIDELITY_VIOLATION/);

// When word/document.xml is explicitly allowed to change:
const allowedDiff = comparePackageEntries(originalEntries, alteredEntries, ['word/document.xml']);
assert.equal(allowedDiff.identical, true);
assert.equal(allowedDiff.allowedModifications.length, 1);
assert.equal(allowedDiff.allowedModifications[0].name, 'word/document.xml');
assert.doesNotThrow(() => assertPackageFidelity(originalEntries, alteredEntries, ['word/document.xml']));

// 5. Added and removed entries reported separately
const addedAndRemovedEntries = new Map(originalEntries);
addedAndRemovedEntries.delete('word/media/image1.png');
addedAndRemovedEntries.set('word/comments.xml', Buffer.from('<w:comments/>'));

const diffStructural = comparePackageEntries(originalEntries, addedAndRemovedEntries);
assert.equal(diffStructural.identical, false);
assert.equal(diffStructural.unexpectedRemoved.length, 1);
assert.equal(diffStructural.unexpectedRemoved[0].name, 'word/media/image1.png');
assert.equal(diffStructural.unexpectedAdded.length, 1);
assert.equal(diffStructural.unexpectedAdded[0].name, 'word/comments.xml');
assert.match(diffStructural.report, /Unexpected added entries/);
assert.match(diffStructural.report, /Unexpected removed entries/);

// Allowing additions/removals explicitly
const allowedStructural = comparePackageEntries(originalEntries, addedAndRemovedEntries, [
    'word/media/image1.png',
    'word/comments.xml'
]);
assert.equal(allowedStructural.identical, true);
assert.equal(allowedStructural.allowedModifications.length, 2);

// 6. Binary media is compared without text decoding
const binaryImageA = Buffer.from([0x00, 0x01, 0x02, 0xfe, 0xff]);
const binaryImageB = Buffer.from([0x00, 0x01, 0x02, 0xfe, 0xfd]);
const pkgBinaryA = new Map([['media.bin', binaryImageA]]);
const pkgBinaryB = new Map([['media.bin', binaryImageB]]);

const diffBinary = comparePackageEntries(pkgBinaryA, pkgBinaryB);
assert.equal(diffBinary.identical, false);
assert.equal(diffBinary.unexpectedChanged[0].name, 'media.bin');
assert.equal(diffBinary.unexpectedChanged[0].before.size, 5);
assert.equal(diffBinary.unexpectedChanged[0].after.size, 5);

console.log('package fidelity inventory tests passed');
