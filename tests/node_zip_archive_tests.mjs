import assert from 'node:assert/strict';
import { buildZip } from '../scripts/lib/minimal-zip.mjs';
import { unzipDocx, zipDocx } from '../node/zip-archive.js';

const entries = new Map([['a.txt', Buffer.from('alpha')], ['nested/b.bin', Buffer.from([0,1,2,255])], ['empty', Buffer.alloc(0)]]);
const roundTrip = unzipDocx(zipDocx(entries));
assert.deepEqual([...roundTrip.keys()], [...entries.keys()]);
for (const [name, data] of entries) assert.deepEqual(roundTrip.get(name), data);

assert.throws(() => unzipDocx(Buffer.from('not a zip')), /ZIP end record/);
const encrypted = buildZip([{ name:'secret.txt', data:'secret' }]);
const central = encrypted.indexOf(Buffer.from([0x50,0x4b,0x01,0x02]));
encrypted.writeUInt16LE(encrypted.readUInt16LE(central + 8) | 1, central + 8);
assert.throws(() => unzipDocx(encrypted), /encrypted ZIP entries/);

const unsupported = buildZip([{ name:'odd.bin', data:'payload' }]);
const unsupportedCentral = unsupported.indexOf(Buffer.from([0x50,0x4b,0x01,0x02]));
unsupported.writeUInt16LE(99, unsupportedCentral + 10);
assert.throws(() => unzipDocx(unsupported), /compression method: 99/);
console.log('node zip archive tests passed');
