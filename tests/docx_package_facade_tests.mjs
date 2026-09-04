import assert from 'node:assert/strict';
import { buildZip } from '../scripts/lib/minimal-zip.mjs';
import { unzipEntries } from '../scripts/lib/zip-reader.mjs';
import { openDocx } from '../node/index.js';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const documentXml = `<w:document xmlns:w="${W}"><w:body><w:p><w:r><w:t>Hello world</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`;
const contentTypes = `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/></Types>`;
const rels = `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/></Relationships>`;
const comments = `<w:comments xmlns:w="${W}"><w:comment w:id="5000" w:author="Prior"><w:p><w:r><w:t>Old</w:t></w:r></w:p></w:comment></w:comments>`;
const untouched = Buffer.from([1,2,3,4]);
const input = buildZip([{name:'[Content_Types].xml',data:contentTypes},{name:'word/document.xml',data:documentXml},{name:'word/_rels/document.xml.rels',data:rels},{name:'word/comments.xml',data:comments},{name:'custom/data.bin',data:untouched}]);

const doc = openDocx(input);
assert.equal(doc.inspect().paragraphs[0].text, 'Hello world');
const applied = await doc.applyOperations([{ type:'comment', target:{ exactText:'Hello world' }, commentContent:'New note', author:'Agent' }], { atomic:true });
assert.equal(applied.written, true);
const outputEntries = unzipEntries(applied.toBuffer());
assert.deepEqual(outputEntries.get('custom/data.bin'), untouched);
assert.match(outputEntries.get('word/comments.xml').toString(), /w:id="5001"/);
assert.match(outputEntries.get('word/comments.xml').toString(), /w:author="Prior"/);
assert.ok(applied.artifactsChanged.includes('word/document.xml'));
assert.ok(Array.isArray(applied.validation.originalIssues));

const failedDoc = openDocx(input);
const failed = await failedDoc.applyOperations([{ type:'comment', target:{ exactText:'missing' }, commentContent:'Nope' }], { atomic:true });
assert.equal(failed.written, false);
assert.equal(failed.rolledBack, true);
assert.deepEqual(failed.toBuffer(), input);

const invalidPackage = buildZip([{name:'word/document.xml',data:documentXml},{name:'word/_rels/document.xml.rels',data:rels}]);
const invalidDoc = openDocx(invalidPackage);
const validationFailure = await invalidDoc.applyOperations([{ type:'replace', target:{ exactText:'Hello world' }, modified:'Changed' }], { author:'Agent', atomic:true });
assert.equal(validationFailure.written, false);
assert.equal(validationFailure.rolledBack, true);
assert.deepEqual(validationFailure.toBuffer(), invalidPackage);

const numberedTypes = contentTypes.replace('</Types>', '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>');
const numberedRels = rels.replace('</Relationships>', '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/></Relationships>');
const existingNumbering = `<w:numbering xmlns:w="${W}"><w:abstractNum w:abstractNumId="41"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum><w:num w:numId="41"><w:abstractNumId w:val="41"/></w:num></w:numbering>`;
const numberedInput = buildZip([{name:'[Content_Types].xml',data:numberedTypes},{name:'word/document.xml',data:documentXml},{name:'word/_rels/document.xml.rels',data:numberedRels},{name:'word/comments.xml',data:comments},{name:'word/numbering.xml',data:existingNumbering}]);
const numberedDoc = openDocx(numberedInput);
const listResult = await numberedDoc.applyOperations([{ type:'replace', target:{ exactText:'Hello world' }, modified:'1. Hello world' }], { author:'Agent', atomic:true });
assert.equal(listResult.written, true);
assert.match(unzipEntries(listResult.toBuffer()).get('word/numbering.xml').toString(), /abstractNumId="41"/);
console.log('docx package facade tests passed');
