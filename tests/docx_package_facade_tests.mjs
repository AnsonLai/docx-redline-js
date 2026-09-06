import assert from 'node:assert/strict';
import { buildZip } from '../scripts/lib/minimal-zip.mjs';
import { unzipEntries } from '../scripts/lib/zip-reader.mjs';
import { openDocx } from '../node/index.js';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const documentXml = `<w:document xmlns:w="${W}"><w:body><w:p><w:r><w:t>Hello world</w:t></w:r></w:p><w:p><w:commentRangeStart w:id="5000"/><w:r><w:t>Prior anchor</w:t></w:r><w:commentRangeEnd w:id="5000"/><w:r><w:commentReference w:id="5000"/></w:r></w:p><w:sectPr/></w:body></w:document>`;
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

const failedAnchorDoc = openDocx(input);
const failedAnchor = await failedAnchorDoc.applyOperations([
    { type:'replace', target:{ exactText:'Hello world' }, modified:'Changed', generateRedlines:false },
    { type:'comment', target:{ exactText:'Hello world' }, textToComment:'missing anchor', commentContent:'Nope' }
], { author:'Agent', atomic:true });
assert.equal(failedAnchor.status, 'error');
assert.equal(failedAnchor.written, false);
assert.equal(failedAnchor.rolledBack, true);
assert.equal(failedAnchor.results[1].error.code, 'ANCHOR_NOT_FOUND');
assert.deepEqual(failedAnchor.toBuffer(), input);

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

const legacyExtendedContentType = 'application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml';
const extendedTypes = contentTypes.replace('</Types>', `<Override PartName="/word/commentsExtended.xml" ContentType="${legacyExtendedContentType}"/></Types>`);
const extendedRels = rels.replace('</Relationships>', '<Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2011/relationships/commentsExtended" Target="commentsExtended.xml"/></Relationships>');
const threadedComments = `<w:comments xmlns:w="${W}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:comment w:id="5000" w:author="Prior"><w:p w14:paraId="ABCDEF12"><w:r><w:t>Old</w:t></w:r></w:p></w:comment></w:comments>`;
const commentsExtended = '<w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"><w15:commentEx w15:paraId="ABCDEF12" w15:done="0"/></w15:commentsEx>';
const legacyExtendedInput = buildZip([
    {name:'[Content_Types].xml',data:extendedTypes},
    {name:'word/document.xml',data:documentXml},
    {name:'word/_rels/document.xml.rels',data:extendedRels},
    {name:'word/comments.xml',data:threadedComments},
    {name:'word/commentsExtended.xml',data:commentsExtended}
]);
const normalizedExtended = await openDocx(legacyExtendedInput).applyOperations(
    [{ type:'replace', target:{ exactText:'Hello world' }, modified:'Hello reliable world' }],
    { author:'Agent', atomic:true }
);
assert.equal(normalizedExtended.written, true, 'safe package metadata normalization should not block unrelated redlines');
const normalizedTypes = unzipEntries(normalizedExtended.toBuffer()).get('[Content_Types].xml').toString();
assert.match(normalizedTypes, /application\/vnd\.ms-word\.commentsExtended\+xml/);
assert.doesNotMatch(normalizedTypes, /application\/vnd\.openxmlformats-officedocument\.wordprocessingml\.commentsExtended\+xml/);
console.log('docx package facade tests passed');
