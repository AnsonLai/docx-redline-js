import assert from 'node:assert/strict';
import { buildZip } from '../scripts/lib/minimal-zip.mjs';
import { unzipEntries } from '../scripts/lib/zip-reader.mjs';
import { openDocx } from '../node/index.js';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const contentTypes = `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`;
const emptyRels = `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`;
function packageBuffer(documentXml, extra = []) { return buildZip([{ name:'[Content_Types].xml', data:contentTypes }, { name:'word/document.xml', data:documentXml }, { name:'word/_rels/document.xml.rels', data:emptyRels }, ...extra]); }

const simpleXml = `<w:document xmlns:w="${W}"><w:body><w:p><w:r><w:t>Alpha</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`;
const simpleInput = packageBuffer(simpleXml, [{ name:'media/unchanged.bin', data:Buffer.from([0,255,1,254]) }]);
const noOpDoc = openDocx(simpleInput);
const noOp = await noOpDoc.applyOperations([], { author:'Editor' });
assert.equal(noOp.status, 'ok'); assert.equal(noOp.written, false); assert.deepEqual(noOp.toBuffer(), simpleInput);

const sequential = openDocx(simpleInput);
const first = await sequential.applyOperations([{ type:'replace', target:{ exactText:'Alpha' }, modified:'Beta', generateRedlines:false }], { author:'Editor' });
assert.equal(first.written, true);
const committed = first.toBuffer();
const afterCommitNoOp = await sequential.applyOperations([], { author:'Editor' });
assert.equal(afterCommitNoOp.written, false); assert.deepEqual(afterCommitNoOp.toBuffer(), committed);
const second = await sequential.applyOperations([{ type:'replace', target:{ exactText:'Missing' }, modified:'Never' }], { author:'Editor', atomic: true });
assert.equal(second.rolledBack, true); assert.deepEqual(second.toBuffer(), committed);
assert.equal(sequential.inspect().paragraphs[0].text, 'Beta');

const commentsXml = `<w:comments xmlns:w="${W}"><w:comment w:id="1" w:author="Alice"><w:p><w:r><w:t>A</w:t></w:r></w:p></w:comment><w:comment w:id="2" w:author="Bob"><w:p><w:r><w:t>B</w:t></w:r></w:p></w:comment></w:comments>`;
const commentDocXml = `<w:document xmlns:w="${W}"><w:body><w:p><w:commentRangeStart w:id="1"/><w:r><w:t>A</w:t></w:r><w:commentRangeEnd w:id="1"/><w:r><w:commentReference w:id="1"/></w:r></w:p><w:p><w:commentRangeStart w:id="2"/><w:r><w:t>B</w:t></w:r><w:commentRangeEnd w:id="2"/><w:r><w:commentReference w:id="2"/></w:r></w:p><w:sectPr/></w:body></w:document>`;
const commentTypes = contentTypes.replace('</Types>', '<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/></Types>');
const commentRels = `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/></Relationships>`;
const commentInput = buildZip([{name:'[Content_Types].xml',data:commentTypes},{name:'word/document.xml',data:commentDocXml},{name:'word/_rels/document.xml.rels',data:commentRels},{name:'word/comments.xml',data:commentsXml}]);

const protectedDoc = openDocx(commentInput);
const protectedPreflight = protectedDoc.preflight([{ type:'delete', target:{ exactText:'A' } }], 'Editor');
assert.equal(protectedPreflight.valid, false);
assert.equal(protectedPreflight.results[0].error.code, 'COMMENTED_CONTENT_DELETE');
assert.deepEqual(protectedPreflight.results[0].error.commentIds, ['1']);
assert.equal(protectedPreflight.results[0].error.comments[0].author, 'Alice');
assert.equal(protectedPreflight.results[0].error.comments[0].text, 'A');
const protectedDelete = await protectedDoc.applyOperations([{ type:'delete', target:{ exactText:'A' } }], { author:'Editor', atomic: true });
assert.equal(protectedDelete.written, false);
assert.equal(protectedDelete.rolledBack, true);
assert.equal(protectedDelete.results[0].error.code, 'COMMENTED_CONTENT_DELETE');
assert.equal(protectedDelete.results[0].error.comments[0].author, 'Alice');
assert.deepEqual(protectedDelete.toBuffer(), commentInput);

const selective = openDocx(commentInput);
const deletedAlice = await selective.deleteComments({ author:'Alice' });
assert.equal(deletedAlice.status, 'ok', JSON.stringify(deletedAlice));
assert.equal(deletedAlice.commentsRemoved, 1); assert.equal(deletedAlice.referencesRemoved, 3);
const selectiveParts = unzipEntries(deletedAlice.toBuffer());
assert.doesNotMatch(selectiveParts.get('word/comments.xml').toString(), /w:author="Alice"/);
assert.match(selectiveParts.get('word/comments.xml').toString(), /w:author="Bob"/);
assert.doesNotMatch(selectiveParts.get('word/document.xml').toString(), /w:id="1"/);
assert.match(selectiveParts.get('word/document.xml').toString(), /w:id="2"/);
const deletedRest = await selective.deleteComments({ allAuthors:true });
assert.equal(deletedRest.commentsRemoved, 1);
assert.doesNotMatch(unzipEntries(deletedRest.toBuffer()).get('word/document.xml').toString(), /commentRange|commentReference/);

const revisedXml = `<w:document xmlns:w="${W}"><w:body><w:p><w:ins w:id="10" w:author="Alice" w:date="2026-01-01T00:00:00Z"><w:r><w:t>A</w:t></w:r></w:ins><w:ins w:id="11" w:author="Bob" w:date="2026-01-01T00:00:00Z"><w:r><w:t>B</w:t></w:r></w:ins></w:p><w:sectPr/></w:body></w:document>`;
const revisions = openDocx(packageBuffer(revisedXml));
const acceptedAlice = await revisions.resolveRevisions('accept', { author:'Alice' });
assert.equal(acceptedAlice.acceptedCount, 1); assert.deepEqual(revisions.inspect().revisionAuthors, ['Bob']);
const rejectedBob = await revisions.resolveRevisions('reject', { author:'Bob' });
assert.equal(rejectedBob.rejectedCount, 1); assert.equal(revisions.inspect().paragraphs[0].text, 'A');

const anchoredHighXml = `<w:document xmlns:w="${W}"><w:body><w:p><w:commentRangeStart w:id="9000"/><w:r><w:t>High anchor</w:t></w:r><w:commentRangeEnd w:id="9000"/><w:r><w:commentReference w:id="9000"/></w:r></w:p><w:sectPr/></w:body></w:document>`;
const highCommentsXml = `<w:comments xmlns:w="${W}"><w:comment w:id="9000" w:author="Existing"><w:p><w:r><w:t>Existing comment</w:t></w:r></w:p></w:comment></w:comments>`;
const highInput = buildZip([{name:'[Content_Types].xml',data:commentTypes},{name:'word/document.xml',data:anchoredHighXml},{name:'word/_rels/document.xml.rels',data:commentRels},{name:'word/comments.xml',data:highCommentsXml}]);
const highDoc = openDocx(highInput);
const highResult = await highDoc.applyOperations([{ type:'comment', target:{ exactText:'High anchor' }, commentContent:'New' }], { author:'Agent' });
assert.equal(highResult.written, true);
assert.match(unzipEntries(highResult.toBuffer()).get('word/comments.xml').toString(), /w:id="9001"/);
assert.ok(highResult.artifactsChanged.includes('word/document.xml'));
assert.ok(highResult.artifactsChanged.includes('word/comments.xml'));
console.log('docx package transaction edge tests passed');
