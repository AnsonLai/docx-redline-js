/**
 * Builds the Word revision-semantics fixture matrix (WP-2 COM backstop).
 *
 *   node scripts/generate-word-revision-semantics-matrix.mjs
 *   powershell -File scripts/word-revision-semantics-matrix.ps1
 *
 * Writes tests/fixtures/word-revision-semantics/cases.json and one minimal .docx
 * per case; the PowerShell script drives Word and writes expectations.json.
 */
import { mkdirSync, writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { buildZip } from './lib/minimal-zip.mjs';

const OUT = join(dirname(fileURLToPath(import.meta.url)), '..', 'tests', 'fixtures', 'word-revision-semantics');
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PKG = 'http://schemas.openxmlformats.org/package/2006/relationships';
const DATE = '2026-01-01T00:00:00Z';

let id = 100;
const rev = (tag, author = 'A') => `<w:${tag} w:id="${id++}" w:author="${author}" w:date="${DATE}"/>`;
const run = t => `<w:r><w:t xml:space="preserve">${t}</w:t></w:r>`;
const insRun = (t, a = 'A') => `<w:ins w:id="${id++}" w:author="${a}" w:date="${DATE}"><w:r><w:t xml:space="preserve">${t}</w:t></w:r></w:ins>`;
const delRun = (t, a = 'A') => `<w:del w:id="${id++}" w:author="${a}" w:date="${DATE}"><w:r><w:delText xml:space="preserve">${t}</w:delText></w:r></w:del>`;
const NUM = '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>';
const LIST = '<w:pStyle w:val="ListParagraph"/>' + NUM;
const mark = (kind, a = 'A') => `<w:rPr>${rev(kind, a)}</w:rPr>`;
// pPr child order: style/numPr..., rPr (mark), sectPr, pPrChange
const para = (content, { ppr = '', mark: m = '', extra = '' } = {}) =>
    `<w:p>${ppr || m || extra ? `<w:pPr>${ppr}${m}${extra}</w:pPr>` : ''}${content}</w:p>`;
const cell = inner => `<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>${inner}</w:tc>`;
const table = rows => `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>${rows}</w:tbl>`;
const row = (inner, trPr = '') => `<w:tr>${trPr ? `<w:trPr>${trPr}</w:trPr>` : ''}${inner}</w:tr>`;
const SECT = '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>';

const defs = [];
function add(name, description, build) { id = 100; defs.push({ name, description, bodyXml: build() + SECT }); }
const P = t => para(run(t));
const pprChange = old => `<w:pPrChange w:id="${id++}" w:author="A" w:date="${DATE}"><w:pPr>${old}</w:pPr></w:pPrChange>`;
const BOLD_MARK = '<w:rPr><w:b/></w:rPr>';

add('ins-mark-mid-plain', 'Mid-body paragraph with inserted content and inserted mark',
    () => P('Alpha.') + para(insRun('Inserted.'), { mark: mark('ins') }) + P('Omega.'));
add('ins-mark-mid-untracked-content', 'Mid-body paragraph, inserted mark, untracked content',
    () => P('Alpha.') + para(run('Untracked.'), { mark: mark('ins') }) + P('Omega.'));
add('ins-mark-list-then-plain', 'List paragraph with inserted mark followed by plain paragraph',
    () => P('Alpha.') + para(insRun('Item.'), { ppr: LIST, mark: mark('ins') }) + P('Plain after.'));
add('ins-mark-last-prev-deleted', 'Last body paragraph inserted mark; preceding paragraph mark deleted (replacement shape)',
    () => P('First.') + para(delRun('Second paragraph.'), { mark: mark('del') })
        + para(insRun('Item one.'), { ppr: LIST, mark: mark('ins') })
        + para(insRun('Item two.'), { ppr: LIST, mark: mark('ins') }));
add('ins-mark-last-prev-plain', 'Last body paragraph inserted mark; preceding paragraph plain',
    () => P('First.') + P('Second paragraph.') + para(insRun('Added last.'), { ppr: LIST, mark: mark('ins') }));
add('ins-mark-last-plain-style', 'Last body paragraph inserted mark, no list props',
    () => P('First.') + para(insRun('Added last.'), { mark: mark('ins') }));
add('ins-mark-cell-sole', 'Sole paragraph in a table cell has inserted mark',
    () => P('Intro.') + table(row(cell(para(insRun('Cell ins.'), { mark: mark('ins') })))) + P('After table.'));
add('ins-mark-cell-last-of-two', 'Last of two paragraphs in a cell has inserted mark',
    () => P('Intro.') + table(row(cell(P('Cell one.') + para(insRun('Cell two.'), { mark: mark('ins') })))) + P('After table.'));
add('ins-mark-before-table', 'Paragraph with inserted mark directly before a table',
    () => P('Intro.') + para(insRun('Inserted.'), { mark: mark('ins') }) + table(row(cell(P('Cell.')))) + P('After table.'));
add('ins-mark-bookmark-between', 'Inserted mark paragraph with bookmark between paragraphs',
    () => P('Alpha.') + para('<w:bookmarkStart w:id="900" w:name="bm1"/>' + insRun('Inserted.') + '<w:bookmarkEnd w:id="900"/>', { mark: mark('ins') }) + P('Omega.'));

add('del-mark-mid-plain', 'Mid-body paragraph with deleted content and deleted mark',
    () => P('Alpha.') + para(delRun('Deleted.'), { mark: mark('del') }) + P('Omega.'));
add('del-mark-mid-untracked-content', 'Mid-body paragraph, deleted mark, untracked content',
    () => P('Alpha.') + para(run('Untracked.'), { mark: mark('del') }) + P('Omega.'));
add('del-mark-list-then-plain', 'List paragraph with deleted mark followed by plain paragraph',
    () => P('Alpha.') + para(delRun('Item.'), { ppr: LIST, mark: mark('del') }) + P('Plain after.'));
add('del-mark-last', 'Last body paragraph with deleted mark',
    () => P('First.') + para(delRun('Gone last.'), { mark: mark('del') }));
add('del-mark-last-list-prev-plain', 'Last body list paragraph with deleted mark; preceding plain',
    () => P('First.') + P('Second.') + para(delRun('Gone last.'), { ppr: LIST, mark: mark('del') }));
add('del-mark-cell-sole', 'Sole cell paragraph with deleted mark',
    () => P('Intro.') + table(row(cell(para(delRun('Cell del.'), { mark: mark('del') })))) + P('After table.'));
add('del-mark-cell-last-of-two', 'Last of two cell paragraphs with deleted mark',
    () => P('Intro.') + table(row(cell(P('Cell one.') + para(delRun('Cell two.'), { mark: mark('del') })))) + P('After table.'));
add('del-mark-before-table', 'Paragraph with deleted mark directly before a table',
    () => P('Intro.') + para(delRun('Deleted.'), { mark: mark('del') }) + table(row(cell(P('Cell.')))) + P('After table.'));
add('del-mark-bookmark-between', 'Deleted mark paragraph with bookmark between paragraphs',
    () => P('Alpha.') + para('<w:bookmarkStart w:id="900" w:name="bm1"/>' + delRun('Deleted.') + '<w:bookmarkEnd w:id="900"/>', { mark: mark('del') }) + P('Omega.'));

add('pprchange-restores-list', 'pPrChange: current Heading1, previous list item; mark rPr has bold',
    () => P('Alpha.') + para(run('Was a list item.'), { ppr: '<w:pStyle w:val="Heading1"/>', mark: BOLD_MARK, extra: pprChange(LIST) }) + P('Omega.'));
add('pprchange-removes-list', 'pPrChange: current list item, previous plain; mark rPr has bold',
    () => P('Alpha.') + para(run('Became a list item.'), { ppr: LIST, mark: BOLD_MARK, extra: pprChange('') }) + P('Omega.'));
add('del-mark-with-sectpr', 'Deleted-mark paragraph carrying a section break',
    () => para(delRun('Section one end.'), { mark: mark('del'), extra: '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>' }) + P('Section two.'));
add('adjacent-ins-then-del-marks', 'Inserted-mark paragraph followed by deleted-mark paragraph',
    () => P('Alpha.') + para(insRun('New.'), { mark: mark('ins') }) + para(delRun('Old.'), { mark: mark('del') }) + P('Keep.'));
add('adjacent-del-then-ins-marks', 'Deleted-mark paragraph followed by inserted-mark paragraph',
    () => P('Alpha.') + para(delRun('Old.'), { mark: mark('del') }) + para(insRun('New.'), { mark: mark('ins') }) + P('Keep.'));
add('mixed-authors-marks', 'Inserted marks by two authors plus a deleted mark by a third',
    () => P('Alpha.') + para(insRun('By A.', 'A'), { mark: mark('ins', 'A') })
        + para(insRun('By B.', 'B'), { mark: mark('ins', 'B') })
        + para(delRun('By C.', 'C'), { mark: mark('del', 'C') }) + P('Omega.'));
add('ins-row', 'Table with an inserted row',
    () => P('Intro.') + table(
        row(cell(P('Existing row.')))
        + row(cell(para(insRun('Inserted row.'), { mark: mark('ins') })), rev('ins'))) + P('After table.'));
add('del-row', 'Table with a deleted row',
    () => P('Intro.') + table(
        row(cell(P('Existing row.')))
        + row(cell(para(delRun('Deleted row.'), { mark: mark('del') })), rev('del'))) + P('After table.'));

// R5: end-of-document table append. insRow = row-level ins with inserted cell marks (library shape).
const insRow = (text, cols = 2) => row(Array.from({ length: cols }, (_, i) => cell(para(insRun(`${text}${i + 1}`), { mark: mark('ins') }))).join(''), rev('ins'));
add('table-append-end-trailing-ins-mark', 'Table appended after the final paragraph: inserted rows, then an empty trailing paragraph with an inserted mark (library shape)',
    () => P('Alpha.') + P('Final paragraph.') + table(insRow('R1C') + insRow('R2C')) + para('', { mark: mark('ins') }));
add('table-append-end-word-native', 'Table appended after the final paragraph as Word writes it: final paragraph mark inserted, inserted rows, untracked empty trailing paragraph',
    () => P('Alpha.') + para(run('Final paragraph.'), { mark: mark('ins') }) + table(insRow('R1C') + insRow('R2C')) + P(''));
add('table-append-end-untracked-trailing', 'Table appended after the final paragraph: inserted rows followed by an untracked empty trailing paragraph',
    () => P('Alpha.') + P('Final paragraph.') + table(insRow('R1C') + insRow('R2C')) + P(''));

const wrap = body => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>${body}</w:body></w:document>`;
const styles = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="${W}"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/><w:pPr><w:ind w:left="720"/></w:pPr></w:style></w:styles>`;
const numbering = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering xmlns:w="${W}"><w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="hybridMultilevel"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="-"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`;
const entries = body => [
    { name: '[Content_Types].xml', data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>` },
    { name: '_rels/.rels', data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${PKG}"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>` },
    { name: 'word/document.xml', data: wrap(body) },
    { name: 'word/_rels/document.xml.rels', data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${PKG}"><Relationship Id="rId1" Type="${R}/styles" Target="styles.xml"/><Relationship Id="rId2" Type="${R}/numbering" Target="numbering.xml"/></Relationships>` },
    { name: 'word/styles.xml', data: styles },
    { name: 'word/numbering.xml', data: numbering }
];

mkdirSync(OUT, { recursive: true });
for (const def of defs) writeFileSync(join(OUT, `${def.name}.docx`), buildZip(entries(def.bodyXml)));
writeFileSync(join(OUT, 'cases.json'), JSON.stringify(defs, null, 2) + '\n');
console.log(`Wrote ${defs.length} cases to ${OUT}`);
