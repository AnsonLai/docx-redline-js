import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import path from 'node:path';

import { unzipDocx, zipDocx } from '../../document/zip-archive.js';

/**
 * Fixture documents for the reject-roundtrip matrix (WP-3).
 *
 * Synthetic fixtures reuse the package parts (numbering: numId 1 bullet, 2 decimal; styles; rels) of the
 * Word-authored nested-lists source and replace only the body, so list and package behavior is real.
 */

const HERE = path.dirname(fileURLToPath(import.meta.url));
const NESTED_LISTS = path.join(HERE, '..', 'fixtures', 'agentic-lists', 'nested-lists-source.docx');

const dec = new TextDecoder();
const enc = new TextEncoder();

const META = (id, author) => `w:id="${id}" w:author="${author}" w:date="2026-09-01T00:00:00Z"`;

const t = text => `<w:t xml:space="preserve">${text}</w:t>`;
const run = (text, rPr = '') => `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}${t(text)}</w:r>`;
const p = (inner, pPr = '') => `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}${inner}</w:p>`;
const bullet = (inner, ilvl = 0) => p(inner, `<w:numPr><w:ilvl w:val="${ilvl}"/><w:numId w:val="1"/></w:numPr>`);
const decimal = (inner, ilvl = 0) => p(inner, `<w:numPr><w:ilvl w:val="${ilvl}"/><w:numId w:val="2"/></w:numPr>`);
const link = (rid, text) => `<w:hyperlink r:id="${rid}" w:history="1"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr>${t(text)}</w:r></w:hyperlink>`;
const ins = (id, author, inner) => `<w:ins ${META(id, author)}>${inner}</w:ins>`;
const del = (id, author, text) => `<w:del ${META(id, author)}><w:r><w:delText xml:space="preserve">${text}</w:delText></w:r></w:del>`;
const tbl = rows => `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>${
    rows.map(cells => `<w:tr>${cells.map(c => `<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>${c}</w:tc>`).join('')}</w:tr>`).join('')
}</w:tbl>`;

export const SYNTHETIC_BODIES = {
    'hyperlinks-trailing-punctuation': [
        p(run('Review the ') + link('rIdMatrixLink1', 'Terms of Service') + run(', and then continue with the payment schedule.')),
        p(run('Contact support (see ') + link('rIdMatrixLink2', 'Annex A') + run(').')),
        p(link('rIdMatrixLink1', 'https://example.com/terms') + run('.')),
        p(run('Plain closing paragraph with no link.'))
    ],
    'breaks-and-tabs': [
        p(`<w:r>${t('Line one of the notice')}<w:br/>${t('Line two of the notice')}</w:r>`),
        p(`<w:r>${t('Name')}<w:tab/>${t('Value of the party')}<w:tab/>${t('Third column')}</w:r>`),
        p(run('Heading ', '<w:b/>') + `<w:r><w:tab/>${t('after tab')}</w:r>`),
        p(run('A plain paragraph that follows.'))
    ],
    'simple-fields': [
        p(run('Page ') + '<w:fldSimple w:instr=" PAGE "><w:r><w:t>3</w:t></w:r></w:fldSimple>' + run(' of the agreement is attached.')),
        p('<w:fldSimple w:instr=" DATE \\@&quot;yyyy&quot; "><w:r><w:t>2026</w:t></w:r></w:fldSimple>' + run(' annual review of terms.')),
        p(run('Closing paragraph after the fields.'))
    ],
    'tables': [
        p(run('Paragraph before the table.')),
        tbl([
            [p(run('Cell alpha one')), p(run('Cell beta one'))],
            [p(run('Cell alpha two')) + p(run('Cell alpha two continued')), p(run('Cell beta two'))]
        ]),
        p(run('Paragraph after the table.')),
        tbl([[p(run('Only cell content')), p('')]]),
        p(run('Final body paragraph after both tables.'))
    ],
    'multi-author-revisions': [
        p(run('Baseline paragraph kept as is.')),
        p(run('Alice rewrote ') + del(901, 'Bob Reviewer', 'the old phrase ') + ins(902, 'Alice Reviewer', run('the new phrase ')) + run('in this clause.')),
        p(ins(903, 'Bob Reviewer', run('A paragraph Bob inserted entirely.')), `<w:rPr><w:ins ${META(904, 'Bob Reviewer')}/></w:rPr>`),
        p(run('Closing text with ') + ins(905, 'Alice Reviewer', run('an addition ')) + run('and a tail.'))
    ],
    'empty-list-items': [
        p(run('Intro before the list.')),
        bullet(run('First bullet item')),
        bullet(''),
        bullet(run('Third bullet item')),
        bullet(''),
        bullet(''),
        decimal(run('Numbered one')),
        decimal(''),
        decimal(run('Numbered three'), 1),
        p(run('Outro after the lists.'))
    ],
    'typed-headers': [
        p(run('A. Scope of the agreement')),
        p(run('Body text that explains the scope.')),
        p(run('B. ', '<w:b/>') + run('Term and termination', '<w:b/>')),
        p(`<w:r>${t('C.')}<w:tab/>${t('Fees and charges')}</w:r>`),
        p(run('1. Numbered heading one')),
        p(run('2) Numbered heading two')),
        p(run('IV. Roman heading four')),
        p(run('- Dash bullet heading')),
        p(link('rIdMatrixLink1', 'D. ') + run('Linked marker heading')),
        tbl([[p(run('a. Cell heading')), p(run('b. Second cell heading'))]]),
        p(run('Closing paragraph without a marker.'))
    ],
    'last-body-paragraph': [
        p(run('First clause of the agreement.')),
        p(''),
        p(''),
        p(run('Second clause with ') + run('bold words', '<w:b/>') + run(' inside.')),
        p(run('Final clause that ends the document body.'))
    ]
};

function buildFromTemplate(bodyXml) {
    const entries = unzipDocx(readFileSync(NESTED_LISTS));
    const original = dec.decode(entries.get('word/document.xml'));
    const sectPr = original.match(/<w:sectPr[\s\S]*<\/w:sectPr>\s*<\/w:body>/)[0];
    const head = original.slice(0, original.indexOf('<w:body>') + '<w:body>'.length);
    entries.set('word/document.xml', enc.encode(`${head}${bodyXml}${sectPr}</w:document>`));
    const rels = dec.decode(entries.get('word/_rels/document.xml.rels'));
    const links = ['rIdMatrixLink1', 'rIdMatrixLink2']
        .map((id, i) => `<Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/${i + 1}" TargetMode="External"/>`)
        .join('');
    entries.set('word/_rels/document.xml.rels', enc.encode(rels.replace('</Relationships>', `${links}</Relationships>`)));
    return zipDocx(entries);
}

/** @returns {Array<{ name: string, bytes: Uint8Array }>} */
export function buildMatrixFixtures() {
    const fixtures = [{ name: 'nested-lists-source', bytes: new Uint8Array(readFileSync(NESTED_LISTS)), options: {} }];
    for (const [name, paragraphs] of Object.entries(SYNTHETIC_BODIES)) {
        const bytes = new Uint8Array(buildFromTemplate(paragraphs.join('')));
        fixtures.push({ name, bytes, options: {} });
        // Foreign reviewers' revisions need the explicit slicing policy to be editable at all.
        if (name === 'multi-author-revisions') {
            fixtures.push({ name: `${name}+slice-cross-author`, bytes, options: { existingRevisions: 'slice-cross-author' } });
        }
    }
    return fixtures;
}
