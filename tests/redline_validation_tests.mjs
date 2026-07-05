import './setup-xml-provider.mjs';

import assert from 'assert/strict';

import { applyRedlineToOxml, validateRedlineOoxml } from '../index.js';

const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

function paragraph(inner) {
    return `<w:p xmlns:w="${NS_W}">${inner}</w:p>`;
}

const META = 'w:id="101" w:author="Validator" w:date="2026-07-04T00:00:00Z"';
const META2 = 'w:id="102" w:author="Validator" w:date="2026-07-04T00:00:00Z"';

function issueCodes(result) {
    return result.issues.map(issue => issue.code);
}

// --- Clean engine output validates ---
{
    const source = paragraph('<w:r><w:t>The quick brown fox.</w:t></w:r>');
    const redlined = await applyRedlineToOxml(source, 'The quick brown fox.', 'The quick red fox.', {
        generateRedlines: true,
        author: 'Validator'
    });
    const result = validateRedlineOoxml(redlined.oxml);
    assert.equal(result.valid, true, `expected engine output to validate, got: ${JSON.stringify(result.issues)}`);
}

// --- Clean untracked paragraph validates ---
{
    const result = validateRedlineOoxml(paragraph('<w:r><w:t>Plain text.</w:t></w:r>'));
    assert.equal(result.valid, true);
    assert.equal(result.issues.length, 0);
}

// --- Parse errors ---
{
    const result = validateRedlineOoxml('<w:p><w:r>unclosed');
    assert.equal(result.valid, false);
    assert(issueCodes(result).includes('PARSE_ERROR'));
}
{
    const result = validateRedlineOoxml('');
    assert.equal(result.valid, false);
    assert(issueCodes(result).includes('PARSE_ERROR'));
}

// --- Nested revisions ---
{
    const result = validateRedlineOoxml(paragraph(
        `<w:ins ${META}><w:del ${META2}><w:r><w:delText>bad</w:delText></w:r></w:del></w:ins>`
    ));
    assert.equal(result.valid, false);
    assert(issueCodes(result).includes('NESTED_REVISION'));
}

// --- w:t inside w:del ---
{
    const result = validateRedlineOoxml(paragraph(
        `<w:del ${META}><w:r><w:t>should be delText</w:t></w:r></w:del>`
    ));
    assert.equal(result.valid, false);
    assert(issueCodes(result).includes('DEL_CONTAINS_T'));
}

// --- Missing revision metadata ---
{
    const result = validateRedlineOoxml(paragraph(
        '<w:ins w:id="7"><w:r><w:t>no author or date</w:t></w:r></w:ins>'
    ));
    assert.equal(result.valid, false);
    assert(issueCodes(result).includes('MISSING_REVISION_METADATA'));
}

// --- Duplicate revision ids ---
{
    const result = validateRedlineOoxml(paragraph(
        `<w:ins ${META}><w:r><w:t>one</w:t></w:r></w:ins>` +
        `<w:ins ${META}><w:r><w:t>two</w:t></w:r></w:ins>`
    ));
    assert.equal(result.valid, false);
    assert(issueCodes(result).includes('DUPLICATE_REVISION_ID'));
}

// --- Boundary whitespace without xml:space="preserve" ---
{
    const result = validateRedlineOoxml(paragraph('<w:r><w:t>trailing space </w:t></w:r>'));
    assert.equal(result.valid, false);
    assert(issueCodes(result).includes('MISSING_SPACE_PRESERVE'));
}

// --- Empty text element and empty wrapper are warnings, not errors ---
{
    const result = validateRedlineOoxml(paragraph(
        `<w:r><w:t></w:t></w:r><w:ins ${META}></w:ins>`
    ));
    assert.equal(result.valid, true, 'warnings alone should not invalidate');
    assert(issueCodes(result).includes('EMPTY_TEXT_ELEMENT'));
    assert(issueCodes(result).includes('EMPTY_REVISION_WRAPPER'));
}

// --- Paragraph-mark revisions inside w:rPr are legitimately empty ---
{
    const result = validateRedlineOoxml(paragraph(
        `<w:pPr><w:rPr><w:ins ${META}/></w:rPr></w:pPr><w:r><w:t>inserted paragraph</w:t></w:r>`
    ));
    assert(!issueCodes(result).includes('EMPTY_REVISION_WRAPPER'),
        'paragraph-mark revision must not be flagged as an empty wrapper');
}

console.log('PASS: redline validation tests');
