import assert from 'node:assert/strict';

import './setup-xml-provider.mjs';
import { acceptTrackedChangesInOoxml, applyRedlineToOxml, rejectTrackedChangesInOoxml } from '../index.js';
import { parseOoxmlSafe } from '../adapters/xml-adapter.js';

// Regression: a recitals paragraph whose "A. / B. / C." items are separated by
// soft line breaks (w:br). Word's Paragraph.text reports the breaks as "\v",
// and the generated upperAlpha Markdown ("A. ...\nB. ...") is byte-identical to
// the manually lettered source text.

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const ITEMS = ['First recital.', 'Second recital.', 'Third recital.'];
const LETTERED = ITEMS.map((text, index) => `${'ABC'[index]}. ${text}`);
const LIST_MARKDOWN = LETTERED.join('\n');
const SOFT_BREAK_OOXML = `<w:document ${W}><w:body><w:p>`
    + `<w:r><w:t xml:space="preserve">${LETTERED[0]}</w:t></w:r>`
    + `<w:r><w:br/><w:t xml:space="preserve">${LETTERED[1]}</w:t></w:r>`
    + `<w:r><w:br/><w:t xml:space="preserve">${LETTERED[2]}</w:t></w:r>`
    + `</w:p></w:body></w:document>`;
const SEPARATE_OOXML = `<w:document ${W}><w:body>`
    + LETTERED.map(text => `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`).join('')
    + `</w:body></w:document>`;
const OPTIONS = { author: 'Soft Break Lists', generateRedlines: true };

const paragraphsOf = xml => {
    const parsed = parseOoxmlSafe(xml);
    assert.ok(parsed.doc, parsed.error?.message || 'parseable OOXML');
    return Array.from(parsed.doc.getElementsByTagNameNS('*', 'p'));
};
const visibleText = paragraph => Array.from(paragraph.getElementsByTagNameNS('*', 'r')).map(run => (
    Array.from(run.childNodes).map(node => (
        node.localName === 't' || node.localName === 'delText' ? node.textContent
            : node.localName === 'br' ? '\n'
                : ''
    )).join('')
)).join('');
const listParagraphs = xml => paragraphsOf(xml).filter(paragraph => paragraph.getElementsByTagNameNS('*', 'numPr').length > 0);

// 1. Word host text uses "\v" for soft breaks; the target must still be found.
{
    const result = await applyRedlineToOxml(SOFT_BREAK_OOXML, LETTERED.join('\u000b'), LIST_MARKDOWN, OPTIONS);
    assert.equal(result.status, 'ok', JSON.stringify(result.error));
    assert.equal(result.hasChanges, true);
    const accepted = acceptTrackedChangesInOoxml(result.oxml, { allAuthors: true }).oxml;
    assert.deepEqual(listParagraphs(accepted).map(visibleText), ITEMS, 'accepted view is a three-item real list');

    // Reject All must restore the original soft breaks, not raw newline characters.
    const rejected = rejectTrackedChangesInOoxml(result.oxml, { allAuthors: true }).oxml;
    const restored = paragraphsOf(rejected).filter(paragraph => visibleText(paragraph).trim());
    assert.equal(restored.length, 1, 'rejected view restores one source paragraph');
    assert.equal(visibleText(restored[0]), LETTERED.join('\n'), 'rejected view keeps w:br line breaks');
    assert.ok(!/[\n\v]/.test(Array.from(restored[0].getElementsByTagNameNS('*', 't')).map(node => node.textContent).join('')),
        'no raw line-break characters inside w:t');
}

// 2. Identical manual markers: implicit requests stay a no-op (an unchanged
// "1. Heading" must not silently become an auto-numbered list) ...
for (const [name, ooxml] of [['soft breaks', SOFT_BREAK_OOXML], ['separate paragraphs', SEPARATE_OOXML]]) {
    const implicit = await applyRedlineToOxml(ooxml, LETTERED.join('\n'), LIST_MARKDOWN, OPTIONS);
    assert.equal(implicit.hasChanges, false, `${name}: implicit identical text is unchanged`);

    // ... while an explicit structured-content request converts to a real list.
    const explicit = await applyRedlineToOxml(ooxml, LETTERED.join('\n'), LIST_MARKDOWN, {
        ...OPTIONS,
        explicitStructuredContent: true
    });
    assert.equal(explicit.status, 'ok', `${name}: ${JSON.stringify(explicit.error)}`);
    assert.equal(explicit.hasChanges, true, `${name}: explicit list request converts identical manual markers`);
    const accepted = acceptTrackedChangesInOoxml(explicit.oxml, { allAuthors: true }).oxml;
    assert.deepEqual(listParagraphs(accepted).map(visibleText), ITEMS, `${name}: accepted view is a real list`);
}

// 3. An explicit request against an already-equivalent real list is still a no-op.
{
    const listOoxml = `<w:document ${W}><w:body>`
        + ITEMS.map(text => `<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="7"/></w:numPr></w:pPr>`
            + `<w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`).join('')
        + `</w:body></w:document>`;
    const result = await applyRedlineToOxml(listOoxml, LETTERED.join('\n'), LIST_MARKDOWN, {
        ...OPTIONS,
        explicitStructuredContent: true
    });
    assert.equal(result.hasChanges, false, 'existing equivalent list is unchanged');
}

console.log('PASS: soft-break list conversion targets Word text, converts explicit identical markers, and rejects back to w:br');
