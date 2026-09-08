import assert from 'assert/strict';

import './setup-xml-provider.mjs';
import {
    acceptTrackedChangesInOoxml,
    applyRedlineToOxml,
    ingestWordOoxmlToPlainText,
    rejectTrackedChangesInOoxml,
    validateRedlineOoxml
} from '../index.js';
import { NS_W } from '../core/types.js';
import { parseOoxmlSafe } from '../adapters/xml-adapter.js';

const NS_W16DU = 'http://schemas.microsoft.com/office/word/2023/wordml/word16du';
const BARRY_DATE = '2026-09-08T09:29:00Z';
const BARRY_DATE_UTC = '2026-09-08T16:29:00Z';

function paragraph(content) {
    return `<w:p xmlns:w="${NS_W}" xmlns:w16du="${NS_W16DU}">${content}</w:p>`;
}

function insertion(content, id = 1, author = 'Barry') {
    return `<w:ins w:id="${id}" w:author="${author}" w:date="${BARRY_DATE}" `
        + `w16du:dateUtc="${BARRY_DATE_UTC}">${content}</w:ins>`;
}

function parse(xml) {
    const parsed = parseOoxmlSafe(xml, 'application/xml');
    assert.equal(parsed.error, null);
    return parsed.doc;
}

function elements(node, localName) {
    return Array.from(node.getElementsByTagNameNS(NS_W, localName));
}

function directElements(node, localName = null) {
    return Array.from(node.childNodes || []).filter(child => {
        return child.nodeType === 1 && (!localName || child.localName === localName);
    });
}

function attr(node, localName) {
    return node.getAttributeNS(NS_W, localName) || node.getAttribute(`w:${localName}`);
}

function revisionText(node) {
    return elements(node, 't').concat(elements(node, 'delText'))
        .map(textNode => textNode.textContent || '')
        .join('');
}

function assertValidAndUnique(xml) {
    const validation = validateRedlineOoxml(xml);
    assert.equal(validation.valid, true, JSON.stringify(validation.issues));
    const ids = elements(parse(xml), 'ins').concat(elements(parse(xml), 'del'))
        .map(node => attr(node, 'id'));
    assert.equal(new Set(ids).size, ids.length, `duplicate revision IDs: ${ids.join(', ')}`);
}

async function apply(source, original, modified, extra = {}) {
    const result = await applyRedlineToOxml(source, original, modified, {
        author: 'Anson',
        existingRevisions: 'slice-cross-author',
        ...extra
    });
    assert.equal(result.status, 'ok', JSON.stringify(result.error));
    assert.equal(result.hasChanges, true);
    assertValidAndUnique(result.oxml);
    return result;
}

// WP03: an insertion inside a foreign insertion is hoisted between split carriers.
{
    const source = paragraph(
        '<w:r><w:t xml:space="preserve">Base </w:t></w:r>'
        + insertion('<w:r><w:t xml:space="preserve">amended by this </w:t></w:r><w:r><w:t>Agreement.</w:t></w:r>')
    );
    const result = await apply(
        source,
        'Base amended by this Agreement.',
        'Base amended by this MASTER Agreement.'
    );
    const doc = parse(result.oxml);
    const p = elements(doc, 'p')[0];
    const insertions = directElements(p, 'ins');
    assert.equal(insertions.length, 3);
    assert.deepEqual(insertions.map(node => attr(node, 'author')), ['Barry', 'Anson', 'Barry']);
    assert.deepEqual(insertions.map(revisionText), ['amended by this ', 'MASTER ', 'Agreement.']);
    assert.equal(attr(insertions[0], 'date'), BARRY_DATE);
    assert.equal(insertions[2].getAttributeNS(NS_W16DU, 'dateUtc'), BARRY_DATE_UTC);

    const accepted = acceptTrackedChangesInOoxml(result.oxml, { allAuthors: true });
    const rejected = rejectTrackedChangesInOoxml(result.oxml, { allAuthors: true });
    assert.equal(ingestWordOoxmlToPlainText(accepted.oxml).trim(), 'Base amended by this MASTER Agreement.');
    assert.equal(ingestWordOoxmlToPlainText(rejected.oxml).trim(), 'Base');
}

// WP03: exact carrier boundaries omit empty fragments.
for (const scenario of [
    { original: 'Base Pending', modified: 'Base New Pending', authors: ['Anson', 'Barry'] },
    { original: 'Base Pending', modified: 'Base Pending New', authors: ['Barry', 'Anson'] }
]) {
    const source = paragraph(
        '<w:r><w:t xml:space="preserve">Base </w:t></w:r>'
        + insertion('<w:r><w:t>Pending</w:t></w:r>')
    );
    const result = await apply(source, scenario.original, scenario.modified);
    const insertions = directElements(elements(parse(result.oxml), 'p')[0], 'ins');
    assert.equal(insertions.length, 2);
    assert.deepEqual(insertions.map(node => attr(node, 'author')), scenario.authors);
    assert(insertions.every(node => revisionText(node).length > 0));
}

// The slicing policy retains merge-same-author semantics for the current author.
{
    const source = paragraph(
        '<w:r><w:t xml:space="preserve">Base </w:t></w:r>'
        + insertion('<w:r><w:t>draft</w:t></w:r>', 1, 'Anson')
    );
    const result = await apply(source, 'Base draft', 'Base final');
    const insertions = directElements(elements(parse(result.oxml), 'p')[0], 'ins');
    assert.equal(insertions.length, 1);
    assert.equal(attr(insertions[0], 'author'), 'Anson');
    assert.equal(revisionText(insertions[0]), 'final');
}

// New cross-author text inherits effective formatting, not the carrier's historical rPrChange ID.
{
    const source = paragraph(insertion(
        `<w:r><w:rPr><w:b/><w:rPrChange w:id="2" w:author="Formatter" w:date="${BARRY_DATE}">`
        + '<w:rPr/></w:rPrChange></w:rPr><w:t>BoldText</w:t></w:r>'
    ));
    const result = await apply(source, 'BoldText', 'BoldNewText');
    const doc = parse(result.oxml);
    const ansonInsertion = elements(doc, 'ins').find(node => attr(node, 'author') === 'Anson');
    assert.equal(elements(ansonInsertion, 'b').length, 1);
    assert.equal(elements(ansonInsertion, 'rPrChange').length, 0);
}

// WP04: a formatted, multi-run deletion becomes one nested w:del in the carrier.
{
    const source = paragraph(
        '<w:r><w:t xml:space="preserve">Base </w:t></w:r>'
        + insertion(
            '<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">The Services </w:t></w:r>'
            + '<w:r><w:t>generate outputs</w:t></w:r>'
        )
    );
    const result = await apply(source, 'Base The Services generate outputs', 'Base The outputs');
    const doc = parse(result.oxml);
    const carrier = directElements(elements(doc, 'p')[0], 'ins')[0];
    assert.equal(directElements(carrier, 'del').length, 1);
    const deletion = directElements(carrier, 'del')[0];
    assert.equal(attr(deletion, 'author'), 'Anson');
    assert.equal(revisionText(deletion), 'Services generate ');
    assert.equal(directElements(deletion, 'r').length, 2, 'one deletion wrapper should preserve two source runs');
    assert.equal(elements(directElements(deletion, 'r')[0], 'b').length, 1);
}

// SYN-02, SYN-11, and SYN-12: generate stacked deletions and exercise every lifecycle branch.
{
    const source = paragraph(
        '<w:r><w:t xml:space="preserve">Background. </w:t></w:r>'
        + insertion('<w:r><w:t>The Services will process the Input to generate outputs</w:t></w:r>')
    );
    const anson = await apply(
        source,
        'Background. The Services will process the Input to generate outputs',
        'Background. The Services will process the Input to outputs'
    );
    const ansonCarrier = directElements(elements(parse(anson.oxml), 'p')[0], 'ins')[0];
    const ansonDeletion = directElements(ansonCarrier, 'del')[0];
    assert.equal(attr(ansonCarrier, 'author'), 'Barry');
    assert.equal(attr(ansonDeletion, 'author'), 'Anson');
    assert.equal(revisionText(ansonDeletion), 'generate ');

    const acceptAll = acceptTrackedChangesInOoxml(anson.oxml, { allAuthors: true });
    assert.equal(ingestWordOoxmlToPlainText(acceptAll.oxml).trim(),
        'Background. The Services will process the Input to outputs');
    assert.equal(elements(parse(acceptAll.oxml), 'ins').length + elements(parse(acceptAll.oxml), 'del').length, 0);

    const acceptBarry = acceptTrackedChangesInOoxml(anson.oxml, { author: 'Barry' });
    assert.equal(elements(parse(acceptBarry.oxml), 'ins').length, 0);
    assert.equal(attr(elements(parse(acceptBarry.oxml), 'del')[0], 'author'), 'Anson');

    const rejectBarry = rejectTrackedChangesInOoxml(anson.oxml, { author: 'Barry' });
    assert.equal(ingestWordOoxmlToPlainText(rejectBarry.oxml).trim(), 'Background.');
    assert.equal(elements(parse(rejectBarry.oxml), 'del').length, 0);

    const rejectAnson = rejectTrackedChangesInOoxml(anson.oxml, { author: 'Anson' });
    assert.equal(revisionText(elements(parse(rejectAnson.oxml), 'ins')[0]),
        'The Services will process the Input to generate outputs');
    assert.equal(elements(parse(rejectAnson.oxml), 'del').length, 0);

    const chris = await apply(
        anson.oxml,
        'Background. The Services will process the Input to outputs',
        'Background. The Services will the Input to outputs',
        { author: 'Davis, Chris' }
    );
    const stackedCarrier = directElements(elements(parse(chris.oxml), 'p')[0], 'ins')[0];
    const stackedDeletions = directElements(stackedCarrier, 'del');
    assert.deepEqual(stackedDeletions.map(node => attr(node, 'author')), ['Davis, Chris', 'Anson']);
    assert.deepEqual(stackedDeletions.map(revisionText), ['process ', 'generate ']);

    const rejectChris = rejectTrackedChangesInOoxml(chris.oxml, { author: 'Davis, Chris' });
    assert.equal(elements(parse(rejectChris.oxml), 'del').length, 1);
    assert.equal(attr(elements(parse(rejectChris.oxml), 'del')[0], 'author'), 'Anson');
    assert.equal(ingestWordOoxmlToPlainText(rejectChris.oxml).trim(),
        'Background. The Services will process the Input to outputs');
}

// WP04: start, end, and complete deletion remain nested without empty carriers.
for (const scenario of [
    { original: 'Base Pending text', modified: 'Base text', deleted: 'Pending ' },
    { original: 'Base Pending text', modified: 'Base Pending', deleted: ' text' },
    { original: 'Base Pending text', modified: 'Base ', deleted: 'Pending text' }
]) {
    const source = paragraph(
        '<w:r><w:t xml:space="preserve">Base </w:t></w:r>'
        + insertion('<w:r><w:t>Pending text</w:t></w:r>')
    );
    const result = await apply(source, scenario.original, scenario.modified);
    const carrier = directElements(elements(parse(result.oxml), 'p')[0], 'ins')[0];
    assert(carrier);
    assert.equal(directElements(carrier, 'del').length, 1);
    assert.equal(revisionText(directElements(carrier, 'del')[0]), scenario.deleted);
}

// WP05: a baseline-to-insertion straddle stays split across schema contexts.
{
    const source = paragraph(
        '<w:r><w:t xml:space="preserve">Baseline start </w:t></w:r>'
        + insertion('<w:r><w:t>inserted finish.</w:t></w:r>')
    );
    const result = await apply(source, 'Baseline start inserted finish.', 'Baseline finish.');
    const p = elements(parse(result.oxml), 'p')[0];
    const topDeletion = directElements(p, 'del')[0];
    const carrier = directElements(p, 'ins')[0];
    assert.equal(revisionText(topDeletion), 'start ');
    assert.equal(revisionText(directElements(carrier, 'del')[0]), 'inserted ');
    assert.equal(revisionText(carrier).replace('inserted ', ''), 'finish.');
}

// WP05: insertion-to-baseline and two-carrier straddles keep one nested del per carrier.
{
    const source = paragraph(
        insertion('<w:r><w:t>Inserted start</w:t></w:r>')
        + '<w:r><w:t xml:space="preserve"> baseline finish.</w:t></w:r>'
    );
    const result = await apply(source, 'Inserted start baseline finish.', 'Inserted finish.');
    const p = elements(parse(result.oxml), 'p')[0];
    assert.equal(revisionText(directElements(directElements(p, 'ins')[0], 'del')[0]), 'start');
    assert.equal(revisionText(directElements(p, 'del')[0]), ' baseline ');
}
{
    const source = paragraph(
        insertion('<w:r><w:t xml:space="preserve">Barry removeA </w:t></w:r>', 1, 'Barry')
        + insertion('<w:r><w:t>removeC Carl</w:t></w:r>', 2, 'Carl')
    );
    const result = await apply(source, 'Barry removeA removeC Carl', 'Barry Carl');
    const carriers = directElements(elements(parse(result.oxml), 'p')[0], 'ins');
    assert.deepEqual(carriers.map(node => attr(node, 'author')), ['Barry', 'Carl']);
    assert.equal(revisionText(directElements(carriers[0], 'del')[0]), 'removeA ');
    assert.equal(revisionText(directElements(carriers[1], 'del')[0]), 'removeC ');
}

// WP05: replacement preserves the deletion in the left carrier and hoists the new insertion.
{
    const source = paragraph(insertion('<w:r><w:t>process generate outputs</w:t></w:r>'));
    const result = await apply(source, 'process generate outputs', 'process synthesize outputs', {
        pairReplacements: true
    });
    const revisions = directElements(elements(parse(result.oxml), 'p')[0]);
    assert.deepEqual(revisions.map(node => node.localName), ['ins', 'ins', 'ins']);
    assert.deepEqual(revisions.map(node => attr(node, 'author')), ['Barry', 'Anson', 'Barry']);
    const deletion = directElements(revisions[0], 'del')[0];
    assert.equal(revisionText(deletion), 'generate');
    assert.equal(revisionText(revisions[1]), 'synthesize');
    assert.equal(attr(deletion, 'date'), attr(revisions[1], 'date'));
}

// WP06: selective lifecycle operations preserve dependencies and normalize fragments.
{
    const source = paragraph(
        '<w:r><w:t xml:space="preserve">Base </w:t></w:r>'
        + insertion('<w:r><w:t>Original text</w:t></w:r>')
    );
    const sliced = await apply(source, 'Base Original text', 'Base Original NEW text');

    const rejectAnson = rejectTrackedChangesInOoxml(sliced.oxml, { author: 'Anson' });
    const rejectAnsonDoc = parse(rejectAnson.oxml);
    const restoredBarry = elements(rejectAnsonDoc, 'ins').filter(node => attr(node, 'author') === 'Barry');
    assert.equal(restoredBarry.length, 1, 'rejecting Anson should coalesce Barry fragments');
    assert.equal(revisionText(restoredBarry[0]), 'Original text');

    const acceptAnson = acceptTrackedChangesInOoxml(sliced.oxml, { author: 'Anson' });
    assert.equal(elements(parse(acceptAnson.oxml), 'ins').filter(node => attr(node, 'author') === 'Barry').length, 2);
    assert.equal(ingestWordOoxmlToPlainText(acceptAnson.oxml).trim(), 'Base Original NEW text');

    const rejectBarry = rejectTrackedChangesInOoxml(sliced.oxml, { author: 'Barry' });
    assert.equal(elements(parse(rejectBarry.oxml), 'ins').filter(node => attr(node, 'author') === 'Barry').length, 0);
    assert.equal(ingestWordOoxmlToPlainText(rejectBarry.oxml).trim(), 'Base NEW');
}
{
    const source = paragraph(
        '<w:r><w:t xml:space="preserve">Base </w:t></w:r>'
        + insertion('<w:r><w:t>keep remove keep</w:t></w:r>')
    );
    const sliced = await apply(source, 'Base keep remove keep', 'Base keep keep');

    const rejectAnson = rejectTrackedChangesInOoxml(sliced.oxml, { author: 'Anson' });
    const restoredCarrier = elements(parse(rejectAnson.oxml), 'ins')[0];
    assert.equal(elements(restoredCarrier, 'del').length, 0);
    assert.equal(revisionText(restoredCarrier), 'keep remove keep');

    const acceptBarry = acceptTrackedChangesInOoxml(sliced.oxml, { author: 'Barry' });
    const remainingDeletion = elements(parse(acceptBarry.oxml), 'del');
    assert.equal(remainingDeletion.length, 1);
    assert.equal(attr(remainingDeletion[0], 'author'), 'Anson');

    const rejectBarry = rejectTrackedChangesInOoxml(sliced.oxml, { author: 'Barry' });
    assert.equal(elements(parse(rejectBarry.oxml), 'del').length, 0, 'dependent nested deletion should cascade away');
    assert.equal(ingestWordOoxmlToPlainText(rejectBarry.oxml).trim(), 'Base');
}

console.log('PASS: WP03-WP06 cross-author slicing synthetic tests');
