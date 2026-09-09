import assert from 'assert/strict';

import './setup-xml-provider.mjs';
import { parseOoxmlSafe } from '../adapters/xml-adapter.js';
import { NS_W, RevisionIdAllocator } from '../core/types.js';
import { splitTrackChangeCarrier } from '../engine/surgical-run-splitting.js';

const NS_W16DU = 'http://schemas.microsoft.com/office/word/2023/wordml/word16du';
const DATE = '2026-09-08T09:29:00Z';
const DATE_UTC = '2026-09-08T16:29:00Z';

function parseCarrier(content, attributes = '') {
    const xml = `<w:p xmlns:w="${NS_W}" xmlns:w16du="${NS_W16DU}">`
        + `<w:ins w:id="7" w:author="Barry Plasteras" w:date="${DATE}" `
        + `w16du:dateUtc="${DATE_UTC}" ${attributes}>${content}</w:ins></w:p>`;
    const { doc, error } = parseOoxmlSafe(xml, 'application/xml');
    assert.equal(error, null);
    return { doc, carrier: doc.getElementsByTagNameNS(NS_W, 'ins')[0] };
}

function wordAttribute(element, localName) {
    return element.getAttributeNS(NS_W, localName) || element.getAttribute(`w:${localName}`);
}

function text(element) {
    return Array.from(element.getElementsByTagNameNS(NS_W, 't'))
        .map(node => node.textContent || '')
        .join('');
}

function revisionIds(element) {
    const descendants = ['ins', 'rPrChange'].flatMap(localName => {
        return Array.from(element.getElementsByTagNameNS(NS_W, localName));
    });
    return [element, ...descendants]
        .filter(node => node.namespaceURI === NS_W && ['ins', 'rPrChange'].includes(node.localName))
        .map(node => Number(wordAttribute(node, 'id')));
}

// Interior split across multiple runs preserves formatting and carrier metadata.
{
    const { doc, carrier } = parseCarrier(
        '<w:r><w:t xml:space="preserve">amended </w:t></w:r>'
        + `<w:r><w:rPr><w:b/><w:rPrChange w:id="8" w:author="Formatter" w:date="${DATE}">`
        + '<w:rPr/></w:rPrChange></w:rPr><w:t>Agreement</w:t></w:r>'
    );
    const allocator = new RevisionIdAllocator();
    const recordedRevisions = [];
    allocator._receiptCollector = {
        recordRevision(id, kind) {
            recordedRevisions.push([id, kind]);
        }
    };
    allocator.seed(doc);
    const result = splitTrackChangeCarrier(doc, carrier, 12, allocator);

    assert.equal(text(result.leftCarrier), 'amended Agre');
    assert.equal(text(result.rightCarrier), 'ement');
    assert.equal(wordAttribute(result.leftCarrier, 'id'), '7');
    assert.equal(wordAttribute(result.rightCarrier, 'id'), '1001');
    assert.equal(wordAttribute(result.rightCarrier, 'author'), 'Barry Plasteras');
    assert.equal(wordAttribute(result.rightCarrier, 'date'), DATE);
    assert.equal(result.rightCarrier.getAttributeNS(NS_W16DU, 'dateUtc'), DATE_UTC);
    assert.equal(result.rightCarrier.getElementsByTagNameNS(NS_W, 'b').length, 1);
    assert.deepEqual(revisionIds(result.leftCarrier), [7, 8]);
    assert.deepEqual(revisionIds(result.rightCarrier), [1001, 1000]);
    assert.deepEqual(recordedRevisions, [[1000, 'rPrChange'], [1001, 'ins']]);
    assert.equal(text(carrier), 'amended Agreement', 'source carrier must not be mutated');
}

// Text-like run children use character offsets consistently.
{
    const { doc, carrier } = parseCarrier(
        '<w:r><w:t>A</w:t><w:tab/><w:t>B</w:t><w:br/><w:t>C</w:t></w:r>'
    );
    const allocator = new RevisionIdAllocator();
    allocator.seed(doc);
    const { leftCarrier, rightCarrier } = splitTrackChangeCarrier(doc, carrier, 3, allocator);
    assert.equal(leftCarrier.getElementsByTagNameNS(NS_W, 'tab').length, 1);
    assert.equal(leftCarrier.getElementsByTagNameNS(NS_W, 'br').length, 0);
    assert.equal(rightCarrier.getElementsByTagNameNS(NS_W, 'br').length, 1);
    assert.equal(text(leftCarrier), 'AB');
    assert.equal(text(rightCarrier), 'C');
}

// Exact boundaries do not emit empty containers or consume an ID.
{
    const { doc, carrier } = parseCarrier('<w:r><w:t>Clause</w:t></w:r>');
    const allocator = new RevisionIdAllocator();
    allocator.seed(doc);
    const atStart = splitTrackChangeCarrier(doc, carrier, 0, allocator);
    assert.equal(atStart.leftCarrier, null);
    assert.equal(wordAttribute(atStart.rightCarrier, 'id'), '7');
    assert.equal(allocator.next(), 1000);

    const endAllocator = new RevisionIdAllocator();
    endAllocator.seed(doc);
    const atEnd = splitTrackChangeCarrier(doc, carrier, 6, endAllocator);
    assert.equal(wordAttribute(atEnd.leftCarrier, 'id'), '7');
    assert.equal(atEnd.rightCarrier, null);
    assert.equal(endAllocator.next(), 1000);
}

// Invalid offsets and non-carriers fail explicitly.
{
    const { doc, carrier } = parseCarrier('<w:r><w:t>Clause</w:t></w:r>');
    assert.throws(() => splitTrackChangeCarrier(doc, carrier, 7), RangeError);
    assert.throws(() => splitTrackChangeCarrier(doc, carrier, 1.5), RangeError);
    assert.throws(() => splitTrackChangeCarrier(doc, carrier.firstChild, 1), TypeError);
}

console.log('PASS: cross-author tracked-change carrier splitting');
