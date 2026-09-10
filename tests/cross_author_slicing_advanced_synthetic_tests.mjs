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
import { extractCanonicalParagraphText } from '../core/paragraph-text.js';
import { applyOperationsToDocumentXml } from '../services/standalone-operation-runner.js';

const NS_W16DU = 'http://schemas.microsoft.com/office/word/2023/wordml/word16du';
const ALICE_DATE = '2026-09-08T09:00:00Z';
const ALICE_DATE_UTC = '2026-09-08T16:00:00Z';

function parse(xml) {
    const parsed = parseOoxmlSafe(xml, 'application/xml');
    assert.equal(parsed.error, null, `XML parse error: ${parsed.error?.message}`);
    return parsed.doc;
}

function elements(node, localName) {
    return Array.from(node.getElementsByTagNameNS(NS_W, localName));
}

function attr(node, localName) {
    return node.getAttributeNS(NS_W, localName) || node.getAttribute(`w:${localName}`);
}

function accepted(xml, author = null) {
    const opts = author ? { author } : { allAuthors: true };
    return ingestWordOoxmlToPlainText(acceptTrackedChangesInOoxml(xml, opts).oxml).trim();
}

function rejected(xml, author = null) {
    const opts = author ? { author } : { allAuthors: true };
    return ingestWordOoxmlToPlainText(rejectTrackedChangesInOoxml(xml, opts).oxml).trim();
}

function assertValidAndUnique(xml, label = '') {
    const validation = validateRedlineOoxml(xml);
    assert.equal(validation.valid, true, `${label} validation failed: ${JSON.stringify(validation.issues)}`);

    const doc = parse(xml);
    const insNodes = elements(doc, 'ins');
    const delNodes = elements(doc, 'del');
    const allNodes = insNodes.concat(delNodes);
    const ids = allNodes.map(node => attr(node, 'id')).filter(Boolean);
    assert.equal(new Set(ids).size, ids.length, `${label}: duplicate revision IDs found: ${ids.join(', ')}`);

    // Strict schema hierarchy checks:
    // 1. No ins inside ins
    for (const ins of insNodes) {
        assert.equal(elements(ins, 'ins').length, 0, `${label}: illegal <w:ins> nested inside <w:ins>`);
        assert.equal(elements(ins, 'moveFrom').length, 0, `${label}: illegal <w:moveFrom> inside <w:ins>`);
        assert.equal(elements(ins, 'moveTo').length, 0, `${label}: illegal <w:moveTo> inside <w:ins>`);
    }
    // 2. No ins, del, move inside del
    for (const del of delNodes) {
        assert.equal(elements(del, 'ins').length, 0, `${label}: illegal <w:ins> nested inside <w:del>`);
        assert.equal(elements(del, 'del').length, 0, `${label}: illegal <w:del> nested inside <w:del>`);
        assert.equal(elements(del, 'moveFrom').length, 0, `${label}: illegal <w:moveFrom> inside <w:del>`);
        assert.equal(elements(del, 'moveTo').length, 0, `${label}: illegal <w:moveTo> inside <w:del>`);
    }
}

async function applySlice(source, original, modified, extra = {}) {
    const result = await applyRedlineToOxml(source, original, modified, {
        author: 'Bob',
        existingRevisions: 'slice-cross-author',
        ...extra
    });
    assert.equal(result.status, 'ok', `applyRedlineToOxml failed: ${JSON.stringify(result.error)}`);
    assert.equal(result.hasChanges, true, 'expected mutation to yield changes');
    assertValidAndUnique(result.oxml, `applySlice(${extra.author || 'Bob'})`);
    return result;
}

// ============================================================================
// SUITE 1: ALREADY EDITED TEXT (STACKED & MULTI-ROUND MUTATIONS)
// ============================================================================

// 1.1: Multi-author stacked edits (Alice ins -> Bob ins inside Alice -> Charlie ins inside Bob -> Dana del from Charlie)
{
    const source = `<w:p xmlns:w="${NS_W}">
        <w:r><w:t xml:space="preserve">Contract: </w:t></w:r>
        <w:ins w:id="1" w:author="Alice" w:date="${ALICE_DATE}">
            <w:r><w:t>Supplier shall deliver products.</w:t></w:r>
        </w:ins>
    </w:p>`;

    // Turn 1: Bob slices in "promptly "
    const resBob = await applySlice(
        source,
        'Contract: Supplier shall deliver products.',
        'Contract: Supplier shall promptly deliver products.',
        { author: 'Bob' }
    );
    assert.equal(accepted(resBob.oxml), 'Contract: Supplier shall promptly deliver products.');

    // Turn 2: Charlie slices into Bob's insertion: "promptly certified "
    const resCharlie = await applySlice(
        resBob.oxml,
        'Contract: Supplier shall promptly deliver products.',
        'Contract: Supplier shall promptly certified deliver products.',
        { author: 'Charlie' }
    );
    assert.equal(accepted(resCharlie.oxml), 'Contract: Supplier shall promptly certified deliver products.');

    // Turn 3: Dana deletes "certified " from Charlie's insertion
    const resDana = await applySlice(
        resCharlie.oxml,
        'Contract: Supplier shall promptly certified deliver products.',
        'Contract: Supplier shall promptly deliver products.',
        { author: 'Dana' }
    );
    assert.equal(accepted(resDana.oxml), 'Contract: Supplier shall promptly deliver products.');
    assert.equal(rejected(resDana.oxml), 'Contract:');

    // Selective lifecycles:
    // Rejecting Dana restores Charlie's "certified "
    assert.equal(rejected(resDana.oxml, 'Dana'), 'Contract: Supplier shall promptly certified deliver products.');
    // Rejecting Charlie restores Bob's state
    assert.equal(rejected(resDana.oxml, 'Charlie'), 'Contract: Supplier shall promptly deliver products.');
}

// 1.2: Straddle deletion across multiple sliced sibling carriers
{
    const sliced = `<w:p xmlns:w="${NS_W}">
        <w:ins w:id="1" w:author="Alice" w:date="${ALICE_DATE}">
            <w:r><w:t xml:space="preserve">amended by </w:t></w:r>
        </w:ins>
        <w:ins w:id="2" w:author="Bob" w:date="2026-09-08T09:10:00Z">
            <w:r><w:t xml:space="preserve">MASTER </w:t></w:r>
        </w:ins>
        <w:ins w:id="3" w:author="Alice" w:date="${ALICE_DATE}">
            <w:r><w:t>Agreement.</w:t></w:r>
        </w:ins>
    </w:p>`;

    // Charlie deletes across Alice's left carrier, Bob's middle carrier, and Alice's right carrier:
    // "by MASTER Agree" -> leaving "amended " and "ment."
    const resStraddle = await applySlice(
        sliced,
        'amended by MASTER Agreement.',
        'amended ment.',
        { author: 'Charlie' }
    );
    assert.equal(accepted(resStraddle.oxml), 'amended ment.');

    const doc = parse(resStraddle.oxml);
    const delNodes = elements(doc, 'del');
    // Charlie should have created nested deletions within Alice's and Bob's carriers
    assert.equal(delNodes.length, 3, 'Charlie deletion should be partitioned into 3 carriers');
    for (const del of delNodes) {
        assert.equal(attr(del, 'author'), 'Charlie');
    }
}

// 1.3: Re-editing by original author in a mixed-author paragraph
{
    const mixed = `<w:p xmlns:w="${NS_W}">
        <w:r><w:t xml:space="preserve">Preamble </w:t></w:r>
        <w:ins w:id="1" w:author="Alice" w:date="${ALICE_DATE}">
            <w:r><w:t xml:space="preserve">first clause </w:t></w:r>
        </w:ins>
        <w:ins w:id="2" w:author="Bob" w:date="2026-09-08T09:10:00Z">
            <w:r><w:t xml:space="preserve">second clause </w:t></w:r>
        </w:ins>
        <w:r><w:t>closing.</w:t></w:r>
    </w:p>`;

    // Case A: Alice edits her own text ("first clause" -> "first revised clause")
    const resAliceOwn = await applySlice(
        mixed,
        'Preamble first clause second clause closing.',
        'Preamble first revised clause second clause closing.',
        { author: 'Alice' }
    );
    assert.equal(accepted(resAliceOwn.oxml), 'Preamble first revised clause second clause closing.');
    const docA = parse(resAliceOwn.oxml);
    const bobCarrier = elements(docA, 'ins').find(n => attr(n, 'author') === 'Bob');
    assert(bobCarrier, 'Bob carrier must be preserved when Alice edits her own insertion');

    // Case B: Alice edits Bob's clause ("second clause" -> "second amended clause")
    const resAliceEditsBob = await applySlice(
        mixed,
        'Preamble first clause second clause closing.',
        'Preamble first clause second amended clause closing.',
        { author: 'Alice' }
    );
    assert.equal(accepted(resAliceEditsBob.oxml), 'Preamble first clause second amended clause closing.');
    const docB = parse(resAliceEditsBob.oxml);
    const insAuthors = elements(docB, 'ins').map(n => attr(n, 'author'));
    assert(insAuthors.includes('Bob') && insAuthors.includes('Alice'));
}

// 1.4: Disjoint multi-replacement within the same already-edited paragraph
{
    const source = `<w:p xmlns:w="${NS_W}">
        <w:r><w:t xml:space="preserve">Spec: </w:t></w:r>
        <w:ins w:id="1" w:author="Alice" w:date="${ALICE_DATE}">
            <w:r><w:t xml:space="preserve">red car </w:t></w:r>
        </w:ins>
        <w:r><w:t xml:space="preserve">with </w:t></w:r>
        <w:ins w:id="2" w:author="Alice" w:date="${ALICE_DATE}">
            <w:r><w:t>blue truck</w:t></w:r>
        </w:ins>
    </w:p>`;

    const resDisjoint = await applySlice(
        source,
        'Spec: red car with blue truck',
        'Spec: crimson vehicle with navy transport',
        { author: 'Bob' }
    );
    assert.equal(accepted(resDisjoint.oxml), 'Spec: crimson vehicle with navy transport');
    assert.equal(rejected(resDisjoint.oxml), 'Spec: with');
}

// 1.5: Insertion at the exact boundary of a foreign deletion (<w:del>)
{
    const sourceWithDel = `<w:p xmlns:w="${NS_W}">
        <w:r><w:t xml:space="preserve">Status: </w:t></w:r>
        <w:del w:id="1" w:author="Alice" w:date="${ALICE_DATE}">
            <w:r><w:delText xml:space="preserve">pending </w:delText></w:r>
        </w:del>
        <w:r><w:t>approved</w:t></w:r>
    </w:p>`;

    // Current accepted view is "Status: approved"
    // Bob inserts "provisionally " before "approved"
    const resDelBoundary = await applySlice(
        sourceWithDel,
        'Status: approved',
        'Status: provisionally approved',
        { author: 'Bob' }
    );
    assert.equal(accepted(resDelBoundary.oxml), 'Status: provisionally approved');
    // Ensure Bob's ins is sibling to Alice's del, NOT nested inside it
    const doc = parse(resDelBoundary.oxml);
    const bobIns = elements(doc, 'ins').find(n => attr(n, 'author') === 'Bob');
    assert(bobIns);
    assert.equal(bobIns.parentNode.localName, 'p', 'Bob ins must be direct child of p');
}

// 1.6: Complete deletion of a previously sliced insertion down to empty
{
    const sliced = `<w:p xmlns:w="${NS_W}">
        <w:r><w:t xml:space="preserve">Item: </w:t></w:r>
        <w:ins w:id="1" w:author="Alice" w:date="${ALICE_DATE}">
            <w:r><w:t xml:space="preserve">first </w:t></w:r>
        </w:ins>
        <w:ins w:id="2" w:author="Bob" w:date="2026-09-08T09:10:00Z">
            <w:r><w:t>second</w:t></w:r>
        </w:ins>
    </w:p>`;

    // Charlie deletes "first second" completely
    const resWipe = await applySlice(
        sliced,
        'Item: first second',
        'Item:',
        { author: 'Charlie' }
    );
    assert.equal(accepted(resWipe.oxml), 'Item:');
    assert.equal(rejected(resWipe.oxml), 'Item:');
}

// ============================================================================
// SUITE 2: TABLES (<w:tbl>, <w:tr>, <w:tc>, <w:p>)
// ============================================================================

// 2.1: Slicing inside table cells with formatting and multi-cell batch operations
{
    const docXml = `<w:document xmlns:w="${NS_W}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <w:body>
            <w:tbl>
                <w:tr>
                    <w:tc>
                        <w:p>
                            <w:pPr><w:pStyle w:val="HeaderStyle"/></w:pPr>
                            <w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">Service: </w:t></w:r>
                            <w:ins w:id="1" w:author="Alice" w:date="${ALICE_DATE}">
                                <w:r><w:rPr><w:b/><w:i/></w:rPr><w:t>Cloud Hosting v1</w:t></w:r>
                            </w:ins>
                        </w:p>
                    </w:tc>
                    <w:tc>
                        <w:p>
                            <w:r><w:t xml:space="preserve">SLA: </w:t></w:r>
                            <w:ins w:id="2" w:author="Alice" w:date="${ALICE_DATE}">
                                <w:r><w:t>99.9% uptime</w:t></w:r>
                            </w:ins>
                        </w:p>
                    </w:tc>
                </w:tr>
            </w:tbl>
            <w:sectPr/>
        </w:body>
    </w:document>`;

    // Bob edits Cell 1 and Cell 2 in a single batch
    const batch = await applyOperationsToDocumentXml(docXml, [
        {
            type: 'redline',
            target: { exactText: 'Service: Cloud Hosting v1' },
            modified: 'Service: Cloud Hosting Enterprise v2',
            author: 'Bob',
            existingRevisions: 'slice-cross-author'
        },
        {
            type: 'redline',
            target: { exactText: 'SLA: 99.9% uptime' },
            modified: 'SLA: 99.99% high-availability uptime',
            author: 'Bob',
            existingRevisions: 'slice-cross-author'
        }
    ], 'Bob', null, { atomic: true, strictTargets: true });

    assert.equal(batch.status, 'ok');
    assert.equal(batch.results[0]?.status, 'applied');
    assert.equal(batch.results[1]?.status, 'applied');
    assertValidAndUnique(batch.documentXml, 'Table multi-cell batch');

    const doc = parse(batch.documentXml);
    assert.equal(elements(doc, 'tbl').length, 1);
    assert.equal(elements(doc, 'tc').length, 2);

    // Verify cell 1 accepted text and format inheritance
    assert(accepted(batch.documentXml).includes('Service: Cloud Hosting Enterprise v2'));
    assert(accepted(batch.documentXml).includes('SLA: 99.99% high-availability uptime'));
}

// 2.2: Multi-turn negotiation in a table cell with selective author lifecycles
{
    const baseTableDoc = `<w:document xmlns:w="${NS_W}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <w:body>
            <w:tbl>
                <w:tr>
                    <w:tc>
                        <w:p>
                            <w:r><w:t xml:space="preserve">Clause: </w:t></w:r>
                            <w:ins w:id="1" w:author="Alice" w:date="${ALICE_DATE}">
                                <w:r><w:t>The Supplier will deliver goods.</w:t></w:r>
                            </w:ins>
                        </w:p>
                    </w:tc>
                </w:tr>
            </w:tbl>
            <w:sectPr/>
        </w:body>
    </w:document>`;

    // Turn 1: Bob inserts "prompt "
    const turn1 = await applyOperationsToDocumentXml(baseTableDoc, [{
        type: 'redline',
        target: { exactText: 'Clause: The Supplier will deliver goods.' },
        modified: 'Clause: The Supplier will deliver prompt goods.',
        author: 'Bob',
        existingRevisions: 'slice-cross-author'
    }], 'Bob', null, { atomic: true, strictTargets: true });
    assert.equal(turn1.status, 'ok');

    // Turn 2: Charlie replaces "prompt goods" with "high-grade materials"
    const turn2 = await applyOperationsToDocumentXml(turn1.documentXml, [{
        type: 'redline',
        target: { exactText: 'Clause: The Supplier will deliver prompt goods.' },
        modified: 'Clause: The Supplier will deliver high-grade materials.',
        author: 'Charlie',
        existingRevisions: 'slice-cross-author'
    }], 'Charlie', null, { atomic: true, strictTargets: true });
    assert.equal(turn2.status, 'ok');
    assertValidAndUnique(turn2.documentXml, 'Table multi-turn turn2');

    // Lifecycle assertions on table cell:
    assert.equal(accepted(turn2.documentXml), 'Clause: The Supplier will deliver high-grade materials.');
    assert.equal(rejected(turn2.documentXml), 'Clause:');
    assert.equal(rejected(turn2.documentXml, 'Charlie'), 'Clause: The Supplier will deliver prompt goods.');
}

// 2.3: Multi-paragraph table cell with cross-author slicing
{
    const docXml = `<w:document xmlns:w="${NS_W}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <w:body>
            <w:tbl>
                <w:tr>
                    <w:tc>
                        <w:p>
                            <w:r><w:t xml:space="preserve">Paragraph 1: </w:t></w:r>
                            <w:ins w:id="1" w:author="Alice" w:date="${ALICE_DATE}">
                                <w:r><w:t>Draft clause 1.</w:t></w:r>
                            </w:ins>
                        </w:p>
                        <w:p>
                            <w:r><w:t xml:space="preserve">Paragraph 2: </w:t></w:r>
                            <w:ins w:id="2" w:author="Alice" w:date="${ALICE_DATE}">
                                <w:r><w:t>Draft clause 2.</w:t></w:r>
                            </w:ins>
                        </w:p>
                    </w:tc>
                </w:tr>
            </w:tbl>
            <w:sectPr/>
        </w:body>
    </w:document>`;

    const batch = await applyOperationsToDocumentXml(docXml, [
        {
            type: 'redline',
            target: { exactText: 'Paragraph 1: Draft clause 1.' },
            modified: 'Paragraph 1: Finalized clause 1.',
            author: 'Bob',
            existingRevisions: 'slice-cross-author'
        },
        {
            type: 'redline',
            target: { exactText: 'Paragraph 2: Draft clause 2.' },
            modified: 'Paragraph 2: Finalized clause 2.',
            author: 'Bob',
            existingRevisions: 'slice-cross-author'
        }
    ], 'Bob', null, { atomic: true, strictTargets: true });

    assert.equal(batch.status, 'ok');
    assertValidAndUnique(batch.documentXml, 'Multi-paragraph table cell');
    const doc = parse(batch.documentXml);
    assert.equal(elements(doc, 'tc').length, 1);
    assert.equal(elements(doc, 'p').length, 2);
}

// ============================================================================
// SUITE 3: BULLETS AND LISTS (<w:numPr>)
// ============================================================================

// 3.1: Multi-item bullet list with cross-author slicing preserving <w:numPr>
{
    const docXml = `<w:document xmlns:w="${NS_W}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <w:body>
            <w:p>
                <w:pPr>
                    <w:numPr><w:ilvl w:val="0"/><w:numId w:val="5"/></w:numPr>
                </w:pPr>
                <w:r><w:t xml:space="preserve">1. First requirement: </w:t></w:r>
                <w:ins w:id="1" w:author="Alice" w:date="${ALICE_DATE}">
                    <w:r><w:t>execute standard security audit.</w:t></w:r>
                </w:ins>
            </w:p>
            <w:p>
                <w:pPr>
                    <w:numPr><w:ilvl w:val="1"/><w:numId w:val="5"/></w:numPr>
                </w:pPr>
                <w:r><w:t xml:space="preserve">a. Audit cadence: </w:t></w:r>
                <w:ins w:id="2" w:author="Alice" w:date="${ALICE_DATE}">
                    <w:r><w:t>conducted annually by third party.</w:t></w:r>
                </w:ins>
            </w:p>
            <w:sectPr/>
        </w:body>
    </w:document>`;

    // Bob edits both bullet items in a batch
    const batch = await applyOperationsToDocumentXml(docXml, [
        {
            type: 'redline',
            target: { exactText: '1. First requirement: execute standard security audit.' },
            modified: '1. First requirement: execute standard SOC-2 security audit.',
            author: 'Bob',
            existingRevisions: 'slice-cross-author'
        },
        {
            type: 'redline',
            target: { exactText: 'a. Audit cadence: conducted annually by third party.' },
            modified: 'a. Audit cadence: conducted semi-annually by certified third party.',
            author: 'Bob',
            existingRevisions: 'slice-cross-author'
        }
    ], 'Bob', null, { atomic: true, strictTargets: true });

    assert.equal(batch.status, 'ok');
    assertValidAndUnique(batch.documentXml, 'Bullet list batch');

    // Verify <w:numPr> is preserved in both paragraphs
    const doc = parse(batch.documentXml);
    const pList = elements(doc, 'p');
    assert.equal(pList.length, 2);
    for (const p of pList) {
        const numPr = elements(p, 'numPr');
        assert.equal(numPr.length, 1, '<w:numPr> must be preserved untouched');
    }
}

// 3.2: Multi-turn negotiation on a bullet item with selective author lifecycles
{
    const baseBulletDoc = `<w:document xmlns:w="${NS_W}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <w:body>
            <w:p>
                <w:pPr>
                    <w:numPr><w:ilvl w:val="0"/><w:numId w:val="99"/></w:numPr>
                </w:pPr>
                <w:r><w:t xml:space="preserve">Term A: </w:t></w:r>
                <w:ins w:id="10" w:author="Alice" w:date="${ALICE_DATE}">
                    <w:r><w:t>Payment within thirty days of invoice.</w:t></w:r>
                </w:ins>
            </w:p>
            <w:sectPr/>
        </w:body>
    </w:document>`;

    // Turn 1: Bob edits "thirty days" to "forty-five calendar days"
    const turn1 = await applyOperationsToDocumentXml(baseBulletDoc, [{
        type: 'redline',
        target: { exactText: 'Term A: Payment within thirty days of invoice.' },
        modified: 'Term A: Payment within forty-five calendar days of invoice.',
        author: 'Bob',
        existingRevisions: 'slice-cross-author'
    }], 'Bob', null, { atomic: true, strictTargets: true });
    assert.equal(turn1.status, 'ok');

    // Turn 2: Charlie edits "Payment within" to "Prompt payment within"
    const turn2 = await applyOperationsToDocumentXml(turn1.documentXml, [{
        type: 'redline',
        target: { exactText: 'Term A: Payment within forty-five calendar days of invoice.' },
        modified: 'Term A: Prompt payment within forty-five calendar days of invoice.',
        author: 'Charlie',
        existingRevisions: 'slice-cross-author'
    }], 'Charlie', null, { atomic: true, strictTargets: true });
    assert.equal(turn2.status, 'ok');

    // Turn 3: Dana deletes "calendar "
    const turn3 = await applyOperationsToDocumentXml(turn2.documentXml, [{
        type: 'redline',
        target: { exactText: 'Term A: Prompt payment within forty-five calendar days of invoice.' },
        modified: 'Term A: Prompt payment within forty-five days of invoice.',
        author: 'Dana',
        existingRevisions: 'slice-cross-author'
    }], 'Dana', null, { atomic: true, strictTargets: true });
    assert.equal(turn3.status, 'ok');
    assertValidAndUnique(turn3.documentXml, 'Bullet multi-turn turn3');

    // Lifecycle assertions:
    assert.equal(accepted(turn3.documentXml), 'Term A: Prompt payment within forty-five days of invoice.');
    assert.equal(rejected(turn3.documentXml), 'Term A:');
    assert.equal(rejected(turn3.documentXml, 'Dana'), 'Term A: Prompt payment within forty-five calendar days of invoice.');
    assert.equal(rejected(turn3.documentXml, 'Charlie'), 'Term A: Payment within forty-five days of invoice.');
}

// 3.3: Bullet item with <w:tab/> character and existing revisions
{
    const bulletWithTab = `<w:p xmlns:w="${NS_W}">
        <w:pPr>
            <w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>
        </w:pPr>
        <w:r><w:t>Section 1</w:t><w:tab/></w:r>
        <w:ins w:id="1" w:author="Alice" w:date="${ALICE_DATE}">
            <w:r><w:t>Scope of Work</w:t></w:r>
        </w:ins>
    </w:p>`;

    const resTab = await applySlice(
        bulletWithTab,
        'Section 1\tScope of Work',
        'Section 1\tDetailed Scope of Work',
        { author: 'Bob' }
    );
    const acceptedDoc = parse(acceptTrackedChangesInOoxml(resTab.oxml).oxml);
    const acceptedP = elements(acceptedDoc, 'p')[0];
    assert.equal(extractCanonicalParagraphText(acceptedP), 'Section 1\tDetailed Scope of Work');
    const doc = parse(resTab.oxml);
    assert.equal(elements(doc, 'tab').length, 1, '<w:tab/> must be preserved');
    assert.equal(elements(doc, 'numPr').length, 1, '<w:numPr> must be preserved');
}

console.log('PASS: cross_author_slicing_advanced_synthetic_tests.mjs');
