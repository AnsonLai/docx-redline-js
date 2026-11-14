/**
 * A pending same-author run-formatting change (a tracked bold, `w:rPrChange`) survives later
 * same-author work on the paragraph:
 *   - a text edit keeps the formatting in Accept All and its snapshot for Reject All;
 *   - a further format operation on the same run keeps the first snapshot, so Reject All restores
 *     the original formatting, and overlapping format operations can share one batch.
 */
import './setup-xml-provider.mjs';
import assert from 'node:assert/strict';
import {
    acceptTrackedChangesInOoxml,
    parseOoxmlSafe,
    rejectTrackedChangesInOoxml
} from '../index.js';
import { applyOperationsToDocumentXml } from '../services/standalone-operation-runner.js';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const options = {
    atomic: true,
    structuredContent: true,
    pairReplacements: true,
    generateRedlines: true,
    strictTargets: true,
    onInfo: () => {},
    onWarn: () => {}
};

/** Runs of the first paragraph as `[text, 'b' | 'i' | 'u' ...]` after resolving every revision. */
function resolvedRuns(documentXml, resolve) {
    const resolved = resolve(documentXml, { allAuthors: true });
    assert.notEqual(resolved.status, 'error');
    assert.doesNotMatch(resolved.oxml, /<w:(ins|del|rPrChange)\b/);
    const { doc } = parseOoxmlSafe(resolved.oxml, 'text/xml');
    const merged = [];
    for (const run of Array.from(doc.getElementsByTagNameNS(W, 'r'))) {
        const text = Array.from(run.getElementsByTagNameNS(W, 't')).map(t => t.textContent).join('');
        if (!text) continue;
        const rPr = run.getElementsByTagNameNS(W, 'rPr')[0];
        const on = name => {
            const el = rPr?.getElementsByTagNameNS(W, name)[0];
            if (!el) return false;
            const val = el.getAttribute('w:val');
            return val !== '0' && val !== 'false' && val !== 'none';
        };
        const format = ['b', 'i', 'u'].filter(on).join('');
        const last = merged[merged.length - 1];
        if (last && last[1] === format) last[0] += text;
        else merged.push([text, format]);
    }
    return merged;
}

function revisionIds(documentXml) {
    return [...documentXml.matchAll(/<w:(?:ins|del|rPrChange)\b[^>]*\bw:id="(\d+)"/g)].map(match => match[1]);
}

function assertUniqueRevisionIds(documentXml) {
    const ids = revisionIds(documentXml);
    assert.equal(new Set(ids).size, ids.length, `revision ids are unique: ${ids.join(',')}`);
    assert.doesNotMatch(documentXml, /w:id="\d+"[^>]*\bw:id="/, 'no element carries two w:id attributes');
}

// --- A text edit in a paragraph with a pending same-author bold -----------------------------

const NOTICE = 'This notice explains how Maple Clinic collects your information.';
const noticeDoc = `<w:document xmlns:w="${W}"><w:body><w:p><w:r><w:t xml:space="preserve">This notice explains how </w:t></w:r><w:r><w:rPr><w:b/><w:rPrChange w:id="5" w:author="A" w:date="2026-10-31T00:00:00Z"><w:rPr/></w:rPrChange></w:rPr><w:t>Maple Clinic</w:t></w:r><w:r><w:t xml:space="preserve"> collects your information.</w:t></w:r></w:p></w:body></w:document>`;
const noticeTarget = { exactText: NOTICE, index: 1 };

// A replacement elsewhere in the paragraph keeps the bold.
{
    const result = await applyOperationsToDocumentXml(noticeDoc, [{
        type: 'replace',
        target: noticeTarget,
        replacements: [{ find: 'collects', replace: 'gathers' }]
    }], 'A', {}, options);
    assert.equal(result.error, undefined, JSON.stringify(result.results?.[0]?.error?.message));
    assertUniqueRevisionIds(result.documentXml);
    assert.deepEqual(resolvedRuns(result.documentXml, acceptTrackedChangesInOoxml), [
        ['This notice explains how ', ''],
        ['Maple Clinic', 'b'],
        [' gathers your information.', '']
    ]);
    assert.deepEqual(resolvedRuns(result.documentXml, rejectTrackedChangesInOoxml), [[NOTICE, '']]);
}

// A replacement inside the bold phrase splits its run: the text stays bold and ids stay unique.
{
    const result = await applyOperationsToDocumentXml(noticeDoc, [{
        type: 'replace',
        target: noticeTarget,
        replacements: [{ find: 'Maple', replace: 'Oak' }]
    }], 'A', {}, options);
    assert.equal(result.error, undefined, JSON.stringify(result.results?.[0]?.error?.message));
    assertUniqueRevisionIds(result.documentXml);
    assert.deepEqual(resolvedRuns(result.documentXml, acceptTrackedChangesInOoxml), [
        ['This notice explains how ', ''],
        ['Oak Clinic', 'b'],
        [' collects your information.', '']
    ]);
    assert.deepEqual(resolvedRuns(result.documentXml, rejectTrackedChangesInOoxml), [[NOTICE, '']]);
}

// A complete-paragraph redline keeps the bold too.
{
    const result = await applyOperationsToDocumentXml(noticeDoc, [{
        type: 'redline',
        target: noticeTarget,
        modified: 'This notice explains how Maple Clinic gathers and stores your information.',
        structuredContent: false
    }], 'A', {}, options);
    assert.equal(result.error, undefined, JSON.stringify(result.results?.[0]?.error?.message));
    assert.deepEqual(resolvedRuns(result.documentXml, acceptTrackedChangesInOoxml), [
        ['This notice explains how ', ''],
        ['Maple Clinic', 'b'],
        [' gathers and stores your information.', '']
    ]);
    assert.deepEqual(resolvedRuns(result.documentXml, rejectTrackedChangesInOoxml), [[NOTICE, '']]);
}

// Another author is still refused by the existing-revisions policy.
{
    const result = await applyOperationsToDocumentXml(noticeDoc, [{
        type: 'replace',
        target: noticeTarget,
        replacements: [{ find: 'collects', replace: 'gathers' }]
    }], 'B', {}, options);
    assert.equal(result.results[0].error?.code, 'EXISTING_REVISIONS');
    assert.equal(result.documentXml, noticeDoc);
}

// --- Format operations on a run whose formatting is already pending --------------------------

const RECITAL = 'RECITALS: WHEREAS, Provider owns it.';
const recitalDoc = `<w:document xmlns:w="${W}"><w:body><w:p><w:r><w:t>${RECITAL}</w:t></w:r></w:p></w:body></w:document>`;
const recitalTarget = { exactText: RECITAL, index: 1 };
const format = (textToFormat, properties) => ({ type: 'format', target: recitalTarget, textToFormat, properties });

async function applyFormats(operations, source = recitalDoc) {
    const result = await applyOperationsToDocumentXml(source, operations, 'A', {}, options);
    assert.equal(result.error, undefined, `${result.error?.code}: ${result.error?.message}`);
    assertUniqueRevisionIds(result.documentXml);
    assert.deepEqual(
        resolvedRuns(result.documentXml, rejectTrackedChangesInOoxml),
        [[RECITAL, '']],
        'Reject All restores the unformatted source'
    );
    return result;
}

// Overlapping spans in one batch (was RECEIPT_RECONCILIATION_FAILED).
{
    const result = await applyFormats([
        format('WHEREAS', { bold: true }),
        format('WHEREAS, Provider', { bold: true })
    ]);
    assert.deepEqual(resolvedRuns(result.documentXml, acceptTrackedChangesInOoxml), [
        ['RECITALS: ', ''],
        ['WHEREAS, Provider', 'b'],
        [' owns it.', '']
    ]);
}

// Bold on, then bold off for the same word, in one batch.
{
    const result = await applyFormats([
        format('WHEREAS', { bold: true }),
        format('WHEREAS', { bold: false })
    ]);
    assert.deepEqual(resolvedRuns(result.documentXml, acceptTrackedChangesInOoxml), [[RECITAL, '']]);
}

// The same two steps as separate calls keep the first snapshot as well.
{
    const first = await applyFormats([format('WHEREAS', { bold: true })]);
    const second = await applyFormats([format('WHEREAS, Provider', { bold: true })], first.documentXml);
    assert.deepEqual(resolvedRuns(second.documentXml, acceptTrackedChangesInOoxml), [
        ['RECITALS: ', ''],
        ['WHEREAS, Provider', 'b'],
        [' owns it.', '']
    ]);
}

// Formatting part of an already-pending run splits it into pieces with their own revision ids.
{
    const result = await applyFormats([
        format('WHEREAS', { bold: true }),
        format('ERE', { italic: true })
    ]);
    assert.deepEqual(resolvedRuns(result.documentXml, acceptTrackedChangesInOoxml), [
        ['RECITALS: ', ''],
        ['WH', 'b'],
        ['ERE', 'bi'],
        ['AS', 'b'],
        [', Provider owns it.', '']
    ]);
}

// --- Pending formatting mixed with other pending same-author revisions -----------------------

const CLINIC = 'Maple Clinic collects your information today.';
const clinicDoc = `<w:document xmlns:w="${W}"><w:body><w:p><w:r><w:t>${CLINIC}</w:t></w:r></w:p></w:body></w:document>`;

async function applyStages(stages) {
    let xml = clinicDoc;
    for (const operation of stages) {
        const result = await applyOperationsToDocumentXml(xml, [operation], 'A', {}, options);
        assert.equal(result.error, undefined, `${operation.type}: ${JSON.stringify(result.results?.[0]?.error?.message)}`);
        xml = result.documentXml;
    }
    assertUniqueRevisionIds(xml);
    return xml;
}
const firstParagraph = xml => xml.slice(0, xml.indexOf('</w:p>'));

// A pending text edit and a pending bold, then a second text edit: the bold stays pending.
{
    const xml = await applyStages([
        { type: 'replace', target: { exactText: CLINIC }, replacements: [{ find: 'collects', replace: 'gathers' }] },
        { type: 'format', target: { exactText: 'Maple Clinic gathers your information today.' }, textToFormat: 'Maple Clinic', properties: { bold: true } },
        { type: 'replace', target: { exactText: 'Maple Clinic gathers your information today.' }, replacements: [{ find: 'today', replace: 'now' }] }
    ]);
    assert.deepEqual(resolvedRuns(xml, acceptTrackedChangesInOoxml), [
        ['Maple Clinic', 'b'],
        [' gathers your information now.', '']
    ]);
    assert.deepEqual(resolvedRuns(xml, rejectTrackedChangesInOoxml), [[CLINIC, '']]);
}

// A pending bold and a pending paragraph alignment, then a text edit: both stay pending.
{
    const xml = await applyStages([
        { type: 'format', target: { exactText: CLINIC }, textToFormat: 'Maple Clinic', properties: { bold: true } },
        { type: 'paragraph-format', target: { exactText: CLINIC }, properties: { alignment: 'center' } },
        { type: 'replace', target: { exactText: CLINIC }, replacements: [{ find: 'today', replace: 'now' }] }
    ]);
    assert.match(xml, /<w:jc w:val="center"\/><w:pPrChange /);
    assert.deepEqual(resolvedRuns(xml, acceptTrackedChangesInOoxml), [
        ['Maple Clinic', 'b'],
        [' collects your information now.', '']
    ]);
    assert.match(acceptTrackedChangesInOoxml(xml, { allAuthors: true }).oxml, /<w:jc w:val="center"\/>/);
    assert.deepEqual(resolvedRuns(xml, rejectTrackedChangesInOoxml), [[CLINIC, '']]);
    assert.doesNotMatch(rejectTrackedChangesInOoxml(xml, { allAuthors: true }).oxml, /<w:jc /);
}

// A pending bold in a paragraph whose mark became an insertion (paragraphs appended after it).
{
    const xml = await applyStages([
        { type: 'format', target: { exactText: CLINIC }, textToFormat: 'Maple Clinic', properties: { bold: true } },
        { type: 'redline', target: { exactText: CLINIC }, modified: `${CLINIC}
Next.`, structuredContent: false },
        { type: 'replace', target: { exactText: CLINIC }, replacements: [{ find: 'today', replace: 'now' }] }
    ]);
    assert.match(firstParagraph(xml), /<w:rPrChange /);
    assert.match(firstParagraph(acceptTrackedChangesInOoxml(xml, { allAuthors: true }).oxml), /<w:b w:val="1"\/>/);
    assert.deepEqual(resolvedRuns(xml, rejectTrackedChangesInOoxml), [[CLINIC, '']]);
}

console.log('PASS: pending same-author run formatting');
