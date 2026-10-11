import assert from 'node:assert/strict';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { openDocx } from '../node/index.js';
import { executeCli } from '../node/cli.js';
import { validateDocumentOperation } from '../services/document-operation-contract.js';
import { normalizeErrorWithRecovery } from '../services/error-recovery.js';
import {
    parseLiteralListMarker,
    readSourceListLevel,
    resolveListDiagnosticPolicy
} from '../services/list-conversion-diagnostics.js';

// Interim list safety diagnostics (canonical list operations plan, WP-1). These are independent of
// the Reject All verifier: none of them may be reported as REVISION_FIDELITY_VIOLATION.

const fixtureFile = fileURLToPath(new URL('./fixtures/agentic-lists/nested-lists-source.docx', import.meta.url));
const fixture = await readFile(fixtureFile);
const PLAIN = 'Plain paragraph before bullet list.';
const BULLET = 'Bullet Root A';
const NUMBERED = 'Number Root A';

const apply = (operations, options = {}) => openDocx(fixture).applyOperations(operations, { author: 'Diag', ...options });
const textOf = async result => openDocx(result.toBuffer()).inspect({}).paragraphs.map(p => p.text);

// Marker detection is conservative: prose that merely begins with a number or an abbreviation is not a marker.
assert.equal(parseLiteralListMarker('7 days after notice'), null);
assert.equal(parseLiteralListMarker('2.1 Definitions'), null);
assert.equal(parseLiteralListMarker('Mr. Smith agrees'), null);
assert.equal(parseLiteralListMarker('Header'), null);
assert.equal(parseLiteralListMarker('first\nA. second'), null);
assert.deepEqual(
    ['A. Header', '1. Header', 'iv. Header', '- Header'].map(text => parseLiteralListMarker(text).numFmt),
    ['upperLetter', 'decimal', 'lowerRoman', 'bullet']
);
assert.equal(resolveListDiagnosticPolicy({}), 'warn');
assert.equal(resolveListDiagnosticPolicy({ atomic: true }), 'error');
assert.equal(resolveListDiagnosticPolicy({ atomic: true, strictListConversion: false }), 'warn');
assert.equal(resolveListDiagnosticPolicy({ strictListConversion: true }), 'error');
assert.equal(resolveListDiagnosticPolicy({ atomic: true, listMarkers: 'literal' }), 'off');

// Source level lookup reads the numbering definition the paragraph keeps.
const sourceNumbering = openDocx(fixture).inspect({}).paragraphs.find(p => p.text === BULLET).list;
assert.equal(sourceNumbering.format, 'bullet');

// (a) Plain "Header" -> "A. Header": ok, but no longer silent.
{
    const result = await apply([{ type: 'redline', target: { exactText: PLAIN }, modified: `A. ${PLAIN}` }]);
    assert.equal(result.status, 'ok');
    const item = result.results[0];
    assert.equal(item.status, 'applied');
    assert.equal(item.diagnostics.length, 1);
    const [diagnostic] = item.diagnostics;
    assert.equal(diagnostic.code, 'LIST_MARKER_NOT_BOUND');
    assert.equal(diagnostic.severity, 'warning');
    assert.equal(diagnostic.listMarker, 'A.');
    assert.deepEqual(diagnostic.requestedList, { kind: 'numbered', format: 'upperLetter' });
    assert.equal(diagnostic.recovery.field, 'listMarkers');
    assert.ok(item.warnings.some(warning => warning.startsWith('LIST_MARKER_NOT_BOUND: ')));
    assert.ok((await textOf(result)).includes(`A. ${PLAIN}`), 'the marker was written as literal text');
    assert.equal(openDocx(result.toBuffer()).inspect({}).paragraphs.find(p => p.text === `A. ${PLAIN}`).list, null);
    assert.ok(!JSON.stringify(result).includes('REVISION_FIDELITY_VIOLATION'));
}

// Under the strict policy (atomic) the same request is refused and nothing is written.
{
    const result = await apply([{ type: 'redline', target: { exactText: PLAIN }, modified: `1. ${PLAIN}` }], { atomic: true });
    assert.equal(result.status, 'error');
    assert.equal(result.written, false);
    const error = result.results[0].error;
    assert.equal(error.code, 'UNSUPPORTED_LIST_CONVERSION');
    assert.equal(error.stage, 'list-conversion');
    assert.equal(error.category, 'request_fixable');
    assert.equal(error.recovery.action, 'use_list_format_convert_plain');
    assert.equal(error.recovery.sameArgumentsSafe, false);
    assert.equal(error.recovery.field, 'listMarkers');
    assert.equal(error.listMarker, '1.');
    assert.deepEqual(error.requestedList, { kind: 'numbered', format: 'decimal' });
    assert.notEqual(error.code, 'REVISION_FIDELITY_VIOLATION');
    assert.deepEqual(result.toBuffer(), fixture);
}

// The explicit opt-out accepts literal text under both policies, without diagnostics.
for (const options of [{}, { atomic: true }]) {
    const result = await apply(
        [{ type: 'redline', target: { exactText: PLAIN }, modified: `A. ${PLAIN}`, listMarkers: 'literal' }],
        options
    );
    assert.equal(result.status, 'ok');
    assert.equal(result.results[0].diagnostics, undefined);
    assert.equal(result.results[0].warnings, undefined);
}

// Only the literal opt-out value is accepted.
assert.equal(validateDocumentOperation({ type: 'redline', target: 'x', modified: 'y', listMarkers: 'bind' }).error.code, 'INVALID_OPERATION');
assert.equal(validateDocumentOperation({ type: 'redline', target: 'x', modified: 'y', listMarkers: 'literal' }).valid, true);

// No false positives: an existing equivalent marker, plain edits, list items and prose numbers.
for (const [target, modified] of [
    ['Untouched bold sentinel', 'Untouched bold sentinel, amended'],
    [PLAIN, `7 days: ${PLAIN}`],
    [PLAIN, `Mr. Smith: ${PLAIN}`]
]) {
    const result = await apply([{ type: 'redline', target: { exactText: target }, modified }], { atomic: true });
    assert.equal(result.status, 'ok', `${target} -> ${modified}`);
    assert.equal(result.results[0].diagnostics, undefined);
}
{
    // A literal "1." heading edited to another "1." heading keeps its (literal) marker without noise.
    const documentXml = (await import('../services/standalone-operation-runner.js'));
    const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
    const bare = `<w:document xmlns:w="${W}"><w:body><w:p><w:r><w:t>1. Old heading</w:t></w:r></w:p></w:body></w:document>`;
    const result = await documentXml.applyOperationsToDocumentXml(bare, [
        { type: 'redline', target: { exactText: '1. Old heading' }, modified: '1. New heading' }
    ], 'Diag', null, { atomic: true });
    assert.equal(result.status, 'ok', JSON.stringify(result.error || null));
    assert.equal(result.results[0].diagnostics, undefined);

    // The bare-XML entry point produces the same diagnostics for the unbound marker.
    const plain = `<w:document xmlns:w="${W}"><w:body><w:p><w:r><w:t>Header</w:t></w:r></w:p></w:body></w:document>`;
    const warned = await documentXml.applyOperationsToDocumentXml(plain, [
        { type: 'redline', target: { exactText: 'Header' }, modified: 'A. Header' }
    ], 'Diag', null, {});
    assert.equal(warned.status, 'ok');
    assert.equal(warned.results[0].diagnostics[0].code, 'LIST_MARKER_NOT_BOUND');
    assert.match(warned.documentXml, /A\. Header|A\. /);
    const refused = await documentXml.applyOperationsToDocumentXml(plain, [
        { type: 'redline', target: { exactText: 'Header' }, modified: 'A. Header' }
    ], 'Diag', null, { atomic: true });
    assert.equal(refused.status, 'error');
    assert.equal(refused.results[0].error.code, 'UNSUPPORTED_LIST_CONVERSION');
}

// (b) A bullet item asked to become numbered keeps its old definition: report it.
{
    const operation = { type: 'list-change', target: { exactText: BULLET }, modified: `1. ${BULLET}\n2. Bullet Insertion Anchor` };
    const result = await apply([operation]);
    assert.equal(result.status, 'ok');
    const [diagnostic] = result.results[0].diagnostics;
    assert.equal(diagnostic.code, 'LIST_FORMAT_NOT_APPLIED');
    assert.equal(diagnostic.severity, 'warning');
    assert.deepEqual(diagnostic.requestedList, { kind: 'numbered', format: 'decimal' });
    assert.deepEqual(diagnostic.sourceList, { numId: '1', level: 0, kind: 'bullet', format: 'bullet' });
    assert.equal(diagnostic.recovery.action, 'use_list_format_change_format');
    assert.ok(result.results[0].warnings.some(warning => warning.startsWith('LIST_FORMAT_NOT_APPLIED: ')));
    const kept = openDocx(result.toBuffer()).inspect({}).paragraphs.find(p => p.text === BULLET);
    assert.equal(kept.list.numId, '1', 'the paragraph still uses the bullet definition');
    assert.equal(kept.list.format, 'bullet');

    const strict = await apply([operation], { atomic: true });
    assert.equal(strict.status, 'error');
    const error = strict.results[0].error;
    assert.equal(error.code, 'UNSUPPORTED_LIST_FORMAT_CHANGE');
    assert.equal(error.recovery.action, 'use_list_format_change_format');
    assert.deepEqual(error.sourceList, diagnostic.sourceList);
    assert.deepEqual(strict.toBuffer(), fixture);
    // WP-5: the recovery names the list-format change-format request; following it (with a target) changes the list.
    assert.deepEqual(error.recovery.suggestedOperation, { type: 'list-format', list: { mode: 'change-format', format: 'decimal', level: 0, scope: 'list' } });
    const followed = await apply([{ ...error.recovery.suggestedOperation, target: { exactText: BULLET } }], { atomic: true });
    assert.equal(followed.status, 'ok', JSON.stringify(followed.results?.[0]?.error || followed.error || null));
    const changed = openDocx(followed.toBuffer()).inspect({}).paragraphs.find(p => p.text === BULLET);
    assert.equal(changed.list.format, 'decimal');
    assert.notEqual(changed.list.numId, '1', 'a new definition, never the shared one edited in place');

    const literal = await apply([{ ...operation, listMarkers: 'literal' }], { atomic: true });
    assert.equal(literal.status, 'ok');
    assert.equal(literal.results[0].diagnostics, undefined);

    // Same kind and format as the source definition is not a format change.
    const same = await apply(
        [{ type: 'list-change', target: { exactText: NUMBERED }, modified: `1. ${NUMBERED}\n2. Added numbered item` }],
        { atomic: true }
    );
    assert.equal(same.status, 'ok', JSON.stringify(same.results[0].error || null));
    assert.equal(same.results[0].diagnostics, undefined);

    // A different numbered format (decimal source asked to become lettered) is also reported.
    const lettered = await apply([{ type: 'list-change', target: { exactText: NUMBERED }, modified: `A. ${NUMBERED}\nB. Added item` }]);
    assert.equal(lettered.results[0].diagnostics[0].code, 'LIST_FORMAT_NOT_APPLIED');
    assert.equal(lettered.results[0].diagnostics[0].requestedList.format, 'upperLetter');
    assert.equal(lettered.results[0].diagnostics[0].sourceList.format, 'decimal');
}

// Source level lookup honours the numbering part.
{
    const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
    const xml = `<w:numbering xmlns:w="${W}"><w:abstractNum w:abstractNumId="5"><w:lvl w:ilvl="0"><w:numFmt w:val="upperRoman"/></w:lvl><w:lvl w:ilvl="1"><w:numFmt w:val="bullet"/></w:lvl></w:abstractNum><w:num w:numId="9"><w:abstractNumId w:val="5"/></w:num></w:numbering>`;
    assert.deepEqual(readSourceListLevel(xml, '9', 0), { numFmt: 'upperRoman', kind: 'numbered' });
    assert.deepEqual(readSourceListLevel(xml, '9', 1), { numFmt: 'bullet', kind: 'bullet' });
    assert.equal(readSourceListLevel(xml, '10', 0), null);
    assert.equal(readSourceListLevel('', '9', 0), null);
}

// paragraph-format: unknown keys are reported (warning first, error in a later minor).
{
    const result = await apply([{
        type: 'paragraph-format',
        target: { exactText: PLAIN },
        properties: { alignment: 'center', spacingAfter: 120, indentLeft: 720 }
    }]);
    assert.equal(result.status, 'ok');
    const [diagnostic] = result.results[0].diagnostics;
    assert.equal(diagnostic.code, 'UNKNOWN_PARAGRAPH_FORMAT_PROPERTY');
    assert.equal(diagnostic.severity, 'warning');
    assert.deepEqual(diagnostic.properties, ['spacingAfter', 'indentLeft']);
    assert.equal(result.results[0].status, 'applied', 'supported keys are still applied');

    const onlyUnknown = await apply([{ type: 'paragraph-format', target: { exactText: PLAIN }, properties: { spacingAfter: 120 } }]);
    assert.equal(onlyUnknown.results[0].status, 'no_change');
    assert.equal(onlyUnknown.results[0].diagnostics[0].code, 'UNKNOWN_PARAGRAPH_FORMAT_PROPERTY');

    const clean = await apply([{ type: 'paragraph-format', target: { exactText: PLAIN }, properties: { alignment: 'center', keepNext: true } }]);
    assert.equal(clean.results[0].diagnostics, undefined);
}

// Recovery metadata for the new codes.
for (const [code, action] of [
    ['UNSUPPORTED_LIST_CONVERSION', 'use_list_format_convert_plain'],
    ['UNSUPPORTED_LIST_FORMAT_CHANGE', 'use_list_format_change_format']
]) {
    const normalized = normalizeErrorWithRecovery({ code, message: 'm' });
    assert.equal(normalized.stage, 'list-conversion');
    assert.equal(normalized.category, 'request_fixable');
    assert.equal(normalized.recovery.action, action);
}

// Compact CLI output keeps the new fields (warning diagnostics and the strict error details).
const temp = await mkdtemp(path.join(tmpdir(), 'docx-list-diagnostics-'));
try {
    const operationsFile = path.join(temp, 'operations.json');
    await writeFile(operationsFile, JSON.stringify([{ type: 'redline', target: { exactText: PLAIN }, modified: `A. ${PLAIN}` }]));

    const warned = await executeCli(['apply', fixtureFile, '--operations', operationsFile, '--output', path.join(temp, 'warned.docx')]);
    assert.equal(warned.status, 'ok');
    assert.equal(warned.results[0].diagnostics[0].code, 'LIST_MARKER_NOT_BOUND');
    assert.match(warned.results[0].warnings[0], /^LIST_MARKER_NOT_BOUND: /);

    const refused = await executeCli([
        'apply', fixtureFile, '--operations', operationsFile, '--profile', 'agent', '--compact',
        '--output', path.join(temp, 'refused.docx')
    ]);
    assert.equal(refused.status, 'error');
    const error = refused.results[0].error;
    assert.equal(error.code, 'UNSUPPORTED_LIST_CONVERSION');
    assert.equal(error.stage, 'list-conversion');
    assert.equal(error.category, 'request_fixable');
    assert.equal(error.listMarker, 'A.');
    assert.deepEqual(error.requestedList, { kind: 'numbered', format: 'upperLetter' });
    assert.equal(error.targetExcerpt, PLAIN);
    assert.equal(error.recovery.action, 'use_list_format_convert_plain');
    assert.equal(refused.completion, false);

    await writeFile(operationsFile, JSON.stringify([{ type: 'redline', target: { exactText: PLAIN }, modified: `A. ${PLAIN}`, listMarkers: 'literal' }]));
    const literal = await executeCli([
        'apply', fixtureFile, '--operations', operationsFile, '--profile', 'agent', '--compact',
        '--output', path.join(temp, 'literal.docx')
    ]);
    assert.equal(literal.status, 'ok');
    assert.equal(literal.completion, true);
    assert.equal(literal.results[0].diagnostics, undefined);
} finally {
    await rm(temp, { recursive: true, force: true });
}

console.log('PASS: list conversion diagnostics');
