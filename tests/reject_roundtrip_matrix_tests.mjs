import './setup-xml-provider.mjs';

import { readFileSync, existsSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import path from 'node:path';

import { parseOoxmlSafe } from '../adapters/xml-adapter.js';
import { openDocx } from '../index.js';
import { unzipDocx } from '../document/zip-archive.js';
import { applyOperationsToDocumentXml } from '../services/standalone-operation-runner.js';
import { projectWordRevisionView } from '../core/word-revision-semantics.js';
import { generateOperationsForFixture, makeRng, resolveBaseSeed } from './helpers/operation-generators.mjs';
import { buildMatrixFixtures } from './helpers/reject-roundtrip-fixtures.mjs';
import { assertPackageFidelity } from './helpers/package-fidelity.mjs';
import {
    canonicalNumberingDefinition,
    mergeNumberingParts,
    validateNumberingIntegrity
} from '../services/numbering-allocation.js';

/**
 * WP-3 reject-roundtrip matrix: every generated (fixture, operation) is applied through both the bare
 * runner (`applyOperationsToDocumentXml`) and the packaged facade (`openDocx`) with default options
 * (revision-fidelity enforcement on). A `REVISION_FIDELITY_VIOLATION` refusal means the library generated a
 * shape Word's Reject All would not undo, which fails the matrix.
 *
 * Deterministic by default. Extended lane / replay:
 *
 *   FUZZ_SEED=<n> MATRIX_ITERATIONS=<n> node tests/reject_roundtrip_matrix_tests.mjs
 */

const BASE_SEED = resolveBaseSeed(20261010);
const ITERATIONS = Math.max(1, Number(process.env.MATRIX_ITERATIONS) || 150);
const MAX_REPORTED_FAILURES = Number(process.env.MATRIX_MAX_REPORTED) || 8;
// Refusal codes that signal a library defect rather than an unsupported request; they fail the matrix.
const INTERNAL_DEFECT_CODES = new Set(['RECEIPT_RECONCILIATION_FAILED']);
const ALLOWED_PACKAGE_CHANGES = ['word/document.xml', 'word/numbering.xml'];
const AUTHOR = 'Matrix Author';
const here = path.dirname(fileURLToPath(import.meta.url));
const ALLOWLIST_PATH = path.join(here, 'fixtures', 'reject-roundtrip-allowlist.json');

const decoder = new TextDecoder();
const partText = (entries, name) => (entries.has(name) ? decoder.decode(entries.get(name)) : undefined);
const parse = xml => {
    const parsed = parseOoxmlSafe(xml);
    if (!parsed.doc) throw new Error(`XML parse failed: ${parsed.error?.message}`);
    return parsed.doc;
};
const view = (xml, which) => projectWordRevisionView(parse(xml), which);
const texts = records => records.map(record => record.text);
const numPresence = records => records.map(record => Boolean(record.numPr));
// Per paragraph: level plus the canonical (id-free) numbering definition its numPr resolves to.
const numberingSignature = (records, numberingXml) => records.map(record => (record.numPr
    ? [record.numPr.ilvl, record.numPr.numId == null || !numberingXml ? null : canonicalNumberingDefinition(numberingXml, record.numPr.numId)]
    : null));
const issueKeys = issues => new Set(issues.map(issue => `${issue.code}:${issue.numId ?? ''}:${issue.ilvl ?? ''}`));
const innerError = result => result?.results?.find(item => item?.error)?.error || result?.error || null;
const errorCode = result => {
    const outer = result?.error?.code;
    const inner = result?.results?.find(item => item?.error)?.error?.code;
    return !outer || outer === 'BATCH_OPERATION_FAILED' ? (inner || outer || null) : outer;
};
const violationsOf = result => innerError(result)?.violations || result?.error?.violations || innerError(result)?.details?.violations || [];

const allowlist = existsSync(ALLOWLIST_PATH)
    ? JSON.parse(readFileSync(ALLOWLIST_PATH, 'utf8')).entries || []
    : [];
const allowlisted = (fixture, kind, path) => allowlist.find(entry =>
    entry.fixture === fixture && entry.kind === kind && (!entry.path || entry.path === path));

async function runBare(entries, operation, options = {}) {
    const documentXml = partText(entries, 'word/document.xml');
    const result = await applyOperationsToDocumentXml(documentXml, Array.isArray(operation) ? operation : [operation], AUTHOR, {
        numberingXml: partText(entries, 'word/numbering.xml'),
        stylesXml: partText(entries, 'word/styles.xml')
    }, options);
    // Bare callers merge the returned numbering parts themselves, through the one allocator merge.
    const numberingXml = mergeNumberingParts(partText(entries, 'word/numbering.xml'), result.numberingXmlParts);
    return { result, documentXml: result.documentXml, numberingXml };
}

async function runPackaged(bytes, operation, options = {}) {
    const result = await openDocx(bytes).applyOperations(Array.isArray(operation) ? operation : [operation], { author: AUTHOR, ...options });
    const output = result.status === 'error' ? null : result.toUint8Array();
    const outputEntries = output ? unzipDocx(output) : null;
    return {
        result,
        output,
        documentXml: outputEntries ? partText(outputEntries, 'word/document.xml') : null,
        numberingXml: outputEntries ? partText(outputEntries, 'word/numbering.xml') : null
    };
}

const UNTRACKED_KINDS = new Set(['replace-word', 'replace-sentence', 'replace-paragraph', 'replacements', 'insert-after', 'delete']);
// Revisions as `element|id|author|date`; a sliced fragment of a source revision keeps element/author/date.
const REVISION_TAG = /<w:(ins|del|moveFrom|moveTo|rPrChange|pPrChange|sectPrChange|tblPrChange|trPrChange|tcPrChange)\b([^>]*)>/g;
const attr = (attrs, name) => attrs.match(new RegExp(`\\bw:${name}="([^"]*)"`))?.[1] ?? '';
const revisionIds = xml => new Set([...xml.matchAll(REVISION_TAG)].map(m => `${m[1]}|${attr(m[2], 'id')}|${attr(m[2], 'author')}|${attr(m[2], 'date')}`));
const sliceSignatures = ids => new Set([...ids].map(key => key.split('|').filter((_, i) => i !== 1).join('|')));

const failures = [];
const skipped = new Map();
const unavailableKinds = new Map();
const counts = { cases: 0, applied: 0, refusedOther: 0, allowlisted: 0, numberingCompared: 0, untracked: 0 };
const startedAt = Date.now();

function fail(context, message, detail = {}) {
    failures.push({ context, message, detail });
}

function checkApplied(context, sourceEntries, sourceDocXml, kindInfo, path, run) {
    const { documentXml } = run;
    const sourceRejected = view(sourceDocXml, 'rejected');
    const resultRejected = view(documentXml, 'rejected');
    const resultAccepted = view(documentXml, 'accepted');
    // (2) Reject All restores the source view (cross-check of the verifier).
    const sourceJson = JSON.stringify(sourceRejected.map(r => [r.text, r.props, r.numPr, r.wrappers]));
    const outputJson = JSON.stringify(resultRejected.map(r => [r.text, r.props, r.numPr, r.wrappers]));
    if (sourceJson !== outputJson) {
        fail({ ...context, path }, 'rejected view differs from the source rejected view', {
            expected: texts(sourceRejected), actual: texts(resultRejected)
        });
    }
    // (3) Accepted view contains the intended text.
    const acceptedJoined = texts(resultAccepted).join('\n');
    for (const expected of kindInfo.expectAccepted) {
        if (!acceptedJoined.includes(expected)) {
            fail({ ...context, path }, `accepted view lacks intended text ${JSON.stringify(expected)}`, { accepted: texts(resultAccepted) });
        }
    }
    // (3b) Two separate continuing conversions in one batch (A. then B.) are ONE list.
    if (kindInfo.kind === 'list-sequence-redline') {
        const [first, second] = kindInfo.expectAccepted.map(text => resultAccepted.find(record => record.text === text));
        const firstId = first?.numPr?.numId;
        const secondId = second?.numPr?.numId;
        if (first && second && first !== second && (firstId == null || String(firstId) !== String(secondId))) {
            fail({ ...context, path }, 'separate continuing conversions did not share one numbering instance', { first: firstId, second: secondId });
        }
    }
    // (6) numbering integrity: no numId/abstractNum/ilvl reference is left dangling that the source did not have.
    const sourceNumbering = partText(sourceEntries, 'word/numbering.xml');
    const preExisting = issueKeys(validateNumberingIntegrity(sourceDocXml, sourceNumbering));
    for (const issue of validateNumberingIntegrity(documentXml, run.numberingXml)) {
        if (!preExisting.has(`${issue.code}:${issue.numId ?? ''}:${issue.ilvl ?? ''}`)) {
            fail({ ...context, path }, `numbering integrity: ${issue.code}: ${issue.message}`);
        }
    }
    return {
        rejected: resultRejected,
        accepted: resultAccepted,
        numbering: {
            accepted: numberingSignature(resultAccepted, run.numberingXml),
            rejected: numberingSignature(resultRejected, run.numberingXml)
        }
    };
}

const fixtures = buildMatrixFixtures();
const perFixture = Math.ceil(ITERATIONS / fixtures.length);

for (let fixtureIndex = 0; fixtureIndex < fixtures.length && counts.cases < ITERATIONS; fixtureIndex++) {
    const fixture = fixtures[fixtureIndex];
    const entries = unzipDocx(fixture.bytes);
    const sourceDocXml = partText(entries, 'word/document.xml');
    const inspection = openDocx(fixture.bytes).inspect({ existingRevisions: undefined });
    const rng = makeRng((BASE_SEED + Math.imul(fixtureIndex + 1, 0x9E3779B1)) >>> 0);
    const { cases, unavailable } = generateOperationsForFixture(rng, inspection.paragraphs, Math.min(perFixture, ITERATIONS - counts.cases));
    for (const kind of unavailable) unavailableKinds.set(`${fixture.name}:${kind}`, true);

    for (const generated of cases) {
        counts.cases++;
        const context = {
            seed: BASE_SEED, fixture: fixture.name, kind: generated.kind, operation: generated.operation
        };
        const outcomes = {};
        for (const [pathName, runner] of [
            ['bare', () => runBare(entries, generated.operation, fixture.options)],
            ['packaged', () => runPackaged(fixture.bytes, generated.operation, fixture.options)]
        ]) {
            let run;
            try {
                run = await runner();
            } catch (error) {
                fail({ ...context, path: pathName }, `threw: ${error?.message || error}`);
                continue;
            }
            const { result } = run;
            if (result.status === 'error') {
                const code = errorCode(result);
                if (code === 'REVISION_FIDELITY_VIOLATION') {
                    if (allowlisted(fixture.name, generated.kind, pathName)) { counts.allowlisted++; continue; }
                    fail({ ...context, path: pathName }, 'REVISION_FIDELITY_VIOLATION refusal (library generated a bad shape)', {
                        violations: violationsOf(result), message: result.error?.message
                    });
                } else if (INTERNAL_DEFECT_CODES.has(code)) {
                    fail({ ...context, path: pathName }, `${code} refusal (library defect)`, { message: innerError(result)?.message });
                } else {
                    counts.refusedOther++;
                    const key = `${generated.kind} -> ${code || 'unknown'}`;
                    if (process.env.MATRIX_VERBOSE) console.error(`skip ${fixture.name} ${pathName} ${key}: ${innerError(result)?.message}
  ${JSON.stringify(generated.operation)}`);
                    skipped.set(key, (skipped.get(key) || 0) + 1);
                }
                outcomes[pathName] = { refused: code };
                continue;
            }
            if (!result.hasChanges || !run.documentXml) { outcomes[pathName] = { noop: true }; continue; }
            counts.applied++;
            outcomes[pathName] = checkApplied(context, entries, sourceDocXml, generated, pathName, run);
            if (pathName === 'packaged' && run.output) {
                // (5) package fidelity: only document.xml / numbering.xml may change.
                try {
                    assertPackageFidelity(fixture.bytes, run.output, ALLOWED_PACKAGE_CHANGES);
                } catch (error) {
                    fail({ ...context, path: pathName }, `package fidelity: ${error.message}`);
                }
            }
        }
        // (4) bare vs packaged parity.
        const bare = outcomes.bare;
        const packaged = outcomes.packaged;
        if (bare && packaged) {
            const bareRefused = bare.refused !== undefined;
            const packagedRefused = packaged.refused !== undefined;
            if (bareRefused !== packagedRefused && !(bare.noop || packaged.noop)) {
                fail(context, 'bare and packaged disagree on refusal', { bare: bare.refused, packaged: packaged.refused });
            } else if (bare.accepted && packaged.accepted) {
                for (const which of ['accepted', 'rejected']) {
                    if (JSON.stringify(texts(bare[which])) !== JSON.stringify(texts(packaged[which]))) {
                        fail(context, `bare vs packaged ${which} paragraph texts differ`, { bare: texts(bare[which]), packaged: texts(packaged[which]) });
                    }
                    if (JSON.stringify(numPresence(bare[which])) !== JSON.stringify(numPresence(packaged[which]))) {
                        fail(context, `bare vs packaged ${which} numPr presence differs`, { bare: numPresence(bare[which]), packaged: numPresence(packaged[which]) });
                    }
                    // WP-5a parity: the numbering definitions each paragraph references are canonically equal.
                    counts.numberingCompared += bare.numbering[which].filter(entry => entry && entry[1]).length;
                    if (JSON.stringify(bare.numbering[which]) !== JSON.stringify(packaged.numbering[which])) {
                        fail(context, `bare vs packaged ${which} referenced numbering definitions differ`, { bare: bare.numbering[which], packaged: packaged.numbering[which] });
                    }
                }
            }
        }
        // (7) Untracked variant (generateRedlines: false) under slice-cross-author: applying must not create revisions.
        if (fixture.options.existingRevisions === 'slice-cross-author' && UNTRACKED_KINDS.has(generated.kind)
            && !Array.isArray(generated.operation)) {
            const untracked = { ...generated.operation, generateRedlines: false };
            const sourceIds = revisionIds(sourceDocXml);
            for (const [pathName, runPath] of [
                ['bare', () => runBare(entries, untracked, fixture.options)],
                ['packaged', () => runPackaged(fixture.bytes, untracked, fixture.options)]
            ]) {
                const run = await runPath();
                if (run.result.status === 'error') {
                    if (violationsOf(run.result).some(v => v.invariant === 'UNTRACKED_CREATED_REVISION')) {
                        fail({ ...context, path: `${pathName}+untracked` }, 'untracked operation created a revision', { violations: violationsOf(run.result) });
                    }
                    continue;
                }
                counts.untracked++;
                const sourceSlices = sliceSignatures(sourceIds);
                const created = [...revisionIds(run.documentXml)].filter(key => !sourceIds.has(key)
                    && !sourceSlices.has(key.split('|').filter((_, i) => i !== 1).join('|')));
                if (created.length) {
                    fail({ ...context, path: `${pathName}+untracked` }, 'untracked operation created revisions', { created });
                }
            }
        }
    }
}

const elapsed = ((Date.now() - startedAt) / 1000).toFixed(1);
if (failures.length > 0) {
    for (const failure of failures.slice(0, MAX_REPORTED_FAILURES)) {
        console.error(`\nMATRIX CASE FAILED: ${failure.message}`);
        console.error(`  reproduce: FUZZ_SEED=${failure.context.seed} MATRIX_ITERATIONS=${ITERATIONS} node tests/reject_roundtrip_matrix_tests.mjs`);
        console.error(`  fixture: ${failure.context.fixture}  kind: ${failure.context.kind}  path: ${failure.context.path || 'both'}`);
        console.error(`  operation: ${JSON.stringify(failure.context.operation)}`);
        console.error(`  detail: ${JSON.stringify(failure.detail, null, 1)}`);
    }
    console.error(`\n${failures.length} matrix failure(s) in ${counts.cases} case(s) (seed ${BASE_SEED}, ${elapsed}s).`);
    process.exit(1);
}

const skipSummary = [...skipped.entries()].map(([key, count]) => `${key}=${count}`).join('; ') || 'none';
console.log(`PASS: ${counts.cases} reject-roundtrip matrix cases x 2 paths (seed ${BASE_SEED}, ${elapsed}s): ` +
    `${counts.applied} applied, ${counts.refusedOther} non-fidelity refusals skipped, ${counts.allowlisted} allowlisted, ` +
    `${counts.numberingCompared} paragraph numbering definitions compared, ${counts.untracked} untracked variants checked`);
console.log(`  skipped refusals: ${skipSummary}`);
if (unavailableKinds.size) console.log(`  generator had no site for: ${[...unavailableKinds.keys()].join(', ')}`);
