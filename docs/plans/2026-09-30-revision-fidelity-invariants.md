# Revision Fidelity Invariants: Turning Gap Bugs into Fail-Closed Refusals

Date: 2026-09-30
Status: PROPOSED. Package baseline 0.8.2 plus the uncommitted list fixes, CLI `contractVersion` 8.
Source: holistic review of the last run of bug fixes (hyperlink boundary, manual line break,
plain-anchor insertion, list-range replacement, historical list inspection, facade numbering).
The review concluded that they were local gap-plugs, not structural fixes. This plan proposes the
structural fix: one post-apply invariant that every mutation path must satisfy, a Word-faithful
library oracle, and the consolidation of the duplicated code that let the gaps open.

Related plans: `2026-09-30-canonical-list-operations.md` (its WP-2 numbering allocator is the
convergence target for WP-5 here; it is referenced, not duplicated),
`completed/2026-09-05-structural-revisions-and-fidelity-oracles.md` (origin of the oracle helpers and
the paragraph-mark model), `completed/2026-09-01-performance-and-complexity-reduction.md` (budget rules).

## Why now

Every defect below was fixed correctly in isolation. Each one shipped because no check existed that
asked the only question that matters for a redline library: "if the reviewer presses Reject All, is
the document the source; if Accept All, is it the intent?"

| # | Defect | Where it lived | What the existing oracle saw | What caught it |
|---|---|---|---|---|
| 3 | A localized replacement pulled an adjacent plain `.` into a hyperlink | `engine/reconstruction-mode.js` `keepUnchangedTextOutsideWrappers` (post-diff patch) | Accepted text correct, so green | Manual review: wrapper membership wrong |
| 4 | Reject All reordered text across a manual `<w:br/>` | tokenization in `pipeline/diff-engine.js` (now `atomicChars` + `anchorSharedReferences`) | Text-only compare normalises breaks | Manual review |
| - | Plain-anchor insertion produced a paragraph with no tracked-inserted paragraph mark | `services/document-operation-mutations.js` document-level shortcut | Library `rejectTrackedChangesInOoxml` did not reveal it | Only Word (Reject All left an empty paragraph) |
| - | List-range replacement collapsed N source paragraphs into one deleted paragraph; then empty-item and all-empty ranges lost paragraphs (three fix rounds) | `pipeline/list-generation.js` | Text equality | Word, then iterative review |
| - | Historical `numPr`/`pStyle`/`outlineLvl` inside `w:pPrChange` reported as current | `services/document-inspection.js`, `core/list-targeting.js` (descendant lookups cross `pPrChange`) | n/a (inspection, not mutation) | Review |
| - | `openDocx` path remapped deleted source paragraphs to a new generic `numId`; Reject (ListType) and Accept (ListValue) broke | `services/numbering-helpers.js`; standalone runner passed | Standalone suite green | Consumer integration |
| - | Bare `document.xml` vs packaged `.docx` divergence (`"1. Header"` yields `RECEIPT_RECONCILIATION_FAILED` only on bare XML) | `services/receipt-collector.js` reconciliation vs bare-XML session | Not exercised | Reproduced in the canonical-list plan; still open |

Three structural causes, all visible in that table.

1. **No central paragraph-mark invariant.** Paragraph-mark revision accounting is implemented
   separately in `engine/reconstruction-writer.js`, `pipeline/list-generation.js`, the document-level
   shortcuts in `services/document-operation-mutations.js`, and the surgical mode. Each re-derives
   "which marks are inserted, which are deleted, which survive Reject".
2. **The oracle is weaker than the product's promise.** Most suites compare text after the library's
   own accept/reject. That oracle ignores structure (paragraph count, wrapper membership, `pPr`/`numPr`,
   numbering parts) and diverges from Word on at least some paragraph-mark cases. Word is the real
   oracle but is reachable only manually or through the Word COM lane.
3. **Parallel implementations of one job.** Facade vs standalone-runner numbering; bare vs packaged
   documents; many ad-hoc "current properties" reads that descend into `*PrChange`.

The existing receipt reconciliation (`services/receipt-collector.js`, `reconcileReceiptsAgainstOutput`,
called from `services/document-operation-applier.js` ~line 497 and
`services/batch-operation-orchestrator.js` ~line 524) checks only that committed ids exist in the
output (revision, comment, numbering and relationship ids). It confirms the receipt is truthful; it
does not confirm the revision structure is right. WP-1 adds that missing check at the same seam, which
already has the savepoint rollback machinery.

## Goals and non-goals

Goals:

- Every mutating operation self-verifies Word-semantics reject and accept views before it is committed.
- A violation is a refusal with recovery guidance under strict/atomic, not a silent success.
- The library's own `rejectTrackedChangesInOoxml` / `acceptTrackedChangesInOoxml` follow Word for
  paragraph marks, so tests and the verifier share one trustworthy implementation.
- One accessor for current properties, one helper for paragraph-mark revisions, one numbering path.

Non-goals:

- No new operations, no change to operation schemas, no change to accepted-text semantics.
- No layout, rendering or pagination checks; the visual lane keeps that role.
- No attempt to verify other authors' pre-existing revisions beyond "unchanged by this operation".
- No rewrite of the reconciliation engine modes; the verifier wraps them, WP-6 only extracts a helper.
- Not replacing the Word COM lane. WP-2 and WP-3 add to it; they do not substitute for it.
- Canonical list operations (`list-apply`, `change-format`) stay in their own plan.

## WP-1: Post-apply revision fidelity invariant (first; blocks nothing, unblocks everything)

### Design

New module `services/revision-fidelity-verifier.js` (leaf service; imports only `core/` and
`adapters/`; no `index.js`, no `node/`). Entry point:

```js
verifyRevisionFidelity({
    before,            // { documentXml, numberingXml } the savepoint state
    after,             // same shape, post-mutation
    touched,           // paragraph scope, see below
    intent,            // { acceptedParagraphs?: string[] } from the operation when it has one
    author
}) -> { ok: true } | { ok: false, violations: Violation[] }
```

It builds two structural projections of `after` by running a Word-semantics simulator, then compares
them. A projection is an array of paragraph records:

```js
{ ordinal, text, pPr: canonicalCurrentPPr, numPr: {numId, ilvl}|null, pStyle,
  wrappers: [{kind: 'hyperlink'|'fldSimple'|'sdt'|..., id/rid, textSpan}],
  breaks: [offsets], container: 'body'|'tc'|'hdr'... }
```

Reject view (`after`, reject all authors that this operation touched, equivalently all revisions
created by `author` in this call): apply the Word rules.

- Paragraph mark inside `w:pPr/w:rPr/w:ins`: the mark is removed; the paragraph's rejected content
  merges into the following paragraph (surviving mark and `pPr` are the following paragraph's).
- Paragraph mark inside `w:pPr/w:rPr/w:del`: the mark is restored.
- `w:ins` content removed, `w:del` content restored, `*PrChange` restores its recorded old
  properties (including `numPr`, `pStyle`).
- Last-paragraph-in-cell and last-paragraph-in-body edge cases follow the documented Word behavior
  (WP-2 pins them with COM evidence).

Invariants checked (names are stable and appear in the error payload):

| Invariant id | Assertion |
|---|---|
| `REJECT_PARAGRAPH_COUNT` | rejected view of the touched scope has the same paragraph count and order as `before` |
| `REJECT_PARAGRAPH_TEXT` | per-paragraph text equal (breaks, tabs, fields compared in the canonical text view, not normalised) |
| `REJECT_PARAGRAPH_PROPS` | canonical current `pPr` equal, including `numPr`, `pStyle`, `outlineLvl` |
| `REJECT_WRAPPERS` | wrapper membership equal: hyperlink/field/sdt boundaries by text span (catches defect #3 class) |
| `REJECT_NUMBERING_REFS` | every `numId` referenced by a restored paragraph exists in the output numbering parts and its `abstractNum` level format equals the source definition's (catches facade numbering class) |
| `ACCEPT_INTENT` | accepted view of the touched scope equals the intended result, when the operation declares one (`modified` text, replacement result, inserted paragraphs) |
| `ACCEPT_NEIGHBORS` | paragraphs outside the touched scope are byte-identical in both views |

Not every operation has an `intent`; `ACCEPT_INTENT` is skipped when absent, the reject-side
invariants are unconditional because "reject equals source" is true for every tracked operation.
Untracked operations (`generateRedlines: false`) skip reject invariants and run only
`ACCEPT_NEIGHBORS`.

### Where it runs

At the existing reconciliation seam, after `reconcileReceiptsAgainstOutput` succeeds, in:

- `services/document-operation-applier.js` (~line 497, inside the `!session.deferSerialization`
  branch; the savepoint is already in scope: `session.restoreSavepoint(savepoint)`),
- `services/batch-operation-orchestrator.js` (~line 524) for the batch/deferred path, run once per
  batch over the union of touched paragraphs.

On deferred serialization (the Phase 1 session optimisation) the verifier works on the live DOM
snapshot rather than a serialized string; its simulator takes DOM nodes, not text, so it must not
force serialization (see Performance).

### Behavior by mode

| Mode | Violation result |
|---|---|
| `atomic: true` or `strictTargets: true` | Operation fails, savepoint restored, error code `REVISION_FIDELITY_VIOLATION`, receipt `finalDisposition: 'rolled_back'` |
| Non-atomic, default agent profile | Same refusal for the failing operation only (progressive mode already isolates operations); later operations proceed. A fidelity violation is a library defect or an unsupported shape, never something to apply and warn about |
| Library host opting out | `verifyRevisionFidelity: 'warn'` option returns the applied result plus `warnings[]` entries with the same payload. Default is `'enforce'`. `'off'` exists only for diagnosing and is rejected by the agent profile |

Rationale for not shipping non-atomic as warn-only: a warn-only default would reproduce the
silent-success defect class this plan removes. The rollout still goes through a `warn` period (see
Sequencing) so false positives are found before `enforce` becomes the default.

### Error contract

Add to `services/error-recovery.js` / the recovery envelope (follow `error.recovery.action` and
`retryPlan` conventions in `docs/AGENT_KNOWLEDGE_BASE.md`):

```json
{ "code": "REVISION_FIDELITY_VIOLATION",
  "message": "Operation would not restore the source on Reject All (REJECT_PARAGRAPH_COUNT: expected 12, got 11 at paragraph 4).",
  "violations": [{ "invariant": "REJECT_PARAGRAPH_COUNT", "paragraph": 4, "expected": "...", "actual": "..." }],
  "recovery": { "action": "narrow_or_change_operation",
                "hint": "Retry with a smaller target range, or use replacements for a small literal change. Do not retry the same arguments." } }
```

`recovery.action` must be a value the existing recovery tests accept; add a case to
`tests/error_recovery_contract_tests.mjs`. Add the code to `docs/AGENT_KNOWLEDGE_BASE.md` failure
examples and `index.d.ts`.

### Reuse (do not duplicate)

- Seed: the Word-semantics paragraph-mark simulator in `tests/list_reject_fidelity_tests.mjs`
  (`wordParagraphs`, `markType`, `rejectedFragment`, `acceptedFragment`). Move its logic into the new
  module and make that test import it, so the test and product share one implementation.
- `tests/helpers/canonical-ooxml.mjs` (`canonicalizeOoxml`): compare `pPr`/numbering definitions
  structurally; keep the helper in tests, but extract its pure canonicalisation into `core/` only if
  WP-1 needs it at runtime (a small `canonicalizeProperties` is enough; do not move the whole file).
- `tests/helpers/roundtrip.mjs` (`assertRoundTripStructure`, `normalizeVisibleText`) and
  `tests/fidelity_oracle_tests.mjs` / `tests/helpers/mutation-envelopes.mjs` (`verifyMutationFidelity`):
  the envelope idea (declare what may change, assert the rest is identical) is exactly
  `ACCEPT_NEIGHBORS`; align naming and extend the envelopes rather than inventing a parallel one.
- `core/paragraph-text.js` `extractParagraphRevisionSegments` and `isNodeVisibleInRevisionView` for
  the canonical per-view text; do not write a third text walker (see Phase 4 parity test).

### Scope and performance budget

- Touched scope = paragraphs the operation created, deleted or modified (known from the receipt's
  `revisionItems` and the savepoint diff) plus one neighbor on each side (paragraph-mark merges move
  content into the following paragraph, so the next neighbor is mandatory) plus the paragraphs whose
  `numId` the operation referenced.
- Complexity is O(touched), not O(document). Do not parse or serialize the whole part. Wrapper
  checks use the paragraph's own spans.
- Budget (consistent with `completed/2026-09-01-performance-and-complexity-reduction.md`: accuracy
  governs, record the measured cost): verifier adds at most 10% to the per-operation time in
  `scripts/benchmark-operation-session.mjs` on its default scenario and zero extra full-document
  serializations (assert with the Phase 1 instrumentation in
  `tests/performance_phase1_session_tests.mjs`). Record the measured cost in the PR.
- Batch path verifies once per batch over the union of scopes, not per operation, when serialization
  is deferred.

### Files

- New: `services/revision-fidelity-verifier.js`, `core/word-revision-semantics.js` (the pure
  reject/accept simulator over DOM paragraph lists; shared with WP-2), `tests/revision_fidelity_verifier_tests.mjs`.
- Edit: `services/document-operation-applier.js`, `services/batch-operation-orchestrator.js`,
  `services/receipt-collector.js` (add `revisionScope` to the receipt if the savepoint diff is not
  enough), `services/error-recovery.js`, `index.d.ts`, `docs/AGENT_KNOWLEDGE_BASE.md`, `CHANGELOG.md`.
- Isolation: run `npm run test:isolation` (`core/` must not import `services/`).

### Tests

1. Positive: wire the verifier into the existing suites by running it in `verifyMutationFidelity`
   so every existing envelope case also exercises it. Expect zero violations; every failure is either
   a real latent bug (file it) or a verifier bug.
2. Regression corpus, one case per row of the "Why now" table, built from `git show HEAD:` sources of
   the pre-fix files in the scratchpad: with the fixes reverted the verifier must report the expected
   invariant id (`REJECT_PARAGRAPH_COUNT` for plain insertion and list range, `REJECT_WRAPPERS` for
   hyperlink, `REJECT_PARAGRAPH_TEXT` for line break, `REJECT_NUMBERING_REFS` for facade numbering).
   These are the proof that WP-1 would have prevented the history.
3. Mutation test of the verifier: programmatically corrupt a correct output (drop an ins mark, move a
   run out of a hyperlink, swap a `numId`) and assert detection.
4. Fail-closed: atomic and non-atomic refusal, savepoint restoration (byte-exact), receipt
   disposition, recovery envelope.
5. Performance: instrumentation assertion of zero extra serializations plus a recorded benchmark.

### Risks

- False positives from legitimate Word behavior differences (merged paragraph properties on mark
  removal). Mitigation: WP-2 pins the semantics with COM evidence before `enforce` is the default;
  until then default to `warn`.
- Other authors' revisions in the touched scope make "reject equals source" ambiguous. Define the
  reject view as "reject only the revisions this operation created" (identified by the receipt's
  revision ids), and compare against `before` with the other authors' revisions left in place.
- Verifier cost on very large batches: cap scope and fall back to a sampled check only with an
  explicit `warnings[]` entry, never silently.

## WP-2: Word-faithful `rejectTrackedChangesInOoxml` / `acceptTrackedChangesInOoxml`

Goal: the library's accept/reject, and the module WP-1 builds on, give the same paragraph structure
Word does. Today the plain-anchor insertion bug passed the library's reject but failed in Word.

### Investigation first (half a day, produces the work list)

Reproduce on the pre-fix sources, not the working tree:

1. Extract `services/document-operation-mutations.js` and dependencies at `HEAD` into the scratchpad
   (`git show HEAD:services/document-operation-mutations.js` etc., or a detached `git worktree` for
   the scratchpad only; do not stash or switch the working tree).
2. Run the plain-anchor insertion from `tests/list_reject_fidelity_tests.mjs` case 1 and record why
   `rejectTrackedChangesInOoxml` returned the source. Candidate explanations to confirm or refute by
   reading `services/revision-comment-management.js` (`rejectTrackedChangesInOoxml`, ~line 391;
   `mergeParagraphIntoNextAndRemove`, ~line 179; `coalesceAdjacentCompatibleInsertions`):
   - the test's `libraryParagraphs` projection (concatenated `w:t` per paragraph) hides an empty
     paragraph or a merged-vs-unmerged difference;
   - removing a whole-paragraph `w:ins` leaves an empty `w:p` that a later normalisation step drops,
     where Word keeps it (the mark is not tracked, so Word keeps the paragraph);
   - an inserted paragraph adjacent to a list item is merged by the surviving-pPr rule in a way the
     assertion cannot see.
3. Write the actual divergence down as a table in this plan's follow-up commit message / release
   notes, with a minimal OOXML fixture per row.

### Likely changes (confirm against step 2)

- Reject: an inserted run-content paragraph with no tracked mark must remain as an empty paragraph,
  exactly as Word does. The library must not heal it.
- Reject/Accept mark merge: the surviving paragraph takes the properties Word gives it (the following
  paragraph's mark and `pPr`, with the empty-first-paragraph exception verified in COM), including the
  last-paragraph-in-cell and last-paragraph-in-body cases currently returning `false` /
  `removeNode` (`mergeParagraphIntoNextAndRemove` lines ~179-198).
- `*PrChange` rejection restores `numPr` and `pStyle` and removes now-empty `pPr` children
  (`rejectPropertyChangeNode`, ~line 305), consistent with WP-4's "never cross `pPrChange`".
- Refactor into `core/word-revision-semantics.js` (the same module WP-1 consumes) with
  `revision-comment-management.js` becoming a thin author-filtering wrapper. One implementation.

### Word COM backstop

Add a fixture matrix to the Word COM lane (`npm run test:word`, `scripts/word-com-suite.ps1`;
fixtures via the existing generators under `scripts/`): for each paragraph-mark shape, Word applies
Reject All / Accept All and the harness serialises Word's resulting paragraph list; the JS simulator
output must equal Word's. Rows: inserted mark on plain/list/table-cell/last-body paragraphs, deleted
mark on same, ins mark with non-empty and empty content, adjacent ins+del marks, mark `pPrChange`,
mixed authors. Store the Word-produced expectations as checked-in JSON so the non-Word lane
(`npm test`) can assert the simulator against recorded Word behavior without Word installed. Record
the Word build used in the JSON header.

### Tests

- `tests/word_revision_semantics_tests.mjs` (new, in `npm test`): simulator vs recorded Word
  expectations; library accept/reject vs the same expectations.
- Update `tests/list_reject_fidelity_tests.mjs` to assert library reject equals the Word simulator
  (it currently asserts both against text lists).

### Risks

- Changing library reject behavior is observable to consumers who accept/reject through the public
  API. Gate behind a changelog entry and, if a behavior change is visible, a `contractVersion`
  note; the Word behavior is the specification, so this is a bug fix, not a policy change.

## WP-3: Reject-roundtrip property/matrix test

Goal: systematically generate the shapes that produced the defect history, instead of discovering
them one consumer report at a time.

### Design

`tests/reject_roundtrip_matrix_tests.mjs` plus `tests/helpers/operation-generators.mjs`.

- **Fixtures** (Word-authored where possible, reuse
  `tests/fixtures/agentic-lists/nested-lists-source.docx` and the existing generator scripts; add
  fixtures through the `scripts/generate-word-*.ps1` family, never hand-edit XML for a Word-authored
  claim): nested lists (bullet/decimal/mixed, empty items), hyperlinks (including trailing
  punctuation adjacent to the wrapper), manual breaks and tabs, fields (`fldSimple` and complex),
  comments and replies, tables (merged cells, last-paragraph-in-cell), headers/footers, footnotes,
  content controls, multi-author existing revisions.
- **Generators** (pure, seeded): for a fixture, enumerate valid operation shapes: `redline` replace
  of a word/sentence/paragraph/range, insert before/after, delete, range replace across list items,
  empty-item and all-empty ranges, `replacements` literal edits, format ops, list-change,
  plain-to-list. Parameters drawn from a seeded PRNG (reuse the `FUZZ_SEED` convention in
  `tests/roundtrip_fuzz_tests.mjs`).
- **Properties**: for each (fixture, operation): apply with the production options, then
  (1) reject view structurally equals source using the WP-1 projection, (2) accept view equals intent,
  (3) neighbors untouched, (4) package fidelity via `tests/helpers/package-fidelity.mjs`
  (`assertPackageFidelity`, allowed changed entries listed), (5) idempotence: apply again to the
  accepted result with the same intent is a no-op.
- Run both entry paths for every case: `applyOperationsToDocumentXml` (standalone/bare) and `openDocx`
  (packaged). Divergence between the two is a failure by construction; this is the permanent guard for
  WP-5.

### Where it runs

- `npm test`: a small deterministic subset (fixed seed, about 150 cases, under 20 s) so that the
  suite stays within the existing budget; included via `scripts/run-tests.mjs`.
- Extended lane: `FUZZ_SEED=<n> MATRIX_ITERATIONS=<n> node tests/reject_roundtrip_matrix_tests.mjs`,
  added to the nightly `fuzz-extended` job described under "Continuous validation" in
  `docs/TESTING.md` with the date-derived seed; failure log prints the seed, fixture, operation JSON
  and the violated invariant so it replays locally.
- Word COM lane: the same generated cases exported by a flag (`--export-word`) and fed to
  `scripts/word-com-suite.ps1` for a periodic Word differential. Not part of `npm test`.

### Risks

- Generator reachability: operations that legitimately fail (target not found) must be filtered by
  the generator, not counted as passes. Count and report skipped cases.
- Flaky nightly noise from fixture-specific limitations; keep a documented allowlist file with
  reasons and an owner, empty by default.

## WP-4: Single "current properties" accessor

Problem: descendant lookups (`getElementsByTagNameNS` / `getWordElementsByLocalName`) cross
`w:pPrChange` and `w:rPrChange`, so historical `numPr`, `pStyle`, `outlineLvl`, `rFonts`, etc. look
current. The two known instances were fixed locally in `services/document-inspection.js` and
`core/list-targeting.js`.

### Design

New `core/current-properties.js`:

```js
getCurrentParagraphProperties(p)   // -> {pPr: Element|null, numPr:{numId,ilvl}|null, pStyle, outlineLvl, ...}
getCurrentRunProperties(r)
getCurrentChild(propsEl, localName) // direct child only, never inside *PrChange
getHistoricalParagraphProperties(p) // explicit opt-in for the pPrChange snapshot, for reject simulation
```

Rule: only direct children of `w:pPr` / `w:rPr`; `*PrChange` snapshots are reachable only through the
`getHistorical*` functions, which exist so the reject simulator (WP-1/WP-2) can use the same code.
Add a lint-style test (`tests/current_properties_usage_tests.mjs`) that fails when a source file
outside an allowlist calls descendant lookups for the property names `numPr`, `numId`, `ilvl`,
`pStyle`, `outlineLvl`, `rFonts`, `sz`, `b`, `i`, `u`.

### Migration

Audit with `rg` and migrate in this order (highest blast radius first). Start list, to be completed
by the audit (the agent that fixed inspection judged the rest safe; the audit confirms that):

```
rg -n "getElementsByTagNameNS\(.*'(numPr|numId|ilvl|pStyle|outlineLvl|rPr|pPr)'" core services engine pipeline orchestration
rg -n "getWordElementsByLocalName\(.*'(numPr|numId|ilvl|pStyle|outlineLvl|rPr|pPr)'" core services engine pipeline orchestration
rg -n "getElementsByTagNameNS\('\*'" core services engine pipeline
```

Known files with descendant lookups to audit (from `rg -l "getElementsByTagNameNS|descendant"`):
`core/paragraph-targeting.js`, `core/word-xml.js`, `core/xml-query.js`, `services/batch-operation-orchestrator.js`,
`services/capture-engine.js`, `services/comment-replies.js`, `services/comment-thread-parts.js`,
`services/document-inspection.js`, `services/document-operation-mutations.js`,
`services/headers-footers.js`, `services/numbering-helpers.js`, `services/operation-batch-compiler.js`,
`services/operation-preflight.js`, `services/revision-comment-management.js`,
`services/standalone-docx-plumbing.js`, `engine/formatting-removal.js`,
`engine/reconstruction-writer.js`, `pipeline/ingestion-export.js`. For each hit record: reads current
vs historical vs "any", and whether a `*PrChange` child can change the answer. Put the table in the
PR description.

Tests: extend `tests/document_inspection_edge_tests.mjs` with a paragraph whose current `pPr` differs
from its `pPrChange` snapshot in `numPr`, `pStyle`, `outlineLvl` and a run whose `rPrChange` holds
different `rFonts`, asserting every migrated reader returns the current value; keep the existing
list-targeting cases.

### Risks

- Behavior change where code accidentally relied on the descendant read (for example, paragraph-mark
  `rPr` lookups that legitimately look at `pPr/rPr`); the accessor must expose the mark `rPr`
  explicitly (`getParagraphMarkProperties`) so those callers do not regress.
- `core/` must not import `services/`; keep the module dependency-free and pass it through
  `tests/core_dependency_graph_check.mjs`.

## WP-5: Converge parallel paths

Two divergences, one rule: one implementation, two thin entry points.

### 5a. Facade vs standalone numbering

Current: `services/standalone-operation-runner.js` and the `openDocx` facade path
(`document/docx-document.js`, `ensureNumberingArtifactsInZip` / `mergeNumberingXmlBySchemaOrder`,
lines ~429 and ~626) allocate and merge numbering differently; the facade produced the broken remap
that the same-kind reuse heuristic (`reuseSourceListNumbering`, `services/numbering-helpers.js`) now
masks.

Plan: the numbering allocator in `2026-09-30-canonical-list-operations.md` ("Numbering definition
allocation and remapping", WP-2 there) is the single implementation; this plan only adds
requirements to it and removes the divergence:

- Requirement R1: tracked-deleted source paragraphs never change `numPr` (keeps Reject exact);
  expressed as a verifier invariant (`REJECT_NUMBERING_REFS`), not a convention.
- Requirement R2: the allocator is invoked by the runner and the facade through the same function,
  with the same arguments; neither path carries its own merge.
- Requirement R3: the same-kind reuse heuristic is replaced by allocator policy once the allocator
  lands; until then keep it and its regression (`tests/facade_list_numbering_tests.mjs`).
- Parity test (permanent): the WP-3 matrix runs both paths per case and asserts identical numbering
  parts (canonical compare) and identical `numPr` per paragraph in both views.

### 5b. Bare `document.xml` vs packaged `.docx`

Open defect: `"1. Header"` on a bare `<w:document>` with `generateRedlines: true` ->
`RECEIPT_RECONCILIATION_FAILED: Committed revision id ... (kind: del) was not found`; not reproduced
on packaged input (see the canonical-list plan, section (a) observation 2).

Steps:

1. Reproduce (scratchpad script from the canonical-list plan, `repro-a3.mjs` family) and find which
   id the receipt commits and why the output lacks it. Check first: the explicit-decimal path
   (`tryExplicitDecimalHeaderListConversion`, `services/document-operation-mutations.js` ~1345)
   recording a revision id for a deletion that the bare-XML serialization path then renumbers or
   drops (`RevisionIdAllocator`, `core/types.js` ~194, seeded differently without a package), and
   `reconcileReceiptsAgainstOutput` matching `partName === 'word/document.xml'` on an item that a bare
   run records differently.
2. Decide the contract: bare XML must behave identically to packaged for document-part operations, or
   fail with an explicit `UNSUPPORTED_INPUT` before mutation. Never a late reconciliation failure.
3. Fix at the source of the divergence, add the bare-vs-packaged parity dimension to the WP-3 matrix.

Files: `services/standalone-operation-runner.js`, `services/standalone-docx-plumbing.js`,
`services/receipt-collector.js`, `core/types.js`. Tests: new case in
`tests/phase3_list_structural_fallback_tests.mjs` for bare XML; parity dimension in WP-3.

## WP-6: One paragraph-mark revision helper

Problem: four code paths each decide mark revisions: `engine/reconstruction-writer.js`,
`pipeline/list-generation.js`, document-level shortcuts in `services/document-operation-mutations.js`,
and the surgical mode (`engine/surgical-mode.js`). The missing ins mark in the plain-anchor shortcut and
the three-round list collapse are both "one path forgot a rule the others have".

### Design

New `engine/paragraph-mark-revisions.js` (engine layer so both `pipeline/` output builders and
`services/` shortcuts may depend on it; it depends only on `core/`), exposing a small intent-based
API instead of ad-hoc attribute poking:

```js
markParagraphInserted(p, ctx)        // pPr/rPr/ins (new id, author, date)
markParagraphDeleted(p, ctx)         // pPr/rPr/del; content moves to w:del wrappers elsewhere
planParagraphMarks({sourceParagraphs, targetParagraphs, alignment}, ctx)
   // -> per paragraph: keep | insertMark | deleteMark | propertyChange
   // guarantees: untracked mark count (source side) == source paragraph count;
   //             every non-source paragraph carries an inserted mark;
   //             one deleted paragraph per deleted source paragraph (never collapse N into 1)
```

Callers migrate one at a time, each gated by the WP-1 verifier in `warn`:

1. Document-level shortcut in `services/document-operation-mutations.js` (the plain-anchor bug).
2. `pipeline/list-generation.js` (range collapse, empty-item and all-empty cases): its three rounds
   become fixtures for the helper's unit tests.
3. `engine/reconstruction-writer.js` and `engine/surgical-mode.js`.

After migration add an architecture test (`tests/paragraph_mark_single_site_tests.mjs`): the string
patterns that construct a `w:ins`/`w:del` under `w:pPr/w:rPr` appear in exactly one module (use `rg`
patterns in the test, same technique as `tests/no_word_api_index_check.mjs`).

Risk: behavior-neutral refactors in four write paths. Mitigation: do it after WP-1 and WP-3 exist,
land one caller per commit, and require the matrix + COM lane green on each. Record the benefit: the
helper's `planParagraphMarks` invariants are the same ones WP-1 asserts, so a future path that skips
the helper is caught at the verifier, not at the consumer.

## Sequencing

1. **WP-1 in `warn` mode** (first, small slice): verifier module, wired at both seams, regression
   corpus from the "Why now" table. Ship in a patch release as warnings. This immediately converts any
   remaining gap bug into a visible refusal candidate without risking existing consumers.
2. **WP-2 investigation** in parallel (it only reads code and reproduces): its output decides the
   Word-semantics rules WP-1 enforces. Land WP-2 changes before flipping WP-1 to `enforce`.
3. **WP-3 deterministic subset** next; it validates WP-1 against many shapes and finds the first
   batch of real violations; fix those before enforce. Add the extended lane to nightly.
4. **Flip WP-1 default to `enforce`** (minor release, changelog + `index.d.ts` + knowledge-base failure
   examples). Agent profile enforces from this point.
5. **WP-4** (accessor + migration + lint test), independent; schedule alongside WP-3 since both are
   test-heavy rather than behavior-heavy.
6. **WP-5** after the canonical-list plan's allocator WP lands (5a); 5b (bare vs packaged) can start
   early because the WP-3 parity dimension needs it.
7. **WP-6** last. It is a refactor best done with the full safety net (verifier enforcing, matrix,
   COM backstop) in place.

Rationale for ordering: WP-1 is the only package that changes the failure mode of every future gap
bug from silent-and-found-in-Word to a fail-closed refusal with recovery guidance; WP-2 makes the
oracle correct; WP-3 makes it exhaustive; WP-4 to WP-6 remove the duplication that creates the gaps.

## Definition of done

- `REVISION_FIDELITY_VIOLATION` exists, is documented, and is enforced by default in the agent
  profile; every row of the "Why now" table has a verifier regression that fails on the pre-fix source.
- Library reject/accept equal the recorded Word expectations for every paragraph-mark shape in the
  WP-2 matrix.
- The matrix runs in `npm test` (subset) and nightly (extended), on both bare and packaged paths.
- No source outside `core/current-properties.js` reads current paragraph/run properties through a
  descendant lookup; `*PrChange` is reachable only through the explicit historical accessors.
- Paragraph-mark revision markup is constructed in one module; numbering allocation in one function.
- `npm test`, `npm run check:types`, `npm run test:isolation`, and the Word COM differential
  (`npm run test:word`) pass; measured verifier cost is recorded and within budget.

## Open questions

- Should `verifyRevisionFidelity: 'warn'` be a public option or an internal environment flag during
  the rollout? Recommended: public option, documented as transitional.
- When other authors' revisions overlap the touched scope, is "reject only this operation's
  revisions" always well-defined for cross-author slicing (`2026-09-08-cross-author-revision-slicing`)?
  Verify with the slicing suites before enforcing.
- Whether Word's merged-paragraph property rule (surviving mark's `pPr`) has exceptions for
  empty paragraphs and table-cell boundaries; answered by the WP-2 COM matrix.
