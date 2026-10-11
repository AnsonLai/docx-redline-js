# Revision Fidelity Invariants: Turning Gap Bugs into Fail-Closed Refusals

Date: 2026-09-30
Status: DONE except the open items below (2026-10-11, started on 0.8.4 + `c98874d`). All WPs landed, including WP-5a
(`4718c36`). Open items: nightly matrix wiring (no CI config changed); `ACCEPT_INTENT` (now worked under
`2026-10-11-word-plugin-reported-issues.md` as I1); numbering-part existence check; accept-first range-end scope
(`resolvedTargetEnd` not captured); optional `planParagraphMarks` wiring into the write paths; plain redline
reconstruction drops `w14:paraId` (found during list WP-6; a composed redline + list-format pair does not keep paragraph
identity); re-measure verifier cost on a quiet machine. See "Progress and hand-off" and the Definition-of-done status below.

Definition-of-done status (details at the end of the plan):
- [x] `REVISION_FIDELITY_VIOLATION` documented and enforced by default (agent profile rejects `off`); regressions for each
  "Why now" row live in `tests/revision_fidelity_verifier_tests.mjs` as synthetic defective outputs (not rebuilt from
  `git show` pre-fix sources).
- [x] Library reject/accept equal recorded Word expectations for all 27 WP-2 matrix shapes.
- [~] Matrix in `npm test` on bare and packaged paths; extended lane documented in `docs/TESTING.md` but not yet wired
  into a nightly CI job (no CI config in this repo was changed).
- [x] No descendant property reads outside `core/current-properties.js` (lint test).
- [x] Paragraph-mark markup constructed in one module (done); numbering allocation in one function (WP-5a, `4718c36`, done).
- [x] `npm test`, `check:types`, `test:isolation`, Word differential (64/64 synthetic) and Word corpus (60/60) pass; verifier cost
  recorded (re-measure on a quiet machine is an open item).

## Progress and hand-off (keep this section current; newest first)

Resume here. Each line says what landed, where, and what is next. Verify with `git log --oneline`.

| WP | State | Commits / files | Next step |
|---|---|---|---|
| WP-1 | Reject-side invariants landed and **enforced by default** (`c07c37b`, owner decision 2026-10-10: refuse by default with clear, retry-oriented errors instead of a warn-only patch release) | `05dcf4c`, `c933dcd` (baseline cache), `bbb23cb` (accept-first baseline scoped to resolved target), `c07c37b` (enforce default, error payload, CLI `--revision-fidelity`, agent profile rejects `off`, capability `revision-fidelity-v1`): `core/word-revision-semantics.js` (Word Reject/Accept projection over DOM), `services/revision-fidelity-verifier.js`, wiring in `services/document-operation-applier.js` (after the `GENERATED_OOXML_INVALID` check, before `markMutationCommitted`), `REVISION_FIDELITY_VIOLATION` in `services/error-recovery.js`, `verifyRevisionFidelity` in `services/standalone-operation-runner.d.ts`, KB failure row, `tests/revision_fidelity_verifier_tests.mjs` | `ACCEPT_NEIGHBORS` landed for untracked single-target operations (`d045e7e`; tracked operations are already covered by the Reject All invariants). Next: `ACCEPT_INTENT` (partly covered by the localized-replacement `acceptedViewMatchesCompiledText` verification), numbering-part existence check, range-end scope for accept-first (`resolvedTargetEnd` is not captured today), re-benchmark on a quiet machine |
| WP-2 | **Done** except the single-module refactor: simulator and library both match Word on all 27 cases (`KNOWN_DIVERGENCES` empty, `cedc68a`: cell last paragraph kept on reject, all-deleted single-cell table removed on accept, `pPrChange` reject keeps mark `rPr`/`sectPr`; regressions verified to fail pre-fix) | `1594add`: `scripts/generate-word-revision-semantics-matrix.mjs` (fixtures) + `scripts/word-revision-semantics-matrix.ps1` (run with `pwsh`, not Windows PowerShell 5.1, which hangs in `SaveAs2`) -> `tests/fixtures/word-revision-semantics/expectations.json` (Word 16.0.20430); `tests/word_revision_semantics_tests.mjs` with a `KNOWN_DIVERGENCES` map that fails when a divergence appears or disappears. `584c83a`: simulator end-of-container rule from COM evidence | Optional follow-ups: make `revision-comment-management.js` a thin author-filtering wrapper over one shared implementation (the plan's "one implementation" goal; the library mutates DOM while the simulator projects, so share the rules, not the code path); extend the matrix (non-empty trailing content with a removed mark, text boxes, footnotes, move revisions) |
| WP-3 | **Done** (default subset + extended lane) | `f5b485d`: `tests/reject_roundtrip_matrix_tests.mjs`, `tests/helpers/operation-generators.mjs` (seeded mulberry32, `FUZZ_SEED`), `tests/helpers/reject-roundtrip-fixtures.mjs` (nested-lists docx + 7 synthetic bodies), `docs/TESTING.md` entry; bare and packaged paths per case under enforce, package-fidelity check. Defects found and fixed: fldSimple dropped on paragraph edit (`c52b2a6`), deleted br/tab/inline objects untracked (`090a537`), list conversion moved text out of hyperlink/content control (`50d48a7`). 150 cases x 2 paths ~2.5 s (266 applied, 34 expected fail-closed skips); `MATRIX_ITERATIONS=1000` passes on seeds 1 and 2 (~18 s). No bare/packaged divergence; allowlist empty. Word synthetic lane re-run after these fixes: 49/49. Since then the canonical-list plan added generator kinds `paragraph-format-numbering`, `list-format-convert-plain`, `-change-format`, `-join` and `list-format-redline-compose` (a two-operation batch; `operation` may be an array) and fixtures `typed-headers`, `adjacent-lists`; 2026-10-11: default 150 and `FUZZ_SEED=8 MATRIX_ITERATIONS=500` pass, Word synthetic 64/64, corpus 60/60 | Not done: idempotence property; comments/footnotes/headers/complex-field fixtures; `--export-word` Word differential; nightly `fuzz-extended` wiring. Open: `PATCH_ROUNDTRIP_MISMATCH` for range-replace / plain-to-list of another author's inserted paragraph under `slice-cross-author` (fail-closed; may deserve support) |
| WP-4 | **Done** (npm test 137/137, types, isolation green) | `1a3eeb6`, `3405bca`, `ff49729`, `cfc9419`, `1f89ae0`: `core/current-properties.js` (+ `getCurrentNumIdElements`, explicit `getAllReferencedNumIdElements`), migrated readers, `tests/current_properties_tests.mjs`, lint `tests/current_properties_usage_tests.mjs` (allowlist: accessor itself, `standalone-docx-plumbing.js` numbering-usage check) | Optional: converge the ~10 private `directChild` duplicates; plan asked for cases in `tests/document_inspection_edge_tests.mjs` (covered instead by `current_properties_tests.mjs`) |
| WP-5a | **Done** (`4718c36`; list plan WP-2) | `services/numbering-allocation.js` is the only numbering merge/allocation policy. R1: `overwriteParagraphNumIds`, the remap and the continue-source policy skip tracked-deleted paragraphs (verifier `REJECT_NUMBERING_REFS` stays the enforcement). R2: `ensureNumberingArtifactsInZip` always calls `mergeNumberingParts` (callback option ignored); `openDocx` passes it only `numberingXmlParts`; a conflicting incoming id throws `NUMBERING_ID_CONFLICT`. R3: `reuseSourceListNumbering` removed, replaced by `allocateNumberingForReplacement` policy (`tests/facade_list_numbering_tests.mjs`, `tests/generated_bullet_numbering_collision_tests.mjs` green). Parity: `tests/reject_roundtrip_matrix_tests.mjs` compares canonical referenced numbering definitions and per-paragraph `numPr` in both views across bare and packaged; `tests/numbering_allocation_tests.mjs` | Fresh bare contexts that are not told the document numbering can still allocate colliding ids; the merge now reports it (see KNOWN_BUGS section 7) instead of binding silently |
| WP-5b | **Done** | `73c67e8`: bare runner always seeds `numberingIdState` (`services/batch-operation-orchestrator.js`); `tryExplicitDecimalHeaderListConversion` reserves numbering before redlining so an abort records no receipt revisions; regression in `tests/phase3_list_structural_fallback_tests.mjs` | Add bare-vs-packaged parity dimension in WP-3 |
| WP-6 | **Done** (npm test 142/142, matrix + `FUZZ_SEED=3 MATRIX_ITERATIONS=500`, types, isolation, Word synthetic 49/49 and corpus 60/60 re-run by the parent after the last commit) | `c356c5c` helper `engine/paragraph-mark-revisions.js` (DOM `markParagraphInserted/Deleted`, `setParagraphPropertiesMarkRevision`, `restoreParagraphMarkInserted`, `clearParagraphMarkRevision`; string `paragraphMarkRevisionXml`/`paragraphMarkPPrXml`; `planParagraphMarks` with the plan's invariants) + `tests/paragraph_mark_revisions_tests.mjs`; callers migrated one per commit: `7cfa6d9` document-level shortcuts, `5bddb82` list-generation, `606f00b` reconstruction-writer, `9a4ef99` table-mode, `0d2b553` table-reconciliation, `83fb19b` oxml-engine; `172fd56` removed `run-builders.js` `markParagraphMark*` and added the single-site architecture test `tests/paragraph_mark_single_site_tests.mjs` (allowlist empty). Surgical mode builds no paragraph marks. Behavior deltas (all suites green): new mark `rPr` placed before `sectPr`/`pPrChange`; list-generation replaces an existing mark revision instead of duplicating; receipt kind for the empty-modified shortcut and generated table cell marks is `ins`/`del` (was `structural`) | Optional: call `planParagraphMarks` from the write paths (reconstruction-writer, list-generation) so they plan before they write |

### Word verification (2026-10-10, Word 16.0.20430)

- **API caveat (important for every Word lane):** `Document.RejectAllRevisions()` / `AcceptAllRevisions()` can leave a
  revision behind. On a document whose last body paragraph has a tracked-inserted mark, `RejectAllRevisions()` keeps
  that mark and an extra empty paragraph (`revisionsLeft=1`), while the ribbon command
  (`CommandBars.ExecuteMso("ReviewRejectAllChangesInDocument")`), `WordBasic.RejectAllChangesInDoc` and
  `Document.Revisions.RejectAll()` all remove it. All Word lanes now use the `Revisions` collection and fail on leftover
  revisions (`d9c2fe1`); the COM matrix already did.
- **Structural Word check:** `scripts/word-com-differential.ps1` (used by `npm run test:word` and the corpus lane) now
  also fails when Word's Reject All paragraph count differs from `<name>.source.docx` (text-only comparison had missed
  the defect below). Synthetic lane: 49/49 with the structural check.
- **Corpus lane:** 60/60 on text, but the enforce verifier refused 3 real-document scenarios
  (`administrative-table-reconciliation-status-update`, `-row-insertion`, `-multi-cell-update`): markdown-table
  reconciliation drops, untracked, the trailing empty paragraph of an untouched sibling cell. Word confirmed (source 86
  paragraphs, Reject All 85). **Fixed in `b4f3bd4`** (`services/table-reconciliation.js`: trailing empty paragraphs no
  longer mark a cell modified; multi-paragraph cells reconcile per paragraph). Corpus prep prepares all 60 under enforce,
  Word corpus lane 60/60 with the structural check (86 vs 86). Known trade-off: when Markdown drops a non-empty
  paragraph of a multi-paragraph cell, its text is tracked-deleted but the paragraph shell stays, so Accept All leaves an
  empty paragraph there (Reject All exact). Stale fixtures numbered 42-47 in `tmp/superdoc-word-fixtures` are leftovers.
- **R5 end-of-document table append (2026-10-11):** the plugin's reported "extra paragraph + 1 revision" after Reject All reproduces
  only with `Document.RejectAllRevisions()`; `Revisions.RejectAll()` and the ribbon give the source's 7 paragraphs. Word's own COM
  authoring (TrackRevisions on, `Tables.Add` at the end) marks the *existing final paragraph's* mark as inserted, writes inserted
  rows and leaves an untracked empty trailing paragraph; it survives both APIs. The library keeps the final paragraph and adds a
  tracked-inserted trailing one (fine under `Revisions.RejectAll`). Word adds an untracked empty paragraph to any body that ends
  with a table (matrix `table-append-end-no-trailing-paragraph`: Reject 3 paragraphs, Accept 7, both views +1), modelled in
  `core/word-revision-semantics.js`. Defect fixed: a pretty-printed body (whitespace before `w:sectPr`) got no trailing paragraph
  (differential 5 vs 4). Matrix is 31 cases; synthetic differential 65/65.
- Scratch helpers (not committed): `%TEMP%\claude\wordcount.ps1` (source vs Reject All paragraph counts) and
  `wordui.ps1` (compares the four Reject All entry points).

### WP-1 implementation decisions (deviations from the design below)

- **Reject-all on both sides instead of filtering by receipt revision ids.** The verifier compares Word Reject All of
  `after` with Word Reject All of `before` (all authors). A tracked operation only adds revisions, so the views must be
  equal; other reviewers' revisions cancel out without id bookkeeping. This answers the cross-author open question for
  every policy except accept-first, which uses `baselineView: 'accepted'` (the policy accepts the source first).
- **Whole-document projection, DOM only.** No serialization; it reuses `readCanonicalRunText` for text (no third text
  walker). Not yet scoped to touched paragraphs; measure before optimising (cost is one projection of `before` and one of
  `after` per operation; the applier already validates the whole document twice per operation).
- **Runs per operation inside the applier**, so the batch/deferred path is covered with per-operation isolation; no
  separate orchestrator hook was needed.
- **Refusal disposition is `refused`** (detected before commit, like `GENERATED_OOXML_INVALID`), not `rolled_back`.
- **Untracked operations (`generateRedlines: false`) are skipped**; `ACCEPT_INTENT` and `ACCEPT_NEIGHBORS` are not
  implemented yet. `REJECT_NUMBERING_REFS` currently means "restored `numId`/`ilvl` differ from the source"; the
  numbering-part existence/format check is not implemented yet.
- **Cost (2026-10-10, `c933dcd`):** zero extra parses/serializations (benchmark counts unchanged). Projection +
  compare is ~8 ms per operation on a 1000-paragraph document without caching; the session caches the previous
  verified `after` projection (`session._rejectedViewCache`, keyed by DOM identity, cleared on any unverified commit)
  so a batch projects once per operation. `scripts/benchmark-operation-session.mjs` medians: off 265/277 ms vs warn
  260 ms (within noise; a concurrent Word COM run made one sample 360 ms). Re-measure on a quiet machine: OPEN (enforce shipped without the re-measure).
- **Mode resolution:** explicit `verifyRevisionFidelity` option (CLI `--revision-fidelity`), else env
  `DOCX_REDLINE_REVISION_FIDELITY` (diagnostic/test lanes), else `'enforce'`. The CLI agent profile rejects `off`;
  `effectiveOptions.verifyRevisionFidelity` reports the resolved mode.
- **Refusal payload (`c07c37b`), designed for an agent's retry loop:** `error.message` is one log-ready line
  ("Refused, document unchanged: ... (INVARIANT at source paragraph N (\"excerpt\"): expected ..., got ...). <explanation>
  Retry: <instruction>"); `violations[]` (max 5) each carry `invariant`, `expected`/`actual` excerpts, `explanation`,
  `source`/`output` `{ sourceParagraphIndex, container, text }` where `sourceParagraphIndex` is the 1-based extraction
  index; `invariants`, `verificationMode`, `recovery.action: narrow_or_change_operation`, `recovery.hint`,
  `recovery.retryGuidance[]` per invariant and `recovery.reportIfNarrowRetryFails`. The same line goes to `onWarn`,
  `warnings` and the refused receipt's warnings; compact CLI JSON keeps `violations`/`invariants`. Per-invariant text
  lives in `REVISION_FIDELITY_INVARIANTS` in the verifier; extend it when adding an invariant.

### Enforce probe (2026-10-10): `DOCX_REDLINE_REVISION_FIDELITY=enforce npm test`

**Final state (`bbb23cb`): `DOCX_REDLINE_REVISION_FIDELITY=enforce npm test` passes (138 suites).** The probe found
and fixed three real generator defects: explicit-start/`1. Header` list conversion bound `numPr` without a
`w:pPrChange` (`f72f60a`, `orchestration/list-structural-fallback.js` `enforceListBindingOnParagraphNodes`
`trackPropertyChange`); text-to-table produced tracked cell text in untracked rows (`c20f33c`,
`engine/table-mode.js` now passes `trackAsBlock: true`; `pipeline/pipeline.js:340` has the same call without it and is
untested, so check it); and the accept-first false positives (verifier scoping, `bbb23cb`). (Historical: enforce is now the default, so this probe is no longer a separate release gate.)

**Update after COM evidence (`584c83a`):** Word removes a trailing body paragraph whose inserted mark is rejected
(or deleted mark accepted), so class 2 below was a simulator error, not a library defect. Only 4 suites still fail
under enforce:
- `phase3_list_structural_fallback` (bare `1. Header`): **real defect**. The explicit-decimal conversion adds
  `numPr` without a `w:pPrChange`, so Reject All leaves a numbered list item.
- `existing_revisions_modes_matrix` (~line 338), `standalone_operation_runner` (text-to-table range), and
  `interagency_agreement_multi_author`: not yet triaged.
A Sonnet agent was assigned all four on 2026-10-10 (generator fixes only, verifier files untouched); check
`git log` for its commits. When enforce passes everywhere: make enforce the default for the agent profile
(Sequencing step 4).

Original probe (before the COM evidence): default warn was all green. Under enforce, 14 suites failed. Two classes found:

1. **Fixed:** `accept-all-first` policies compared against the wrong baseline (now `baselineView: 'accepted'`).
2. **Open, dominant (~all remaining):** an operation that replaces/extends the **last body paragraph** (directly before
   `w:body/w:sectPr`) emits the final paragraph with a tracked-inserted mark, e.g. "Second paragraph." -> "1. Item
   one\n2. Item two" yields `[del-mark P][ins-mark Item one][ins-mark Item two]<sectPr/>`. The simulator keeps the last
   paragraph of a container (its mark cannot be removed), so Reject All shows one extra empty paragraph. Whether Word
   does the same is the WP-2 COM question "last-paragraph-in-body"; if Word keeps it, this is a real library defect in
   every list/paragraph generator (the inserted mark belongs on the paragraph *before* the final one, as Word itself
   writes it); if Word removes it, relax the simulator. Reproduce: `tests/receipt_collector_tests.mjs` case 5,
   `tests/list_reject_fidelity_tests.mjs` case 3, `tests/table_append_fidelity_tests.mjs`.

Failing suites under enforce (before class 1 fix): docxjs_dashboard_rendering, existing_revisions_modes_matrix,
generated_bullet_numbering_collision, interagency_agreement_multi_author, list_reject_fidelity, mutation_receipt,
phase3_list_structural_fallback, receipt_collector, standalone_operation_runner, structured_content_planner,
table_append_fidelity (+ existing_revision_mutation_policy, merge_same_author, revision_scope, now passing).

### WP-2 findings (investigation, 2026-10-10)

- The plain-anchor premise was wrong: pre-fix (`c44f227^`) library reject already matched Word (it showed the stray
  empty paragraph). The defect was generation-only (untracked mark), so library reject is not "healing" anything there.
- Divergences in `services/revision-comment-management.js` against the plan's Word rules (to be confirmed by COM):
  1. last body paragraph with ins mark (reject) / del mark (accept): `mergeParagraphIntoNextAndRemove` (~line 186)
     removes the whole paragraph, losing untracked content;
  2. sole paragraph in a cell: `return false` (~182-184) leaves the tracked mark in place;
  3. last of two paragraphs in a cell: untracked text lost (~186);
  4. ins mark directly before a table: `getNextWordParagraph` (~170-177) scans past the table and merges across it;
  5. `rejectPropertyChangeNode` (~321-336) replaces the whole `pPr`, dropping the mark `rPr` (and `sectPr`);
  6. del mark on a paragraph carrying `pPr/sectPr`: the section break is dropped on accept merge (~190).
  Already Word-correct: untracked-mark paragraph with inserted content stays empty; mid-body ins mark merges into the
  following paragraph keeping its `pPr`; `pPrChange` reject restores `pStyle`/removes `numPr`.

### WP-4 audit summary (2026-10-10)

Class (a) readers that crossed `*PrChange`: `pipeline/ingestion-export.js` (rStyle/b/i, pStyle/outlineLvl/numPr),
`pipeline/ingestion-paragraph.js` `detectNumberingContext`, `services/numbering-helpers.js`
(`overwriteParagraphNumIds`, `extractFirstParagraphNumId`, `paragraphNumIdNodes`), `orchestration/list-structural-fallback.js`
numId lookup, `engine/oxml-engine.js` `sourceMatchesListShape`, `services/capture-engine.js` `hasList`. Intentionally
historical: numbering remap (`numbering-helpers.js` ~451), `collectNumberingIdsFromNodes` (merge needs all referenced
ids), `standalone-docx-plumbing.js` "uses numbering". ~10 private direct-child helpers duplicate the new accessor.

---
Source: holistic review of the last run of bug fixes (hyperlink boundary, manual line break,
plain-anchor insertion, list-range replacement, historical list inspection, facade numbering).
The review concluded that they were local gap-plugs, not structural fixes. This plan proposes the
structural fix: one post-apply invariant that every mutation path must satisfy, a Word-faithful
library oracle, and the consolidation of the duplicated code that let the gaps open.

Related plans (keep in step; update them in the same change when an overlapping item moves):
`2026-09-30-canonical-list-operations.md` (its WP-2 numbering allocator is the convergence target for WP-5a here, which
adds R1-R3; its "Cross-plan dependencies" section makes this plan's verifier the acceptance gate for `list-format` /
`change-format` / `paragraph-format.numbering` and requires `core/current-properties.js` for property reads; its section
(a) observation 2 was fixed by WP-5b, `73c67e8`; its id-collision residual is the same open item as WP-5a R2),
`2026-10-11-word-plugin-reported-issues.md` (I1 adds `UNTRACKED_CREATED_REVISION` and accept-intent enforcement; R5 may extend the Word-semantics model for end-of-document table appends),
`2026-10-01-consumer-reported-library-issues.md` (LOCAL ONLY: `docs/plans/*` is gitignored and this file was never
force-added, so other agents cannot rely on it; its WP-3 id-collision residual is WP-5a R2), `KNOWN_BUGS.md` (section 7 mirrors verifier limits and the
fixed gap bugs; update it when a limit is lifted),
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
| - | Bare `document.xml` vs packaged `.docx` divergence (`"1. Header"` yields `RECEIPT_RECONCILIATION_FAILED` only on bare XML) | `services/receipt-collector.js` reconciliation vs bare-XML session | Not exercised | Reproduced in the canonical-list plan; fixed by WP-5b (`73c67e8`) |

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

**Historical design; as built see "WP-1 implementation decisions" above.** Shipped as the enforce default (`c07c37b`); the warn-first rollout was skipped by owner decision. The `touched` scope, `intent`, O(touched) cost and once-per-batch verification below are superseded by the as-built decisions block (whole-document projection, per-operation inside the applier, reject-all on both sides).

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

*Historical: the `'warn'` period was skipped; default is `'enforce'` (`c07c37b`), refusal disposition is `refused`, not `rolled_back`.*

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

*Superseded by the as-built decisions block (whole-document projection, cached per operation, ~8 ms on 1000 paragraphs). Open: re-measure on a quiet machine.*

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

*Historical: mitigations (default `warn` until WP-2) were not needed; WP-2 landed and enforce shipped as the default. The cross-author risk was resolved by comparing Reject All on both sides.*

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

Tests (done in `tests/current_properties_tests.mjs`, not `document_inspection_edge_tests.mjs`): extend with a paragraph whose current `pPr` differs
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

**Historical design text (R1-R3 are done in `4718c36`, see the WP-5a row above); kept for the rationale.**

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

*Historical plan. As built: WP-1 shipped directly as enforce (`c07c37b`, owner decision), skipping the warn release; steps 2-7 all landed.*

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

Answered 2026-10-11 (original text kept):
- `warn` option: public (`verifyRevisionFidelity`, CLI `--revision-fidelity`) plus env override `DOCX_REDLINE_REVISION_FIDELITY`.
- Cross-author scope: the verifier compares Reject All on both sides (all authors), so no per-operation revision-id filtering; accept-first uses an accepted baseline.
- Merged-paragraph property rule and exceptions: answered by the WP-2 COM matrix (`584c83a`, `cedc68a`).

- Should `verifyRevisionFidelity: 'warn'` be a public option or an internal environment flag during
  the rollout? Recommended: public option, documented as transitional.
- When other authors' revisions overlap the touched scope, is "reject only this operation's
  revisions" always well-defined for cross-author slicing (`2026-09-08-cross-author-revision-slicing`)?
  Verify with the slicing suites before enforcing.
- Whether Word's merged-paragraph property rule (surviving mark's `pPr`) has exceptions for
  empty paragraphs and table-cell boundaries; answered by the WP-2 COM matrix.
