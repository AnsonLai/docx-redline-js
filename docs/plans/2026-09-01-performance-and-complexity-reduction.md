# Performance and Complexity Reduction Plan

**Status:** Complete — Phases 1 through 6 finished (2026-09-04)
**Original date:** 2026-09-01  
**Revised:** 2026-09-04  

This plan reduces batch latency and internal complexity without weakening the
public operation contracts, target safety, package transactions, or exact-text
workflow added by the 2026-09-03 agent-friendly document plan.

The performance work must optimize the implementation behind those boundaries;
it must not create a second runner, inspector, text model, or package mutation
path.

Accuracy is the governing constraint. A performance target is observational,
not permission to weaken tracked-change fidelity, target safety, operation
isolation, accepted/rejected text, or exact rollback. When speed and redline
quality conflict, keep the accuracy safeguard and record the measured cost.

---

## 1. Current Baseline

As of 2026-09-04:

- Before Phase 2, `npm test` ran 55 isolated test files successfully. A local
  sequential run took roughly 29 seconds, with process startup a material part
  of the total. The Phase 1 and 2 regressions raise the suite to 57 files.
- The pre-Phase-2 coverage baseline was 89.05% statements/lines, 77.04%
  branches, and 92.63% functions. The checked Phase 1 result is 89.32%
  statements/lines, 77.09% branches, and 92.78% functions.
- Before Phase 2, `services/standalone-operation-runner.js` was 1,564 lines /
  67,994 bytes. It is now a 13-line compatibility facade; the historical size
  remains the baseline for the decomposition work.
- `applyOperationsToDocumentXml(...)` now parses the full document once and
  serializes it once after a successful changed batch. Scoped engine payloads
  still use their existing parse/serialize paths to preserve behavior.
- `RevisionIdAllocator.seed(...)` performs one universal element query and
  copies the result with `Array.from(...)`.
- `core/paragraph-text.js` is now the authoritative accepted/rejected/current
  paragraph-text API used by targeting, export ingestion, and structured
  inspection.
- Some engine walkers remain intentionally specialized because they also build
  offset maps, formatting spans, structural sentinels, and reference maps.
- The stable operation contract lives in
  `services/document-operation-contract.js`; strict read-only diagnostics live
  in `services/operation-preflight.js`.
- Complete DOCX mutation is owned by the separate Node facade in `node/`; the
  CLI builds on that facade and requires strict, atomic, validated writes.

These replace the older 38-test, four-extractor, and 1,451-line assumptions in
the original plan.

## 2. Non-Regression Contracts

Every phase must preserve the following behavior:

1. The root API, `./standalone-runner`, and `./node` package exports and their
   declarations remain stable.
2. Operation normalization, `INVALID_OPERATION`, per-operation author
   precedence, original operation indexes, `resolvedBy`, `resolvedTarget`, and
   `authorsUsed` remain unchanged.
3. Strict targeting never silently selects one of multiple exact matches.
   Fingerprints, paragraph IDs, occurrences, table context, and transient
   indexes retain their documented resolution semantics.
4. `preflightOperations(...)` remains read-only, deterministic, serializable,
   and consistent with the subsequent mutation path.
5. Comments still execute before text-changing operations when required for
   stable anchors. Atomic rollback must not commit allocator, numbering,
   comment, target-snapshot, or runtime-context state.
6. `inspectDocumentParts(...)`, targeting, ingestion export, and CLI `extract`
   continue to use canonical exact text. Performance work must not normalize
   whitespace or change accepted/rejected/current-view semantics.
7. Package operations keep comment IDs package-scoped, merge numbering by
   schema order, preserve unrelated part contents, validate before commit, and
   return original bytes on atomic failure or no-op.
8. The browser/root dependency graph remains free of Node filesystem/ZIP
   dependencies. Node-specific optimizations stay under `node/`.
9. Public pipeline exports such as `ReconciliationPipeline`,
   `serializeToOoxml`, and `wrapInDocumentFragment` cannot be removed in a
   minor release merely because the main engine stops using them internally.

## 3. Measurement Before Refactoring

Add deterministic performance fixtures before changing architecture:

- document sizes: approximately 100, 1,000, and 10,000 paragraphs;
- operation counts: 1, 10, 50, and 100;
- target shapes: unique text, duplicate text with stable descriptors, table
  cells, lists, comments, and mixed batches;
- operation outcomes: success, no-op, strict ambiguity, and atomic rollback;
- package cases: no artifacts, existing comments, existing numbering, and both.

Record wall time, parse count, serialize count, peak heap, and output size. Run
warm-up iterations and report median plus p95 rather than one timing. Store
machine-readable results under ignored `tmp/benchmarks/`; keep fixture builders
and thresholds in version control.

Performance acceptance should use relative comparisons on the same machine.
Initial targets are aspirational after all correctness gates pass:

- at least 5x lower median latency for 10+ paragraph operations on a large
  document;
- one full-document parse and one final full-document serialization for a
  successful batch, excluding validation/package reads;
- no more than 10% regression for single-operation calls;
- no material peak-heap regression after removing parse/serialize churn.

The 2026-09-04 Phase 1 benchmark (`npm run benchmark:session`) used 1,000
paragraphs, 10 replacements, two warmups, and seven measured iterations on
Node 24/Windows x64. The live batch recorded a 167.19 ms median and 224.25 ms
p95 versus 264.96 ms and 306.17 ms for sequential single-operation execution,
or a 1.58x median speedup. Full-document parse/serialize counts fell from 10/10
to 1/1 and output sizes matched. The 5x aspiration was not met because each
operation retains a full DOM savepoint so a thrown error or false no-op cannot
leak a partial mutation. That is an intentional accuracy tradeoff, not a reason
to remove the safeguard. Heap deltas were too GC-sensitive in this short run to
support a claim; retain them as observational output only.

Do not require byte-identical XML serialization. Require semantic equivalence,
accepted/rejected text parity, structural validation, and unchanged package
parts; namespace and attribute ordering may legitimately differ.

---

## Phase 1 — In-Memory Document Operation Session

**Status:** Complete (2026-09-04). Phase 2 created the boundary first; Phase 1
then added live-DOM execution with accuracy-preserving savepoints.

### Problem

Before Phase 1, the batch runner parsed the complete document at the batch
boundary, reparsed inside every single-operation call, serialized after each
mutation, and then parsed again to continue. The new agent APIs add more
metadata and preflight requirements, so the live DOM must also preserve target
identity and transaction state.

### Design

The internal `DocumentOperationSession` owned by the batch orchestrator
contains:

- the live `xmlDoc` and serializer;
- the original input string for rollback;
- one document-scoped revision allocator;
- the immutable initial target-reference snapshot;
- lazily built paragraph/ID/text indexes with explicit invalidation after
  mutations;
- numbering/list fallback context;
- accumulated comment and numbering payloads;
- operation results, authors, warnings, and execution order.

This is an internal execution object, not a second public API. Existing
`applyOperationToDocumentXml(...)` creates a one-operation session;
`applyOperationsToDocumentXml(...)` creates one session for the entire batch.
The Node facade continues to call the stable runner and then commits package
artifacts transactionally.

### Work

1. Split DOM-native operation functions from string compatibility wrappers.
2. Resolve targets against the live DOM while retaining the initial snapshot
   used by transient references.
3. Apply paragraph/table replacements by importing scoped results into the
   live document without serializing the entire document between operations.
   Scoped paragraph serialization is acceptable until individual engines gain
   DOM-native entry points.
4. Invalidate only affected paragraph indexes after replacement/insertion;
   never reuse stale node references.
5. Serialize once after the batch succeeds. On atomic failure, discard the
   session and return the exact original string and no generated artifacts.
6. Commit external runtime context only after the batch and package transaction
   succeed.

### Acceptance

- The benchmark targets above are met or the measured limitation is recorded.
- Existing operation, preflight, inspection, CLI, comment, numbering, and
  rollback tests pass unchanged.
- A new parse/serialize instrumentation test proves the full document is not
  reparsed per operation.
- Mixed batches retain original indexes, comment-first execution, resolution
  metadata, and per-operation authors.

### Implementation record

- A batch parses `word/document.xml` once, keeps the initial target-reference
  snapshot and one revision allocator, mutates one live DOM, and serializes
  once only after a successful changed batch.
- The existing scoped redline, list, table, highlight, and comment engines are
  unchanged. They still serialize only the selected scope and import their
  result into the live document, preserving their established accuracy.
- Every operation starts with a cloned-DOM and revision-allocator savepoint.
  Errors and reported no-ops restore that savepoint. This is deliberately more
  conservative than maximum-throughput mutation and prevents partial state or
  consumed revision IDs from leaking across operations.
- Atomic failure and all-no-op batches return the exact original XML without a
  final serialization. Non-atomic batches serialize only successfully
  committed mutations. Runtime context is committed only after successful
  batch serialization; the Node facade continues to isolate package context in
  its package transaction.
- The session owns lazy paragraph metadata and invalidates it after every
  committed mutation or restore, so future indexing work cannot reuse stale DOM
  nodes.
- `performance_phase1_session_tests.mjs` verifies real provider parse/serialize
  counts, expected accepted and rejected text, structural redline validation,
  list/table/comment/highlight behavior, equivalence with sequential execution,
  byte-exact no-op output, and zero-serialization atomic rollback.
- The checked benchmark reports the accuracy-first 1.58x result and the
  documented 10/10 to 1/1 parse/serialize reduction above.
- All 57 isolated tests and coverage pass. Coverage is 89.32%
  statements/lines, 77.09% branches, and 92.78% functions. Isolation, types,
  lint, build, the 143-file package dry run, package self-imports, and
  `git diff --check` also pass.
- Desktop Word, corpus, visual, XSD, and LibreOffice lanes were not rerun. The
  scoped reconciliation engines and emitted revision construction were not
  changed, while the new equivalence test validates exact accepted/rejected
  outcomes and the existing package/OOXML suite exercises the integrated path.

---

## Phase 2 — Modularize the Runner Around Existing Boundaries

**Status:** Complete (2026-09-04), intentionally completed before the Phase 1
in-memory optimization.

### Problem

Before this phase, `standalone-operation-runner.js` mixed compatibility
exports, scheduling, targeting, DOM mutation, list/table heuristics, and result
assembly. The operation-contract and preflight modules already established
boundaries that the decomposition reused.

### Work

1. Add `services/document-operation-session.js` for live DOM state, indexes,
   allocators, invalidation, serialization, and rollback.
2. Add `services/batch-operation-orchestrator.js` for scheduling, dependency
   order, context commit, atomic policy, and batch result assembly.
3. Add `services/document-operation-applier.js` for canonical single-operation
   dispatch and result metadata.
4. Add `services/operation-heuristics.js` for adjacency insertion, explicit
   range insertion, and list/table scope-expansion heuristics.
5. Keep `services/document-operation-contract.js` and
   `services/operation-preflight.js` authoritative; do not duplicate their
   normalization or target diagnostics in the new modules.
6. Turn `services/standalone-operation-runner.js` into a small compatibility
   facade re-exporting the same runtime functions. Preserve
   `services/standalone-operation-runner.d.ts` and the package subpath.
7. Replace the runner's import from `../index.js` with direct leaf-module
   imports. This removes the current circular dependency without changing the
   public root exports.

### Acceptance

- The facade is preferably under 250 lines and new responsibility modules are
  preferably under 500 lines; exceptions require a short rationale.
- No operation-normalization, preflight, or package-transaction fork exists.
- Root, standalone-runner, and Node self-import type fixtures pass.
- Coverage does not fall below the checked baseline for extracted behavior.

### Implementation record

- `standalone-operation-runner.js` is a compatibility-only facade and preserves
  the same runtime exports, declaration file, and package subpath.
- `document-operation-applier.js` owns validation, author resolution, dispatch,
  and uniform result metadata; `batch-operation-orchestrator.js` owns stable
  scheduling, atomic policy, artifact aggregation, and context commit.
- `document-operation-session.js` owns parse state, exact rollback input,
  revision allocation, and isolated runtime-context clone/commit helpers. Its
  live-DOM optimization remains Phase 1 work; Phase 2 does not claim the batch
  parse/serialize performance target.
- `operation-heuristics.js` owns the extracted adjacency and explicit-range
  decisions. Existing shared list/table scope decisions remain authoritative in
  `core/list-targeting.js` and `core/table-targeting.js` rather than being
  duplicated.
- Coupled OOXML construction and replacement helpers remain together in
  `document-operation-mutations.js` (about 1,000 lines). This is the documented
  size exception: splitting those interdependent mutation paths before the
  live-session work would add interfaces and semantic churn without reducing
  reparsing. The public facade, applier, orchestrator, session, and heuristics
  modules are each below 250 lines.
- The operation implementation no longer imports `../index.js`; it imports leaf
  modules directly. `performance_phase2_boundary_tests.mjs` locks the facade
  identities, leaf-import rule, exact rollback, isolated context, and stable
  scheduling.
- All 56 isolated tests and coverage pass. The post-refactor coverage totals
  are 89.16% statements/lines, 77.18% branches, and 92.71% functions, meeting
  or exceeding every recorded baseline dimension.
- `npm run test:isolation`, `npm run check:types`, `npm run lint`, `npm run
  build`, `npm pack --dry-run`, package self-import checks, and `git diff
  --check` pass. The dry-run package contains all new service modules while the
  stable standalone subpath still exposes exactly its four prior functions.
- Desktop Word, corpus, visual, XSD, and LibreOffice lanes were not rerun for
  this boundary-only phase because reconciliation routing and emitted OOXML
  behavior did not change; the existing script-level and package tests cover
  the moved behavior.

---

## Phase 3 — Target and Traversal Hot Paths

**Status:** Complete (2026-09-04). The session and preflight paths now share
one document-scoped metadata representation, while all caches remain bounded
to their owning operation or preflight call.

### Work

1. Build paragraph metadata once per operation session: canonical text,
   normalized text, paragraph ID, fingerprint inputs, table context, and
   document index. Share the immutable representation with preflight where
   possible.
2. Replace repeated `textSpans.filter(...)`/paragraph scans in target detection
   with a single grouped map, reusing existing paragraph-info builders where
   their semantics match.
3. Change revision-ID seeding to a single pointer-based tree traversal that
   recognizes revision-bearing elements without allocating an array of every
   element. Do not replace one full traversal with nine independent tag
   traversals unless benchmarks show that is faster in both browser DOM and
   `@xmldom/xmldom`.
4. Replace `Array.from(...)` only in measured hot recursive loops. Keep it where
   it improves clarity and does not dominate profiles.
5. Benchmark diff preprocessing before adding a manual common-prefix/suffix
   trim: `diff-match-patch` already performs that optimization internally.
   Add another fast path only if tokenization before DMP is proven to dominate.
6. Bound caches to one session. Do not retain DOM nodes or document-derived text
   in process-global maps.

### Acceptance

- Target resolution remains byte-for-byte deterministic in its JSON metadata.
- Canonical text and fingerprint compatibility tests pass.
- Profiles demonstrate reduced traversal/allocation cost on large fixtures.
- Browser and Node benchmarks show no meaningful single-operation regression.

### Implementation record

- `buildParagraphMetadataIndex(...)` computes canonical text, normalized text,
  paragraph ID, fingerprint, table context, and the one-based document index in
  one pass. It exposes grouped paragraph-ID and normalized-text maps plus a
  node-keyed metadata map; entries are immutable and retain document order.
- `DocumentOperationSession` constructs the index once and uses it to build the
  initial target-reference snapshot. It preserves the earlier paragraph-array
  accessor for internal compatibility, exposes the richer index separately,
  and invalidates both after any committed mutation or savepoint restore.
- Batch preflight builds one metadata index and shares it with every strict
  target resolution. Mutation targeting and resolution metadata use the same
  cached entries, avoiding repeated canonical-text, ID, table-context, and
  fingerprint traversal without changing the resolver's public JSON shapes.
- Strict duplicate detection uses the grouped normalized-text map. Exact-text
  filtering still occurs within the matched group, so whitespace normalization
  does not broaden strict selection and duplicate targets remain ambiguous.
- `RevisionIdAllocator.seed(...)` now uses a pointer-based depth-first walk and
  examines revision-bearing elements without materializing every element into
  an array. A regression deliberately disables universal tag queries to prove
  the allocator uses the pointer traversal.
- `performance_phase3_target_hot_paths_tests.mjs` covers cached/uncached
  resolution parity for exact text, duplicate occurrences, paragraph IDs,
  fingerprints, transient indexes, and table context; it also checks snapshot
  construction, session reuse/invalidation, and revision-ID allocation.
- `npm run benchmark:targeting` is a deterministic 10,000-paragraph benchmark
  with two warmups and seven measured iterations. On Node 24/Windows x64, one
  operation improved from a 45.10 ms median to 36.53 ms (1.23x), satisfying
  the no-regression constraint. One hundred resolutions improved from
  1,948.49 ms to 41.31 ms (47.16x). The uncached comparison includes the same
  initial target snapshot required before this phase, so the comparison does
  not omit session setup work.
- `scripts/benchmark-targeting-browser.html` runs the same fixture, warmups,
  percentiles, snapshot work, and resolver comparison against a native browser
  DOM. Query parameters select paragraph, operation, and iteration counts so
  single-operation and batch profiles can be reproduced without a Node DOM.
  On Chrome 151/Windows x64, the 10,000-paragraph single-operation median moved
  from 43.10 ms to 42.30 ms (1.02x), satisfying the browser no-regression
  constraint. One hundred resolutions moved from 1,363.20 ms to 37.40 ms
  (36.45x). Both browser measurements used two warmups and seven measured runs.
- No extra diff prefix/suffix fast path was introduced: the work did not show
  token preprocessing to be the dominant cost, and `diff-match-patch` already
  performs prefix/suffix trimming internally.

---

## Phase 4 — List-Domain Consolidation Without Replacing Canonical Text

**Status:** Complete (2026-09-04). Canonical consumers now use the established
paragraph-text API, specialized offset/sentinel walkers remain specialized,
and all list-facing parsers share one dependency-light grammar.

### Current state

Canonical extraction is already implemented in `core/paragraph-text.js`; this
phase must not create `getParagraphVisibleText(...)` in `core/word-xml.js` or
move the source of truth again.

Engine mapping walkers may remain specialized when they produce more than
visible text. The goal is shared semantics, not forcing offset/sentinel logic
through a string-only helper.

### Work

1. Inventory every remaining paragraph-text walker and label it as either:
   - canonical visible-text consumption; or
   - specialized mapping with offsets, formatting, fields, references, or
     structural sentinels.
2. Replace only the first category with `extractCanonicalParagraphText(...)`.
   For specialized walkers, add parity tests for the visible projection while
   retaining their richer data models.
3. Extract shared list-token grammar and marker classification into a small
   dependency-light module used by markdown parsing, routing, and fallback.
4. Keep targeting, structural planning, OOXML writing, numbering allocation,
   and inspection numbering resolution in focused modules. Avoid replacing six
   files with one oversized `list-service.js`.
5. Make single-line and multi-line list paths consume the same parsed list-item
   representation and numbering-style vocabulary.
6. Preserve `inspectDocumentParts(...)` labels, exact excerpts, target
   descriptors, numbering merge behavior, and CLI JSON shapes.

### Acceptance

- `core/paragraph-text.js` remains the only public canonical text definition.
- Specialized walkers have explicit parity tests for tabs, breaks, moves,
  fields, notes, and revision views.
- List syntax is defined once without merging unrelated package/targeting/writer
  responsibilities.
- All list, numbering, inspection, package, and Word differential fixtures pass.

### Implementation record

- `docs/PERFORMANCE-CONSOLIDATION.md` inventories the remaining text walkers as
  canonical consumers or specialized mappings and records the semantic boundary
  of each. `core/paragraph-text.js` remains the sole canonical visible-text
  definition.
- `engine/table-cell-context.js`, a string-only consumer, now calls
  `extractCanonicalParagraphText(...)`. Mapping walkers in ingestion, surgical
  editing, formatting, reconstruction, and comment anchoring keep their richer
  offset, run, sentinel, or reference models.
- The accepted visible projections in ingestion and surgical span mapping now
  include soft hyphens consistently with the canonical walker. Existing
  structural ownership and zero-width handling for fields and note references
  are retained.
- `pipeline/list-markers.js` now owns list-marker matching, stripping,
  classification, numbering-style inference, outline-depth parsing, and the
  common parsed list-item representation. Pipeline analysis, orchestration
  parsing/normalization, and list targeting consume those helpers; numbering
  allocation, structural planning, OOXML writing, and package merging remain
  in their focused modules.
- The shared vocabulary recognizes bullet markers (`-`, `*`, `+`, and common
  Unicode bullets), decimal/composite decimal markers, alphabetic markers, and
  Roman markers. It reports `bullet`/`numbered`, the established numbering
  style, indentation level, and optional outline level.
- `performance_phase4_list_and_text_parity_tests.mjs` exercises marker parity
  across every consumer, including redundant marker stripping, and compares
  canonical, ingestion, and surgical visible projections for tabs, breaks,
  hyperlinks, soft/non-breaking hyphens, fields, notes, insertions, deletions,
  and moves.
- The complete automated suite and the 47-fixture Microsoft Word differential
  pass. The Word set includes ten list cases plus tables, comments, hyperlinks,
  fields, footnotes/endnotes, formatting-only edits, and independent Accept All
  and Reject All expectations.

---

## Phase 5 — Diff-Engine Consolidation With a Compatibility Window

**Status:** Complete (2026-09-04). The preferred focused list generator now
serves the safe single-paragraph route, while multi-paragraph marked-list edits
remain on the compatibility pipeline until equally strong structural evidence
supports moving them.

### Problem

Reconstruction, surgical, table, and legacy pipeline paths overlap, but they do
not yet have identical structural capabilities. The legacy pipeline also has
public exports, so deleting its files in a nominally non-breaking phase would
conflict with the current API contract.

### Work

1. Create a capability matrix for paragraphs, lists, tables, hyperlinks,
   fields, comments, notes, formatting-only edits, and structural insertions.
2. Instrument route selection and establish corpus frequency before changing
   defaults.
3. Add missing list/numbering behavior to the preferred engine behind existing
   result shapes, including `numberingXml`, `sourceType`, warnings, and
   `useNativeApi` behavior.
4. Move internal list routing off `ReconciliationPipeline` only after focused,
   fuzz, package, Word accept/reject, and visual tests demonstrate parity.
5. Keep `ReconciliationPipeline`, `serializeToOoxml`, and
   `wrapInDocumentFragment` as compatibility exports for the remainder of the
   current major version. Mark genuinely obsolete APIs deprecated in docs and
   the changelog before removal.
6. Remove public legacy modules only in a planned major release. Internal dead
   code may be removed earlier when no export or deep-import contract is
   affected.

### Acceptance

- Route changes are supported by benchmark and correctness data, not line-count
  targets alone.
- Accepted/rejected text, revision metadata, structural references, formatting,
  comments, tables, and numbering match the established oracle fixtures.
- No public export disappears in a minor release.
- Any eventual major-version removal has a migration example and deprecation
  period.

### Implementation record

- `engine/route-selection.js` contains an executable capability matrix for
  format-only, surgical, reconstruction, table, direct-list, and compatibility-
  pipeline routes. It records primary structures, retained structures, artifact
  behavior, and the reason each route remains distinct.
- Route instrumentation is opt-in and internal: `_routeInstrumentation.onRoute`
  observes the selected route and context without changing public result
  shapes or providing a routing-policy override. `npm run profile:routes`
  provides a deterministic synthetic frequency profile; its checked four-case
  sample selects table, reconstruction, direct-list, and compatibility-list
  exactly once each.
- A list edit sourced from one paragraph now calls the focused
  `executeListGeneration(...)` path directly with the established numbering
  service and document-scoped revision allocator. This preserves wrapping,
  `numberingXml`, `sourceType`, warnings, and native-API signaling.
- Multi-paragraph marked-list edits continue through `ReconciliationPipeline`.
  That path performs run-aware patching across existing paragraph boundaries;
  retaining it is an intentional accuracy boundary, not unfinished deletion.
- The structural list fallback uses the focused generator directly during
  normal execution but continues to honor an explicitly injected pipeline for
  compatibility tests and existing internal integrations.
- `performance_phase5_route_consolidation_tests.mjs` compares direct and legacy
  list generation with independent accepted/rejected extraction, verifies
  numbering-package output, locks single- versus multi-paragraph route choices,
  checks the capability matrix, and proves the public compatibility export
  identities remain available.
- `ReconciliationPipeline`, `serializeToOoxml`, and `wrapInDocumentFragment`
  remain public and supported throughout the current major version. No public
  API was deprecated or removed by this phase; future removal still requires a
  major version, a prior deprecation notice, and a migration example.
- All automated, coverage, isolation, type, lint, build, package-dry-run, and
  fixture-export checks pass. The 62-file coverage result is 89.81% statements/
  lines, 77.34% branches, and 93.04% functions, improving every recorded
  baseline dimension. The 47/47 Microsoft Word differential also passes.
- The 60-scenario SuperDoc Microsoft Word corpus passes after the route change,
  including multi-paragraph and nested lists, list insertions/deletions, table
  reconciliation, page headers, long documents, and mixed batches. Its package
  guard also verifies that every unrelated DOCX part remains byte-identical.
- The local ECMA-376 XSD command could not run because this Windows environment
  has the WSL launcher but no usable `/bin/bash`. Schema validation remains in
  the nightly independent-oracle workflow; runtime structural validation, the
  synthetic Word differential, and the real-document Word corpus all passed.

---

## Phase 6 — Faster Tests While Preserving Isolation

**Status:** Complete (2026-09-04). The JavaScript suite now uses bounded
subprocess concurrency while every test file retains a fresh Node process.

### Problem

The sequential runner starts 57 Node processes. Importing every test into one
process would be faster, but it would also combine global XML providers,
loggers, default authors, revision counters, module caches, environment changes,
and top-level test side effects that are currently isolated.

### Work

1. Replace synchronous shell-string execution with bounded parallel
   `child_process.execFile`/`spawn` calls. Keep one process per test file.
2. Default concurrency conservatively (for example, CPU count capped at 4) and
   allow `DOCX_TEST_CONCURRENCY=1` for deterministic reproduction.
3. Preserve sorted result reporting, per-file timeout handling, captured
   stdout/stderr, failure-marker detection, and nonzero exit behavior.
4. Keep Word COM, visual, corpus, and filesystem-collision-sensitive lanes
   serial unless their runners explicitly coordinate resources.
5. Confirm `c8` still merges subprocess coverage correctly. If it does not,
   retain a coverage-specific serial command.
6. Consider Node's native test runner only after tests migrate from top-level
   script assertions and their global-state assumptions are documented.

### Acceptance

- Median `npm test` time improves by at least 40% on the same machine.
- Serial and parallel modes execute the same files and report the same failures.
- No flaky failures occur across at least 20 repeated parallel runs.
- `npm run test:coverage`, isolation checks, and Windows paths remain valid.

### Implementation record

- `scripts/run-tests.mjs` now launches test files with asynchronous
  `execFile(...)` workers instead of synchronous shell command strings. It uses
  `process.execPath` plus an argument array, avoiding shell quoting and keeping
  Windows paths safe.
- The default worker count is the runtime's available parallelism capped at
  four. `DOCX_TEST_CONCURRENCY=1` provides deterministic serial reproduction;
  any positive integer is accepted explicitly and is bounded by the discovered
  test count. Invalid values fail before starting tests.
- Test discovery and final reporting remain filename-sorted regardless of
  completion order. Every child retains its own process, 30-second timeout,
  captured stdout/stderr, failure-marker detection, and nonzero-exit handling.
  Success summaries stay on stdout and detailed failure output stays on stderr.
- The Word COM, synthetic visual, corpus, corpus-visual, and other independent
  oracle commands remain separate serial runners. JavaScript tests that write
  fixtures use distinct temporary directories; the runner does not combine
  their module state or files into one process.
- `performance_phase6_parallel_runner_tests.mjs` locks sorted discovery,
  setup-helper exclusion, the concurrency bound, separate child PIDs, serial/
  parallel result parity, successful-exit failure markers, nonzero exits,
  stdout/stderr capture, timeouts, and deterministic reporting.
- `npm run benchmark:tests` runs one warmup and three paired measurements and
  writes machine-readable output to ignored
  `tmp/benchmarks/test-runner-latest.json`. On Node 24/Windows x64, concurrency
  four reduced the median from 18,018.68 ms to 7,964.47 ms: a 55.80% reduction
  and 2.26x speedup, exceeding the 40% acceptance target.
- Twenty consecutive four-worker runs passed all 63 test files with no flaky
  failure. Their observed range was 7.49–8.24 seconds, with a 7.86-second
  average. A separate `DOCX_TEST_CONCURRENCY=1` run passed the same 63 files.
- c8 successfully merged the parallel subprocess data, so the regular coverage
  command remains parallel. The checked result is 89.90% statements/lines,
  77.40% branches, and 93.10% functions, meeting or improving the Phase 5
  totals. No coverage-specific serial fork was added.
- Node's native test runner was not adopted: these files remain top-level
  assertion scripts with intentional process-global isolation. A future
  migration should first document and change that execution contract.

---

## 4. Revised Execution Roadmap

1. **Measurement and safety baseline**
   - Add benchmarks, parse/serialize instrumentation, and result-equivalence
     fixtures.
2. **In-memory session foundation — complete 2026-09-04**
   - Introduce the focused operation-session module behind the existing runner,
     retain compatibility wrappers, and measure the improvement. Per-operation
     savepoints remain mandatory unless an equally strong isolation proof
     replaces them.
3. **Boundary-preserving decomposition — complete 2026-09-04**
   - Extract the applier, orchestrator, and heuristics around the proven session
     without changing public behavior. This step was completed first; the
     session boundary exists but still uses legacy per-operation serialization.
4. **Measured hot paths and parallel test runner — complete 2026-09-04**
   - Optimize indexes/traversals and add bounded subprocess concurrency.
5. **List-domain cleanup**
   - Share syntax and visible-text semantics without collapsing distinct
     responsibilities.
6. **Engine consolidation**
   - Change internal routing only after capability, corpus, and Word parity;
     defer public removals to a major release.

## 5. Required Validation Per Phase

At minimum run:

```bash
npm test
npm run test:coverage
npm run test:isolation
npm run check:types
npm run lint
npm run build
npm pack --dry-run
```

For phases that change reconciliation routes, list construction, comments,
numbering, or package output, also run the applicable Word differential, corpus,
visual, XSD, and LibreOffice lanes described in `docs/TESTING.md` and
`docs/VALIDATION.md`.

Each phase should update benchmark results, `ARCHITECTURE.md`, `CHANGELOG.md`,
`docs/TESTING.md`, and this plan as implementation decisions become concrete.
