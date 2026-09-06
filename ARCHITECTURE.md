# Reconciliation Core Architecture

This document describes the standalone, host-agnostic OOXML reconciliation package and how to work inside it safely.

## Scope

This repository contains only the publishable package surface:

- `adapters/`
- `core/`
- `engine/`
- `pipeline/`
- `services/`
- `orchestration/`
- `node/` (separate Node-only package facade)
- `index.js`
- `index.d.ts`
- `dist/`

No Word add-in entrypoints or host-specific integration layers are part of this package.

## Goals

- Preserve Word-compatible redlines by editing OOXML directly.
- Keep core logic host-independent (no Office.js globals, no Word API calls).
- Reuse the same engine in browser, Node.js, and other JavaScript runtimes.
- Keep generated OOXML schema-safe through shared Word element creation and
  centralized revision metadata helpers.

## Folder Layout

```text
.
├── adapters/
│   ├── config.js
│   ├── logger.js
│   └── xml-adapter.js
├── core/
│   ├── paragraph-text.js
│   └── word-xml.js
├── engine/
│   ├── oxml-engine.js
│   ├── surgical-mode.js
│   ├── reconstruction-mode.js
│   ├── run-builders.js
│   └── formatting-removal.js
├── orchestration/
├── pipeline/
├── services/
│   ├── comment-engine.js
│   ├── numbering-helpers.js
│   ├── revision-comment-management.js
│   ├── document-inspection.js
│   ├── document-operation-session.js
│   ├── document-operation-applier.js
│   ├── document-operation-mutations.js
│   ├── batch-operation-orchestrator.js
│   ├── operation-heuristics.js
│   ├── standalone-docx-plumbing.js
│   └── standalone-operation-runner.js
├── node/
│   ├── docx-document.js
│   ├── cli.js
│   ├── zip-archive.js
│   └── index.js
├── index.js
└── index.d.ts
```

## Entry Points

- `index.js` (Root exports containing the reconciliation logic)

## Module Responsibilities

- `adapters/config.js`
  - Runtime configuration for defaults (`setDefaultAuthor`, `getDefaultAuthor`, `setPlatform`, `getPlatform`).
- `adapters/xml-adapter.js`
  - XML parser/serializer injection for browser or Node.js runtimes.
- `adapters/logger.js`
  - Runtime logger injection and shared logging methods.
- `core/*`
  - Shared types, OOXML identity helpers, target resolution, list/table targeting heuristics, and XML query helpers.
- `core/word-xml.js`
  - Namespace-safe Word element creation, tracked-change detection, and OOXML payload source-shape helpers.
- `core/types.js`
  - Shared model enums/types plus revision metadata generation and document-aware revision ID seeding.
- `core/paragraph-text.js`
  - Canonical accepted/rejected/current-view text shared by targeting,
    ingestion, and inspection.
- `core/paragraph-targeting.js`
  - Target descriptors and session-scoped paragraph metadata indexes. The
    index groups IDs and normalized text while retaining canonical text,
    fingerprints, document order, and table context for deterministic reuse.
- `core/redline-validation.js`
  - Runtime structural validation (`validateRedlineOoxml`) mirroring the test-suite invariants: no nested revisions, `w:delText` inside `w:del`, complete revision metadata, unique revision ids, preserved boundary whitespace.
- `engine/oxml-engine.js`
  - Main reconciliation router, mode selection, existing-revision policy gate, and status/error result handling.
- `engine/route-selection.js`
  - Internal reconciliation capability matrix and opt-in diagnostic route
    instrumentation. It observes policy; it does not expose a route override.
- `engine/run-builders.js`
  - Shared builders for insertion/deletion wrappers, paragraph-mark revisions, visible run content, and run-property changes.
- `engine/surgical-*.js`
  - Surgical run splitting, diff application, and span helpers for localized edits that preserve surrounding markup.
- `engine/formatting-removal.js`
  - Shared formatting removal and highlight helpers.
- `pipeline/list-markers.js`
  - Dependency-light canonical list-marker grammar, classification, numbering
    vocabulary, and parsed list-item representation.
- `pipeline/list-generation.js`
  - Structural paragraph-to-list generation, paragraph-mark revisions, and
    selective inheritance of source typography for inserted list runs.
- `pipeline/structured-content.js`
  - Strict decomposition and planning for atomic mixed heading/paragraph/list/
    table replacements. It rejects ambiguous table-like pipe text rather than
    allowing a fallback to literal paragraphs.
- `pipeline/*`
  - Ingestion, markdown preprocessing, diffing, patching, and serialization stages. Ingestion treats deleted and moved-from content as non-visible text and inserted/moved-to content as visible text.
- `services/comment-engine.js`
  - Comment creation and package-level comment XML handling.
- `services/numbering-helpers.js`
  - Dynamic numbering ID allocation, numbering payload remapping, and schema-order-safe numbering merges.
- `services/standalone-docx-plumbing.js`
  - Package-level extraction/wiring/validation for `word/document.xml`, `word/numbering.xml`, and `word/comments.xml`.
- `services/revision-comment-management.js`
  - OOXML transforms for accepting/rejecting insertion, deletion, move, paragraph-mark, and property-change revisions by author/all-authors, plus deleting comments by author/all-authors.
- `services/standalone-operation-runner.js`
  - Stable compatibility facade for full-document operations. It only re-exports
    preflight, single-operation application, batch application, and scheduling.
- `services/document-operation-applier.js`
  - Canonical single-operation validation, author resolution, dispatch, and
    result metadata assembly.
- `services/document-operation-mutations.js`
  - Coupled OOXML mutation implementations for redline, highlight, and comment
    operations. These use leaf-module imports and never import the root entry.
- `services/batch-operation-orchestrator.js`
  - Comment-first stable scheduling, atomic policy, artifact aggregation,
    per-operation results, one final document serialization, and deferred
    runtime-context commit.
- `services/document-operation-session.js`
  - One live document DOM and revision allocator per invocation, immutable
    start-of-batch target snapshot, lazy paragraph metadata, per-operation DOM
    and allocator savepoints, exact original XML rollback, index invalidation,
    artifact/result accumulation, and isolated runtime-context helpers.
- `services/operation-heuristics.js`
  - DOM-light decisions for list/plain adjacency insertion and explicit-range
    list insertion. Canonical list/table targeting remains in `core/*`.
- `services/document-operation-contract.js`
  - Compatibility normalization, canonical operation kinds, author precedence,
    and stable runtime validation for document operations.
- `services/operation-preflight.js`
  - Read-only strict target and anchor resolution, revision-policy diagnostics,
    artifact prediction, and same-paragraph conflict reporting.
- `services/document-inspection.js`
  - Read-only paragraphs/comments inventory with target identity, headings,
    revision authors, table context, and advisory visible numbering.
- `node/docx-document.js`
  - Transactional whole-DOCX editing, artifact wiring, validation, and rollback.
    This surface is excluded from the browser/root dependency graph.
- `node/cli.js` and `bin/docx-redline.js`
  - Cross-platform, JSON-only agent command boundary. Read commands never
    mutate; write commands require attribution, use package transactions, and
    only overwrite source files under explicit `--in-place` authorization.
- `orchestration/*`
  - Route planning and list fallback orchestration utilities.

## End-to-End Flow

1. Caller imports from `index.js`.
2. Caller configures XML provider/logger/defaults when needed via `adapters/*`.
3. Caller invokes reconciliation APIs (`applyRedlineToOxml`, operation runner, ingestion/export helpers).
4. A full-document batch parses once, captures its initial target snapshot, and
   creates an operation savepoint before each mutation.
5. `engine/oxml-engine.js` routes each selected scope to format, table, list,
   surgical, or reconstruction flows; successful scoped output is imported
   into the live document.
6. Failed and no-op operations restore their savepoint. A successful changed
   batch serializes the full document once; atomic failures and all-no-op
   batches return the exact input string without serialization.
7. Pipeline/services return OOXML, optional package artifacts (`numberingXml`, comments payloads), and non-breaking `status`/`error` fields where applicable.
8. Optional revision/comment management transforms can accept/reject revisions, including move revisions, or delete comments by author.
9. Caller writes resulting XML back to package/document boundaries.

## Public Surfaces

- Primary: `index.js`
- Types: `index.d.ts`
- Stable XML operation runner: `@ansonlai/docx-redline-js/standalone-runner`
- Node-only DOCX facade: `@ansonlai/docx-redline-js/node`

Keep public exports centralized through `index.js`; deep imports are supported by
the package `exports` map for advanced consumers, but new public APIs should
still be re-exported from `index.js`.

## Reliability Guardrails

- Create Word namespace elements with `createWordElement(xmlDoc, 'w:...')`.
  Avoid direct `document.createElement('w:*')` or `createElementNS(NS_W, 'w:*')`
  outside `core/word-xml.js`.
- Generate tracked-change metadata through `createRevisionMetadata(author)` so
  `w:id`, `w:author`, and `w:date` stay consistent and document-unique.
- Seed revision IDs from parsed input with `seedRevisionIdsFromDocument(xmlDoc)`
  before emitting new tracked changes.
- Treat revision metadata inside cloned content as identity-bearing state, not
  ordinary formatting. When a run containing `w:rPrChange` is split or cloned,
  preserve an existing revision ID on at most one output run and allocate a
  fresh document-scoped ID for every additional copy.
- Use `containsTrackedChanges(xmlDoc)` before redlining existing revisions unless
  the caller explicitly chooses the `existingRevisions: 'accept-all-first'` policy.
- Do not write unknown `result.oxml` payloads directly into `word/document.xml`;
  normalize with `extractReplacementNodesFromOoxml(...)` or use
  `applyOperationToDocumentXml(...).documentXml` for full-document replacement.
- Run `validateRedlineOoxml(oxml)` on generated output before packaging it;
  it reports structural invariant violations as `{ valid, issues }`.

## Integration Contracts

### Targeting and text fidelity

- Paragraph targeting and replacement have different contracts. Targeting may
  use normalized text and fallback heuristics; replacement content is literal
  and must preserve tabs, line breaks, non-breaking spaces, repeated spaces,
  and boundary whitespace.
- Text extraction used for targeting represents the visible accepted view of
  tracked content: insertions are visible and deletions are excluded.
- Paragraph indexes are transient integration references, not user-facing
  document identifiers. Prefer paragraph IDs plus exact text where available.
- Ambiguous matches are unsafe for document mutation. New targeting surfaces
  should report candidate matches or `AMBIGUOUS_TARGET` rather than silently
  choosing the first paragraph.
- Do not remove operation-session savepoints merely to improve throughput.
  Any replacement must prove that a thrown error or false no-op cannot leak a
  partial DOM mutation or revision-ID allocation into later operations.

### Structural paragraph-to-list replacement

- In WordprocessingML, `w:numId w:val="0"` suppresses numbering. It represents
  no reusable list context and must never be returned by numbering allocation
  or assigned to a generated list item. Generated lists use a positive `numId`
  bound to a compatible `w:abstractNum` definition.
- Expanding one non-list paragraph into multiple list paragraphs is a block
  replacement, not an inline replacement. The original content remains in its
  own paragraph with a deleted paragraph mark; every replacement item occupies
  its own paragraph with an inserted paragraph mark and inserted text. This is
  required so Word can accept the edit as the new paragraphs or reject it as
  the exact original paragraph without joining the old heading to item one.
- Inserted list runs selectively inherit source typography (`w:rFonts`, size,
  language, and related script properties). Semantic emphasis such as bold,
  underline, italic, or strike is emitted only when requested by the replacement
  markup; heading emphasis must not leak into ordinary list body text.
- Low-level list output can carry a package-fragment sentinel needed by its
  standalone payload shape. Full-document mutation removes that sentinel while
  importing replacement nodes, so it cannot become an empty document paragraph.
- Tests for this route must verify positive numbering IDs, separate physical
  paragraphs, formatting inheritance/non-inheritance, valid revision markup,
  and exact accepted and rejected paragraph sequences.

### Atomic mixed-content replacement

- `planStructuredReplacement(...)` converts agent-authored Markdown into one
  replacement operation carrying `structuredContent: true`. The block plan is
  diagnostic metadata; execution remains one mutation so replacing the anchor
  cannot invalidate later blocks.
- Mixed-content parsing recognizes explicit Markdown headings, blank-line
  paragraph boundaries, adjacent list items, and contiguous table rows. Tables
  require a header separator and consistent column counts. Invalid input fails
  with `STRUCTURED_CONTENT_INVALID`; silently degrading table pipes into visible
  paragraph text is forbidden on this route.
- Newly inserted tables inside a mixed replacement are tracked once at block
  scope. Cell runs are not nested in additional `w:ins` wrappers, preserving the
  no-nested-revisions invariant while allowing Reject All to remove the table.
- Agents must use explicit heading markers and valid table grammar rather than
  relying on capitalization or layout inference. The engine preserves the
  declared block types and does not guess legal-document semantics.

### Package artifacts and transactions

- A real `.docx` stores document markup, comments, numbering, relationships,
  and content types in separate parts. APIs operating on one XML string cannot
  claim to update artifacts held in another part.
- Comment IDs must be allocated against both document anchors and the existing
  `word/comments.xml` part. Revision IDs must be allocated against the complete
  document revision scope.
- Numbering payloads must be merged with `mergeNumberingXmlBySchemaOrder` when
  a package already contains numbering definitions; replacement is not a merge.
- Safe package mutation is transactional: retain the original package, apply
  operations, merge all artifacts, validate redline markup and package wiring,
  and commit only if every required check succeeds.
- `openDocx(buffer)` implements this transaction for Node without adding a ZIP
  dependency to browser or XML-only consumers. Unmodified part contents remain
  byte-identical after extraction, although the ZIP container is reserialized.

### Operation results

- `hasChanges: false` does not imply success. Callers must check `status`,
  `error`, and warnings to distinguish errors, missing anchors, and true no-ops.
- Batch results must preserve the caller's original operation indexes even when
  execution is reordered for stable anchors.
- Author attribution is externally visible document data. Agent-facing APIs
  should require an explicit author and report the author used for every
  operation rather than relying silently on a configured fallback.
- Operation-level authors override the batch author. Runtime results expose
  `authorUsed`, `authorsUsed`, `operationType`, `resolvedBy`, and resolved target
  metadata so integrations can audit what the engine actually selected.
- `preflightOperations` is the read-only safety boundary for agent-generated
  batches. It uses strict targeting by default; mutation APIs retain permissive
  legacy targeting unless `strictTargets: true` is requested. In v1.0.0,
  application will default to strict targeting; in the current warning cycle,
  permissive resolution that chooses among multiple identical paragraphs emits
  `AMBIGUOUS_TARGET_HEURISTIC_USED` with candidate count and migration guidance.
  Both preflight and mutation runners share the same candidate resolver
  (`resolveTargetParagraph`).

### Mutation receipts and output reconciliation oracle

- `ReceiptCollector` tracks exact allocations (revision IDs with kind and target part,
  comment IDs, numbering IDs, relationship IDs, affected targets, and warnings)
  directly at the point of allocation/attachment during operation execution.
- Collector state is snapshotted within operation savepoints. Failed or rolled-back
  operations cleanly restore prior collector state without leaking orphaned allocations.
- Single operations expose `result.receipt`; batches expose per-item `results[i].receipt`
  and top-level `result.receipts`, with dispositions (`applied`, `refused`, `no_change`,
  `rolled_back`, `not_attempted`) and `committed: true/false`.
- The output reconciliation oracle (`reconcileReceiptsAgainstOutput`) performs an
  independent, non-negotiable verification: every durable ID reported as committed
  is parsed from the serialized output XML parts (`word/document.xml`, `word/comments.xml`,
  `word/numbering.xml`). Any discrepancy immediately fails the transaction and triggers
  atomic rollback.


## Build Output

`npm run build` generates CDN-ready ESM bundles under `dist/`:

- `dist/docx-redline-js.esm.js`
- `dist/docx-redline-js.esm.js.map`
- `dist/docx-redline-js.esm.min.js`
- `dist/docx-redline-js.esm.min.js.map`

The bundle inlines `diff-match-patch` and keeps `@xmldom/xmldom` external.

## Testing

- `npm test`
  - Runs all `tests/*.mjs` except setup helpers in separate Node processes with
    a conservative four-worker cap. Use `DOCX_TEST_CONCURRENCY=1` to reproduce
    the same sorted suite serially.
- `npm run benchmark:tests`
  - Compares paired serial and bounded-parallel suite runs and records JSON
    under ignored `tmp/benchmarks/`.
- `npm run test:isolation`
  - Runs boundary checks for Word API markers and dependency-graph isolation.
- `npm run check:types`
  - Smoke-checks `index.d.ts`.
- `node scripts/export-validation-fixtures.mjs`
  - Writes release-time validation fixtures to `tmp/validation-docx/` as
    `word/document.xml` parts, assembled `.docx` files, and expected-text sidecars.
- `tests/roundtrip_fuzz_tests.mjs` (part of `npm test`)
  - Seeded fuzz sweep of the accept/reject round-trip invariant; tune with
    `FUZZ_SEED` / `FUZZ_ITERATIONS`.
- `npm run smoke:word -- path/to/file.docx`
  - Optional Windows/Word COM smoke test for a completed `.docx`.
- `npm run smoke:word:diff`
  - Windows/Word COM differential test: Word itself accepts/rejects the
    generated fixtures and the resulting text is compared to expectations.
- `bash scripts/validate-fixtures-xsd.sh`
  - Validates exported fixtures against the ECMA-376 transitional `wml.xsd`.
- `.github/workflows/validation.yml`
  - Nightly independent-oracle validation: XSD schema check, LibreOffice
    conversion, and an extended 20k-case fuzz sweep with a fresh seed.

Use these checks before publishing or tagging. See [docs/TESTING.md](./docs/TESTING.md#release-validation-and-independent-oracles).

## Fast Orientation For Contributors

Use this sequence to understand or modify behavior without reading everything:

1. Start at `index.js` to locate the exported API.
2. Follow exports into `engine/oxml-engine.js` or relevant `services/*` module.
3. For targeting bugs, inspect `core/paragraph-targeting.js`, `core/list-targeting.js`, and `core/table-targeting.js`.
4. For package wiring issues, inspect `services/standalone-docx-plumbing.js`.
5. For revision/comment cleanup behavior, inspect `services/revision-comment-management.js`.
6. For numbering/list issues, inspect `services/numbering-helpers.js` and orchestration list-fallback modules.
7. For reliability regressions, start with `tests/roundtrip_invariant_tests.mjs`,
   `tests/engine_reliability_tests.mjs`, `tests/paragraph_mark_revision_tests.mjs`,
   `tests/move_revision_tests.mjs`, and `tests/hardening_status_tests.mjs`.
