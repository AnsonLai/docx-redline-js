# Changelog

## Unreleased

### ⚠️ Breaking changes

The following changes can require caller or contributor updates. No public
export or valid existing function signature was removed.

#### Runtime and result-contract changes

- **Canonical paragraph text:** `getParagraphText(...)` now returns the
  accepted/current view instead of the earlier simplified `w:t`/`w:tab`
  concatenation. It excludes deleted and `w:moveFrom` content and includes
  structural breaks, soft hyphens, and non-breaking hyphens.
  **Migration:** regenerate compared/cached text with the upgraded library. Use
  `extractCanonicalParagraphText(...)` when accepted/rejected-view semantics
  are intended; integrations that require the former raw traversal must retain
  their own legacy extractor before upgrading.
- **Paragraph fingerprints:** fingerprints now incorporate canonical paragraph
  text. Stored fingerprints can become stale for paragraphs containing
  revisions, moves, breaks, soft hyphens, or non-breaking hyphens.
  **Migration:** do not carry fingerprints across this upgrade; re-extract the
  document and regenerate target descriptors.
- **Invalid operations:** malformed or field-incompatible document-operation
  objects now return a structured `INVALID_OPERATION` error at the runner
  boundary instead of reaching later failure/no-op paths.
  **Migration:** validate operation files against the published schema and
  handle `status === 'error'` plus `error.code === 'INVALID_OPERATION'`.
- **Commented paragraph deletion:** deleting a complete paragraph that contains
  an existing comment now fails with `COMMENTED_CONTENT_DELETE` instead of
  producing OOXML with dangling comment references.
  **Migration:** preserve or deliberately remove/relocate the comment before
  deleting the paragraph, and handle the structured error. Atomic callers keep
  the original document unchanged.
- **Comment anchor failures:** missing or ambiguous comment anchors now return
  `ANCHOR_NOT_FOUND` or `AMBIGUOUS_ANCHOR` and fail/roll back an atomic batch;
  they are no longer reported as successful `no_change` operations.
  **Migration:** treat these codes as failed operations and use a unique exact
  anchor or stricter target descriptor before retrying.
- **Stricter package validation:** DOCX package validation now rejects
  unbalanced comment ranges, dangling references/usages, orphan definitions,
  and duplicate comment-definition IDs that earlier validation could allow.
  **Migration:** repair the comment anchors and `word/comments.xml` definitions
  before applying or validating further changes.

#### Development workflow change

- **Parallel tests by default:** `npm test` now runs up to four isolated child
  processes concurrently. This does not change the published runtime API, but
  contributor automation that depends on test execution order or shared files
  must be updated. **Migration:** make tests independent or set
  `DOCX_TEST_CONCURRENCY=1` for the previous serial scheduling behavior.

Strict targeting remains opt-in on existing low-level mutation APIs; it is the
default only on the newly added Node facade and CLI.

### Deprecations and future breaking changes

- Replacing an existing numbering part without a merge callback now emits a
  deprecation warning. The Node facade and CLI already merge safely by default.
  **This remains supported in this release but will become an error in the next
  major version.** Migrate low-level callers to
  `mergeNumberingXmlBySchemaOrder` before that release.

### Added

- Added session-scoped paragraph metadata indexing for canonical text,
  normalized text, paragraph IDs, fingerprints, table context, and document
  indexes, plus deterministic targeting and route-profiling benchmarks.
- Added a shared list-marker grammar and parsed list-item vocabulary used by
  pipeline analysis, orchestration, normalization, and list targeting.
- Added an internal reconciliation capability matrix and opt-in route
  instrumentation. Public reconciliation result shapes and route selection
  policy remain unchanged.

- Added `scripts/extract_text.mjs` and `scripts/apply_changes.mjs` as thin
  compatibility entrypoints for legacy skill installations. They delegate to
  the supported CLI, retain strict atomic validated writes, and accept legacy
  operation files with a top-level `changes` array.
- Added typed operations, per-operation authors, strict target descriptors,
  deterministic preflight diagnostics, and auditable resolution metadata.
- Added `inspectDocumentParts(...)`, a structured document inventory using the
  same canonical text extractor as targeting and ingestion.
- Added `@ansonlai/docx-redline-js/node`. `openDocx(...)` applies batches to
  complete DOCX buffers transactionally and validates before commit.
- Added the cross-platform `docx-redline` CLI with JSON `inspect`, `extract`,
  `preflight`, `apply`, `accept`, `reject`, `delete-comments`, and `validate`
  commands, plus a published operation-file JSON Schema.

### Behavior and reliability

- Target preflight and live mutation reuse one document-scoped paragraph index;
  strict duplicate detection and resolution metadata retain their prior
  semantics while avoiding repeated whole-document text traversal.
- Revision-ID seeding now walks the DOM with child/sibling pointers instead of
  allocating an array containing every element.
- Single-source-paragraph list expansion uses the focused list generator
  directly. Multi-paragraph marked-list edits deliberately retain the legacy
  run-aware reconciliation pipeline for structural parity.
- Canonical-only table-cell targeting now consumes the shared paragraph-text
  extractor, while offset- and sentinel-producing walkers remain specialized.
  Their accepted visible projections now agree on soft hyphens as well as tabs,
  breaks, hyperlinks, and non-breaking hyphens.

- Comment preflight and application now share one anchor resolver. Exact
  matches take priority, unique ordinary-space/NBSP differences preserve raw
  offsets, and missing or repeated anchors return structured
  `ANCHOR_NOT_FOUND` or `AMBIGUOUS_ANCHOR` errors.
- Full-document operation batches now share one live DOM and revision-ID
  allocator, reducing complete-document parsing and serialization from once per
  operation to once per changed batch. Scoped reconciliation engines and their
  targeting decisions are unchanged.
- Each live-session operation uses a DOM and allocator savepoint. Errors and
  reported no-ops restore the savepoint; atomic failures and all-no-op batches
  still return the exact original XML without serializing it.
- Decomposed the standalone document-operation runner into a stable 13-line
  compatibility facade plus focused session, applier, batch-orchestrator,
  heuristic, and OOXML-mutation modules. Runtime exports and declarations are
  unchanged, and operation internals now use leaf imports instead of importing
  the root entry point.
- Package comment IDs are seeded from existing anchors and definitions, and
  the Node facade merges numbering without discarding prior definitions.
- Canonical paragraph text consistently handles revisions, moves, tabs,
  breaks, soft hyphens, and non-breaking hyphens; see the breaking-change note
  above for `getParagraphText(...)` and fingerprint compatibility.
- Structured inspection now explicitly supports the `current` view as the
  accepted/current document view and resolves comment ranges spanning multiple
  paragraphs as one exact anchor string.
- Transactional no-op results preserve the original DOCX bytes exactly,
  including when the same facade instance has already committed an edit.

### Testing and development

- The JavaScript test runner now uses bounded asynchronous subprocess workers,
  capped at four by default, while retaining one fresh Node process per file,
  sorted reporting, timeouts, captured diagnostics, and failure-marker checks.
  Set `DOCX_TEST_CONCURRENCY=1` for serial reproduction.
- Added `npm run benchmark:tests` and a Phase 6 runner regression. The checked
  benchmark reduced median suite time from 18.02 seconds to 7.96 seconds
  (55.80%); 20 consecutive parallel runs and a separate serial run passed all
  63 files. Parallel c8 collection remains valid at 89.90% statements/lines,
  77.40% branches, and 93.10% functions.

- Added Phase 3 targeting/traversal, Phase 4 list/text parity, and Phase 5 route
  compatibility regressions. Together with the Phase 6 runner regression, the
  suite now contains 63 passing test files.
- Added `npm run benchmark:targeting` and `npm run profile:routes`. The checked
  10,000-paragraph benchmark showed no single-operation regression and a 47.16x
  Node median improvement for 100 cached resolutions. The equivalent native
  Chrome benchmark showed no single-operation regression and a 36.45x batch
  improvement. All 47 Microsoft Word differential fixtures and all 60 real-
  document Word corpus scenarios passed after the routing changes.

- Added subprocess coverage for legacy wrapper selection, operation-file
  compatibility, successful output, atomic anchor failure, source-byte
  preservation, and protection of pre-existing output files.
- Added a live-session accuracy and instrumentation suite covering accepted and
  rejected text, structural validity, sequential equivalence, comments, lists,
  tables, highlights, one-parse/one-serialize execution, exact no-ops, and
  atomic rollback. Added `npm run benchmark:session`; the checked accuracy-first
  benchmark reduced parse/serialize counts from 10/10 to 1/1 and measured a
  1.58x median speedup with per-operation savepoints retained.
- Added a Phase 2 architecture regression covering facade export identity,
  direct leaf imports, exact session rollback, isolated context commit, and
  stable comment-first scheduling.
- Added five edge-case suites for canonical text views, cross-paragraph comment
  inspection, table/list/reference context, sequential package transactions,
  selective multi-author cleanup, ZIP rejection behavior, CLI filters, JSON
  exit contracts, and destructive-output safeguards. Together with the Phase 2
  Phase 1 and Phase 2 regressions, these established the earlier 57-file
  baseline expanded by the subsequent performance phases.

## 0.3.0

### Breaking changes

- Malformed OOXML no longer escapes public transform APIs as a raw parser
  exception. In particular, revision accept/reject and comment-deletion callers
  that previously used `try`/`catch` must now inspect `status === 'error'` and
  `error.code === 'PARSE_ERROR'` on the returned result.
- Caller content is no longer sanitized by default. Pass `sanitizeInput: true`
  to remove a standalone leading assistant-preface line. Dollar-delimited text
  and literal `\\n` / `\\r\\n` sequences are never removed implicitly.
- The exported `sanitizeAiResponse` helper no longer removes dollar-delimited
  spans or converts literal `\\n` / `\\r\\n` sequences. Direct callers that
  relied on those transformations must perform them explicitly.
- `applyOperationsToDocumentXml` is atomic by default. Failed batches return
  the original document with `rolledBack: true`; pass `atomic: false` to retain
  partial-result behavior.
- Missing multi-line targets now return `TARGET_NOT_FOUND` instead of being
  indistinguishable from a no-op. Stale batch anchors also return
  `TARGET_NOT_FOUND` instead of falling back to a visibly different paragraph.

### Behavior changes and fixes

- Revision IDs are now allocated per document. Near-limit prior IDs restart in
  a safe low range, and unrelated bookmark/relationship IDs no longer seed the
  revision counter.
- Concurrent calls with explicit authors keep their revision attribution
  isolated from process-global defaults.
- `existingRevisions: 'accept-all-first'` now preserves the original OOXML when
  the requested edit is a no-op.
- Added `existingRevisions: 'accept-all-first-keep-normalized'` for callers that
  explicitly want accepted existing revisions returned even when no redline is
  added.
- Diff output is deterministic by default. Inputs exceeding 262,144 unique
  tokens return `DIFF_TOKEN_LIMIT` with the original OOXML unchanged instead of
  risking truncated text; leading whitespace is preserved exactly.
- Reconstruction preserves structural line breaks and body section-property
  placement. A target that names only part of a reconstruction range now
  returns `PARTIAL_TARGET` instead of risking deletion of untargeted content.
- Paragraph targeting and reconstruction now preserve leading, middle, and
  trailing `w:tab` elements as literal tab characters instead of omitting or
  trimming them.
- Adjacent redline edits now preserve complex-field instructions, begin/
  separate/end markers, and unchanged cached display results in their original
  field boundary. Field instructions remain inert and are not evaluated.
- Added structured XML parse results and result-returning ingestion helpers.
- Hardened deterministic diff token handling and leading-whitespace fidelity.

### Testing and development

- Expanded the desktop Word differential from 20 to 25 cases, adding structural
  checks for bookmarks, internal hyperlinks, mixed formatting, content
  controls, and tables.
- Expanded the desktop Word differential from 25 to 28 cases with structural
  tabs and a locked PAGE field, and added tab/field shapes to focused and fuzz
  regression testing.
- Expanded the desktop Word differential from 28 to 33 cases with comments,
  footnotes, endnotes, headers/footers, and an external hyperlink.
- Extended the deterministic script-only DOCX packager with opt-in related
  parts, relationship/content-type validation, reusable fixture constructors,
  and SHA-256 verification that supplied parts remain byte-identical. Runtime
  library code and dependencies are unchanged.
- Added validated task, structure, oracle, and manual-review metadata across the
  33 synthetic and 20 SuperDoc Word cases. `npm run report:word:coverage`
  produces a deterministic coverage matrix with explicit high-priority gap
  dispositions, while `npm run review:word:prepare` creates a pending rotating
  human-review manifest without self-certifying visual results.
- Added detailed production-function coverage reporting and a checked per-file
  regression baseline. Five behavior-focused Phase 3 suites exercise numbering
  collisions/remapping, every reconciliation route, list construction and
  fallback, patch/format boundaries, table decisions, pipeline modes, and
  standalone rollback/highlight paths. Production function coverage moved from
  437/540 to 496/542 and all reachable P0 functions are covered; no runtime
  behavior or compatibility contract changed.
