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
- **Permissive target resolution deprecation (`AMBIGUOUS_TARGET_HEURISTIC_USED`):**
  Lower-level document operation runners currently default to permissive targeting.
  In v1.0.0, application will default to `strictTargets: true`. In this release,
  when permissive resolution encounters multiple identical candidate paragraphs and
  heuristically chooses the first candidate, it emits an `AMBIGUOUS_TARGET_HEURISTIC_USED`
  warning with candidate count and migration guidance.
  **Migration:** Disambiguate operations using `paragraphId`, `index`, `occurrence`,
  or `fingerprint` and pass `{ strictTargets: true }`.

### ℹ️ Noteworthy Default & Usability Changes (Non-breaking)

The following improvements update default behaviors to maximize agent speed, reduce friction, and eliminate unnecessary errors during iterative editing. These changes are **not technically breaking changes** (no public APIs or exports were removed, and all existing options remain configurable), but represent important behavioral refinements that callers should note:

- **Fallback Author Default (`'AI Redliner'`):** `setDefaultAuthor`, `node/cli.js`, and `openDocx` now default to `'AI Redliner'` (configurable via the `DOCX_REDLINE_AUTHOR` environment variable or `--author`). The CLI `apply` command and core APIs no longer error with `AUTHOR_REQUIRED` when operations omit an author.
- **Output Overwrite by Default:** Destination files provided via `--output` now overwrite by default instead of failing with `OUTPUT_EXISTS`. Callers wanting protection against accidental overwrites can pass `--no-overwrite` or `--no-clobber`. The source document remains protected and is never overwritten unless `--in-place` is explicitly passed.
- **Replacement Event Pairing (`pairReplacements: true` by default):** Adjacent `<w:del>` and `<w:ins>` revisions now default to sharing linked revision metadata and identical timestamps so Microsoft Word groups them as a single replacement in the Reviewing Pane. Pass `pairReplacements: false` for independent revision timestamps.
- **Structured Content Auto-Detection (`structuredContent: true` by default):** Markdown tables, headings (`#`), and lists in replacement text automatically render as native Word elements (`w:tbl`, `w:pStyle`, `w:numPr`). Single outline-numbered legal clauses (e.g. `13.2.1.1`) continue to be treated as ordinary paragraphs without spurious list conversion. Pass `structuredContent: false` to treat replacement text strictly as plain text.
- **Progressive Batch Mode (`atomic: false` default, `atomic: true` opt-in):** Batch operations and CLI `apply` now apply edits progressively by default (`atomic: false`): valid edits commit directly, while failing edits report structured errors in `results` for faster debugging. Pass `{ atomic: true }` (or `--atomic` on the CLI) for all-or-nothing rollback when any operation fails.
- **Clean Direct Edits (`generateRedlines: false` / `--no-redlines`):** Prominently documented that tracked redlines are not always the preferred method. Callers can pass `generateRedlines: false` (or `--no-redlines` on the CLI) for clean execution drafts, document restructuring, or minor typo fixes without tracked changes markup.
- **Existing Revisions Policy (`existingRevisions: 'merge-same-author'` by default):** When editing a paragraph that already contains tracked changes from the same author, prior revisions by that author are merged against the pre-revision baseline (prior changes by that author are reverted and re-diffed to the new modified text), avoiding revision accumulation and nested markup. If a paragraph contains tracked changes from another reviewer, the operation fails closed with `EXISTING_REVISIONS` to protect third-party review marks. Commented revision content fails closed with `COMMENTED_CONTENT_MERGE` so comment anchors cannot be removed or orphaned. Callers can explicitly pass `'accept-all-first'` to normalize prior revisions or `'reject-input'` to refuse editing any revised paragraph.
- **Inline One-Liner CLI Edits:** Added `--target <text>` with `--modified <text>` or `--comment <text>` on `docx-redline apply` for fast 1–2 edit workflows without creating a JSON operations file.

### Added

- Added **Mutation Receipts** (`receipt` on `DocumentOperationResult` and `receipts`
  on `BatchOperationResult`). Every operation execution returns structured,
  commit-aware telemetry capturing durable revision IDs (with kind and target part),
  comment IDs, numbering IDs, relationship IDs, affected targets, and warnings.
  Supports `not_attempted`, `applied`, `refused`, `no_change`, and `rolled_back`
  dispositions.
- Added **Commit-Aware Output Reconciliation Oracle** (`reconcileReceiptsAgainstOutput`).
  Before completing any transaction, reported durable IDs are checked against a
  fresh parse of the generated output OOXML parts. Any reconciliation discrepancy
  immediately fails the transaction and triggers atomic rollback.
- Exported receipt primitives: `ReceiptCollector`, `createEmptyReceipt`, and
  `reconcileReceiptsAgainstOutput` along with `MutationReceipt` and
  `MutationReceiptRevisionItem` TypeScript declarations.
- Added **Paragraph Mark Revisions** (`w:pPrChange`, `markParagraphMarkInserted`,
  `markParagraphMarkDeleted`) ensuring whole-paragraph deletions and insertions
  cleanly track paragraph mark lifecycles for full Accept All / Reject All symmetry.
- Added **Paragraph Boundary Mutation Validation** (`validateParagraphBoundaryMutation`)
  guarding against invalid cross-paragraph boundary merges.

- Added `analyzeStructuredContent(...)` and `planStructuredReplacement(...)`
  for agent-authored attachments and schedules containing mixed headings,
  paragraphs, lists, and tables. The planner returns typed blocks and produces
  one atomic `structuredContent` operation only when the Markdown is valid.

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

- Explicit structured replacements now reject missing Markdown table separator
- Generated Markdown tables repeat their header row and prevent logical data rows from splitting across pages.
  rows, missing data rows, and inconsistent column counts with
  `STRUCTURED_CONTENT_INVALID` instead of inserting literal pipe-delimited text.
- Tables created inside mixed replacements are tracked once at block scope,
  avoiding nested revisions while keeping Accept All and Reject All symmetric.

- Target preflight and live mutation reuse one document-scoped paragraph index;
  strict duplicate detection and resolution metadata retain their prior
  semantics while avoiding repeated whole-document text traversal.
- Revision-ID seeding now walks the DOM with child/sibling pointers instead of
  allocating an array containing every element.
- Single-source-paragraph list expansion uses the focused list generator
  directly. Multi-paragraph marked-list edits deliberately retain the legacy
  run-aware reconciliation pipeline for structural parity.
- Fixed paragraph-to-list expansion when a manually numbered heading carries
  `w:numId w:val="0"`. Numbering suppression is no longer reused as a generated
  list ID, so inserted bullet and numbered items receive valid positive IDs.
- Tracked heading-to-list replacements now keep the deleted source heading in
  its own paragraph and track the paragraph marks of every inserted list item.
  This prevents deleted heading text from joining the first item and preserves
  exact Accept/Reject structure without adding a full-document sentinel
  paragraph.
- Inserted list items now inherit source font family, size, language, and
  related script typography without inheriting heading-only bold, underline,
  italic, or strike formatting unless requested by the replacement markdown.
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

- Added focused and semantic visual-failure regressions for suppressed
  numbering, paragraph separation, paragraph-mark revisions, selective
  typography inheritance, and exact accepted/rejected paragraph sequences,
  plus seeded fuzz variants across list length, marker, font, and size.
  Added the independently specified
  `legal-suppressed-heading-to-bullet-list` and
  `legal-structured-attachment-mixed-blocks` Microsoft Word differential
  fixtures; the synthetic catalogue now contains 49 cases and 49 task types.

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
