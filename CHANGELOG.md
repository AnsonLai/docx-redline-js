# Changelog

## Unreleased

### Breaking changes

- `getParagraphText(...)` now returns canonical accepted-view text instead of
  the earlier simplified `w:t`/`w:tab` concatenation. For affected OOXML it now
  excludes deleted and `w:moveFrom` content and includes structural breaks,
  soft hyphens, and non-breaking hyphens. Callers that compare, cache, or store
  its output may observe different strings. Use the new
  `extractCanonicalParagraphText(...)` explicitly when accepted/rejected-view
  semantics are intended; callers that require the former raw traversal must
  preserve that behavior in their integration before upgrading.
- Paragraph fingerprints include canonical paragraph text. Previously stored
  fingerprints can therefore become stale for paragraphs containing revisions,
  moves, breaks, soft hyphens, or non-breaking hyphens. Regenerate fingerprints
  from the upgraded library rather than persisting them across the upgrade.
- Malformed or field-incompatible document-operation objects now return a
  structured `INVALID_OPERATION` error at the runner boundary. Integrations
  that treated invalid operations as later failures or no-ops must handle this
  explicit error result.

No public export or valid existing function signature was removed. Strict
targeting also remains opt-in on existing low-level mutation APIs; it is the
default only on the newly added Node facade and CLI.

### Added

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

- Comment preflight and application now share one anchor resolver. Exact
  matches take priority, unique ordinary-space/NBSP differences preserve raw
  offsets, and missing or repeated anchors return structured
  `ANCHOR_NOT_FOUND` or `AMBIGUOUS_ANCHOR` errors.
- Comment operations that place no comment can no longer be reported as
  successful `no_change` items. They fail the operation, roll back atomic
  batches, and allocate no comment ID when anchor resolution fails.
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
- Replacing an existing numbering part without a merge callback now emits a
  deprecation warning. The Node facade and CLI already merge safely by default;
  the low-level replacement behavior will become an error in the next major version.

### Testing and development

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
  Phase 1 and Phase 2 regressions, the automated suite now contains 57 passing
  test files.

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
