# Changelog

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
