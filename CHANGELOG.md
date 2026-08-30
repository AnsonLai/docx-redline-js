# Changelog

## 0.3.0

- Breaking: `applyOperationsToDocumentXml` is atomic by default. Failed batches
  return the original document; pass `atomic: false` to retain partial results.
- Stale batch anchors now return `TARGET_NOT_FOUND` instead of falling back to a
  visibly different paragraph.
- Breaking: caller content is no longer sanitized by default. Pass
  `sanitizeInput: true` to remove a standalone leading assistant-preface line.
- Removed unsafe dollar-delimiter removal and literal `\\n` / `\\r\\n`
  conversion from `sanitizeAiResponse`.
- `existingRevisions: 'accept-all-first'` now preserves the original OOXML when
  the requested edit is a no-op.
- Added `existingRevisions: 'accept-all-first-keep-normalized'` for callers that
  explicitly want accepted existing revisions returned even when no redline is
  added.
- Added structured XML parse results and result-returning ingestion helpers.
- Hardened deterministic diff token handling and leading-whitespace fidelity.
