# Agent Fast Start

Use this page for ordinary `.docx` edits. For source changes or unusual Word structures, follow the links at the end.

## Structured agent tool

If your host provides a document-session wrapper:

1. Inspect/search once with enough surrounding context to draft the change.
2. Apply using the returned revision-bound target handle. Send either the
   complete desired paragraph or exact replacements supported by that wrapper.
3. Treat only `ok: true` as complete. A failed handle must be refreshed by
   inspection; never reuse it against a changed document.

The repository's `examples/agent-session-wrapper.mjs` demonstrates this pattern.
It is a development sample, not a package API.

## CLI fallback

1. Extract only the relevant clause or range:

   ```bash
   docx-redline extract contract.docx --search "termination" --around 3
   ```

   Search is case-insensitive. Direct hits are capped; follow `selection.nextAfter` with `--after`, and cite `humanReference`, not `P42`.

2. Copy `exactText` with `paragraphId` or `fingerprint`. Use a UTF-8 operations
   file or serializer-backed stdin, then apply once:

   ```bash
   node emit-operations.mjs | docx-redline apply contract.docx --operations - --profile agent --compact --output reviewed.docx
   ```

   `emit-operations.mjs` should use `JSON.stringify`; do not interpolate legal
   text through shell quoting. `modified` is the complete desired accepted-view
   paragraph, not only the inserted words.

The `agent` profile preserves progressive execution and the ordinary revision policy while making incomplete work exit nonzero. Add `--atomic` deliberately; check `effectiveOptions`. Run `docx-redline apply --help` for redline, comment, and rejected-view restore shapes.

## Batch and result rules

- Strong inspected targets are bound to the batch-start document. Independent
  edits do not need bottom-up sorting around earlier paragraph splits.
- Consolidate multiple desired texts for the same source paragraph. Follow an
  explicit capture dependency for non-unique content created earlier in a batch.
- Require `completion: true`, `written: true`, a non-null `outputPath`, and no
  per-operation error. The source is not overwritten unless `--in-place` is
  explicit.
- On failure, follow `error.recovery.action`. Never retry unchanged arguments.
- `retryPlan.base: "original"` means correct and replay the complete batch.
  `base: "output"` means keep committed progressive work and retry only the
  reported failed or unattempted indexes.
- Never accept/reject revisions or remove comments without user authorization.
  `slice-cross-author` is the history-preserving option for a surgical edit
  inside another reviewer's pending insertion.

Advanced operations and recovery: [Agent Knowledge Base](AGENT_KNOWLEDGE_BASE.md); canonical contract: [document-operations.schema.json](schemas/document-operations.schema.json)  
Wrapper design: [README](../README.md#example-agent-session-wrapper-development-only)
