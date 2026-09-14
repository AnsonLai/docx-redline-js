# Agent Fast Start

Use this page for ordinary `.docx` edits. For source changes or unusual Word structures, follow the links at the end.

## Structured agent tool

If your host provides a document-session wrapper:

1. Inspect/search once with enough surrounding context to draft the change.
2. Apply with the revision-bound handle and complete desired text or exact replacements.
3. Treat only `ok: true` as complete; refresh failed handles before reuse.

The repository's `examples/agent-session-wrapper.mjs` demonstrates this pattern.
It is a development sample, not a package API.

## CLI fallback

1. For a literal mechanical edit with known old and new text, try one-turn apply:

   ```bash
   docx-redline apply contract.docx --find "thirty (30) days" --replace "sixty (60) days" --profile agent --output reviewed.docx
   ```

   The source must occur in one eligible paragraph. To narrow it, use a longer, fairly unique phrase: `--search "distinctive nearby heading or phrase" --context-range 1:3`.
   Generic/repeated anchors can cause `AMBIGUOUS_TARGET` and waste a recovery turn. If `anchorMatchCount > 1`, confirm the returned location and excerpts. `--around 3` is symmetric.

2. For semantic drafting such as “make this provision mutual,” extract enough
   context, then copy `exactText` with `paragraphId` or `fingerprint`:

   ```bash
   docx-redline extract contract.docx --search "termination" --around 3
   ```

   Search is case-insensitive; follow `selection.hint` and `selection.nextAfter`, and report `humanReference`. With `--revised`, use `hasRevisions`, `revisionAuthors`, and `deleted` or `inserted` to interpret paragraphs whose active-view `exactText` is empty. Apply once with a UTF-8 operations file or serializer-backed stdin:

   ```bash
   node emit-operations.mjs | docx-redline apply contract.docx --operations - --profile agent --compact --output reviewed.docx
   ```

   Use `JSON.stringify`, not shell interpolation. `modified` is the complete desired accepted-view paragraph.

The `agent` profile preserves progressive execution and the ordinary revision policy while making incomplete work exit nonzero. Add `--atomic` deliberately; check `effectiveOptions`. Run `docx-redline apply --help` for redline, comment, and rejected-view restore shapes.

## Batch and result rules

- Strong inspected targets bind to the batch start; independent edits need no bottom-up sorting.
- Consolidate writes to one source; use captures for created-content dependencies.
- Require `completion: true`, `written: true`, a non-null `outputPath`, and no
  per-operation error. For localized edits also require
  `results[i].change.committed: true`, `finalDisposition: "applied"`, and
  `verification.acceptedViewMatchesCompiledText: true`. The source is
  not overwritten unless `--in-place` is explicit.
- On failure, follow `error.recovery.action`. Never retry unchanged arguments.
- `retryPlan.base: "original"` means replay the batch; `"output"` means retain committed work and retry reported indexes.
- Never accept/reject revisions or remove comments without user authorization. `slice-cross-author` preserves history for edits inside foreign insertions.

Advanced operations and recovery: [Agent Knowledge Base](AGENT_KNOWLEDGE_BASE.md); canonical contract: [document-operations.schema.json](schemas/document-operations.schema.json)  
Wrapper design: [README](../README.md#example-agent-session-wrapper-development-only)
