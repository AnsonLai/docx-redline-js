# Agent Fast Start

Use this page for ordinary `.docx` edits. For unusual Word structures or optional shortcuts, follow the links at the end.

## Structured agent tool

If the host provides a document-session wrapper (or directly calls `openDocx` from `@ansonlai/docx-redline-js` with `Uint8Array` binary payloads):

1. Inspect once with enough surrounding context to draft the change.
2. Apply once with the returned revision-bound handle.
3. Accept only the wrapper's complete-success result; refresh failed handles.

The universal `DocxDocument` and `openDocx` facade operates natively on `Uint8Array` across Node, browsers, edge workers, and sandboxes without requiring Node built-ins.
The repository's `examples/agent-session-wrapper.mjs` is a development sample, not a package API.

## CLI workflow

1. Extract the relevant clause or range:

   ```bash
   docx-redline extract contract.docx --search "termination" --around 3
   ```

   Search is case-insensitive. Follow `selection.nextAfter` when truncated. Preserve `selection`, `humanReference`, target descriptors, and revision cues when reshaping output. Report `humanReference`, not a machine paragraph number.

2. Retain inspected `exactText` for drafting and target with its `paragraphId` plus `fingerprint`, then apply once with a UTF-8 operations file or serializer-backed stdin:

   ```bash
   node emit-operations.mjs | docx-redline apply contract.docx --operations - --profile agent --compact --output reviewed.docx
   ```

   Use complete `modified` text for a broad semantic revision. For a small literal change in that inspected target, use
   `replacements: [{ "find": "old text", "replace": "new text" }]` instead.
   Use `JSON.stringify`; never interpolate document text through shell quoting. Preserve intentionally relative wording unless the user asks for a concrete value.

## Restore branch

If the user requests deleted text, search with `--view rejected` first. Otherwise follow `selection.hint` only when it identifies an alternate revision view. Restore the inspected strong target; do not infer adjacent content. Run `docx-redline apply --help` for canonical restore shapes.

## Completion and recovery

- Require `completion: true` and a non-null `outputPath`; the CLI computes this from write, operation, commit, and localized-verification results.
- On failure, follow `error.recovery.action`; never retry unchanged arguments.
- Use `retryPlan.base: "original"` to replay a batch and `"output"` to retain
  committed progressive work and retry the reported indexes.
- The source is not overwritten unless `--in-place` is explicit.
- Never accept/reject revisions or remove comments without authorization.
- Add `--atomic` or a non-default revision policy only when deliberately needed;
  confirm `effectiveOptions`.
- Strong inspected targets bind to batch start. Consolidate writes to one source
  and use captures only for intentional created-content dependencies.

Advanced speculative apply, directional scoping, operation constraints, and
recovery: [Agent Knowledge Base](AGENT_KNOWLEDGE_BASE.md). Canonical schema:
[document-operations.schema.json](schemas/document-operations.schema.json).
Wrapper design: [README](../README.md#example-agent-session-wrapper-development-only)
