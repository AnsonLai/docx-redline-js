# AGENTS.md — Fast Start

Use this file as the launch card for `@ansonlai/docx-redline-js`. Do not explore
the whole repository before acting. Open [the full agent knowledge base](docs/AGENT_KNOWLEDGE_BASE.md)
only for the specific topic you need.

## Pick the Route

| Task | Start here |
|---|---|
| Edit or review a complete `.docx` | `docx-redline` CLI; follow the five-step workflow below |
| Change paragraph/range reconciliation | `index.js` → `engine/oxml-engine.js` → selected `engine/*-mode.js` |
| Change complete-document operations | `services/standalone-operation-runner.js` → `services/document-operation-*.js` |
| Change DOCX ZIP handling or CLI behavior | `node/index.js`, `node/docx-document.js`, `node/cli.js` |
| Choose or add tests | Closest `tests/*.mjs`, then [docs/TESTING.md](docs/TESTING.md) |
| Understand ownership or dependency flow | [ARCHITECTURE.md](ARCHITECTURE.md) |

Never inspect `dist/`, a vendored CLI bundle, or an installed plugin bundle to
learn public behavior. Use this file, the operation schema, and unbundled source.

## Fast DOCX Workflow

For ordinary document editing, use this route:

1. Run one focused `extract` for the relevant clauses or range.
2. Copy `exactText` plus `paragraphId` or `fingerprint` into the final operations.
3. Apply once per stable batch. Split batches only when a later edit targets text
   created by an earlier edit.
4. Check every result. Require `completion: true`, `written: true`, a non-null
   `outputPath`, and no `results[i].status === "error"`.
5. Re-extract only changed clauses when list/table placement or exact text needs
   confirmation.

```bash
docx-redline extract contract.docx --range 10:30
docx-redline apply contract.docx --operations operations.json --output reviewed.docx
```

`apply` already performs package and revision validation. Do not add `preflight`
or baseline `validate` unless the user explicitly asks for staged verification.
Do not probe behavior with disposable apply commands before building the real
operation. On failure, correct the reported cause; never repeat the same command.

Batches are progressive by default (`atomic: false`). Pass `--atomic` only when
all-or-nothing rollback is desired. The source is never overwritten unless
`--in-place` is explicit.

Create operation files with a structured file-writing tool or JSON serializer.
Do not use shell heredocs; legal text commonly contains Unicode, backslashes,
and exact whitespace that heredocs can alter.

## Operation TL;DR

For every ordinary text-bearing operation, `modified` is the complete desired
accepted-view content of the target—not merely the fragment being inserted.

| Desired result | Use |
|---|---|
| Edit text in a paragraph | `{ "type": "redline", "target": ..., "modified": "complete desired text" }` |
| Delete a whole paragraph | `{ "type": "delete", "target": ... }` |
| Change a native list | `{ "type": "list-change", "target": ..., "modified": "complete Markdown list" }` |
| Reconcile a table | `{ "type": "table-reconciliation", "target": ..., "modified": "complete Markdown table" }` |
| Add a comment | `comment` with `commentContent`; omit `textToComment` to comment the whole paragraph |
| Highlight text | `highlight` with `textToHighlight` |
| Format text | `character-format`/`format` with `textToFormat` and `properties` |
| Counterpropose a wholly foreign-deleted paragraph | `restore` with a rejected-view target |
| Insert inside rejected-view text | `insert` with rejected-view target, exact `anchor`, and `slice-cross-author` |

Important: ordinary `insert` is a compatibility alias of the redline path. It
does not create a new sibling paragraph unless it uses the special rejected-view
anchor form. `replace`, `list-change`, and `table-reconciliation` also normalize
through the redline operation path while adding intent for routing.

Canonical contract: [document-operations.schema.json](docs/schemas/document-operations.schema.json).

### Append a Native List Item

Use a `list-change` and provide the complete affected list as Markdown:

```json
{
  "type": "list-change",
  "target": {
    "exactText": "Review the report.",
    "paragraphId": "1A2B3C4D"
  },
  "modified": "1. Review the report.\n2. Record the approval decision."
}
```

Markdown markers describe structure. Computed Word labels such as `m)` are for
display only and must not appear in target text. For one adjacent item, a
one-line `redline` may use the exact current item followed by at least six words
of unnumbered new-item text. Use `list-change` for multiple items or nesting.

## Targeting and Failure Rules

- Copy extracted `exactText` verbatim. Do not paraphrase a failed target.
- Use `paragraphId`, `fingerprint`, `occurrence`, or `index` to disambiguate.
- Machine handles such as `P42` and bare paragraph indexes are not human Word
  references. In user-facing prose, use `humanReference`, `provision`, or heading
  context from inspection output.
- `EXISTING_REVISIONS`: use `slice-cross-author` only when editing inside another
  reviewer's pending insertion is intended. Never accept revisions implicitly.
- `PATCH_ROUNDTRIP_MISMATCH`: re-extract exact text and narrow the operation.
- `COMMENTED_CONTENT_MERGE` / `COMMENTED_CONTENT_DELETE`: report the comment and
  resolve it; do not silently remove reviewer content.
- `TARGET_NOT_FOUND`, `AMBIGUOUS_TARGET`, or anchor errors: re-extract and add the
  required exact discriminator.
- Always inspect `status` and `error`, not only `hasChanges` or `written`.

## Code Map

```text
index.js       host-independent public API
adapters/      injected XML/config/logging
core/          OOXML primitives, text views, targeting, validation
pipeline/      ingestion, diffing, Markdown, lists, serialization
engine/        reconciliation modes and run-level mutation
orchestration/ route planning and structural conversion
services/      document operations, comments, receipts, package artifacts
node/          Node-only ZIP and whole-DOCX facade
bin/           CLI launcher
tests/*.mjs    directly runnable suites
scripts/       build, fixture, benchmark, and Word automation
dist/          generated; never edit by hand
```

Keep shared modules from importing `index.js`; keep host-independent code from
importing `node/`. Start at the public entry point, follow only the symbol being
changed, and use `rg` instead of listing or reading whole directories.

## Focused Verification

Run the closest test first:

```bash
node tests/<focused-suite>.mjs
```

Use `npm test` when the change crosses subsystems or for release-level handoff.
Word COM, visual, corpus, coverage, and fixture-export commands are separate
lanes; use [docs/TESTING.md](docs/TESTING.md) to select them. Do not generate new
fixtures when an inline synthetic OOXML case can reproduce the behavior.

## More Detail

- [Agent knowledge base](docs/AGENT_KNOWLEDGE_BASE.md): full CLI workflow,
  operation examples, recovery matrix, APIs, options, gotchas, and validation.
- [Architecture](ARCHITECTURE.md): module ownership and dependency contracts.
- [Testing guide](docs/TESTING.md): test lanes and definitions of done.
- [README](README.md): consumer-facing install and API reference.
