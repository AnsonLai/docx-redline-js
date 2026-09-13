# AGENTS.md — Fast Start

Use this file as the launch card for `@ansonlai/docx-redline-js`. Do not explore
the whole repository before acting. Open [the full agent knowledge base](docs/AGENT_KNOWLEDGE_BASE.md)
only for the specific topic you need.

## Pick the Route

| Task | Start here |
|---|---|
| Edit or review a complete `.docx` | `docx-redline` CLI; follow the five-step workflow below |
| Build a thin agent/tool wrapper | Use the wrapper blueprint below; normally wrap `@ansonlai/docx-redline-js/node` or the CLI |
| Change paragraph/range reconciliation | `index.js` → `engine/oxml-engine.js` → selected `engine/*-mode.js` |
| Change complete-document operations | `services/standalone-operation-runner.js` → `services/document-operation-*.js` |
| Change DOCX ZIP handling or CLI behavior | `node/index.js`, `node/docx-document.js`, `node/cli.js` |
| Choose or add tests | Closest `tests/*.mjs`, then [docs/TESTING.md](docs/TESTING.md) |
| Understand ownership or dependency flow | [ARCHITECTURE.md](ARCHITECTURE.md) |

Never inspect `dist/`, a vendored CLI bundle, or an installed plugin bundle to
learn public behavior. Use this file, the operation schema, and unbundled source.

## Build a Thin Wrapper

Keep wrappers thin: select an integration surface, translate ergonomic tool
arguments into the library's operation schema, call the library once, and return
its structured result. Do not reimplement DOCX ZIP handling, paragraph targeting,
revision allocation, comment relationships, numbering merges, or validation.

### Choose the Integration Surface

| Wrapper environment | Wrap | Why |
|---|---|---|
| Shell-based skill or coding-agent command | `docx-redline` | Complete file I/O, compact JSON stdout, exit codes, safe output paths |
| Node tool receiving/returning DOCX bytes | `openDocx` from `@ansonlai/docx-redline-js/node` | Complete package inspection, mutation, validation, rollback, and output buffer |
| Host that already owns `word/document.xml` and related parts | `@ansonlai/docx-redline-js/standalone-runner` | Document-XML operations without ZIP or filesystem policy |
| Browser/host editing paragraph OOXML | Root `@ansonlai/docx-redline-js` exports | Lowest-level host-independent paragraph/range transforms |

For most new wrappers, prefer the Node facade. Use lower layers only when the
host already owns the corresponding package boundary.

If a skill vendors the CLI, run its `version` command before the first document
operation. Declare the minimum `contractVersion` and capabilities the wrapper
actually depends on, and fail closed with an upgrade instruction when they are
missing. Do not compensate for a stale bundle by inspecting minified source or
editing ZIP/XML parts directly.

### Recommended Agent Tool Surface

A general-purpose wrapper usually needs three tools:

| Tool | Side effect | Underlying call | Return |
|---|---|---|---|
| `inspect_docx` | None | `openDocx(bytes).inspect(filters)` or CLI `extract`/`inspect` | Exact targets, IDs, structure, revisions, comments, and a document-version token |
| `apply_docx_operations` | Produces new bytes/file; never overwrite source implicitly | `document.applyOperations(operations, options)` or CLI `apply` | Status, every operation result, receipts, warnings, validation, output |
| `resolve_docx_review` | Produces new bytes/file | `resolveRevisions` / `deleteComments`, or CLI `accept`/`reject`/`delete-comments` | Counts, status, output |

Add narrower convenience tools only for frequent workflows, such as
`comment_docx`, `redline_clause`, or `restore_deleted_paragraph`. They should
construct canonical operations and delegate to the same apply path. A restore
tool should inspect the rejected view internally and emit an explicit
`revisionView: "rejected"` target. Avoid a generic shell tool that forces the
model to learn command syntax, and avoid exposing internal engine functions as a
large tool catalog.

Keep wrapper policy distinct from library behavior. A skill may deliberately
choose its own author, revision policy, atomicity, or output-naming defaults, but
its tool description should call those wrapper defaults out and pass them
explicitly rather than presenting them as the library's defaults.

Each wrapper tool description should state:

- when to use it and when another wrapper tool is appropriate;
- required inputs and whether text must come from inspection;
- whether it is read-only or produces a new document;
- that ordinary `modified` text is the complete desired target content;
- success fields and common error codes;
- whether retrying unchanged arguments is safe (an unchanged failed operation
  must not be retried).

### Minimal Node Wrapper

```js
import { openDocx } from '@ansonlai/docx-redline-js/node';

export function inspectDocx(inputBytes, filters = {}) {
  const document = openDocx(inputBytes);
  return {
    inspection: document.inspect(filters),
    packageRevision: document.getRevisionToken()
  };
}

export async function applyDocxOperations(inputBytes, operations, options = {}) {
  const document = openDocx(inputBytes);
  const result = await document.applyOperations(operations, {
    author: options.author || 'AI Redliner',
    atomic: options.atomic === true,
    strictTargets: true,
    validate: true,
    ...(options.expectedRevision
      ? { expectedRevision: options.expectedRevision }
      : {}),
    ...(options.existingRevisions
      ? { existingRevisions: options.existingRevisions }
      : {})
  });

  const failed = (result.results || []).filter(item => item.status === 'error');
  const ok = result.status !== 'error' && result.status !== 'partial' && failed.length === 0;

  return {
    ok,
    changed: result.hasChanges === true,
    outputBytes: ok ? result.toBuffer() : null,
    status: result.status,
    error: result.error || null,
    results: result.results || [],
    receipts: result.receipts || [],
    warnings: result.warnings || [],
    validation: result.validation || null
  };
}
```

The facade's `written` means the in-memory package changed; it does not mean a
wrapper wrote a filesystem destination. A file-based wrapper must perform its
own destination write after `ok`, or delegate file I/O to the CLI. On an atomic
or top-level failure, `toBuffer()` returns the rolled-back/original package, but
a wrapper should expose `outputBytes: null` so callers cannot mistake failed
work for completed output.

Do not collapse the response to a boolean. Preserve `status`, top-level `error`,
all per-operation `results`, `receipts`, warnings, and validation diagnostics.
If a wrapper intentionally supports progressive partial output, expose it under
an explicitly partial field and keep `ok: false`; never present a partially
applied document as the completed result.

Target descriptors are scoped to the exact package version that was inspected.
Return the package-scoped token from `document.getRevisionToken()` and require it
as `expectedRevision` when applying a planned operation. The token embedded in
`document.inspect()` is scoped to document parts and is not interchangeable with
the Node facade's package token. A shell wrapper that cannot bind a token must at
least extract and apply against the same unchanged path, and re-extract whenever
it switches to a derived working copy.

The operation input should use the canonical
[document-operations schema](docs/schemas/document-operations.schema.json).

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

### Restore a Wholly Deleted Paragraph

A paragraph wholly deleted by another reviewer is empty in the accepted/current
view. Its source text is discoverable in the rejected view:

```bash
docx-redline extract working.docx --range 50:60 --view rejected
```

Build the restore target from that result:

```json
{
  "type": "restore",
  "target": {
    "exactText": "The deleted source paragraph.",
    "paragraphId": "1A2B3C4D",
    "fingerprint": "fnv1a32:12345678",
    "revisionView": "rejected"
  },
  "modified": "The restored and revised paragraph.",
  "author": "Jane Doe"
}
```

`restore` preserves the foreign deletion and inserts the counterproposal after
it as tracked content. It always generates tracked changes. Do not accept the
deletion, use an ordinary `redline`, or use rejected-view `insert` for a whole
paragraph restoration.

Never reuse a restore descriptor extracted from another input or an earlier
version of the working file. Re-extract from the exact file passed to `apply`.
For a contiguous deleted range, add `targetEnd` and supply exactly one
replacement string per source paragraph in `modified`.

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
- `TARGET_TEXT_MISMATCH` on `restore`: confirm the target explicitly uses the
  rejected view and re-extract it from the exact DOCX version being applied.
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
