# Agent Knowledge Base

> Detailed reference for @ansonlai/docx-redline-js. All code and command paths
> are relative to the repository root. Start with `../AGENTS.md` for repository
> work or `AGENT_FAST_START.md` for ordinary document edits; open this file only
> when the quick guide points here or the task needs deeper behavior.

## Start Here: Repository Layout

This is one JavaScript package with three supported public surfaces. Choose the
surface before following imports:

| Use case | Public entry point | Primary implementation |
|---|---|---|
| Paragraph, range, list, table, comment, and OOXML transforms in any DOM-capable runtime | `index.js` (`@ansonlai/docx-redline-js`) | `engine/`, `pipeline/`, `core/`, and focused `services/` |
| Operations against complete `word/document.xml` strings | `services/standalone-operation-runner.js` (`@ansonlai/docx-redline-js/standalone-runner`) | `services/document-operation-*.js` and `services/batch-operation-orchestrator.js` |
| Complete `.docx` buffers and the CLI in Node.js | `node/index.js` (`@ansonlai/docx-redline-js/node`) and `bin/docx-redline.js` | `node/docx-document.js`, `node/zip-archive.js`, and package-plumbing services |

The source tree is layered as follows:

```text
index.js                 root, host-independent public exports
adapters/                injected XML, configuration, and logging adapters
core/                    shared OOXML primitives, text views, targeting, validation
pipeline/                ingestion, diffing, markdown, lists, and serialization stages
engine/                  paragraph/range reconciliation and surgical/reconstruction modes
orchestration/           route planning and structural list operation conversion
services/                document operations, comments, receipts, package artifacts
node/                    Node-only ZIP and whole-DOCX facade; keep out of root imports
bin/                     CLI launcher; behavior belongs in node/ or services/
tests/*.mjs              directly runnable suites discovered by scripts/run-tests.mjs
tests/helpers/           shared test utilities; not standalone suites
tests/fixtures/          checked-in synthetic/golden inputs and expected outputs
scripts/                 build, fixture generation, benchmarks, and Word automation
docs/                    schemas, testing guidance, plans, and generated reports
dist/                    generated bundle; do not edit by hand
```

`AGENTS.md` is the committed repository-wide agent guide. `.agent/` is ignored
local-only agent configuration and must not be treated as package source or as
instructions that will exist in another clone or in CI.

### Route Changes Without Exploring Everything

1. Identify the public surface in the table above.
2. Read its entry point, then follow only the symbol being changed.
3. Start from the closest existing test named for the behavior. Do not scan all
   of `tests/` or unpack every fixture to understand one code path.
4. Use `rg` for symbols and filenames. Consult `ARCHITECTURE.md` for ownership
   and `docs/TESTING.md` for test-lane selection before inventing a new route,
   helper, test harness, or fixture generator.
5. Keep dependency direction inward: shared implementation modules must not
   import `index.js`, and host-independent code must not import `node/`.
6. Do not inspect `dist/`, a vendored CLI bundle, or an installed plugin bundle
   to discover public behavior. Use this file, `README.md`,
   `docs/schemas/document-operations.schema.json`, and the unbundled source.

### Task-to-Module Shortcuts

| Task | Start here | Common focused tests |
|---|---|---|
| Basic text redline, formatting, or route choice | `engine/oxml-engine.js`, then the selected `engine/*-mode.js` | `tests/engine_reliability_tests.mjs`, `tests/formatting_tests.mjs` |
| Run splitting or edits inside revisions | `engine/surgical-mode.js`, `engine/surgical-run-splitting.js`, `engine/surgical-diff-application.js` | `tests/revision_split_injection_tests.mjs`, `tests/cross_author_slicing_synthetic_tests.mjs` |
| Accepted/rejected/current text or paragraph targeting | `core/paragraph-text.js`, `core/paragraph-targeting.js` | `tests/canonical_paragraph_text_tests.mjs`, `tests/revision_view_target_tests.mjs` |
| Lists, numbering, or markdown structure | `pipeline/list-generation.js`, `pipeline/structured-content.js`, `services/numbering-helpers.js` | `tests/list_tests.mjs`, `tests/list_replacement_structure_tests.mjs`, `tests/structured_content_planner_tests.mjs` |
| Tables | `engine/table-mode.js`, `core/table-targeting.js`, `services/table-reconciliation.js` | `tests/table_tests.mjs`, `tests/table_targeting_and_format_flags.mjs` |
| Comments and replies | `services/comment-engine.js`, `services/comment-replies.js`, `services/comment-package.js` | `tests/comment_tests.mjs`, `tests/comment_reply_tests.mjs` |
| Full-document operation scheduling or rollback | `services/document-operation-applier.js`, `services/batch-operation-orchestrator.js`, `services/document-operation-session.js` | `tests/standalone_operation_runner_tests.mjs`, `tests/performance_phase1_session_tests.mjs` |
| DOCX ZIP wiring or CLI behavior | `node/docx-document.js`, `node/zip-archive.js`, `node/cli.js` | `tests/docx_package_facade_tests.mjs`, `tests/node_zip_archive_tests.mjs`, `tests/agent_cli_tests.mjs` |
| Validation, receipts, and lifecycle oracles | `core/redline-validation.js`, `services/receipt-collector.js`, `services/revision-comment-management.js` | `tests/redline_validation_tests.mjs`, `tests/mutation_receipt_tests.mjs`, `tests/roundtrip_oracle_tests.mjs` |

Run one focused suite with `node tests/<name>.mjs`. Run `npm test` only when the
change crosses several subsystems or before release-level handoff. The Word COM,
visual, corpus, coverage, and fixture-export commands are separate lanes; use
them only when `docs/TESTING.md` says that lane proves the behavior in question.

## What This Package Does

Converts text/markdown edits into valid Office Open XML (OOXML) with Word-native tracked changes. Feed it original OOXML + desired text and it returns OOXML with `w:ins`/`w:del` revision markup.

## Conceptual Model

```
Input: (paragraph OOXML, original text, modified text, options)
  |
  v
Engine routes to: format-only | surgical | reconstruction | list | table mode
  |
  v
Output: { oxml: string, hasChanges: boolean, status?: string, error?: object, warnings?: string[] }
```

The engine usually works at paragraph/range/table scope. For full-document
operations, use the standalone operation runner so the result is safe to write
back to `word/document.xml`.

## Entry Point

```js
import { applyRedlineToOxml, configureXmlProvider } from '@ansonlai/docx-redline-js';
```

`index.js` is the primary host-independent entry point. Complete document XML
and `.docx` package workflows use the standalone runner and Node facade listed
in the repository-layout table above.

## Required Setup (Node.js only)

```js
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
configureXmlProvider({ DOMParser, XMLSerializer });
```

Browsers have native DOM APIs, so no provider injection is typically needed.

## Key APIs by Use Case

### Apply a text edit with tracked changes

```js
const result = await applyRedlineToOxml(oxml, originalText, modifiedText, {
  generateRedlines: true,
  author: 'Agent Name'
});
```

`existingRevisions` defaults to `'merge-same-author'`. When a target paragraph
contains tracked changes from the same author, prior revisions by that author are
reverted to the pre-revision baseline and re-diffed to the new text, cleanly
merging the edits without accumulating intermediate revisions or nesting markup.
If the paragraph contains revisions from a different reviewer, the edit fails
with `EXISTING_REVISIONS` to safeguard third-party marks. Pass
`existingRevisions: 'slice-cross-author'` (or `--existing-revisions slice-cross-author`)
to preserve the other reviewer's attribution while applying Word-native
insertions and deletions inside their pending insertion. Pass
`existingRevisions: 'accept-all-first'` (or `--existing-revisions accept-all-first`
via CLI) to normalize all prior revisions first, or `'reject-input'` to refuse any
paragraph with open revisions. Use `'accept-all-first-keep-normalized'` only when
accepted revisions should be returned as a real change even on a no-op edit.
Same-author merging also fails with `COMMENTED_CONTENT_MERGE` when the revised
paragraph contains comment anchors, because reverting the prior revision could
remove or orphan those comments. Resolve the comments before re-editing.

### Apply a text edit without tracked changes (Direct Edits)

> [!IMPORTANT]
> **Tracked redlines are not always the preferred method.** When finalizing execution copies of contracts, restructuring documents, correcting minor typos, or whenever the user specifically desires clean document text without tracked changes markup clutter, pass `generateRedlines: false` (or `--no-redlines` via CLI).

```js
const result = await applyRedlineToOxml(oxml, originalText, modifiedText, {
  generateRedlines: false
});
```

### Convert OOXML to readable text or markdown

```js
import { ingestWordOoxmlToPlainText, ingestWordOoxmlToMarkdown } from '@ansonlai/docx-redline-js';
const plainText = ingestWordOoxmlToPlainText(documentXml);
const markdown = ingestWordOoxmlToMarkdown(documentXml);
```

### Add a comment to OOXML

```js
import { injectCommentsIntoOoxml } from '@ansonlai/docx-redline-js';
const result = injectCommentsIntoOoxml(paragraphOoxml, [
  {
    paragraphIndex: 1,
    textToFind: 'force majeure',
    commentContent: 'Review this clause'
  }
], { author: 'Agent' });
```

`paragraphIndex` is 1-based within the supplied OOXML payload. The comment
author belongs in the options object and applies to the injected comments.

### Accept tracked changes from one user (or all users)

```js
import { acceptTrackedChangesInOoxml } from '@ansonlai/docx-redline-js';
const acceptedMine = acceptTrackedChangesInOoxml(documentXml, { author: 'Agent' });
const acceptedAll = acceptTrackedChangesInOoxml(documentXml, { allAuthors: true });
```

### Reject tracked changes from one user (or all users)

```js
import { rejectTrackedChangesInOoxml } from '@ansonlai/docx-redline-js';
const rejectedMine = rejectTrackedChangesInOoxml(documentXml, { author: 'Agent' });
const rejectedAll = rejectTrackedChangesInOoxml(documentXml, { allAuthors: true });
```

Move revisions are consumed too: accept removes `w:moveFrom` and unwraps
`w:moveTo`; reject unwraps `w:moveFrom` and removes `w:moveTo`.

### Delete comments from one user (or all users)

```js
import { deleteCommentsByAuthorInOoxml } from '@ansonlai/docx-redline-js';
const removedMine = deleteCommentsByAuthorInOoxml(packageOrDocumentOoxml, { author: 'Agent' });
const removedAll = deleteCommentsByAuthorInOoxml(packageOrDocumentOoxml, { allAuthors: true });
```

### Apply multiple operations to full document XML

```js
import {
  applyOperationToDocumentXml,
  applyOperationsToDocumentXml
} from '@ansonlai/docx-redline-js/standalone-runner';

const result = await applyOperationsToDocumentXml(documentXml, operations, 'Agent', runtimeContext, options);
```

The operation runner uses these field names:

```js
const operations = [
  { type: 'redline', target: 'Old paragraph text', modified: 'New paragraph text', targetRef: 12 },
  { type: 'comment', target: 'Paragraph text', textToComment: 'anchor text', commentContent: 'Comment body', targetRef: 18 },
  { type: 'highlight', target: 'Paragraph text', textToHighlight: 'anchor text', color: 'yellow', targetRef: 24 }
];
```

To counterpropose text for a paragraph wholly deleted by another reviewer,
use explicit restoration intent. A normal `redline` remains fail-closed with
`FOREIGN_PARAGRAPH_MARK_DELETION`:

```js
const restoration = {
  type: 'restore',
  target: { paragraphId: '1A2B3C4D', exactText: 'Original deleted paragraph text.' },
  modified: 'Restored or adjusted paragraph text.',
  author: 'Editor'
};
```

`restore` targets default to `revisionView: 'rejected'`, because that is where
the deleted source text is visible. Inspection/extraction fingerprints are
view-scoped and returned with a matching `revisionView`; copy the fingerprint,
exact text, and view together. Set `revisionView: 'accepted'` explicitly only
when intentionally using an accepted-view descriptor.

For a contiguous range, provide `targetEnd`/`targetEndRef` and one string per
source paragraph in `modified`. Restoration always uses tracked changes,
preserves the deleted source paragraph, and inserts the counterproposal after
the complete deleted source block with a fresh paragraph ID. Unchanged
pre-existing validation defects remain baseline diagnostics; a restore fails
with `GENERATED_OOXML_INVALID` only when it introduces a new validation error.

To insert run-level text inside content visible only in the rejected view, use
an explicit rejected-view `insert` operation:

```js
const insertion = {
  type: 'insert',
  target: { paragraphId: '1A2B3C4D', revisionView: 'rejected' },
  anchor: { exactText: 'must pay', occurrence: 1, offset: 5 },
  modified: '[clarification] ',
  author: 'Editor',
  existingRevisions: 'slice-cross-author'
};
```

The anchor offset is relative to `anchor.exactText`. Repeated anchors require an
explicit `occurrence`. The engine preserves the foreign deletion as sibling
`w:del` carriers around a top-level `w:ins`; unsupported comments, bookmarks,
fields, hyperlinks, moves, or non-text split boundaries fail closed.

`targetRef` is an optional 1-based paragraph reference used to disambiguate
duplicate text. An operation-level `author` overrides the batch author; batch
results report both `authorUsed` per item and the aggregate `authorsUsed` list.

For safer targeting, `target` may be a descriptor:

```js
{
  type: 'replace',
  target: {
    exactText: 'Repeated paragraph text',
    paragraphId: '1A2B3C4D', // when present in the source OOXML
    index: 12,
    occurrence: 2,
    inTable: false,
    fingerprint: 'fnv1a32:...'
  },
  modified: 'Replacement text',
  author: 'Editor'
}
```

Call `preflightOperations(documentXml, operations, author)` when you want a
read-only inspection of an agent-generated batch before applying it. Preflight is
read-only and strict by default: duplicate exact text returns `AMBIGUOUS_TARGET`,
approximate text is not selected, and the result reports candidate targets,
missing anchors, existing revisions, authors, required artifacts, and
same-paragraph conflicts. Direct execution is progressive by default
(`atomic: false`); pass `{ atomic: true }` when the batch must roll back as a
unit. When permissive resolution
encounters duplicate candidate paragraphs, it emits an
`AMBIGUOUS_TARGET_HEURISTIC_USED` warning; migrate to `{ strictTargets: true }`
with strict descriptors (`paragraphId`, `index`, `occurrence`, or `fingerprint`)
before v1.0.0.

Whole-paragraph deletions targeting paragraphs with existing comments fail with
`COMMENTED_CONTENT_DELETE`. Resolve or remove the comments first.

Use `result.documentXml` from these APIs when replacing full `word/document.xml`.
For mixed batches, prefer `applyOperationsToDocumentXml(...)`; it applies comments
before replacements so earlier edits cannot invalidate their anchors.

Batches default to `atomic: false` for maximum speed and progressive execution: valid operations are applied directly, while problematic operations report structured errors in `results` (with `continueOnError: true` by default).

When all-or-nothing transactional protection is desired (e.g. in high-stakes legal contracts, large automated migrations, or strict CI pipelines where partial edits are inadmissible), pass `{ atomic: true }` (or `--atomic` on the CLI). In atomic mode, if any operation fails or yields invalid markup, the entire batch is rolled back to the original untouched document (`rolledBack: true`, `hasChanges: false`, `documentXml: original`).

Internally, a batch uses one live document DOM and one revision allocator, then
serializes the full document once. Every operation has a DOM/allocator savepoint;
do not remove this isolation merely for speed. Redline accuracy, accepted and
rejected text, and exact rollback take precedence over throughput.

Every operation produces a commit-aware `receipt` (and batch-level `receipts`)
enumerating exact allocated `revisionItems`, `commentIds`, `numberingIds`,
`relationshipIds`, `affectedTargets`, and `warnings`. The output reconciliation
oracle (`reconcileReceiptsAgainstOutput`) validates that all reported durable IDs
are present in the serialized output; any discrepancy triggers rollback and fails closed.

Always inspect `status` and `error`, not only `hasChanges`. A failed transform
can return `{ hasChanges: false, status: 'error', error: ... }`. Missing or
ambiguous comment anchors are structured errors and roll back atomic batches;
`no_change` is reserved for genuine no-ops. Continue to inspect warnings for
non-fatal diagnostics.

### Detect existing tracked changes

```js
import { containsTrackedChanges } from '@ansonlai/docx-redline-js';
const hasTrackedChanges = containsTrackedChanges(xmlDoc);
```

### Inspect document parts before editing

```js
import { inspectDocumentParts } from '@ansonlai/docx-redline-js';
const inspection = inspectDocumentParts({ documentXml, commentsXml, numberingXml });
```

Reuse `exactText` plus `paragraphId` or `fingerprint` in an operation. Computed
list labels and excerpts are for display, not replacements for exact targets.

### Safely edit a complete DOCX in Node

```js
import { openDocx } from '@ansonlai/docx-redline-js/node';
const document = openDocx(inputBuffer);
const result = await document.applyOperations(operations, {
  author: 'Agent', atomic: true, validate: true
});
if (!result.written) throw new Error(result.error?.message || 'No output written');
const outputBuffer = result.toBuffer();
```

This facade defaults to strict targets, allocates package-safe comment IDs,
merges numbering, updates relationships/content types, and rolls back to the
original buffer when an atomic transaction fails.

### Designing a thin agent wrapper

New wrappers should delegate at a package boundary instead of copying internal
algorithms:

- Shell/file wrappers invoke the `docx-redline` CLI and preserve its JSON stdout
  and exit code.
- Node byte-oriented wrappers use `openDocx`, `inspect`, `applyOperations`,
  `resolveRevisions`, `deleteComments`, and `toBuffer`.
- XML-only hosts use the standalone runner and remain responsible for package
  artifacts returned beside `documentXml`.
- Paragraph/range hosts use root exports and remain responsible for deciding how
  the returned OOXML is inserted into a larger document.

A wrapper that vendors the CLI should perform a startup compatibility handshake
with `docx-redline version`. Pin the minimum CLI `contractVersion` and only the
capabilities that the wrapper's workflow actually requires. If the runtime is
too old, fail closed with an upgrade instruction; do not inspect the bundled
implementation or fall back to direct ZIP/XML mutation.

A wrapper may choose product defaults for author, atomic mode, output naming,
and accepted operation subsets. It must preserve strict targeting, exact text,
structured errors, per-operation results, receipts, warnings, validation, and
rollback behavior. It must not infer success from `hasChanges` alone or turn a
partial progressive result into an unconditional success.

Document those choices as wrapper policy and pass them explicitly. In
particular, do not describe a wrapper's preferred revision or atomicity policy
as though it were the underlying CLI or facade default.

Target handles are scoped to the exact package version that produced them. A
Node wrapper should return the package-scoped token from
`document.getRevisionToken()` with every inspection and pass it back as
`expectedRevision` when applying the planned operations. Do not substitute the
document-parts token included in `document.inspect()`; that token has a different
scope and the Node facade rejects it. A shell wrapper should ensure extraction
and application use the same unchanged path, and re-extract after switching to a
derived working copy.

For agent-facing function tools, prefer separate inspection, application, and
review-resolution tools. Describe each tool's use case, required inputs, side
effects, retry safety, success criteria, and common error codes. Convenience
tools should build operations defined by
`docs/schemas/document-operations.schema.json` and delegate to the same apply
path rather than implementing custom mutations.

For a `restore_deleted_paragraph` convenience tool, force inspection to
`revisionView: 'rejected'`, copy the exact target descriptor from that view, and
emit a canonical `restore` operation with the same explicit revision view.
Wholly foreign-deleted paragraphs appear empty in accepted/current inspection;
this is expected, not evidence that the paragraph is untargetable. Never reuse a
restore descriptor from a different source or earlier working-copy version. On
`TARGET_TEXT_MISMATCH`, re-extract from the exact package being applied rather
than inspecting implementation bundles or retrying the same operation.

### Agent Document Workflow (CLI)

Use the `docx-redline` CLI for complete `.docx` files. It emits JSON on stdout,
keeps exact text intact, and never overwrites the source unless `--in-place` is
explicitly supplied.

#### Operation Model: Choose by Output Shape

Every text-bearing operation targets existing content. Its `modified` field is
the complete desired accepted-view content for that target, not merely the new
fragment to insert. Do not choose a type from its English name alone.

| Requested result | Operation shape |
|---|---|
| Change text within one paragraph or replace its content | `{ type: 'redline', target, modified }` (`replace` is a compatibility alias) |
| Delete a whole paragraph | `{ type: 'delete', target }` (normalized to a redline with `modified: ''`) |
| Change native list structure | `{ type: 'list-change', target, modified: '<complete Markdown list>' }` |
| Reconcile a Word table | `{ type: 'table-reconciliation', target, modified: '<complete Markdown table>' }` |
| Comment or highlight existing text | `comment` with `commentContent`, or `highlight` with `textToHighlight` |
| Change character or paragraph formatting | `character-format`/`format` with `textToFormat` and `properties`, or `paragraph-format` with `properties` |
| Counterpropose a paragraph wholly deleted by another author | `restore` with a rejected-view target |
| Insert text inside another author's rejected-view content | `insert` with `target.revisionView: 'rejected'`, an exact `anchor`, and `existingRevisions: 'slice-cross-author'` |

`insert`, `list-change`, `table-reconciliation`, `replace`, and text-bearing
`format` are accepted compatibility types, but they normally normalize to the
same redline operation path. In particular, ordinary `{ type: 'insert' }` does
not mean “create a new sibling paragraph”; without a rejected-view target and
anchor, `modified` is still interpreted as the complete replacement text.

To append a native sibling item after an existing list item, prefer an explicit
`list-change` whose `modified` value contains the complete affected list block:

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

The Markdown markers describe structure; visible labels are generated from the
document's numbering. Do not put `m)` or another computed label into target text.
For a single adjacent item in an existing list, the runner also accepts a
one-line `redline` where `modified` is the exact current item followed by the
new item's unnumbered text. Preserve the current item verbatim, omit the new
label, and make the new item at least six words so the adjacency form is
unambiguous. Use the explicit Markdown list form when adding multiple items,
nesting, or changing levels.

The canonical machine-readable contract is
`docs/schemas/document-operations.schema.json`. Read that schema or the examples
here instead of grepping bundled implementation code.

#### Standard Workflow (Fast & Direct)

Use this for everything by default. `apply` is fast, progressive, and self-validating by default—it validates the resulting package and revision markup internally before writing. **Do not insert a `preflight` or baseline `validate` step on top of it "to be safe"**; `apply` already covers that internally. It supports inline one-liners as well as batch operations files:

The short route for a document-editing request is:

1. For an exact mechanical change with known old and new literals, use localized
   `--find`/`--replace`. Omit the target for global fail-closed resolution, or
   scope it with a longer, fairly unique nearby `--search` phrase and a
   directional range such as `--context-range 1:3`. Generic or repeated anchors
   widen the unioned scope and may fail with `AMBIGUOUS_TARGET`, costing another
   turn. `--occurrence` selects only within one uniquely resolved paragraph.
2. For semantic drafting, run one focused `extract`. Add `--around 3` when
   surrounding context on either side is needed, then copy `exactText` plus
   `paragraphId` or `fingerprint`.
3. Build the final operations from the operation table above. Use one operation
   per target paragraph, and consolidate multiple changes to that paragraph.
4. Run `apply` once per stable batch. Strong inspected targets are bound against
   the batch-start document, so independent operations do not need manual
   bottom-up sorting around structural edits. Consolidate multiple complete
   desired states for the same source. A unique exact reference to paragraph
   text created elsewhere in the batch is scheduled automatically; use explicit
   captures/selectors for non-unique or advanced created-content dependencies.
5. Walk every result and require `completion: true`, `written: true`, and no
   per-operation error. For localized patches also require
   `results[i].change.committed: true`, `finalDisposition: "applied"`, and
   positive accepted-view verification. If `anchorMatchCount > 1`, confirm the
   returned location and excerpts because the anchor was repeated.
6. Run a focused `extract` on changed clauses only when placement or list/table
   structure needs confirmation.

Do not probe operation behavior with disposable apply commands or read a vendor
bundle before this route. If `apply` returns an error, use the recovery matrix
below and make one cause-specific correction.

```bash
# 1. Globally unique mechanical edit in one call
docx-redline apply contract.docx --find "thirty (30) days" --replace "sixty (60) days" --profile agent --output reviewed.docx

# 2. Directional edit using a longer phrase with a higher chance of uniqueness
docx-redline apply contract.docx --search "distinctive nearby heading or phrase" --context-range 1:3 --find "thirty (30) days" --replace "sixty (60) days" --profile agent --output reviewed.docx

# 3. Focused contextual discovery for semantic drafting
docx-redline extract contract.docx --search "termination" --around 3

# 4. Inline complete-paragraph edit
docx-redline apply contract.docx --target "Original clause" --modified "New clause" --output reviewed.docx

# 5. Direct edit without tracked changes (clean text, no revision clutter)
docx-redline apply contract.docx --target "Typo fix" --modified "Fixed typo" --no-redlines --output clean.docx

# 6. Cross-author edit inside another reviewer's pending insertion
docx-redline apply contract.docx --target "Pending clause text" --modified "Updated clause text" --existing-revisions slice-cross-author --output reviewed.docx

# 7. Batch operations with ops.json
docx-redline apply contract.docx --operations operations.json --output reviewed.docx

# 8. Serializer-backed stdin with compact stdout
node emit-operations.mjs | docx-redline apply contract.docx --operations - --profile agent --compact --output reviewed.docx
```

Key CLI defaults and behaviors:
- **Author**: Automatically defaults to `'AI Redliner'` (overridable via `--author` or `DOCX_REDLINE_AUTHOR` environment variable).
- **Existing revisions**: Defaults to `'merge-same-author'`. Pass `--existing-revisions slice-cross-author` to edit inside another reviewer's pending insertions with native carrier slicing.
- **Overwrite behavior**: Destination files provided via `--output` overwrite by default. To protect existing destination files, pass `--no-overwrite` or `--no-clobber`. The source document is never overwritten unless `--in-place` is specified.
- **Tracked changes**: Defaults to `generateRedlines: true`. When clean direct text is needed, pass `--no-redlines`.
- **Atomic rollback (optional)**: Operations apply progressively by default (`atomic: false`). For all-or-nothing transactional rollback where any error halts and reverts all changes, pass `--atomic`.
- **Complete-success exit (optional)**: Pass `--require-complete` when `partial` must exit nonzero (`3`). Errors exit `2`; the legacy zero exit for partial results remains when the flag is omitted.
- **Agent profile**: `--profile agent` enables complete-success exit behavior but preserves progressive execution and the ordinary revision policy. Compose it with `--atomic` or an explicit `--existing-revisions` choice when intended; resolved values appear in `effectiveOptions`.
- **Operation transport**: A UTF-8 operations file and serializer-backed `--operations -` are peers. Use whichever the host can construct without interpolating legal text in the shell.
- **Compact mutation JSON**: `apply`, `accept`, `reject`, and `delete-comments` omit document/package XML, full validation arrays, and duplicate nested receipts. Top-level `receipts` is authoritative. Pass `--compact` for one-line JSON; run `validate` for full issue records.
- Check `completion: true`, `written: true`, and a non-null `outputPath` on stdout. `completion` is derived from the write result, top-level status, and every operation status, so failed, partial, and unwritten work cannot appear complete. If an error occurs, inspect `error.code` or `results[i].error.code` (e.g. `TARGET_NOT_FOUND`, `ANCHOR_NOT_FOUND`) before correcting the cause and re-applying.

For multi-clause or multi-page reviews, apply edits **section-by-section** or clause-by-clause (e.g., using `--in-place` on a working copy) rather than bundling dozens of edits into one massive batch. This keeps context compact, simplifies error diagnosis, and prevents cascading anchor drift.

#### High-Assurance / Staged Verification Workflow (Optional)

This is an opt-in, higher-latency path for cases like large automated batch
migrations or workflows where the user specifically requests a non-mutating dry run
and an independent baseline audit report. **Never switch into it on your own initiative**
(not even for "high-stakes" contracts); unless the user explicitly requests it, stick with the
Standard workflow above. Use the extended verification cycle:

```bash
docx-redline inspect contract.docx --non-empty
docx-redline extract contract.docx --range 10:30 > paragraphs.json
docx-redline preflight contract.docx --operations operations.json --author "Editor"
docx-redline apply contract.docx --operations operations.json --author "Editor" --output reviewed.docx
docx-redline validate reviewed.docx --baseline contract.docx
```

Copy `exactText`, `paragraphId`, and `fingerprint` from `extract` into operation
targets. For most unique clauses, `"target": "exact paragraph text"` is
sufficient; use discriminators (`paragraphId`, `fingerprint`, `index`, or
`occurrence`) when duplicate paragraph text appears in the document. Never
normalize or reconstruct `exactText`. Operation files follow
[`docs/schemas/document-operations.schema.json`](schemas/document-operations.schema.json).

#### Commands

- `inspect` returns the structured inventory, comments, authors, and counts.
- `extract` returns a compact target inventory with exact text.
- `preflight` checks targets, anchors, revisions, conflicts, authors, and needed artifacts without mutation (read-only).
- `apply` applies an operation file transactionally with automatic rollback and internal markup validation.
- `accept` and `reject` resolve revisions selected by `--author` or `--all-authors`.
- `delete-comments` removes matching definitions and document anchors together.
- A whole-paragraph delete stops with `COMMENTED_CONTENT_DELETE` when the
  paragraph has an existing comment. Surface the returned reviewer and comment
  text for human follow-up; do not silently convert this into comment removal.
- `validate` audits revision markup and DOCX package wiring, optionally comparing against a `--baseline`.

Paragraph indexes are 1-based. Inspection filters are `--index 12`,
`--range 10:30`, `--indexes 2,5,8`, `--search text`, `--revised`, `--table`,
`--body`, `--non-empty`, and `--view accepted|rejected|current`. Search is
case-insensitive. When a search returns 0 matches in the active view but matches
exist in the alternate view (such as searching for a deleted clause to restore),
`selection.hint` provides immediate guidance to re-run with `--view rejected`.
Compact `extract --revised` results retain `hasRevisions`, `revisionAuthors`,
and `deleted`/`inserted` cues, so an empty `exactText` is identifiable as content
hidden by the selected revision view rather than a failed filter.
Add `--around N` (`--context N` or `-C N`) to a search;
context records are labeled separately and do not consume the direct-hit
`--limit`. Continue a bounded result with `--after <paragraph-index>`. Ordinary
unscoped CLI inspection defaults to 20 direct records and a 48 KiB soft budget;
`--all` deliberately opts out. A malformed filter or unknown option is an error
rather than an unfiltered fallback. CLI help commands (`apply --help`,
`extract --help`, `--help`) report canonical GitHub repository URLs in their
`documentation` array.

Mutating commands use `--author`, operation-level authors, then
`DOCX_REDLINE_AUTHOR`, falling back visibly to `AI Redliner`; review-resolution
commands may use `--all-authors` where applicable. Without `--output`, a sibling
such as `contract.redlined.docx` is chosen. Existing outputs are refused unless
`--force` is present. `--in-place` is the only way to overwrite the input.

Treat a nonzero exit code or JSON `status: "error"` as failure. A failed atomic
operation reports `written: false` and does not write an output file.
Missing or repeated comment anchors are errors rather than no-ops. Explicit
anchors match exact text first and then a unique ordinary-space/NBSP equivalent;
omit `textToComment` to comment the entire resolved paragraph.

To reply inside an existing Word comment thread, use the comment ID returned by
`inspect` and do not supply a paragraph target:

```json
{ "type": "comment_reply", "parentCommentId": "8", "commentContent": "Agreed; updated.", "author": "Editor" }
```

Replies are represented in `word/commentsExtended.xml` and deliberately add no
new `commentRangeStart`, `commentRangeEnd`, or `commentReference` to the body.

#### Legacy skill wrapper migration

Older skills that invoke `scripts/extract_text.mjs` and
`scripts/apply_changes.mjs` should use the compatibility entrypoints published
with this package rather than carrying copied targeting or ZIP logic. The
legacy positional apply form remains supported:

```bash
node scripts/apply_changes.mjs input.docx changes.json output.docx --author "Editor"
```

Operation files may contain an array, an `operations` array, or a legacy
`changes` array. The wrapper delegates to the same strict, atomic, validated
CLI described above. If `--author` and operation authors are absent, its
compatibility fallback is `DOCX_REDLINE_AUTHOR` and then `Agent`. Consumers
must use the JSON status and process exit code; failed atomic work has
`written: false`, `outputPath: null`, and does not modify the output path.

#### Safe Operations File Creation (JSON vs. Shell Heredocs)

When composing batch operation JSON, whether for a file or stdin:

- **Use structured file-writing tools or JSON serializers**: Write operations files via your environment's file-creation tools or a language JSON serializer (`JSON.stringify`).
- **Never compose operations in raw shell heredocs** (e.g., `cat << 'EOF'` in bash or PowerShell `@" ... "@`): Legal clauses routinely contain curly quotes (`“ ”`), smart apostrophes (`’`), em-dashes (`—`), section symbols (`§`), non-breaking spaces, and backslashes. Shell heredocs frequently mangle Unicode character encodings, quote escaping, and whitespace formatting, causing immediate `TARGET_NOT_FOUND` failures.

#### Walking Progressive Batch Results (Status & Partial Execution)

In default progressive mode (`atomic: false`), operations execute independently: valid operations commit to the document while failing operations report errors without aborting the batch:

- **Do not rely solely on top-level `written: true` or `status !== "error"`**: A progressive batch can return `status: "partial"` with `written: true` when some operations succeed and others fail.
- **Walk every entry in `results`**: Check `results[i].status` and `results[i].error`. Any `status: "error"` entry in `results` represents an unapplied change that must be investigated and resolved.
- **`written: false`**: Indicates that zero operations were committed (or an atomic rollback occurred). Never treat or present an unwritten or partial output file as complete.
- **Follow `retryPlan`**: `base: "original"` means use the unchanged source and replay the corrected batch. `base: "output"` means retain committed progressive edits and submit only failed/unattempted indexes. `sameArgumentsSafe` is always false for failures.
- **Follow the recovery envelope**: Use `error.recovery.action`, `requiresReinspection`, and `requiresUserAuthorization` rather than deriving a retry from the prose message. The envelope is versioned by `recoveryVersion`.

#### Human-Readable References vs. Internal Machine Handles

Target handles such as `ref` (`P<index>`), `targetRef`, and bare paragraph `index` numbers are **strictly internal machine handles** for the CLI and engine. They do not correspond to any visual or followable marker in Microsoft Word:

- **Never surface `P11`, `P42`, or bare paragraph numbers** in user-facing prose, comments, redline summaries, or negotiation notes.
- Instead, cite locations using the human-readable fields provided by `inspect` / `extract`:
  - **`provision`**: Lead with section/clause numbers when present (e.g., `§14.1 Entire Agreement`).
  - **`nearestHeading` + ordinal offset**: When `provision` is absent, describe position relative to the nearest heading (e.g., `under "Limitation of Liability", 2nd paragraph`).
  - **Structural context**: For unnumbered clauses prior to the first heading, use plain language (e.g., `opening recital, before Section 1`).
  - **`humanReference`**: Use the pre-joined citation string provided directly on inspected paragraph objects.

#### Actionable Error Recovery Matrix

When the CLI or runner returns an error code, follow these specific recovery actions:

| Error Code | Meaning | Actionable Recovery |
|---|---|---|
| `TARGET_NOT_FOUND` | Target text did not match any paragraph. | **Do NOT retry with paraphrased text.** Re-run `extract`/`inspect`, copy `exactText` verbatim (including exact whitespace/punctuation), and add a discriminator (`paragraphId`, `fingerprint`, or `occurrence`). |
| `AMBIGUOUS_TARGET` | Multiple paragraphs match identical text. | Disambiguate by supplying `paragraphId`, `fingerprint`, `occurrence`, or `index` in the target descriptor. |
| `ANCHOR_NOT_FOUND` / `AMBIGUOUS_ANCHOR` | A comment or rejected-view insertion anchor was not uniquely matched. | For comments, narrow `textToComment` or omit it to anchor the whole paragraph. For rejected-view insertion, copy exact rejected text and provide `anchor.occurrence`. |
| `OVERLAPPING_SOURCE_TARGETS` / `OVERLAPPING_TEXT_EDITS` | Multiple complete text operations target the same batch-start source. | Consolidate all changes to that paragraph into one `redline` or `replace`; the library cannot choose between incompatible complete desired states. |
| `REVISION_ORDER_CONFLICT` | A text rewrite and formatting/highlight operation overlap the same source. | Consolidate or split the work at an intentional created-content dependency; do not try a different arbitrary order. |
| `CAPTURE_FANOUT_CONFLICT` | Multiple mutating consumers share one capture without distinct selectors. | Give consumers distinct selectors, chain them explicitly, or split the batch. |
| `EXISTING_REVISIONS` | Target paragraph contains tracked changes from another author. | Fails closed to protect third-party review marks. If editing inside that reviewer's pending insertion is intended, pass `--existing-revisions slice-cross-author` (or `existingRevisions: 'slice-cross-author'`). Do not pass `accept-all-first` without explicit user authorization. |
| `PATCH_ROUNDTRIP_MISMATCH` | A cross-author surgical edit did not reconstruct the requested modified text exactly. | Treat the operation as unapplied. Re-extract the exact paragraph text and split the edit into a narrower operation that does not cross the reported structural boundary. |
| `FOREIGN_PARAGRAPH_MARK_DELETION` | A normal edit attempted to write into a paragraph wholly deleted by another reviewer. | Use an explicit `restore` operation if the user intends to counterpropose that paragraph; otherwise leave the deletion unresolved. |
| `RESTORATION_STATE_REQUIRED` / `RESTORATION_COUNT_MISMATCH` | A `restore` target is not a wholly foreign-deleted paragraph, or its replacement count does not match the paragraph range. | Re-inspect the document and target the deleted paragraph by stable descriptor; provide exactly one replacement string per source paragraph. |
| `REJECTED_INSERTION_STATE_REQUIRED` / `UNSAFE_REVISION_BOUNDARY` | An explicit rejected-view insertion did not resolve to supported plain run text inside a wholly foreign-deleted paragraph. | Do not fall back to a generic edit. Narrow the exact anchor/offset, or handle comments, bookmarks, fields, hyperlinks, moves, or other structural boundaries manually. |
| `GENERATED_OOXML_INVALID` | The operation introduced a new validation error relative to its baseline. | Treat the operation as unapplied and inspect `generatedIssues`; correct the generating operation or builder rather than repairing or accepting the source document's unrelated baseline defects. |
| `UNSAFE_DELETED_TABLE_ROW` / `UNSUPPORTED_MOVE_REVISION` / `SECTION_BREAK_PARAGRAPH` / `UNSAFE_PARAGRAPH_PLACEMENT` | Paragraph restoration cannot preserve the source structural boundary safely. | Do not retry as an ordinary redline. Resolve the row/move/section/placement condition manually or narrow the restoration to a safe paragraph. |
| `COMMENTED_CONTENT_MERGE` / `COMMENTED_CONTENT_DELETE` | Operation would overwrite, revert, or delete content with comments. | Fails closed to prevent orphaned comment threads. Report the comment author and text to the user; resolve the comment before re-editing. |
| `INVALID_OPERATION` | Operation object violates schema or has incompatible fields. | Validate the JSON structure against [`document-operations.schema.json`](schemas/document-operations.schema.json) before targeting is attempted. |
| `STRUCTURED_CONTENT_INVALID` | Malformed Markdown table or structure in replacement text. | Ensure tables include a separator row (`\| --- \| --- \|`) and consistent column counts; do not downgrade to raw text. |

**Important Rule:** Never repeat the exact same failing command without correcting the reported cause. If an error persists after one correction attempt, stop and report the diagnostic code to the user.

#### Document Scope & Boundary Invariants

The `docx-redline` engine and CLI operate specifically on the **main document body**:

- **Supported Content**: Body paragraphs, numbered/bulleted lists, tables and table cells, comments, and comment replies.
- **Unsupported Content**: Headers, footers, footnotes, endnotes, floating text boxes, shape drawings, watermarks, and embedded macros.
- Do not attempt to target, edit, or comment on header/footer text or footnote citations using `docx-redline`. Use specialized document manipulation tools or manual editing for layout frames outside the body text.

### Convert paragraph text into a Word list

```js
const result = await applyRedlineToOxml(oxml, 'Item text', '1. Item text', {
  generateRedlines: true
});
```

### Insert a large mixed-content block safely

Do not send a long attachment containing literal pipe rows, headings, lists,
and paragraphs as an unchecked replacement. Plan it first:

```js
import { planStructuredReplacement } from '@ansonlai/docx-redline-js';

const plan = planStructuredReplacement(targetDescriptor, markdown, {
  author: 'Agent'
});
if (!plan.valid || !plan.operation) {
  throw new Error(plan.issues.map(issue => issue.message).join(' '));
}
const result = await document.applyOperations([plan.operation], {
  author: 'Agent', atomic: true, validate: true
});
```

Use blank lines between paragraphs, `#`/`##` for headings, normal Markdown
markers for lists, and a separator row immediately after every table header:

```markdown
| Agency | Contact |
| --- | --- |
| BCHD | Dr. Jenkins |
```

The planner returns typed `blocks`, counts, normalized Markdown, and structured
issues. `TABLE_SEPARATOR_REQUIRED` is an error: never remove `structuredContent`
or retry the same content as plain text merely to make the operation pass. Keep
the result as one atomic replacement operation so the first inserted block does
not invalidate the anchor for later blocks. After applying, require real
`w:tbl`, positive list `w:numId` values, valid redline OOXML, and independent
Accept/Reject checks.

### Reconcile a table

```js
import { reconcileMarkdownTableOoxml } from '@ansonlai/docx-redline-js';
const result = await reconcileMarkdownTableOoxml(tableOoxml, originalText, markdownTable);
```

## Detailed Architecture

Use the layout and task shortcuts at the top of this file for normal work.
`ARCHITECTURE.md` is the maintained detailed module-ownership map and describes
the end-to-end dependency flow. Do not infer ownership from directory names or
reconstruct the architecture by repeatedly listing the repository.

## Common Patterns

### Options and Defaults Reference

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `generateRedlines` | `boolean` | `true` | When `true`, emits Word-native tracked changes (`w:ins`/`w:del`). When `false`, applies direct text edits without revision markup. **Note: Redlines are not always the preferred method** — pass `generateRedlines: false` (or `--no-redlines` via CLI) when producing clean execution drafts, restructuring documents, or when revision clutter is unwanted. |
| `author` | `string` | `'AI Redliner'` | Reviewer/author name stamped on generated tracked changes and comments. Overridable via `DOCX_REDLINE_AUTHOR` environment variable. |
| `atomic` | `boolean` | `false` | Batch transaction mode. By default (`false`), valid edits are applied directly and failing operations report errors. When `true`, any operation failure rolls back the entire batch to the original document state (`rolledBack: true`, `hasChanges: false`). Use `atomic: true` (or `--atomic` in CLI) for high-assurance workflows. |
| `structuredContent` | `boolean` | `true` | Auto-detects Markdown tables, headings (`#`), and lists in replacement text and renders them as native Word elements (`w:tbl`, `w:pStyle`, `w:numPr`). Pass `false` to treat replacement text strictly as plain text. |
| `pairReplacements` | `boolean` | `true` | Links adjacent `<w:del>` and `<w:ins>` revisions with matching timestamps so Word groups them as a single replacement in the Reviewing Pane. |
| `strictTargets` | `boolean` | `true` (CLI/facade) | Requires exact target descriptors (`exactText`, `paragraphId`, `index`, `occurrence`, `fingerprint`) and forbids ambiguous matching. Defaults to `false` in low-level runner for backwards compatibility. |
| `existingRevisions` | `string` | `'merge-same-author'` | How to handle paragraphs with existing tracked changes. `'merge-same-author'` merges the same author's work and protects other authors with `EXISTING_REVISIONS`. `'slice-cross-author'` retains same-author merging while allowing Word-native edits inside another author's pending insertion. Pass `'accept-all-first'` to normalize prior revisions or `'reject-input'` to refuse editing revised paragraphs. |
| `removeFormatting` | `boolean` | `false` | When `true` and the text is unchanged with no Markdown hints, strips existing bold/italic/underline/strikethrough formatting. |
| `sanitizeInput` | `boolean` | `false` | Opt-in removal of standalone leading assistant-preface lines. Literal dollar signs and `\n` sequences are always preserved. |

### Options shape

```js
{
  generateRedlines: true,
  author: 'AI Redliner',
  atomic: false,
  structuredContent: true,
  pairReplacements: true,
  existingRevisions: 'merge-same-author',
  removeFormatting: false,
  sanitizeInput: false
}
```

### Typical return shape

```js
{
  oxml: string,
  hasChanges: boolean,
  status?: 'ok' | 'no-op' | 'error',
  error?: { code: string, message: string },
  warnings?: string[],
  numberingXml?: string,
  useNativeApi?: boolean
}
```

Known error codes include `PARSE_ERROR`, `TARGET_NOT_FOUND`, `PARTIAL_TARGET`,
`EXISTING_REVISIONS`, `COMMENTED_CONTENT_MERGE`, `UNSAFE_REVISION_NESTING`, `UNSUPPORTED_REVISION_VIEW_MUTATION`,
`UNSAFE_PARAGRAPH_BOUNDARY`, `DIFF_TOKEN_LIMIT`, and `BATCH_OPERATION_FAILED`.

For ingestion that must distinguish an empty document from malformed OOXML,
use `ingestWordOoxmlToPlainTextResult` or
`ingestWordOoxmlToMarkdownResult`. The legacy ingestion helpers intentionally
retain their string-only return type and return `''` for parse failures.

### Target text versus replacement text

Target resolution may normalize surrounding or repeated whitespace while
matching a paragraph. Replacement text is not normalized: tabs, line breaks,
non-breaking spaces, repeated spaces, and leading/trailing whitespace become
part of the requested edit. When editing extracted document text, copy the
exact paragraph text and modify it in place rather than round-tripping it
through a formatter that may change whitespace.

The normalized caller target is never used for mutation offsets in a text-bearing
edit. After target selection, the engine uses the resolved paragraph's byte-exact
JavaScript string as the source coordinate system. The legacy format-only
fallback for field-code paragraphs with no extractable accepted-view spans is
not a text-replacement path. If an ASCII-space target selected an NBSP
source, `resolvedTarget.targetTextMatch` reports `space_equivalent`, escaped
source/request excerpts, and differing code points. An NBSP-to-space request is
tracked as a replacement; it must not retain the NBSP and append another space.
The CLI keeps these bounded diagnostics but removes resolved clause text.

For ordinary insertions and deletions, target the visible accepted view:
inserted `w:t` text is visible and deleted `w:delText` is not. Move revisions
and other complex structures require additional care until targeting and
ingestion share one canonical text extractor. Prefer a `targetRef` plus the full
paragraph text when duplicate paragraphs are possible. Current text-only
matching can select the first matching paragraph, so callers that cannot
disambiguate safely should stop instead of guessing.

### OOXML wrapping for Word insertOoxml scenarios

```js
import { wrapInDocumentFragment } from '@ansonlai/docx-redline-js';
const wrapped = wrapInDocumentFragment(rawOoxml, { includeNumbering: true, numberingXml });
```

### Output shape guardrail (important for packaging)

When consuming `result.oxml`, do not assume the payload is always safe to write
directly into `word/document.xml`.

- Paragraph/range/table APIs can return a fragment, `<w:document>`, or package payload (`<pkg:package>`).
- `applyOperationToDocumentXml(...).documentXml` is the document-safe path when you need a full `word/document.xml` replacement.
- Use `extractReplacementNodesFromOoxml(payload)` to normalize unknown payloads.
- If `sourceType === 'package'` or the payload starts with `<pkg:package`, do not write it into `word/document.xml` as-is.

## Gotchas

1. Call `configureXmlProvider` first in Node.js.
2. `applyRedlineToOxml` is async.
3. Paragraph APIs expect paragraph-level OOXML, not full `word/document.xml` in all cases.
4. List operations may return `numberingXml` that must be merged into package parts. When `word/numbering.xml` already exists, pass `mergeNumberingXmlBySchemaOrder` to `ensureNumberingArtifactsInZip`; without a merge callback the helper replaces the prior payload.
   That replacement behavior is deprecated and will become an error in the next major version.
5. `useNativeApi: true` means standalone mode cannot fully handle that operation path.
6. `deleteCommentsByAuthorInOoxml` removes definitions and linked anchors only when they are present in the same OOXML payload. In a real `.docx`, `word/comments.xml` and `word/document.xml` are separate parts and must both be updated by the package integration layer.
7. If output begins with `<pkg:package`, treat it as package-level OOXML and normalize it before writing anything back to `word/document.xml`.
8. Existing revisions from the same author are merged by default against the pre-revision baseline (`merge-same-author`), while third-party revisions fail closed with `EXISTING_REVISIONS`. Pass `existingRevisions: 'slice-cross-author'` to preserve third-party attribution while editing inside pending insertions, `'accept-all-first'` to normalize all prior revisions first, or `'reject-input'` to refuse any revised paragraph.
9. Caller content is not sanitized by default. Pass `sanitizeInput: true` only for raw assistant output; literal dollar delimiters and `\\n` sequences are never rewritten.
10. Hyperlinks, bookmarks, comment markers, tabs/breaks, and footnote/endnote references are structural OOXML and should survive adjacent redline edits.
11. Internally, create Word elements through `createWordElement` and tracked-change metadata through `createRevisionMetadata`.
12. Revision IDs are document-scoped in public operation paths. Thread the
    internal allocator through new string-serialization paths; generated
    `w:id` values are not stable across documents.
13. Splitting or cloning a run can duplicate nested `w:rPrChange` metadata.
    Preserve the original ID on at most one resulting run and allocate fresh
    IDs for every additional clone through the document-scoped allocator.
14. Run `validateRedlineOoxml` on generated markup before packaging it, then
    run `validateDocxPackage` after merging comments and numbering artifacts.

## Validation Commands

```bash
npm test
npm run test:isolation
npm run check:types
node scripts/export-validation-fixtures.mjs
```

Optional Windows/Word smoke test for a completed `.docx`:

```bash
npm run smoke:word -- path/to/file.docx
```
