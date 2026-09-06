# AGENTS.md - AI Agent Quick Reference

> This file helps AI coding agents understand @ansonlai/docx-redline-js quickly.
> Read this instead of exploring the full source tree.

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

`index.js` is the single package entry point.

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
same-paragraph conflicts. For direct execution, `applyOperationsToDocumentXml` is
already transactional and atomic by default. When permissive resolution
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

### Agent Document Workflow (CLI)

Use the `docx-redline` CLI for complete `.docx` files. It emits JSON on stdout,
keeps exact text intact, and never overwrites the source unless `--in-place` is
explicitly supplied.

#### Standard Workflow (Fast & Direct)

`apply` is fast, progressive, and self-validating by default. It supports inline one-liners as well as batch operations files:

```bash
# 1. Inline one-liner edit (fastest for 1–2 edits; no JSON file needed)
docx-redline apply contract.docx --target "Original clause" --modified "New clause" --output reviewed.docx

# 2. Direct edit without tracked changes (clean text, no revision clutter)
docx-redline apply contract.docx --target "Typo fix" --modified "Fixed typo" --no-redlines --output clean.docx

# 3. Batch operations with ops.json
docx-redline apply contract.docx --operations operations.json --output reviewed.docx
```

Key CLI defaults and behaviors:
- **Author**: Automatically defaults to `'AI Redliner'` (overridable via `--author` or `DOCX_REDLINE_AUTHOR` environment variable).
- **Overwrite behavior**: Destination files provided via `--output` overwrite by default. To protect existing destination files, pass `--no-overwrite` or `--no-clobber`. The source document is never overwritten unless `--in-place` is specified.
- **Tracked changes**: Defaults to `generateRedlines: true`. When clean direct text is needed, pass `--no-redlines`.
- **Atomic rollback (optional)**: Operations apply progressively by default (`atomic: false`). For all-or-nothing transactional rollback where any error halts and reverts all changes, pass `--atomic`.
- Check `written: true` on stdout. If an error occurs, inspect `error.code` or `results[i].error.code` (e.g. `TARGET_NOT_FOUND`, `ANCHOR_NOT_FOUND`) to correct the target text and re-apply.

For multi-clause or multi-page reviews, apply edits **section-by-section** or clause-by-clause (e.g., using `--in-place` on a working copy) rather than bundling dozens of edits into one massive batch. This keeps context compact, simplifies error diagnosis, and prevents cascading anchor drift.

#### High-Assurance / Staged Verification Workflow (Optional)

For high-stakes legal contracts, large automated batch migrations, or workflows
requiring explicit non-mutating pre-checks and an independent baseline audit
report, use the extended verification cycle:

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
[`docs/schemas/document-operations.schema.json`](docs/schemas/document-operations.schema.json).

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
`--body`, `--non-empty`, and `--view accepted|rejected|current`. A malformed
filter or unknown option is an error rather than an unfiltered fallback.

Mutating commands require `--author`, authors on every operation, or
`--all-authors` where applicable. Without `--output`, a sibling such as
`contract.redlined.docx` is chosen. Existing outputs are refused unless
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

## Module Map

```
index.js
adapters/
  config.js
  xml-adapter.js
  logger.js
core/
  types.js
  paragraph-text.js
  word-xml.js
  paragraph-targeting.js
  list-targeting.js
  table-targeting.js
engine/
  oxml-engine.js
  surgical-mode.js
  surgical-run-splitting.js
  surgical-diff-application.js
  surgical-spans.js
  reconstruction-mode.js
  reconstruction-writer.js
  format-application.js
  formatting-removal.js
  run-builders.js
  table-mode.js
pipeline/
  pipeline.js
  ingestion.js
  ingestion-export.js
  diff-engine.js
  markdown-processor.js
  serialization.js
  list-generation.js
services/
  document-operation-session.js
  document-operation-applier.js
  document-operation-mutations.js
  batch-operation-orchestrator.js
  operation-heuristics.js
  standalone-operation-runner.js
  standalone-operation-runner.d.ts
  document-operation-contract.js
  operation-preflight.js
  standalone-docx-plumbing.js
  numbering-helpers.js
  comment-engine.js
  revision-comment-management.js
  table-reconciliation.js
  package-builder.js
  document-inspection.js
node/
  docx-document.js
  zip-archive.js
orchestration/
  route-plan.js
  list-markdown.js
  list-structural-fallback.js
```

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
| `existingRevisions` | `string` | `'merge-same-author'` | How to handle paragraphs with existing tracked changes. `'merge-same-author'` automatically merges subsequent edits from the same author against the pre-revision baseline while protecting different authors' revisions with `EXISTING_REVISIONS`. Pass `'accept-all-first'` to normalize prior revisions or `'reject-input'` to refuse editing revised paragraphs. |
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
8. Existing revisions from the same author are merged by default against the pre-revision baseline (`merge-same-author`), while third-party revisions fail closed with `EXISTING_REVISIONS`. Pass `existingRevisions: 'accept-all-first'` to normalize all prior revisions first, or `'reject-input'` to refuse any revised paragraph.
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
