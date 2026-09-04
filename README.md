# @ansonlai/docx-redline-js

Host-independent OOXML reconciliation engine for `.docx` manipulation with track changes (redlines).

Converts AI-generated or programmatic text/markdown edits into valid Office Open XML (OOXML) with `w:ins`/`w:del` revision markup that Microsoft Word renders as native tracked changes.

## Features

- Text reconciliation with word-level diffing and native-looking redlines
- Formatting updates (bold, italic, underline, strikethrough) via surgical `w:rPrChange`
- Lists: generate and edit real Word lists (`w:numPr`) from markdown
- Tables: virtual-grid diffing for cell-level edits with merge safety
- Comments: inject OOXML comments anchored to text ranges
- Revision management: detect existing revisions, consume move revisions, and accept/reject tracked changes by author or for all authors
- Comment management: delete comments by author or for all authors
- Highlights: apply highlight colors to runs
- Markdown and OOXML conversion in both directions
- Status/error result fields for parse, targeting, and existing-revision failures
- Package plumbing helpers for numbering.xml, comments.xml, content types, and relationships
- Zero host dependencies: works in Node.js, browsers, Deno, and similar JS runtimes with DOM parsing support
- TypeScript declarations included via `index.d.ts`

## Install

### npm / Node.js

```bash
npm install @ansonlai/docx-redline-js
```

### CDN (browser `<script type="module">`)

```html
<script type="module">
  import { applyRedlineToOxml } from 'https://esm.sh/@ansonlai/docx-redline-js';
</script>
```

Or use the pre-bundled file (no import map needed, `diff-match-patch` is inlined):

```html
<script type="module">
  import { applyRedlineToOxml } from 'https://cdn.jsdelivr.net/npm/@ansonlai/docx-redline-js/dist/docx-redline-js.esm.min.js';
</script>
```

### Local git clone

```bash
git clone https://github.com/AnsonLai/docx-redline-js.git
```

```js
import { applyRedlineToOxml } from './docx-redline-js/index.js';
```

## Quick Start

### Node.js

```js
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import {
  configureXmlProvider,
  setDefaultAuthor,
  applyRedlineToOxml
} from '@ansonlai/docx-redline-js';

configureXmlProvider({ DOMParser, XMLSerializer });
setDefaultAuthor('My App');

const result = await applyRedlineToOxml(
  paragraphOoxml,
  'Original sentence.',
  'Updated sentence.',
  { generateRedlines: true, author: 'Editor' }
);

console.log(result.hasChanges);
console.log(result.oxml);
```

### Browser

```js
import {
  setDefaultAuthor,
  applyRedlineToOxml
} from '@ansonlai/docx-redline-js';

setDefaultAuthor('Browser Editor');

const result = await applyRedlineToOxml(oxml, original, modified, {
  generateRedlines: true
});
```

## API Reference

### Agent-friendly inspection and complete DOCX editing

```js
import { inspectDocumentParts } from '@ansonlai/docx-redline-js';
const inventory = inspectDocumentParts({ documentXml, commentsXml, numberingXml });
```

Inspection returns exact paragraph text, target IDs/fingerprints, headings,
table/list context, revision authors, and joined comment anchors. Filters such
as `search`, `indexes`, `range`, `revisedOnly`, `inTable`, and `skipEmpty`
limit output. `revisionView` accepts `accepted`, `rejected`, or `current`.

For complete `.docx` buffers in Node:

```js
import { openDocx } from '@ansonlai/docx-redline-js/node';
const document = openDocx(inputBuffer);
const result = await document.applyOperations(operations, {
  author: 'Editor', atomic: true, validate: true
});
const outputBuffer = result.toBuffer();
```

The Node facade performs edits, artifact merges, package wiring, validation,
and commit as one transaction. It defaults to strict targets and returns the
untouched input with `written: false` on atomic failure. It is isolated from
the root/browser dependency graph.

Install `@xmldom/xmldom` alongside the package when using the Node facade or
CLI; it remains an optional peer so browser consumers do not install a DOM shim.

### Agent CLI

```bash
docx-redline extract contract.docx --range 10:30
docx-redline preflight contract.docx --operations operations.json --author "Editor"
docx-redline apply contract.docx --operations operations.json --author "Editor" --output reviewed.docx
docx-redline validate reviewed.docx
```

Paragraph indexes are 1-based. Use `--index 12` for one paragraph,
`--indexes 2,5,8` for a set, or `--range 10:30` for an inclusive range.
Malformed filters and unknown options return an error instead of silently
falling back to an unfiltered extraction.

All commands emit JSON. Mutations require explicit authors and preserve the
input unless `--in-place` is supplied. Existing output paths are refused unless
`--force` is supplied. See [the agent workflow](docs/AGENT-WORKFLOW.md) and the
[operation JSON Schema](docs/schemas/document-operations.schema.json).

### Configuration (call once at startup)

| Function | Purpose |
|----------|---------|
| `configureXmlProvider({ DOMParser, XMLSerializer })` | Inject XML parser. Required in Node.js; browsers usually provide native support. |
| `configureLogger({ log, warn, error })` | Replace default console logger. |
| `setDefaultAuthor(name)` | Set fallback track-change author (default: `'Author'`). |
| `setPlatform(label)` | Set platform label for diagnostics (default: `'Unknown'`). |

### Engine (primary reconciliation APIs)

| Function | Purpose |
|----------|---------|
| `applyRedlineToOxml(oxml, original, modified, options)` | Core engine entry point for text/markdown reconciliation with optional redlines. |
| `applyRedlineToOxmlWithListFallback(oxml, original, modified, options)` | Core engine with automatic single-line list structural fallback. |
| `reconcileMarkdownTableOoxml(oxml, original, markdownTable, options)` | Table-specific reconciliation helper. |

Common `applyRedlineToOxml` options:

| Option | Purpose |
|--------|---------|
| `generateRedlines` | When `true`, emit Word-native tracked changes; when `false`, apply clean text changes. |
| `author` | Track-change author used for generated revisions. |
| `existingRevisions` | Existing-revision policy. `'reject-input'` is the default. `'accept-all-first'` normalizes before a real edit but returns the untouched input on no-op. `'accept-all-first-keep-normalized'` explicitly returns accepted revisions as a change even on no-op. |
| `removeFormatting` | When `true` and the text is unchanged with no Markdown hints, explicitly remove existing bold/italic/underline/strikethrough formatting. Defaults to `false`. |
| `sanitizeInput` | Opt-in removal of a standalone leading assistant-preface line. Defaults to `false`; dollar-delimited text and literal `\\n` sequences are always preserved. |

Common result fields:

| Field | Purpose |
|-------|---------|
| `status` | Optional non-breaking status: `'ok'`, `'no-op'`, or `'error'`. |
| `error` | Present when `status === 'error'`; includes a stable `code` such as `PARSE_ERROR`, `TARGET_NOT_FOUND`, `PARTIAL_TARGET`, `EXISTING_REVISIONS`, `DIFF_TOKEN_LIMIT`, or `BATCH_OPERATION_FAILED`. |

Word diffs are deterministic by default (no wall-clock timeout). Inputs above
the safe ceiling of 262,144 unique diff tokens return `DIFF_TOKEN_LIMIT` with
the original OOXML unchanged so callers can split the operation without risking
silent text loss.

### Pipeline (lower-level access)

| Function | Purpose |
|----------|---------|
| `ReconciliationPipeline` | Direct pipeline access (ingest, diff, patch, serialize). |
| `ingestWordOoxmlToPlainText(oxml)` | Extract plain text from OOXML. |
| `ingestWordOoxmlToMarkdown(oxml)` | Convert OOXML to markdown. |
| `ingestWordOoxmlToPlainTextResult(oxml)` | Extract text as `{ text, status, error?, warnings? }`, distinguishing malformed input from an empty document. |
| `ingestWordOoxmlToMarkdownResult(oxml)` | Markdown counterpart to the result-returning plain-text helper. |
| `ingestOoxml(oxml)` | Flatten OOXML into an internal run model with offsets. |
| `preprocessMarkdown(text)` | Normalize markdown and extract format hints. |
| `containsTrackedChanges(xmlDoc)` | Detect `w:ins`, `w:del`, move revisions, property changes, and paragraph-mark revision markup in a parsed OOXML document/fragment. |
| `validateRedlineOoxml(oxml)` | Validate generated redline OOXML against the package's structural invariants (no nested revisions, `w:delText` inside `w:del`, complete metadata, unique revision ids, preserved boundary whitespace). Returns `{ valid, issues }`; run it before writing output into a package. |

### Services

| Function | Purpose |
|----------|---------|
| `injectCommentsIntoOoxml(oxml, comments, options)` | Add comments anchored to text ranges. |
| `acceptTrackedChangesInOoxml(oxml, { author?, allAuthors? })` | Accept `w:ins` / `w:del` / `w:moveFrom` / `w:moveTo` / `*PrChange` revisions for one author or all authors. |
| `rejectTrackedChangesInOoxml(oxml, { author?, allAuthors? })` | Reject `w:ins` / `w:del` / `w:moveFrom` / `w:moveTo` / `*PrChange` revisions for one author or all authors. |
| `deleteCommentsByAuthorInOoxml(oxml, { author?, allAuthors? })` | Delete matching comment definitions and anchors present in the supplied OOXML payload. Real `.docx` packages require updating both `word/comments.xml` and `word/document.xml`. |
| `generateTableOoxml(headers, rows, options)` | Generate a `w:tbl` from tabular data. |
| `createDynamicNumberingIdState(numberingXml)` | Allocate numbering IDs without collisions. |
| `ensureNumberingArtifactsInZip(zip, numberingXml, options)` | Add numbering artifacts to a `.docx` package. Replacement of existing numbering without `mergeNumberingXmlBySchemaOrder` is deprecated and will throw in the next major version. |
| `ensureCommentsArtifactsInZip(zip, commentsXml)` | Merge comments artifacts into a `.docx` package. |
| `validateDocxPackage(zip)` | Validate `.docx` structural consistency. |

Malformed OOXML never escapes these public transform APIs as a raw parser
exception. Transforms return `status: 'error'` with `error.code === 'PARSE_ERROR'`;
validators return a `PARSE_ERROR` issue. Recoverable XML parser
diagnostics are forwarded through the configured logger and included in
`warnings` where the result shape supports them.

### Deep Imports

For advanced usage, import specific submodules:

```js
import {
  applyOperationToDocumentXml,
  applyOperationsToDocumentXml,
  preflightOperations,
  orderOperationsForStableTargets
} from '@ansonlai/docx-redline-js/standalone-runner';
import { getParagraphText } from '@ansonlai/docx-redline-js/core/paragraph-targeting.js';
```

Use `applyOperationsToDocumentXml(...)` for mixed batches. It stably runs comments before text-changing operations so replacements cannot invalidate their original anchors. Other operation types retain their relative order. Batch results retain each operation's original 1-based index and expose the actual `executionOrder`.

The batch runner keeps one live document DOM and performs one final full-document
serialization. Accuracy remains the controlling constraint: each operation has
an internal savepoint so an error or no-op cannot leak a partial edit or consumed
revision ID into later operations.

Batches are atomic by default: any operation error returns the original
`documentXml`, `hasChanges: false`, empty package artifacts, and
`rolledBack: true`. The default `continueOnError: true` still attempts the full
batch so `results` describes what would have applied. Callers that intentionally
consume partial results must pass `{ atomic: false }`; use
`{ continueOnError: false }` to stop after the first error.

Comment anchors use exact matching first, then a unique ASCII-space/NBSP
equivalent match that preserves source offsets and text. Missing anchors return
`ANCHOR_NOT_FOUND`; repeated matches return `AMBIGUOUS_ANCHOR`. Both are
operation errors and therefore roll back atomic batches. When `textToComment`
is omitted, the exact text of the resolved paragraph is used.

Operations may override the batch author and may use a strict target descriptor:

```js
const operations = [{
  type: 'replace',
  author: 'Contract Editor',
  target: {
    exactText: 'Either party may terminate on notice.',
    paragraphId: '1A2B3C4D',
    index: 12,
    fingerprint: 'fnv1a32:...'
  },
  modified: 'Either party may terminate on 30 days written notice.'
}];

const preflight = preflightOperations(documentXml, operations, 'Fallback Author');
if (!preflight.valid) {
  // Resolve missing/ambiguous targets, anchors, revision policies, or conflicts.
}

const result = await applyOperationsToDocumentXml(
  documentXml,
  operations,
  'Fallback Author',
  null,
  { strictTargets: true }
);
```

Preflight is read-only and uses strict targeting by default. It reports
`AMBIGUOUS_TARGET` with candidates instead of selecting the first duplicate,
does not use fuzzy fallback, checks comment/highlight anchors and existing
revision policy, identifies same-paragraph operation conflicts, and reports
authors plus required comments/numbering artifacts. Application remains
permissive by default for compatibility; pass `strictTargets: true` for the same
strict target behavior.

### Output Shape Matrix

Different APIs return different OOXML shapes. Use this as a packaging safety check.

| API | Typical input scope | Output field | Possible root/output shape | Safe to write directly into `word/document.xml` |
|-----|----------------------|--------------|----------------------------|--------------------------------------------------|
| `applyRedlineToOxml(...)` | Paragraph, range, or table-scope OOXML | `result.oxml` | Fragment, `<w:document>`, or package payload (`<pkg:package>`) | No. Inspect first. |
| `applyRedlineToOxmlWithListFallback(...)` | Paragraph or range-scope OOXML | `result.oxml` | Fragment, `<w:document>`, or package payload (`<pkg:package>`) | No. Inspect first. |
| `reconcileMarkdownTableOoxml(...)` | Table or paragraph-scope OOXML | `result.oxml` | Same shapes as `applyRedlineToOxml(...)` for the supplied scope | No. Inspect first. |
| `applyOperationToDocumentXml(...)` | Full `word/document.xml` string | `result.documentXml` | `<w:document>` | Yes. This is the document-safe helper. |
| `applyOperationsToDocumentXml(...)` | Full `word/document.xml` plus an operation batch | `result.documentXml` | `<w:document>` | Yes. Atomic by default; comments are applied before text-changing operations. |
| `extractReplacementNodesFromOoxml(...)` | Any OOXML payload | `{ replacementNodes, numberingXml, sourceType }` | Normalized to `fragment`, `document`, or `package` | Yes. Use this when consuming `result.oxml`. |

### Do / Don't for Packaging

- Do use `applyOperationToDocumentXml(...).documentXml` when your intent is to replace `word/document.xml`.
- Do use `applyOperationsToDocumentXml(...)` rather than an unsorted loop for batches containing comments and replacements that target the same original paragraph.
- Redline application strips proofing markers (`w:proofErr`) from the matched target paragraph before diffing, while preserving complex-field scaffolding (`w:fldChar`, `w:instrText`) and its cached visible result as inert structure. Adjacent edits do not revise or move an unchanged field result.
- Hyperlinks, bookmarks, comment range markers, tabs/breaks, and footnote/endnote references are treated as structural OOXML that should survive adjacent redline edits instead of being orphaned or wrapped in deletions.
- Do use `extractReplacementNodesFromOoxml(...)` when you are consuming `result.oxml` from paragraph/range/table APIs.
- Do merge numbering/comments artifacts with `ensureNumberingArtifactsInZip(...)` and `ensureCommentsArtifactsInZip(...)` when those parts are present. Supply `mergeNumberingXmlBySchemaOrder` when numbering already exists.
- Don't write payloads that start with `<pkg:package` directly into `word/document.xml`.
- Don't assume every `result.oxml` payload is a raw paragraph fragment.

## Working With `.docx` Files

This package operates on OOXML strings (XML parts inside `.docx` zip archives), not raw `.docx` binaries.

Typical flow:

1. Extract the `.docx` zip (for example with JSZip, fflate, or similar)
2. Read `word/document.xml`
3. Apply reconciliation APIs to XML strings
4. Merge numbering/comments artifacts when needed
5. Write the archive back to a `.docx` file

```js
import JSZip from 'jszip';
import {
  applyRedlineToOxml,
  extractReplacementNodesFromOoxml,
  ensureNumberingArtifactsInZip,
  mergeNumberingXmlBySchemaOrder,
  validateDocxPackage
} from '@ansonlai/docx-redline-js';
import { applyOperationToDocumentXml } from '@ansonlai/docx-redline-js/standalone-runner';

const zip = await JSZip.loadAsync(docxBuffer);
const documentXml = await zip.file('word/document.xml').async('string');

const opResult = await applyOperationToDocumentXml(
  documentXml,
  { type: 'redline', target: 'old text', modified: 'new text' },
  'Editor'
);

// applyOperationToDocumentXml(...) returns a full w:document payload.
zip.file('word/document.xml', opResult.documentXml);

const fragmentResult = await applyRedlineToOxml(
  paragraphOoxml,
  'Item text',
  '1. Item text',
  { generateRedlines: true, author: 'Editor' }
);
const normalized = extractReplacementNodesFromOoxml(fragmentResult.oxml);

// If sourceType === 'package', merge extracted content/artifacts instead of
// writing the raw pkg:package payload into word/document.xml.
if (normalized.numberingXml) {
  await ensureNumberingArtifactsInZip(zip, normalized.numberingXml, {
    mergeNumberingXml: mergeNumberingXmlBySchemaOrder
  });
}

await validateDocxPackage(zip);
const output = await zip.generateAsync({ type: 'nodebuffer' });
```

## Validating Output

For the test-lane design and instructions for adding regression, synthetic
Word, and real-corpus cases, see [docs/TESTING.md](./docs/TESTING.md).

Run the automated package checks:

```bash
npm test
npm run test:isolation
npm run check:types
npm run lint
npm run test:coverage
```

For release-time fixture export:

```bash
node scripts/export-validation-fixtures.mjs
```

On Windows with desktop Word installed, you can smoke-test a completed `.docx`:

```bash
npm run smoke:word -- path/to/file.docx
```

To validate against Word as an independent oracle (Word itself accepts and
rejects the generated revisions and the resulting text is compared to the
expected outcomes):

```bash
npm run test:word
```

This Windows-only test command generates an English legal/administrative task
suite under `tmp/word-validation/` and drives installed desktop Microsoft Word
through COM. Its 33 cases include targeted reliability checks for literal
content, multi-paragraph replacement, prior-revision no-op, atomic rollback,
hostile revision IDs, bookmarks, internal hyperlinks, mixed formatted runs,
content controls, table cells, structural tabs, locked complex fields,
comments, footnotes/endnotes, headers/footers, and external hyperlinks.
Structure-focused cases also assert required
OOXML elements before Word independently checks Accept All and Reject All. The
published library remains clean, host-independent JavaScript; Word automation
exists only in development scripts.

Use `npm run report:word:coverage` to print the validated task-by-structure
matrix across all 33 synthetic and 31 SuperDoc scenarios. Before a release,
`npm run review:word:prepare -- --cycle=0` creates a pending human-review
manifest with changed cases, a rotating 20% synthetic sample, and legal plus
administrative corpus representatives. See [docs/TESTING.md](./docs/TESTING.md)
and [docs/WORD-MANUAL-REVIEW.md](./docs/WORD-MANUAL-REVIEW.md); preparation and
AI preflight never count as human sign-off.

A nightly GitHub Actions workflow additionally validates generated fixtures
against the ECMA-376 transitional schemas (`xmllint`), opens them with
LibreOffice, and runs an extended fuzz sweep of the accept/reject round-trip
invariant with a fresh seed. See [docs/VALIDATION.md](./docs/VALIDATION.md).

## Architecture

See [ARCHITECTURE.md](./ARCHITECTURE.md) for module layout, data flow, and contributor guidance.

See [AGENTS.md](./AGENTS.md) for a concise reference for AI coding agents.

See [docs/VALIDATION.md](./docs/VALIDATION.md) for release-time validation steps.

See [docs/TESTING.md](./docs/TESTING.md) for how the test lanes work and how to
add new cases.

## Test Corpus Attribution

Real-document reliability testing uses selected references from
[docx-corpus](https://docxcorp.us/), built by
[SuperDoc](https://superdoc.dev/). The dataset is licensed under the
[Open Data Commons Attribution License (ODC-By) 1.0](https://opendatacommons.org/licenses/by/1-0/).

Only explicitly pinned English legal and administrative documents are eligible
for the initial corpus lane. References and provenance live in
`tests/corpus/superdoc-english-legal-administrative.json`; downloaded documents
are hash-verified and kept in ignored `tmp/` storage rather than committed. On
Windows with desktop Word installed, run the reviewed 31-scenario/23-document lane with:

```bash
npm run test:corpus:word
```

ODC-By applies to the database; individual documents may carry additional
rights, so each selected document must be reviewed before becoming a test case.
