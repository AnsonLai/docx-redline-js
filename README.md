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
| `existingRevisions` | Policy for source OOXML that already contains tracked changes: `'reject-input'` (default) returns `status: 'error'` with code `EXISTING_REVISIONS`; `'accept-all-first'` accepts existing revisions before applying the new edit. |
| `removeFormatting` | When `true` and the text is unchanged with no Markdown hints, explicitly remove existing bold/italic/underline/strikethrough formatting. Defaults to `false`. |

Common result fields:

| Field | Purpose |
|-------|---------|
| `status` | Optional non-breaking status: `'ok'`, `'no-op'`, or `'error'`. |
| `error` | Present when `status === 'error'`; includes a stable `code` such as `PARSE_ERROR`, `TARGET_NOT_FOUND`, or `EXISTING_REVISIONS`. |

### Pipeline (lower-level access)

| Function | Purpose |
|----------|---------|
| `ReconciliationPipeline` | Direct pipeline access (ingest, diff, patch, serialize). |
| `ingestWordOoxmlToPlainText(oxml)` | Extract plain text from OOXML. |
| `ingestWordOoxmlToMarkdown(oxml)` | Convert OOXML to markdown. |
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
| `deleteCommentsByAuthorInOoxml(oxml, { author?, allAuthors? })` | Delete comments and matching anchors/references for one author or all authors. |
| `generateTableOoxml(headers, rows, options)` | Generate a `w:tbl` from tabular data. |
| `createDynamicNumberingIdState(numberingXml)` | Allocate numbering IDs without collisions. |
| `ensureNumberingArtifactsInZip(zip, numberingXml)` | Merge numbering artifacts into a `.docx` package. |
| `ensureCommentsArtifactsInZip(zip, commentsXml)` | Merge comments artifacts into a `.docx` package. |
| `validateDocxPackage(zip)` | Validate `.docx` structural consistency. |

### Deep Imports

For advanced usage, import specific submodules:

```js
import {
  applyOperationToDocumentXml,
  applyOperationsToDocumentXml,
  orderOperationsForStableTargets
} from '@ansonlai/docx-redline-js/services/standalone-operation-runner.js';
import { getParagraphText } from '@ansonlai/docx-redline-js/core/paragraph-targeting.js';
```

Use `applyOperationsToDocumentXml(...)` for mixed batches. It stably runs comments before text-changing operations so replacements cannot invalidate their original anchors. Other operation types retain their relative order. Batch results retain each operation's original 1-based index and expose the actual `executionOrder`.

### Output Shape Matrix

Different APIs return different OOXML shapes. Use this as a packaging safety check.

| API | Typical input scope | Output field | Possible root/output shape | Safe to write directly into `word/document.xml` |
|-----|----------------------|--------------|----------------------------|--------------------------------------------------|
| `applyRedlineToOxml(...)` | Paragraph, range, or table-scope OOXML | `result.oxml` | Fragment, `<w:document>`, or package payload (`<pkg:package>`) | No. Inspect first. |
| `applyRedlineToOxmlWithListFallback(...)` | Paragraph or range-scope OOXML | `result.oxml` | Fragment, `<w:document>`, or package payload (`<pkg:package>`) | No. Inspect first. |
| `reconcileMarkdownTableOoxml(...)` | Table or paragraph-scope OOXML | `result.oxml` | Same shapes as `applyRedlineToOxml(...)` for the supplied scope | No. Inspect first. |
| `applyOperationToDocumentXml(...)` | Full `word/document.xml` string | `result.documentXml` | `<w:document>` | Yes. This is the document-safe helper. |
| `applyOperationsToDocumentXml(...)` | Full `word/document.xml` plus an operation batch | `result.documentXml` | `<w:document>` | Yes. Comments are applied before text-changing operations. |
| `extractReplacementNodesFromOoxml(...)` | Any OOXML payload | `{ replacementNodes, numberingXml, sourceType }` | Normalized to `fragment`, `document`, or `package` | Yes. Use this when consuming `result.oxml`. |

### Do / Don't for Packaging

- Do use `applyOperationToDocumentXml(...).documentXml` when your intent is to replace `word/document.xml`.
- Do use `applyOperationsToDocumentXml(...)` rather than an unsorted loop for batches containing comments and replacements that target the same original paragraph.
- Redline application now strips non-visible field scaffolding (`w:fldChar`, `w:instrText`) and proofing markers (`w:proofErr`) from the matched target paragraph before diffing, while preserving the visible field result text. This avoids a class of Word-open failures caused by tracked changes spanning hidden field instruction runs.
- Hyperlinks, bookmarks, comment range markers, tabs/breaks, and footnote/endnote references are treated as structural OOXML that should survive adjacent redline edits instead of being orphaned or wrapped in deletions.
- Do use `extractReplacementNodesFromOoxml(...)` when you are consuming `result.oxml` from paragraph/range/table APIs.
- Do merge numbering/comments artifacts with `ensureNumberingArtifactsInZip(...)` and `ensureCommentsArtifactsInZip(...)` when those parts are present.
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
  validateDocxPackage
} from '@ansonlai/docx-redline-js';
import { applyOperationToDocumentXml } from '@ansonlai/docx-redline-js/services/standalone-operation-runner.js';

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
  await ensureNumberingArtifactsInZip(zip, normalized.numberingXml);
}

await validateDocxPackage(zip);
const output = await zip.generateAsync({ type: 'nodebuffer' });
```

## Validating Output

Run the automated package checks:

```bash
npm test
npm run test:isolation
npm run check:types
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
node scripts/export-validation-fixtures.mjs
npm run smoke:word:diff
```

A nightly GitHub Actions workflow additionally validates generated fixtures
against the ECMA-376 transitional schemas (`xmllint`), opens them with
LibreOffice, and runs an extended fuzz sweep of the accept/reject round-trip
invariant with a fresh seed. See [docs/VALIDATION.md](./docs/VALIDATION.md).

## Architecture

See [ARCHITECTURE.md](./ARCHITECTURE.md) for module layout, data flow, and contributor guidance.

See [AGENTS.md](./AGENTS.md) for a concise reference for AI coding agents.

See [docs/VALIDATION.md](./docs/VALIDATION.md) for release-time validation steps.
