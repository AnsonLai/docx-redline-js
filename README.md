# @gsd/docx-reconciliation

Host-independent OOXML reconciliation engine for `.docx` manipulation with track changes (redlines).

Converts AI-generated or programmatic text/markdown edits into valid Office Open XML (OOXML) with `w:ins`/`w:del` revision markup that Microsoft Word renders as native tracked changes.

## Features

- Text reconciliation with word-level diffing and native-looking redlines
- Formatting updates (bold, italic, underline, strikethrough) via surgical `w:rPrChange`
- Lists: generate and edit real Word lists (`w:numPr`) from markdown
- Tables: virtual-grid diffing for cell-level edits with merge safety
- Comments: inject OOXML comments anchored to text ranges
- Highlights: apply highlight colors to runs
- Markdown and OOXML conversion in both directions
- Package plumbing helpers for numbering.xml, comments.xml, content types, and relationships
- Zero host dependencies: works in Node.js, browsers, Deno, and similar JS runtimes with DOM parsing support

## Install

### npm / Node.js

```bash
npm install @gsd/docx-reconciliation
```

### CDN (browser `<script type="module">`)

```html
<script type="module">
  import { applyRedlineToOxml } from 'https://esm.sh/@gsd/docx-reconciliation';
</script>
```

Or use the pre-bundled file (no import map needed, `diff-match-patch` is inlined):

```html
<script type="module">
  import { applyRedlineToOxml } from 'https://cdn.jsdelivr.net/npm/@gsd/docx-reconciliation/dist/docx-reconciliation.esm.min.js';
</script>
```

### Local git clone

```bash
git clone https://github.com/YOUR_ORG/docx-reconciliation.git
```

```js
import { applyRedlineToOxml } from './docx-reconciliation/index.js';
```

## Quick Start

### Node.js

```js
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import {
  configureXmlProvider,
  setDefaultAuthor,
  applyRedlineToOxml
} from '@gsd/docx-reconciliation';

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
} from '@gsd/docx-reconciliation';

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

### Pipeline (lower-level access)

| Function | Purpose |
|----------|---------|
| `ReconciliationPipeline` | Direct pipeline access (ingest, diff, patch, serialize). |
| `ingestWordOoxmlToPlainText(oxml)` | Extract plain text from OOXML. |
| `ingestWordOoxmlToMarkdown(oxml)` | Convert OOXML to markdown. |
| `ingestOoxml(oxml)` | Flatten OOXML into an internal run model with offsets. |
| `preprocessMarkdown(text)` | Normalize markdown and extract format hints. |

### Services

| Function | Purpose |
|----------|---------|
| `injectCommentsIntoOoxml(oxml, comments, options)` | Add comments anchored to text ranges. |
| `generateTableOoxml(headers, rows, options)` | Generate a `w:tbl` from tabular data. |
| `createDynamicNumberingIdState(numberingXml)` | Allocate numbering IDs without collisions. |
| `ensureNumberingArtifactsInZip(zip, numberingXml)` | Merge numbering artifacts into a `.docx` package. |
| `ensureCommentsArtifactsInZip(zip, commentsXml)` | Merge comments artifacts into a `.docx` package. |
| `validateDocxPackage(zip)` | Validate `.docx` structural consistency. |

### Deep Imports

For advanced usage, import specific submodules:

```js
import { applyOperationToDocumentXml } from '@gsd/docx-reconciliation/services/standalone-operation-runner.js';
import { getParagraphText } from '@gsd/docx-reconciliation/core/paragraph-targeting.js';
```

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
  configureXmlProvider,
  applyRedlineToOxml,
  ensureNumberingArtifactsInZip,
  validateDocxPackage
} from '@gsd/docx-reconciliation';

const zip = await JSZip.loadAsync(docxBuffer);
const documentXml = await zip.file('word/document.xml').async('string');

// Apply edits with applyRedlineToOxml(...)
// Merge artifacts with ensureNumberingArtifactsInZip(...) as needed

const output = await zip.generateAsync({ type: 'nodebuffer' });
```

## Architecture

See [ARCHITECTURE.md](./ARCHITECTURE.md) for module layout, data flow, and contributor guidance.

See [AGENTS.md](./AGENTS.md) for a concise reference for AI coding agents.
