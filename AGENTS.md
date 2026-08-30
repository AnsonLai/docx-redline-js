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
  author: 'Agent Name',
  existingRevisions: 'reject-input'
});
```

`existingRevisions` defaults to `'reject-input'`. Use `'accept-all-first'` only
when the caller intentionally wants to accept prior tracked changes before
applying a new edit. A no-op still returns the untouched input; use
`'accept-all-first-keep-normalized'` only when accepted revisions should be
returned as a real change even without a new redline.

### Apply a text edit without tracked changes

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
  { text: 'Review this clause', targetText: 'force majeure', author: 'Agent' }
]);
```

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
} from '@ansonlai/docx-redline-js/services/standalone-operation-runner.js';

const result = await applyOperationsToDocumentXml(documentXml, operations, 'Agent', runtimeContext, options);
```

Use `result.documentXml` from these APIs when replacing full `word/document.xml`.
For mixed batches, prefer `applyOperationsToDocumentXml(...)`; it applies comments
before replacements so earlier edits cannot invalidate their anchors.

Batches are atomic by default. If any operation fails, the batch returns the
original `documentXml`, `hasChanges: false`, no comment/numbering artifacts, and
`rolledBack: true`; `results` still describes every attempted operation because
`continueOnError` defaults to `true`. Pass `{ atomic: false }` only when a
partially applied document is intentional. Pass `{ continueOnError: false }` to
stop attempting operations after the first error.

### Detect existing tracked changes

```js
import { containsTrackedChanges } from '@ansonlai/docx-redline-js';
const hasTrackedChanges = containsTrackedChanges(xmlDoc);
```

### Convert paragraph text into a Word list

```js
const result = await applyRedlineToOxml(oxml, 'Item text', '1. Item text', {
  generateRedlines: true
});
```

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
  standalone-operation-runner.js
  standalone-docx-plumbing.js
  numbering-helpers.js
  comment-engine.js
  revision-comment-management.js
  table-reconciliation.js
  package-builder.js
orchestration/
  route-plan.js
  list-markdown.js
  list-structural-fallback.js
```

## Common Patterns

### Options shape

```js
{
  generateRedlines: true,
  author: 'Name',
  existingRevisions: 'reject-input',
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

Known error codes include `PARSE_ERROR`, `TARGET_NOT_FOUND`,
`EXISTING_REVISIONS`, and `DIFF_TOKEN_LIMIT`.

For ingestion that must distinguish an empty document from malformed OOXML,
use `ingestWordOoxmlToPlainTextResult` or
`ingestWordOoxmlToMarkdownResult`. The legacy ingestion helpers intentionally
retain their string-only return type and return `''` for parse failures.

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
4. List operations may return `numberingXml` that must be merged into package parts.
5. `useNativeApi: true` means standalone mode cannot fully handle that operation path.
6. `deleteCommentsByAuthorInOoxml` removes matching `comments.xml` entries and linked comment anchors/references in the document.
7. If output begins with `<pkg:package`, treat it as package-level OOXML and normalize it before writing anything back to `word/document.xml`.
8. Existing revisions are rejected by default; `accept-all-first` preserves the original OOXML on no-op, while `accept-all-first-keep-normalized` explicitly returns normalization as a change.
9. Caller content is not sanitized by default. Pass `sanitizeInput: true` only for raw assistant output; literal dollar delimiters and `\\n` sequences are never rewritten.
10. Hyperlinks, bookmarks, comment markers, tabs/breaks, and footnote/endnote references are structural OOXML and should survive adjacent redline edits.
11. Internally, create Word elements through `createWordElement` and tracked-change metadata through `createRevisionMetadata`.

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
