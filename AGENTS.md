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
Output: { oxml: string, hasChanges: boolean, warnings?: string[] }
```

The engine works at paragraph scope. For full-document operations, callers iterate paragraph targets or use the standalone operation runner.

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

### Apply multiple operations to full document XML

```js
import { applyOperationToDocumentXml } from '@ansonlai/docx-redline-js/services/standalone-operation-runner.js';
const result = await applyOperationToDocumentXml(documentXml, operation, options);
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
  paragraph-targeting.js
  list-targeting.js
  table-targeting.js
engine/
  oxml-engine.js
  surgical-mode.js
  reconstruction-mode.js
  format-application.js
  formatting-removal.js
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
  author: 'Name'
}
```

### Typical return shape

```js
{
  oxml: string,
  hasChanges: boolean,
  warnings?: string[],
  numberingXml?: string,
  useNativeApi?: boolean
}
```

### OOXML wrapping for Word insertOoxml scenarios

```js
import { wrapInDocumentFragment } from '@ansonlai/docx-redline-js';
const wrapped = wrapInDocumentFragment(rawOoxml, { includeNumbering: true, numberingXml });
```

### Output shape guardrail (important for full-document packaging)

When using full-document workflows, do not assume the returned payload can always be
written directly to `word/document.xml`.

- Some operation paths can return package payload (`<pkg:package ...>`).
- Use `extractReplacementNodesFromOoxml(payload)` to normalize and detect source type.
- Treat `sourceType === 'package'` as package-level output; do not write it into
  `word/document.xml` as-is.

## Gotchas

1. Call `configureXmlProvider` first in Node.js.
2. `applyRedlineToOxml` is async.
3. Paragraph APIs expect paragraph-level OOXML, not full `word/document.xml` in all cases.
4. List operations may return `numberingXml` that must be merged into package parts.
5. `useNativeApi: true` means standalone mode cannot fully handle that operation path.
6. If output begins with `<pkg:package`, it is package-level OOXML, not a direct replacement for `word/document.xml`.
