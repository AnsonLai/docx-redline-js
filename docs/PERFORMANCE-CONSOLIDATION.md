# Performance Consolidation Reference

This reference records the accuracy boundaries used by Phases 3–5 of the
performance and complexity plan. Optimizations may share metadata, grammar, or
route implementation only when the observable text and OOXML contracts remain
the same.

## Paragraph-text walker inventory

`core/paragraph-text.js` is the only canonical visible-text definition.

| Walker | Category | Purpose and parity rule |
| --- | --- | --- |
| `extractCanonicalParagraphText` / `readCanonicalRunText` | Canonical | Accepted, current, and rejected visible text for targeting and inspection. |
| `pipeline/ingestion-paragraph.js` | Specialized mapping | Builds run models and offsets. Its accepted visible projection matches canonical tabs, breaks, hyperlinks, soft hyphens, and non-breaking hyphens. Deleted content remains a zero-width model entry. |
| `engine/surgical-spans.js` | Specialized mapping | Retains run ownership and offsets for edits. Its supported visible projection matches canonical text; existing revisions are rejected or normalized before this route. |
| `engine/format-extraction.js` and `format-paragraph-targeting.js` | Specialized formatting map | Associates visible characters with runs and formatting spans. It is not a replacement canonical extractor. |
| `engine/reconstruction-mapper.js` | Specialized structural map | Uses sentinels for fields, drawings, content controls, comments, notes, and breaks so structures survive reconstruction. Sentinel text is intentionally not canonical visible text. |
| `services/comment-locator.js` | Specialized anchor map | Maps exact visible characters back to individual text nodes for comment range insertion. |
| `engine/table-cell-context.js` | Canonical consumer | Uses `extractCanonicalParagraphText` for target matching; it does not maintain a second text definition. |
| `services/document-inspection.js` | Canonical consumer | Uses canonical paragraph and run text for externally visible inventory. |

Parity coverage lives in
`tests/performance_phase4_list_and_text_parity_tests.mjs`, with broader
revision, field, note, comment, and structural preservation coverage in the
canonical text, structural field/tab, engine reliability, move revision, and
round-trip suites.

## Shared list vocabulary

`pipeline/list-markers.js` owns marker grammar, marker classification,
numbering-style inference, outline-depth parsing, and the common parsed
list-item representation. Pipeline parsing, orchestration parsing, targeting,
and fallback consume these helpers. Numbering allocation, target selection,
OOXML writing, and package merging remain separate responsibilities.

The shared marker vocabulary is:

- marker type: `bullet` or `numbered`;
- numbering style: `bullet`, `decimal`, `lowerAlpha`, `upperAlpha`,
  `lowerRoman`, or `upperRoman`;
- level: indentation-derived level from 0 through 8;
- outline level: optional depth encoded by composite decimal markers.

## Reconciliation capability matrix

The executable internal record is
`engine/route-selection.js::RECONCILIATION_CAPABILITY_MATRIX`.

| Route | Primary capability | Accuracy boundary |
| --- | --- | --- |
| `formatOnly` | Formatting-only edits/removal | Retains run-scoped formatting and table-cell scoping. |
| `surgical` | Localized edits in table-bearing scopes | Preserves surrounding runs, hyperlinks, fields, comments, and notes. |
| `reconstruction` | General paragraph reconstruction | Preserves structural nodes through sentinels and reference maps. |
| `table` | Markdown table creation/reconciliation | Owns table structure; formatting behavior remains cell-dependent. |
| `listDirect` | One-source-paragraph expansion into a list | Uses focused list generation with numbering artifacts and established result wrapping. |
| `listCompatibilityPipeline` | Multi-paragraph marked-list edits | Retains run-aware patching and paragraph-boundary behavior until broader parity evidence permits migration. |

Route selection can be observed without changing result shapes by passing the
internal `_routeInstrumentation.onRoute` callback. Run `npm run profile:routes`
for the checked synthetic frequency set. This instrumentation is diagnostic;
it is not a public routing-policy switch.

## Compatibility window

`ReconciliationPipeline`, `serializeToOoxml`, and `wrapInDocumentFragment`
remain public and supported for the current major version. Internal callers may
use focused modules directly, but removal of these exports requires a future
major release, deprecation notice, and migration example.
