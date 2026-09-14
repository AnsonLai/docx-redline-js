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
- Localized exact replacements and fail-closed one-turn CLI edits with verified before/after evidence
- Package plumbing helpers for numbering.xml, comments.xml, content types, and relationships
- Zero host dependencies: works in Node.js, browsers, Deno, and similar JS runtimes with DOM parsing support

## What's New in 0.7.0

Version 0.7.0 adds a contract-8 fast path for exact mechanical edits. Agents can
send a small `find`/`replace` request instead of reproducing a complete legal
paragraph, and may omit a strong target when the source literal resolves to one
accepted-view paragraph. A longer, fairly unique nearby phrase can narrow the
attempt with a signed directional window. Generic or repeated anchors widen the
unioned scope and may cause a safe ambiguity failure and another tool turn.

Speculative edits fail without writing when the anchor, paragraph, or source
span is missing or ambiguous. Successful localized results include a committed
`change` object with the selected legal location, bounded before/after excerpts,
and actual accepted-view verification. Existing full-paragraph operations,
tracked-change semantics, validation, rollback, and review policies remain
available unchanged. See the [complete 0.7.0 release notes](./docs/releases/0.7.0.md).
## Documentation Index

| Document | Description |
|---|---|
| **[README.md](./README.md)** | Library overview, installation, quick start, and public API reference |
| **[docs/AGENT_FAST_START.md](./docs/AGENT_FAST_START.md)** | Compact ordinary-edit protocol for structured tools and shell-only agents |
| **[AGENTS.md](./AGENTS.md)** | Short repository launch card for task routing and contributor verification |
| **[docs/AGENT_KNOWLEDGE_BASE.md](./docs/AGENT_KNOWLEDGE_BASE.md)** | Full agent reference, CLI workflow, operation examples, error recovery, options, and gotchas |
| **[ARCHITECTURE.md](./ARCHITECTURE.md)** | Contributor architecture, module responsibilities, end-to-end data flow, and contracts |
| **[docs/TESTING.md](./docs/TESTING.md)** | Complete testing guide, test lanes, independent oracle validation, and Word visual review checklist |
| **[CHANGELOG.md](./CHANGELOG.md)** | Release history, breaking changes, and migration notes |

## Repository Layout

The package exposes three levels of API:

| Level | Entry point | Use it for |
|---|---|---|
| Host-independent OOXML API | `index.js` | Paragraph/range transforms and exported OOXML utilities in browsers, Node.js, or another DOM-capable runtime |
| Standalone document XML runner | `services/standalone-operation-runner.js` | Applying operations to a complete `word/document.xml` string |
| Node/DOCX API and CLI | `node/index.js`, `bin/docx-redline.js` | Reading, changing, validating, and writing complete `.docx` ZIP packages |

Implementation folders have distinct roles: `core/` holds shared OOXML and
targeting primitives; `pipeline/` handles ingestion, diffing, markdown, lists,
and serialization; `engine/` performs reconciliation; `services/` coordinates
document operations and package artifacts; `node/` contains Node-only ZIP and
whole-document code. Tests are directly runnable `tests/*.mjs` files, while
`tests/helpers/` and `tests/fixtures/` contain support code and data.

Contributors and coding agents should start with the routing table in
[AGENTS.md](./AGENTS.md#pick-the-route) before exploring the tree. The full
dependency and ownership map is in [ARCHITECTURE.md](./ARCHITECTURE.md).

For document-operation JSON, choose operations by the desired output structure,
not by the everyday meaning of the type name. `redline` and `replace` provide
ordinary text replacement; `list-change` and `table-reconciliation` provide
structural intent; and ordinary `insert` is a compatibility alias of the
redline path unless it includes a rejected-view target and anchor. In every
ordinary text-bearing operation, `modified` is the complete desired content for
the target. See the [operation model](./docs/AGENT_KNOWLEDGE_BASE.md#operation-model-choose-by-output-shape)
and the [JSON schema](./docs/schemas/document-operations.schema.json).

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
limit output. Search is a case-insensitive substring match; `around` adds nearby drafting
context, while `limit` and `after` page direct hits. When a search yields 0
matches in the active view, `selection.hint` provides guidance if matches exist
in the alternate revision view (such as searching for deleted text to restore).
Compact `extract --revised` output retains revision authors and marks paragraphs
whose content is wholly hidden in the selected view as `deleted` or `inserted`.
One such paragraph can be restored without restating it:
`docx-redline apply input.docx --restore --target-id ID --output reviewed.docx`.
Add one exact `--find`/`--replace` pair to make a localized correction while
restoring. The shortcut requires a strong rejected-view paragraph ID and does
not infer a range or adjacent content.
`revisionView` accepts `accepted`, `rejected`, or `current`. Returned
paragraphs identify matches versus context and include `humanReference` for
user-facing summaries.

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

#### Example agent session wrapper (development only)

[`examples/agent-session-wrapper.mjs`](./examples/agent-session-wrapper.mjs)
demonstrates how a custom agent harness can keep one `DocxDocument` open, return
short revision-bound target handles, apply safe defaults once, and translate a
narrow edit request into canonical document operations:

```js
import { readFile } from 'node:fs/promises';
import { createExampleAgentSession } from './examples/agent-session-wrapper.mjs';

const session = createExampleAgentSession(await readFile('contract.docx'), {
  profile: { author: 'Editor' }
});
const inspection = session.inspect({ search: 'termination', around: 2 });
const clause = inspection.targets.find(target => target.role === 'match');
const result = await session.applyEdits([{
  target: clause.handle,
  replacements: [{
    find: 'The Company may terminate',
    replace: 'Either party may terminate'
  }]
}]);
```

This file is a testable sample, not a package export or supported alternate
mutation engine. It delegates to `@ansonlai/docx-redline-js/node`, is excluded
from the published package files, and is intended to help MCP servers, Claude
skills, OpenCode tools, and other custom harnesses design thin integrations.
The sample binds each handle to the inspected package revision and view, returns
new handles after successful mutations, and expands localized exact replacements
into complete desired paragraph text before delegating to the canonical redline
operation. Duplicate matches require an explicit `occurrence`; missing,
ambiguous, overlapping, and conflicting patches fail before document mutation.
From a source checkout, run `npm run benchmark:agent` to compare its native
execution and serialized request size with a canonical stateless Node workflow.
The example and its benchmark are excluded from the published package. The
benchmark explicitly does not claim to measure LLM reasoning or provider/tool
latency. Checked comparative results are in the
[agent protocol rollout audit](./docs/validation-reports/2026-09-12-agent-protocol-rollout.md).

> **Pre-1.0 integration warning:** Agent skills, MCP servers, and harnesses
> should pin the exact `@ansonlai/docx-redline-js` release they were tested
> against and declare the CLI contract version/capabilities they require. Do
> not treat a newly installed release or contract as verified until the
> integration's compatibility tests pass. See the
> [skill-authoring contract](./docs/SKILL_AUTHORING.md#runtime-negotiation).

### Agent CLI

```bash
# One-turn exact mechanical edit; fails closed unless one paragraph matches
docx-redline apply contract.docx --find "thirty (30) days" --replace "sixty (60) days" --profile agent --output reviewed.docx

# Directional context using a longer phrase with a higher chance of uniqueness
docx-redline apply contract.docx --search "distinctive nearby heading or phrase" --context-range 1:3 --find "thirty (30) days" --replace "sixty (60) days" --profile agent --output reviewed.docx

# Semantic drafting still starts with focused context
docx-redline extract contract.docx --search "termination" --around 3
docx-redline preflight contract.docx --operations operations.json --author "Editor"
docx-redline apply contract.docx --operations operations.json --author "Editor" --output reviewed.docx
docx-redline validate reviewed.docx
```

```bash
# Inline one-liner edit (no operations file needed)
docx-redline apply contract.docx --target "Original clause text" --modified "New clause text" --output reviewed.docx

# Direct edit without tracked changes
docx-redline apply contract.docx --target "Typo fix" --modified "Fixed typo" --no-redlines --output clean.docx

# Cross-author edit inside another reviewer's pending insertion
docx-redline apply contract.docx --target "Another author's clause" --modified "Revised clause" --existing-revisions slice-cross-author --output reviewed.docx

# High-assurance atomic batch with nonzero exit on any incomplete result
docx-redline apply contract.docx --operations operations.json --atomic --require-complete --output reviewed.docx

# Agent shell path: JSON is emitted by a serializer, not interpolated by the shell
node emit-operations.mjs | docx-redline apply contract.docx --operations - --profile agent --compact --output reviewed.docx
```

All commands emit JSON on stdout. `apply` defaults:
- **Author**: Defaults to `'AI Redliner'` (or `DOCX_REDLINE_AUTHOR` environment variable).
- **Output overwrite**: Destination files provided via `--output` overwrite by default. Pass `--no-overwrite` or `--no-clobber` to safeguard existing destination files. The source input is never overwritten unless `--in-place` is specified.
- **Existing revisions**: Defaults to `'merge-same-author'`. Pass `--existing-revisions slice-cross-author` to edit inside another reviewer's pending insertions with native carrier slicing.
- **Transactionality**: Defaults to `atomic: false` (applies valid operations and reports any failures). Pass `--atomic` for all-or-nothing rollback on any operation error.
- **Complete-success exit**: Pass `--require-complete` when a progressive `partial` result must exit with code `3`; errors exit with code `2`. Without the flag, partial results retain the legacy zero exit code, so always inspect `completion`.
- **Agent profile**: `--profile agent` enables complete-success exit behavior while retaining the ordinary progressive transaction and existing-revision defaults. Compose it with `--atomic` or an explicit `--existing-revisions` policy when intended, and verify the resolved `effectiveOptions`.
- **Operations from stdin**: `--operations -` reads the same array or `{ operations, expectedRevision }` envelope accepted from a file. Feed it from a JSON serializer or structured process API, not shell-interpolated legal text.
- **Compact stdout**: `--compact` emits one-line mutation JSON. It changes serialization only, not operation semantics.
- **Tracked changes**: Defaults to `generateRedlines: true`. Pass `--no-redlines` when clean direct text edits are desired.
- **Inline edits**: Use `--target <text>` with `--modified <text>` or `--comment <text>` for quick one-liners without creating a JSON file.
- **Localized edits**: Use `--find`/`--replace` with a known strong target, or omit the target for fail-closed global resolution. Add a longer, fairly unique `--search` phrase with `--context-range 1:3` for directional scope; `--around 3` is symmetric. Generic or repeated anchors can widen the unioned scope until multiple paragraphs contain `find`, producing a safe `AMBIGUOUS_TARGET` failure and an extra recovery turn. On success, `results[i].change.context.anchorMatchCount > 1` signals a non-unique anchor; confirm the returned location and bounded before/after evidence.
- **Compact mutation results**: `apply`, `accept`, `reject`, and `delete-comments` omit full OOXML/package payloads and inspection text from stdout. CLI operation results omit duplicate nested receipt bodies; the ordered top-level `receipts` array is authoritative. Successful exact target-match diagnostics are omitted and equivalent-whitespace matches retain only mode/count. Node and standalone-runner results remain unchanged. Use `validate` when full issue arrays are needed.

Inspection commands default to 20 direct matches and a 48 KiB soft response
budget when no explicit positional scope is supplied. When a search returns 0
matches in the requested view, `selection.hint` provides immediate guidance if
matches exist in the alternate view (such as searching for a deleted clause to
restore). Use `--limit` with `--after <paragraph-index>` to continue, `--around N`
(aliases `--context` and `-C`) to include nearby paragraphs, and `--all` only
for deliberate unbounded inspection. Every retained target is complete; an
individually oversized paragraph is returned whole with an `oversizeItem`
marker. Machine `index` and `ref` fields are for targeting and pagination, not
user-facing Word locations; use `humanReference`, `provision`, or
`nearestHeading` in reports.

`docx-redline version` reports contract version 8 and the additive
`command-help-v1`, `inspection-context-v1`, `bounded-inspection-v1`,
`human-document-references-v1`, `batch-start-source-binding`,
`recovery-envelope-v1`, `require-complete-exit`, `operations-stdin`,
`agent-safety-profile-v2`, `deduplicated-cli-receipts`, and
`compact-cli-json-v1`, `localized-replacements-v1`,
`speculative-search-apply-v1`, `localized-change-summary-v1`, and
`restore-shortcuts-v1` capabilities.
Wrappers should negotiate only the
capabilities they use. Run `docx-redline <command> --help` for that command's
machine-readable options, behavior, exit codes, canonical GitHub documentation links, and compact examples.

See the [compact agent fast start](./docs/AGENT_FAST_START.md), the
[skill/harness authoring contract](./docs/SKILL_AUTHORING.md), and the
[operation JSON Schema](docs/schemas/document-operations.schema.json).

### Configuration (call once at startup)

| Function | Purpose |
|----------|---------|
| `configureXmlProvider({ DOMParser, XMLSerializer })` | Inject XML parser. Required in Node.js; browsers usually provide native support. |
| `configureLogger({ log, warn, error })` | Replace default console logger. |
| `setDefaultAuthor(name)` | Set fallback track-change author (default: `'AI Redliner'`, configurable via `DOCX_REDLINE_AUTHOR` environment variable). |
| `setPlatform(label)` | Set platform label for diagnostics (default: `'Unknown'`). |

### Options and Defaults Reference

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `generateRedlines` | `boolean` | `true` | When `true`, emit Word-native tracked changes (`w:ins`/`w:del`). When `false`, apply clean direct edits without revision markup. **Note: Redlines are not always the preferred method** — pass `generateRedlines: false` (or `--no-redlines` via CLI) when producing clean execution drafts, restructuring documents, or when revision clutter is unwanted. |
| `author` | `string` | `'AI Redliner'` | Reviewer/author name stamped on generated tracked changes and comments. Overridable via `DOCX_REDLINE_AUTHOR` env variable. |
| `atomic` | `boolean` | `false` | Batch transaction mode. By default (`false`), valid edits are applied and failing operations report errors. When `true`, any operation failure rolls back the entire batch to the original document state (`rolledBack: true`, `hasChanges: false`). Feature prominently in high-assurance workflows. |
| `structuredContent` | `boolean` | `true` | Auto-detects Markdown tables, headings (`#`), and lists in replacement text and renders them as native Word elements (`w:tbl`, `w:pStyle`, `w:numPr`). Pass `false` to treat replacement text strictly as plain text. |
| `pairReplacements` | `boolean` | `true` | Links adjacent `<w:del>` and `<w:ins>` revisions with matching timestamps so Word groups them as a single replacement in the Reviewing Pane. |
| `strictTargets` | `boolean` | `true` (CLI/facade) | Requires exact target descriptors (`exactText`, `paragraphId`, `index`, `occurrence`, `fingerprint`) and forbids ambiguous matching. Defaults to `false` in low-level runner for backwards compatibility. |
| `existingRevisions` | `string` | `'merge-same-author'` | How to handle paragraphs with existing tracked changes. `'merge-same-author'` merges revisions from the same author and protects other authors with `EXISTING_REVISIONS`. `'slice-cross-author'` keeps that same-author merge behavior while allowing Word-native edits inside another author's pending insertion. Pass `'accept-all-first'` to normalize prior revisions or `'reject-input'` to refuse editing revised paragraphs. |
| `removeFormatting` | `boolean` | `false` | When `true` and the text is unchanged with no Markdown hints, strips existing bold/italic/underline/strikethrough formatting. |
| `sanitizeInput` | `boolean` | `false` | Opt-in removal of standalone leading assistant-preface lines. Literal dollar signs and `\n` sequences are always preserved. |

Same-author revision merging refuses paragraphs containing comment anchors with
`COMMENTED_CONTENT_MERGE`; resolve those comments first so the merge cannot
remove or orphan their anchors.

Common result fields:

| Field | Purpose |
|-------|---------|
| `status` | Operation status: `'ok'`, `'partial'`, `'no-op'`, or `'error'`. |
| `error` | Present on failure; retains a stable `code` and adds recovery envelope version, stage, category, bounded context, and a machine-readable recovery action. |
| `written` | CLI/facade boolean indicating whether the output file was successfully written to disk. |
| `completion` | CLI-only boolean that is `true` only when a destination was written, top-level status is neither error nor partial, and no operation result failed. |
| `rolledBack` | Present and `true` when an atomic batch encountered an error and rolled back all changes. |

Word diffs are deterministic by default (no wall-clock timeout). Inputs above
the safe ceiling of 262,144 unique diff tokens return `DIFF_TOKEN_LIMIT` with
the original OOXML unchanged so callers can split the operation without risking
silent text loss.

### Editing inside existing revisions (cross-author slicing)

During multi-round legal negotiations, a reviewer often needs to edit text that was previously inserted by another reviewer whose revision is still pending. Pass `existingRevisions: 'slice-cross-author'` (or `--existing-revisions slice-cross-author` via CLI) to edit inside another author's pending insertion without erasing their attribution or requiring prior acceptance:

```js
const result = await applyRedlineToOxml(paragraphOoxml, originalText, modifiedText, {
  generateRedlines: true,
  author: 'Reviewer B',
  existingRevisions: 'slice-cross-author'
});
```

The engine applies Microsoft Word Desktop-native tracked change structures:
- **Insertions inside pending insertions**: The carrier `<w:ins>` is split into sibling `<w:ins>` containers at the paragraph level (`[ins(Author A), ins(Author B), ins(Author A)]`), maintaining strict schema compliance without illegal `ins/ins` nesting.
- **Deletions inside pending insertions**: The new `<w:del>` is nested directly inside the carrier `<w:ins>` (valid under ECMA-376 Part 1 `CT_RunTrackChange`), ensuring that if Author A's insertion is rejected, Author B's dependent deletion is cleanly removed with it.
- **Straddle deletions**: Deletions spanning between baseline text and pending insertions cleanly partition across their respective container contexts without invalid coalescing.
- **Lifecycle parity**: Accepting or rejecting either reviewer independently produces identical results to Microsoft Word Desktop's native review pane.

### Replacing a heading with a tracked list

The list route treats a one-paragraph heading expanded into multiple markdown
items as a structural block replacement. The deleted heading stays in its own
tracked paragraph, and every inserted item becomes a separate Word list
paragraph. Accepting the revisions produces only the list items; rejecting them
restores the original heading exactly.

```js
const result = await applyRedlineToOxml(
  headingParagraphOoxml,
  'A.\tPURPOSE',
  '* Article A. Purpose and Interagency Alignment\n' +
    '* Key Focus: Joint Street Outreach & Medical Triage',
  { generateRedlines: true, author: 'Editor' }
);
```

`w:numId w:val="0"` means numbering is explicitly suppressed; it is not a list
definition and is never reused for generated bullets. New list items receive a
positive numbering ID. Font family, size, language, and related script
properties are inherited from the source paragraph, while heading emphasis
(such as bold or underline) is not copied unless the replacement markdown asks
for it. For a complete `word/document.xml`, prefer
`applyOperationToDocumentXml(...)` so replacement nodes and numbering artifacts
are imported at the correct scope.

### Planning large mixed-content insertions

For an attachment or schedule containing several kinds of content, use
`planStructuredReplacement(...)` before applying the edit. It decomposes the
Markdown into typed heading, paragraph, list, and table blocks, normalizes the
block boundaries, and returns one atomic operation with
`structuredContent: true`.

```js
import { planStructuredReplacement } from '@ansonlai/docx-redline-js';

const plan = planStructuredReplacement(
  { exactText: 'Date', index: 172 },
  `# ATTACHMENT 4

Introductory paragraph.

| Agency | Contact |
| --- | --- |
| BCHD | Dr. Jenkins |

## Protocol

1. Joint clearance
2. Rapid escalation`,
  { author: 'Editor' }
);

if (!plan.valid || !plan.operation) {
  throw new Error(plan.issues.map(issue => issue.message).join(' '));
}

const result = await document.applyOperations([plan.operation], {
  author: 'Editor', atomic: true, validate: true
});
```

Every Markdown table must include a separator row immediately below its header.
Missing separators, missing data rows, and inconsistent column counts are
reported as structured errors instead of being inserted as visible pipe text.
Use `#` through `#########` to request Word heading paragraphs; blank lines
separate paragraphs; adjacent list markers form a real Word list. Keep the
planned content in one operation so later blocks do not depend on an anchor that
an earlier block has already replaced.

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
| `analyzeStructuredContent(markdown)` | Decompose mixed Markdown into typed blocks and report malformed table syntax without creating an operation. |
| `planStructuredReplacement(target, markdown, options)` | Validate mixed content and return one atomic `structuredContent` replacement operation, or `operation: null` with issues. |
| `containsTrackedChanges(xmlDoc)` | Detect `w:ins`, `w:del`, move revisions, property changes, and paragraph-mark revision markup in a parsed OOXML document/fragment. |
| `validateRedlineOoxml(oxml)` | Validate generated redline OOXML against the package's structural invariants (no nested revisions, `w:delText` inside `w:del`, complete metadata, unique revision ids, preserved boundary whitespace). Returns `{ valid, issues }`; run it before writing output into a package. |

### Services

| Function | Purpose |
|----------|---------|
| `injectCommentsIntoOoxml(oxml, comments, options)` | Add comments anchored to text ranges. |
| `applyCommentReplyToParts(options)` | Build a threaded reply in `comments.xml` and `commentsExtended.xml` without adding a document-body anchor. |
| `acceptTrackedChangesInOoxml(oxml, { author?, allAuthors? })` | Accept `w:ins` / `w:del` / `w:moveFrom` / `w:moveTo` / `*PrChange` revisions for one author or all authors. |
| `rejectTrackedChangesInOoxml(oxml, { author?, allAuthors? })` | Reject `w:ins` / `w:del` / `w:moveFrom` / `w:moveTo` / `*PrChange` revisions for one author or all authors. |
| `deleteCommentsByAuthorInOoxml(oxml, { author?, allAuthors? })` | Delete matching comment definitions and anchors present in the supplied OOXML payload. Real `.docx` packages require updating both `word/comments.xml` and `word/document.xml`. |
| `generateTableOoxml(headers, rows, options)` | Generate a `w:tbl` from tabular data. |
| `createDynamicNumberingIdState(numberingXml)` | Allocate numbering IDs without collisions. |
| `ensureNumberingArtifactsInZip(zip, numberingXml, options)` | Add numbering artifacts to a `.docx` package. Replacement of existing numbering without `mergeNumberingXmlBySchemaOrder` is deprecated and will throw in the next major version. |
| `ensureCommentsArtifactsInZip(zip, commentsXml)` | Merge comments artifacts into a `.docx` package. |
| `ensureCommentsExtendedArtifactsInZip(zip, commentsExtendedXml)` | Add or replace modern Word comment-thread metadata in a `.docx` package. |
| `validateDocxPackage(zip)` | Validate `.docx` structural consistency. |

Malformed OOXML never escapes these public transform APIs as a raw parser
exception. Transforms return `status: 'error'` with `error.code === 'PARSE_ERROR'`;
validators return a `PARSE_ERROR` issue. Recoverable XML parser
diagnostics are forwarded through the configured logger and included in
`warnings` where the result shape supports them.

Cross-author slicing also verifies its exact accepted-view text before success.
If a structural boundary prevents exact reconstruction, the transform returns
`status: 'error'` with `error.code === 'PATCH_ROUNDTRIP_MISMATCH'`,
`hasChanges: false`, and the original OOXML unchanged.
Pure insertion-only slicing uses an exact character-local diff so repeated words
cannot move an insertion to a different occurrence. Leading/trailing spaces,
tabs, and non-breaking spaces are treated as real changes rather than no-ops.
For text-bearing replacements, the exact accepted-view text of the resolved
paragraph—not a space-normalized caller target—defines mutation offsets. Word-level replacement
hunks that differ only by ordinary spaces and NBSPs are refined to character
edits so unchanged hyperlinks and their relationship attributes stay in place.
Runner results expose bounded `resolvedTarget.targetTextMatch` code-point
diagnostics when equivalent whitespace was used to identify the target.

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

Before mutation, the runner resolves strong source descriptors against the
immutable batch-start document and binds them to session-local source
identities. Independent edits therefore do not need to be manually sorted when
an earlier structural rewrite changes later paragraph indexes or fingerprints.
Targets that deliberately refer to uniquely created paragraph text are compiled
into an internal capture dependency. True overlap is not guessed: incompatible
writes to one source fail before mutation with `OVERLAPPING_SOURCE_TARGETS` or
`REVISION_ORDER_CONFLICT`, and mutating capture fan-out without distinct
selectors fails with `CAPTURE_FANOUT_CONFLICT`.

Threaded replies use a comment operation with no body target:

```js
{ type: 'comment_reply', parentCommentId: 8, commentContent: 'Agreed; I revised this.', author: 'Editor' }
```

For complete `.docx` files, use `openDocx(...).applyOperations(...)`; it reads
the existing comments parts and writes the required `commentsExtended.xml`
relationship and content type. Inspection reports `paraId` and
`parentCommentId` so callers can discover and verify the thread hierarchy.

A whole-paragraph `delete` that targets existing comment markup fails with
`COMMENTED_CONTENT_DELETE`. In the Node facade and CLI, the error also includes
the affected comment author and text. Resolve the feedback or explicitly remove
the comment first; the library will not silently discard reviewer context.

The batch runner keeps one live document DOM and performs one final full-document
serialization. Accuracy remains the controlling constraint: each operation has
an internal savepoint so an error or no-op cannot leak a partial edit or consumed
revision ID into later operations.

Batches are progressive by default (`atomic: false`): valid operations commit
while failed operations remain unapplied and are reported in `results`. Pass
`{ atomic: true }` when any operation error must return the original
`documentXml`, `hasChanges: false`, empty package artifacts, and
`rolledBack: true`. The default `continueOnError: true` still attempts the full
batch so `results` describes every operation; use `{ continueOnError: false }`
to stop after the first error.

Every failed or partial mutation includes `retryPlan`. Its `base` is `original`
after rollback/no commit and `output` after a progressive partial commit;
`committedIndexes`, `failedIndexes`, and `unattemptedIndexes` identify the safe
replay scope. Errors use recovery envelope version 1 and always report
`sameArgumentsSafe: false`. Follow `error.recovery.action`; do not infer a retry
from prose. Authorization-sensitive actions, such as resolving comments, are
marked with `requiresUserAuthorization: true`. For `EXISTING_REVISIONS`, the
non-normalizing surgical recommendation is `slice-cross-author`; accepting or
rejecting prior revisions still requires explicit authority.

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
authors plus required comments/numbering artifacts.

Application currently defaults to permissive targeting for backward compatibility,
but will default to `strictTargets: true` in v1.0.0. When permissive resolution
chooses among multiple candidate paragraphs heuristically, it emits an
`AMBIGUOUS_TARGET_HEURISTIC_USED` warning containing candidate count and migration
guidance. Callers should pass `{ strictTargets: true }` and use strict descriptors
(`paragraphId`, `index`, `occurrence`, or `fingerprint`) to prepare for v1.0.0.

### Mutation Receipts

Both single-operation (`applyOperationToDocumentXml`) and batch
(`applyOperationsToDocumentXml`) results expose commit-aware **Mutation Receipts**
(`result.receipt` on single results and per-item `results[i].receipt`, plus `result.receipts`
for the full batch).

This full receipt shape is retained by the Node facade and standalone runner.
Only compact CLI JSON removes duplicate `results[i].receipt` bodies and keeps
the ordered top-level `receipts` array.

```js
const result = await applyOperationsToDocumentXml(documentXml, operations, 'Agent');
for (const receipt of result.receipts) {
  console.log(receipt.operationIndex, receipt.finalDisposition, receipt.committed);
  console.log('Revisions:', receipt.revisionItems);
  console.log('Comments:', receipt.commentIds);
}
```

Receipts report:
- `operationIndex` (1-based), `operationId`, and `authorUsed`
- `attemptedDisposition` and `finalDisposition` (`applied`, `refused`, `no_change`, `rolled_back`, or `not_attempted`)
- `committed` (boolean: verified committed into serialized package output)
- `revisionItems` (exact allocated revision IDs with kind and target part)
- `commentIds`, `numberingIds`, and `relationshipIds`
- `affectedTargets` (resolved target coordinates) and `warnings`

Before completing an operation or batch transaction, `reconcileReceiptsAgainstOutput`
verifies every reported committed durable ID against a fresh parse of the output OOXML.
Any discrepancy fails closed and triggers immediate rollback.

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
- Treat `w:numId w:val="0"` as numbering suppression, never as a reusable list
  ID. Generated bullet and numbered paragraphs must reference a positive ID
  whose definition is merged into `word/numbering.xml`.
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

// Restoring another reviewer's pending whole-paragraph deletion requires
// explicit intent. The restored counterproposal becomes a separately tracked
// sibling paragraph; a normal redline operation remains fail-closed.
const restoration = await applyOperationToDocumentXml(
  documentXml,
  {
    type: 'restore',
    // restore targets default to the rejected view, where deleted text exists
    target: { paragraphId: '1A2B3C4D', exactText: 'Original deleted paragraph text.' },
    modified: 'Restored or adjusted paragraph text.'
  },
  'Editor'
);

// A single restoration follows its deleted source paragraph. A range
// restoration follows the complete deleted source block. Unchanged legacy
// validation defects are retained as baseline issues; newly generated errors
// fail closed before commit.
// inspect/extract descriptors report their revisionView, and each fingerprint
// is computed from the same view as exactText. Keep those fields together.

// To insert run-level text at a location visible only in the rejected view,
// provide explicit rejected-view intent and an exact anchor-relative offset.
const deletedTextInsertion = await applyOperationToDocumentXml(
  documentXml,
  {
    type: 'insert',
    target: { paragraphId: '1A2B3C4D', revisionView: 'rejected' },
    anchor: { exactText: 'must pay', occurrence: 1, offset: 5 },
    modified: '[clarification] ',
    existingRevisions: 'slice-cross-author'
  },
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
administrative corpus representatives. See the [Word visual review guide in docs/TESTING.md](./docs/TESTING.md#microsoft-word-visual-review-guide);
preparation and AI preflight never count as human sign-off.

A nightly GitHub Actions workflow additionally validates generated fixtures
against the ECMA-376 transitional schemas (`xmllint`), opens them with
LibreOffice, and runs an extended fuzz sweep of the accept/reject round-trip
invariant with a fresh seed. See [Release validation in docs/TESTING.md](./docs/TESTING.md#release-validation-and-independent-oracles).

## Architecture & Contributing

- **[ARCHITECTURE.md](./ARCHITECTURE.md)**: Detailed module layout, end-to-end reconciliation flow, and contributor fast orientation.
- **[AGENTS.md](./AGENTS.md)**: Fast-start routing and operational guardrails for AI coding agents.
- **[docs/AGENT_FAST_START.md](./docs/AGENT_FAST_START.md)**: Minimal ordinary-edit contract for agent integrations.
- **[docs/AGENT_KNOWLEDGE_BASE.md](./docs/AGENT_KNOWLEDGE_BASE.md)**: Full agent reference for APIs, operations, CLI automation, recovery, and gotchas.
- **[docs/TESTING.md](./docs/TESTING.md)**: Comprehensive testing model, test lanes, independent oracle checks, and visual review checklist.
- **[CHANGELOG.md](./CHANGELOG.md)**: Version history, migration guides, and deprecation schedules.

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
