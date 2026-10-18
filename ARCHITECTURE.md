# Reconciliation Core Architecture

This document describes the standalone, host-agnostic OOXML reconciliation package and how to work inside it safely.

## Scope

This repository contains only the publishable package surface:

- `adapters/`
- `core/`
- `document/` (universal complete-DOCX facade & ZIP archive layer)
- `engine/`
- `pipeline/`
- `services/`
- `orchestration/`
- `node/` (Node CLI & backward-compatibility facade)
- `index.js`
- `index.d.ts`
- `dist/`

No Word add-in entrypoints or host-specific integration layers are part of this package.

## Goals

- Preserve Word-compatible redlines by editing OOXML directly.
- Keep core logic and document-level operations host-independent (no Office.js globals, no Word API calls, no runtime Node built-ins like `node:zlib` or `node:crypto`).
- Standardize complete-document packages on `Uint8Array` binary payloads, pure-JS SHA-256 revision tokens, and `fflate` ZIP container manipulation.
- Reuse the same engine across Node.js, browsers, Cloudflare Workers, Deno, Bun, and sandboxed runtimes (e.g. n8n Code nodes).
- Keep generated OOXML schema-safe through shared Word element creation and
  centralized revision metadata helpers.

## Folder Layout

```text
.
├── adapters/
│   ├── config.js
│   ├── logger.js
│   └── xml-adapter.js
├── core/
│   ├── current-properties.js
│   ├── paragraph-text.js
│   ├── sha256.js
│   ├── word-revision-semantics.js
│   └── word-xml.js
├── document/
│   ├── docx-document.js
│   ├── docx-document.d.ts
│   └── zip-archive.js
├── engine/
│   ├── oxml-engine.js
│   ├── surgical-mode.js
│   ├── reconstruction-mode.js
│   ├── run-builders.js
│   ├── paragraph-mark-revisions.js
│   └── formatting-removal.js
├── orchestration/
├── pipeline/
├── services/
│   ├── comment-engine.js
│   ├── numbering-helpers.js
│   ├── numbering-allocation.js
│   ├── list-conversion-diagnostics.js
│   ├── list-format.js
│   ├── list-format-mutation.js
│   ├── list-format-change.js
│   ├── operation-batch-compiler.js
│   ├── localized-replacement-compiler.js
│   ├── revision-comment-management.js
│   ├── document-inspection.js
│   ├── inspection-comparison.js
│   ├── comment-parts-merge.js
│   ├── document-operation-session.js
│   ├── document-operation-applier.js
│   ├── document-operation-mutations.js
│   ├── batch-operation-orchestrator.js
│   ├── revision-fidelity-verifier.js
│   ├── operation-heuristics.js
│   ├── standalone-docx-plumbing.js
│   └── standalone-operation-runner.js
├── node/
│   ├── docx-document.js
│   ├── cli.js
│   ├── zip-archive.js
│   └── index.js
├── dist/
│   ├── docx-redline-js.esm.js
│   ├── docx-redline.bundle.js
│   └── docx-redline.bundle.cjs
├── index.js
└── index.d.ts
```

## Entry Points

- `index.js`: primary host-independent exports for OOXML reconciliation and universal complete-document operations (`openDocx`, `DocxDocument`, `computePackageRevisionToken`).
- `document/docx-document.js`: universal `Uint8Array`-based document facade and package-level operation runner.
- `dist/docx-redline.bundle.js` / `dist/docx-redline.bundle.cjs`: zero-dependency standalone sandbox bundles, exported via `@ansonlai/docx-redline-js/bundle`.
- `services/standalone-operation-runner.js`: stable complete-document XML
  operations, exported as `@ansonlai/docx-redline-js/standalone-runner`.
- `node/index.js`: Node-compatible complete-DOCX facade re-exports, exported as
  `@ansonlai/docx-redline-js/node`.
- `bin/docx-redline.js`: CLI launcher; command implementation lives in
  `node/cli.js` and shared operation behavior lives in `services/`.

## Module Responsibilities

- `adapters/config.js`
  - Runtime configuration for defaults (`setDefaultAuthor`, `getDefaultAuthor`, `setPlatform`, `getPlatform`).
- `adapters/xml-adapter.js`
  - XML parser/serializer injection for browser or Node.js runtimes.
- `adapters/logger.js`
  - Runtime logger injection and shared logging methods.
- `core/*`
  - Shared types, OOXML identity helpers, target resolution, list/table targeting heuristics, and XML query helpers.
- `core/word-xml.js`
  - Namespace-safe Word element creation, tracked-change detection, OOXML payload source-shape helpers, and
    `resolveExistingRevisionsPolicy`, the single place `existingRevisions: 'auto'` is resolved (per target
    paragraph or per `targetEnd` range; unattributed changes count as another author's).
- `core/types.js`
  - Shared model enums/types plus revision metadata generation and document-aware revision ID seeding.
- `core/paragraph-text.js`
  - Canonical accepted/rejected/current-view text shared by targeting,
    ingestion, and inspection.
- `core/paragraph-targeting.js`
  - Target descriptors and session-scoped paragraph metadata indexes. The
    index groups IDs and normalized text while retaining canonical text,
    fingerprints, document order, and table context for deterministic reuse.
- `core/redline-validation.js`
  - Runtime structural validation (`validateRedlineOoxml`) mirroring the test-suite invariants: no nested revisions, `w:delText` inside `w:del`, complete revision metadata, unique revision ids, preserved boundary whitespace.
- `core/validation-delta.js`
  - Stable issue signatures and multiset subtraction for classifying baseline
    versus generated validation issues without hiding added duplicate errors.
- `core/word-revision-semantics.js`
  - Pure DOM projection of what Word shows after Reject All or Accept All
    (`projectWordRevisionView`, `wordRevisionParagraphTexts`,
    `paragraphMarkRevision`). It encodes Word's paragraph-mark rules: a
    tracked-inserted mark vanishes on Reject and its content merges into the
    following paragraph of the same container, a tracked-deleted mark does the
    same on Accept, a table cell keeps its last paragraph, rows marked inserted
    or deleted vanish, and Reject restores `w:pPrChange` snapshots. It never
    mutates or serializes. The rules are validated against recorded desktop
    Word 16 COM results (`tests/fixtures/word-revision-semantics/`).
- `core/current-properties.js`
  - Direct-child accessors for the current paragraph/run properties
    (`getCurrentParagraphProperties`, `getCurrentRunProperties`,
    `getCurrentChild`, `getCurrentNumIdElements`, ...). Descendant lookups cross
    `w:pPrChange` / `w:rPrChange`, so a historical `numPr` or `pStyle` would look
    current; the snapshot is reachable only through the explicit
    `getHistoricalParagraphProperties` / `getHistoricalRunProperties` and
    `getAllReferencedNumIdElements`. Depends only on `core/types.js`.
- `core/revision-cloning.js`
  - Shared effective-property cloning strips historical revision descendants;
    intentional revision-bearing splits refresh cloned property-change IDs.
- `engine/oxml-engine.js`
  - Main reconciliation router, mode selection, existing-revision policy gate, and status/error result handling.
- `engine/route-selection.js`
  - Internal reconciliation capability matrix and opt-in diagnostic route
    instrumentation. It observes policy; it does not expose a route override.
- `engine/run-builders.js`
  - Shared builders for insertion/deletion wrappers, visible run content, and run-property changes. It does not build paragraph-mark revisions.
- `engine/paragraph-mark-revisions.js`
  - The single construction site for a tracked paragraph mark (`w:ins` / `w:del`
    under `w:pPr/w:rPr`). DOM entry points (`markParagraphInserted`,
    `markParagraphDeleted`, `setParagraphPropertiesMarkRevision`,
    `clearParagraphMarkRevision`, `restoreParagraphMarkInserted`) and string
    builders (`paragraphMarkRevisionXml`, `paragraphMarkPPrXml`) share one
    metadata path, so ids are allocated through the document's allocator and
    recorded in receipts. `planParagraphMarks` is a pure check that an
    alignment gives every source paragraph exactly one untracked or deleted
    mark and every new target paragraph an inserted mark, so Reject All yields
    the source paragraph count and a deleted paragraph is never collapsed into
    another. Reconstruction, list and table generation, and document-level
    shortcuts all call it. Depends only on `core/`.
- `engine/surgical-*.js`
  - Surgical run splitting, diff application, and span helpers for localized edits that preserve surrounding markup. Plain-text edit groups execute from right to left with a fresh live span index, while space/NBSP-only replacement hunks are refined to character-local changes so unchanged hyperlink containers survive.
    The shared carrier splitter also handles explicit rejected-view deletion
    splits, retaining `w:delText`, tabs, breaks, soft/non-breaking hyphens,
    formatting, foreign metadata, and fresh trailing/property-change IDs.
- `engine/formatting-removal.js`
  - Shared formatting removal and highlight helpers.
- `pipeline/list-markers.js`
  - Dependency-light canonical list-marker grammar, classification, numbering
    vocabulary, and parsed list-item representation.
- `pipeline/list-generation.js`
  - Structural paragraph-to-list generation, paragraph-mark revisions, and
    selective inheritance of source typography for inserted list runs.
- `pipeline/structured-content.js`
  - Strict decomposition and planning for atomic mixed heading/paragraph/list/
    table replacements. It rejects ambiguous table-like pipe text rather than
    allowing a fallback to literal paragraphs.
- `pipeline/*`
  - Ingestion, markdown preprocessing, diffing, patching, and serialization stages. Ingestion treats deleted and moved-from content as non-visible text and inserted/moved-to content as visible text.
- `services/comment-engine.js`
  - Comment creation and package-level comment XML handling.
- `services/numbering-helpers.js`
  - Document-scoped numbering ID state (`createDynamicNumberingIdState`,
    `reserveNextNumberingIdPair`), numbering payload remapping, and shared DOM
    predicates. It no longer owns a merge.
- `services/numbering-allocation.js`
  - The single place that creates, continues, merges and validates numbering
    definitions. `allocateNumberingDefinition` builds a new `w:abstractNum` +
    `w:num` (source levels copied, chosen levels redefined, start overrides)
    from ids reserved through the document state; `allocateNumberingForReplacement`
    holds the reuse-source-list policy (`continue-source`, `allocate`,
    `passthrough`) so no call site decides it; `continueBatchListSequence` lets
    separate operations of one batch that generate items of the same kind,
    format and level continue one numbering instance (reused only while its
    `numId` is still bound in the live document). `mergeNumberingParts` /
    `mergeNumberingXmlBySchemaOrder` are the one numbering merge, used by the
    runner's consumers and the `openDocx` facade; an incoming `w:abstractNum` /
    `w:num` id that exists with different content throws `NUMBERING_ID_CONFLICT`.
    `checkNumberingBinding` validates a direct `{ numId, level }` binding and
    `validateNumberingIntegrity` reports dangling `numId`, `abstractNumId` and
    `ilvl` (`NUMBERING_*` issues) for the facade's baseline-delta validation.
    Old definitions are never deleted or overwritten, and tracked-deleted source
    paragraphs keep their `numPr`.
- `services/list-conversion-diagnostics.js`
  - Safety diagnostics for text edits that retype list markers
    (`LIST_MARKER_NOT_BOUND` / `UNSUPPORTED_LIST_CONVERSION`,
    `LIST_FORMAT_NOT_APPLIED` / `UNSUPPORTED_LIST_FORMAT_CHANGE`; warning by
    default, error under `strictListConversion`, silenced by
    `listMarkers: 'literal'`), the closed `paragraph-format` key set (unknown
    keys are `INVALID_OPERATION`) and the `numbering` shape check behind
    `paragraph-format` numbering binding. Recovery points at `list-format`
    through `recovery.suggestedOperation`.
- `services/standalone-docx-plumbing.js`
  - Package-level extraction/wiring/validation for `word/document.xml`, `word/numbering.xml`, and `word/comments.xml`.
- `services/revision-comment-management.js`
  - OOXML transforms for accepting/rejecting insertion, deletion, move, paragraph-mark, and property-change revisions by author/all-authors, plus deleting comments by author/all-authors.
- `services/standalone-operation-runner.js`
  - Stable compatibility facade for full-document operations. It only re-exports
    preflight, single-operation application, batch application, and scheduling.
- `services/document-operation-applier.js`
  - Canonical single-operation validation, author resolution, dispatch, and
    result metadata assembly. Before commit it validates the entire live
    document against the operation savepoint and refuses newly generated
    structural errors, including duplicate revision IDs. For localized
    `replacements` it also checks, before commit, that the target's accepted text
    (read from the successor paragraph via `SourceTargetRegistry.previewSuccessor`,
    or the restored paragraph) equals the compiled desired text; otherwise the
    savepoint is restored and `PATCH_ROUNDTRIP_MISMATCH` is returned, because a
    tracked whole-paragraph deletion rejects back to the source and the Reject
    All verifier alone cannot see it.
- `services/document-operation-mutations.js`
  - Coupled OOXML mutation implementations for redline, highlight, and comment
    operations. For text-bearing edits, the resolved paragraph's exact
    accepted-view text is the source coordinate system even when target
    selection used normalized whitespace;
    bounded match-mode/code-point diagnostics are attached to resolved target
    metadata. Explicit rejected-view insertion splits a supported direct
    foreign `w:del` at an anchor-relative offset into sibling deletion,
    insertion, and deletion carriers. Paragraph restoration emits its inserted
    block after the untouched source range and verifies baseline-delta,
    mutation-envelope, and lifecycle postconditions. These use leaf-module
    imports and never import the root entry.
- `services/batch-operation-orchestrator.js`
  - Comment-first stable scheduling, atomic policy, artifact aggregation,
    per-operation results, one final document serialization, and deferred
    runtime-context commit.
- `services/document-operation-session.js`
  - One live document DOM and revision allocator per invocation, immutable
    start-of-batch target snapshot, lazy paragraph metadata, per-operation DOM
    and allocator savepoints, exact original XML rollback, index invalidation,
    artifact/result accumulation, and isolated runtime-context helpers.
- `services/operation-heuristics.js`
  - DOM-light decisions for list/plain adjacency insertion and explicit-range
    list insertion. Canonical list/table targeting remains in `core/*`.
- `services/document-operation-contract.js`
  - Compatibility normalization, canonical operation kinds, author precedence,
    and stable runtime validation for document operations.
- `services/list-format.js` and `services/list-format-mutation.js`
  - The `list-format` operation (`convert-plain`, plus the shared request shape
    for `change-format` and `join`): contract shape, the narrow
    typed-marker grammar (letter versus roman disambiguation by sequence or an
    explicit `format`), and the in-place DOM core. Every refusal is decided
    before mutation; the binding reuses the `paragraph-format.numbering` path
    (`snapshotAndAttachPPrChange` + `applyParagraphPropertiesToPPr`), the marker
    becomes one tracked `w:del`, definitions come from
    `allocateNumberingDefinition`, and batch continuity uses
    `runtimeContext.listFallbackSequenceState` (an entry is reused only while
    its `numId` is bound in the live document). Text-equal redline header
    conversions for letters, roman numerals and bullets call the same core.
    The `paragraph-format` numbering binding (`properties.numbering`,
    `{ numId, level }` or `null`) lives in `document-operation-mutations.js`: it
    snapshots the earliest same-author `w:pPrChange`, refuses another author's
    pending snapshot (`PARAGRAPH_FORMAT_REVISION_CONFLICT`), and checks the
    binding with `checkNumberingBinding`. `list-format` is body-only; `openDocx`
    refuses it with a `part` selector.
- `services/list-format-change.js`
  - `list-format` `change-format` and `join` (`scope: 'list'` or `'range'`): rebind existing list items with
    one tracked `w:pPrChange` each, never editing a shared definition.
    `change-format` allocates a new definition through
    `allocateNumberingDefinition` (source levels copied, one redefined, start
    values on the new abstract); `join` reuses the earlier list's `numId`.
    Planning models Word's counters (shared abstracts continue each other, a
    level start re-applies on restart, deeper items start unused levels) and
    refuses splits, joins and style-linked or instance-overridden definitions
    it cannot reproduce exactly. Definitions are read from the source part plus
    the batch's pending numbering parts.
- `services/operation-batch-compiler.js`
  - Binds every operation to batch-start source paragraphs
    (`SourceTargetRegistry`), compiles localized `replacements`, and detects
    overlap before any mutation. Composition rule for `redline` + `list-format`:
    exactly one single-paragraph `redline` (no `targetEnd`, anchor, part or
    structured content) that keeps the typed marker and changes the text after
    it composes with one `list-format` on the same paragraph; the compiler fixes
    the order by setting `_compiledRunAfter` on the list-format, which the batch
    orchestrator honours, so the redline always runs first. Everything else on a
    list-format target is refused with `LIST_FORMAT_COMPOSITION_CONFLICT`
    (`reason`: `marker-changed`, `redline-unchanged`, `redline-shape`, or
    `list-rebind` when a `scope: 'list'` `change-format`/`join` rebinds an item
    that another operation also writes). `list.continueFrom` is bound at batch
    start as a read-only target. `previewSuccessor` lets the applier find a
    target's replacement paragraph before commit.
- `services/operation-preflight.js`
  - Read-only strict target and anchor resolution, revision-policy diagnostics
    (reporting the policy `existingRevisions: 'auto'` resolves to per target via
    `resolveExistingRevisionsPolicy` in `core/word-xml.js`),
    artifact prediction, and same-paragraph conflict reporting.
- `services/document-inspection.js`
  - Read-only paragraphs/comments inventory with target identity, headings,
    revision authors, table context, and advisory visible numbering.
- `services/comment-parts-merge.js`
  - `mergeCommentParts`: the comment half of writing a batch result back into a package whose parts
    are XML strings (no ZIP). Shares `mergeCommentsPartXml` with `ensureCommentsArtifactsInZip` and uses
    `reconcileCommentSiblingParts`, so it merges exactly as `openDocx` does. Relationships and content
    types stay with the host.
- `services/inspection-comparison.js`
  - Pure comparison of two inspections by text, table position, list label,
    merged review segments and comments, ignoring ids, dates and run boundaries.
    Hosts that write a prepared package through their own API use it to verify
    the committed result. Imports nothing.
- `document/docx-document.js`
  - Universal transactional whole-DOCX editing, artifact wiring, validation, and rollback.
    Standardized on `Uint8Array` binary payloads and pure-JS SHA-256 revision tokens.
    Exposes `openDocx`, `DocxDocument`, and `computePackageRevisionToken` directly across all runtimes.
- `document/zip-archive.js`
  - Universal, pure-JS ZIP archive reader and writer based on `fflate`. Operates on
    `Uint8Array` byte arrays with zero Node built-ins (`node:zlib`, `Buffer`).
- `core/sha256.js`
  - Pure-JS synchronous implementation of FIPS 180-4 SHA-256 with zero dependencies,
    enabling deterministic revision tokens in sandboxed runtimes without `node:crypto`.
- `node/docx-document.js` and `node/zip-archive.js`
  - Node-compatible re-exports maintaining `Buffer` compatibility for `@ansonlai/docx-redline-js/node`.
- `node/cli.js` and `bin/docx-redline.js`
  - Cross-platform, JSON-only agent command boundary. Read commands never
    mutate; write commands require attribution, use package transactions, and
    only overwrite source files under explicit `--in-place` authorization.
    Mutation commands expose a compact contract: package/XML payloads and full
    validation arrays remain internal, while stdout contains durability fields,
    per-operation evidence, validation counts, and a derived completion flag.
- `services/revision-fidelity-verifier.js`
  - Post-apply invariant check built on `core/word-revision-semantics.js`.
    `verifyRevisionFidelity` compares Reject All of the output with Reject All
    of the savepoint (both reject every author's revisions) and reports
    `REJECT_PARAGRAPH_COUNT`, `REJECT_PARAGRAPH_TEXT`, `REJECT_PARAGRAPH_PROPS`,
    `REJECT_NUMBERING_REFS` and `REJECT_WRAPPERS`. Under accept-first
    existing-revision policies the baseline is the accepted view, scoped to the
    resolved target range when known (otherwise the changed region between the
    unchanged prefix and suffix). `REJECT_NUMBERING_REFS`
    (`verifyNumberingReferences`) also refuses a numId that the output binds or
    Reject All restores when the source did not use it and no `w:num` ->
    `w:abstractNum` -> `w:lvl` defines it in the numbering available to the
    operation (numbering existence). `verifyUntrackedLocality` covers
    `generateRedlines: false` single-target operations with `ACCEPT_NEIGHBORS`:
    no paragraph outside the target may change in either view.
    `verifyNoCreatedRevisions` adds `UNTRACKED_CREATED_REVISION`: an untracked
    operation may not introduce a revision id that was not in the source
    (fragments of a sliced source revision, same element, author and date, are
    allowed). `verifyAcceptIntent` adds `ACCEPT_INTENT`: a tracked plain-text
    `redline` with complete `modified` (single paragraph or `targetEnd` range)
    must read as `modified` in the Accept All view of its target; accept-first
    range targets use the exact accepted scope.
    `buildRevisionFidelityError` produces the `REVISION_FIDELITY_VIOLATION`
    payload (`violations[].explanation`, `source.sourceParagraphIndex`,
    `recovery.retryGuidance`). Modes are `enforce` (default, refuse and roll
    back), `warn` (apply and append a warning) and `off`, resolved by
    `resolveRevisionFidelityMode`: explicit `verifyRevisionFidelity` option,
    then the diagnostic `DOCX_REDLINE_REVISION_FIDELITY` environment variable,
    then `enforce`. The view memo is `createFidelityViews(document)`: lazily
    filled `accepted` / `rejected` projections, revision-id signatures and
    numbering references for one DOM. The applier builds a memo for the savepoint
    (taking over `session._fidelityViews` when it still describes the live
    document) and one for the result; after a committed tracked or untracked
    operation the result's memo is stored back in `session._fidelityViews` as
    the next operation's baseline, and a refused operation restores the prior
    memo, so each view is computed at most once per DOM.
- `orchestration/*`
  - Route planning and list fallback orchestration utilities.

## End-to-End Flow

1. Caller imports from `index.js`.
2. Caller configures XML provider/logger/defaults when needed via `adapters/*`.
3. Caller invokes reconciliation APIs (`applyRedlineToOxml`, operation runner, ingestion/export helpers).
4. A full-document batch parses once, captures its initial target snapshot,
   compiles the operations against it (`operation-batch-compiler.js`: batch-start
   source binding, localized replacements, overlap and list-format composition
   checks, all before any mutation), and creates an operation savepoint before
   each mutation.
5. `engine/oxml-engine.js` routes each selected scope to format, table, list,
   surgical, or reconstruction flows; successful scoped output is imported
   into the live document.
6. After each mutation the applier (`services/document-operation-applier.js`)
   validates the generated OOXML against the savepoint, checks the localized
   replacement intent, then runs the revision fidelity verifier (Reject All,
   Accept All intent, neighbor, created-revision and numbering-existence
   invariants); a violation in `enforce` mode restores the savepoint and
   refuses the operation before anything is committed. Failed and no-op
   operations restore their savepoint. A successful changed
   batch serializes the full document once; atomic failures and all-no-op
   batches return the exact input string without serialization.
7. Pipeline/services return OOXML, optional package artifacts (`numberingXml`, comments payloads; numbering parts are merged by `mergeNumberingParts` and
   checked by `validateNumberingIntegrity`), and non-breaking `status`/`error` fields where applicable.
8. Optional revision/comment management transforms can accept/reject revisions, including move revisions, or delete comments by author.
9. Caller writes resulting XML back to package/document boundaries.

## Public Surfaces

- Primary: `index.js`
- Types: `index.d.ts`
- Stable XML operation runner: `@ansonlai/docx-redline-js/standalone-runner`
- Node-only DOCX facade: `@ansonlai/docx-redline-js/node`

Keep public exports centralized through `index.js`; deep imports are supported by
the package `exports` map for advanced consumers, but new public APIs should
still be re-exported from `index.js`.

## Reliability Guardrails

- Create Word namespace elements with `createWordElement(xmlDoc, 'w:...')`.
  Avoid direct `document.createElement('w:*')` or `createElementNS(NS_W, 'w:*')`
  outside `core/word-xml.js`.
- Generate tracked-change metadata through `createRevisionMetadata(author)` so
  `w:id`, `w:author`, and `w:date` stay consistent and document-unique.
- Seed revision IDs from parsed input with `seedRevisionIdsFromDocument(xmlDoc)`
  before emitting new tracked changes.
- Treat revision metadata inside cloned content as identity-bearing state, not
  ordinary formatting. When a run containing `w:rPrChange` is split or cloned,
  preserve an existing revision ID on at most one output run and allocate a
  fresh document-scoped ID for every additional copy.
- Use `containsTrackedChanges(xmlDoc)` and `getTrackedChangeAuthors(xmlDoc)`
  to inspect existing revisions. By default (`existingRevisions: 'merge-same-author'`),
  subsequent edits from the same author are merged against the pre-revision baseline
  while revisions from different authors fail closed with `EXISTING_REVISIONS`.
  Revised paragraphs containing comment anchors fail closed with
  `COMMENTED_CONTENT_MERGE`; never discard anchors as a side effect of restoring
  the pre-revision baseline.
  Pass `existingRevisions: 'accept-all-first'` to normalize prior revisions explicitly.
- Every operation is checked after it is applied
  (`services/revision-fidelity-verifier.js`, default `enforce`): Word's Reject
  All of a tracked output must equal Word's Reject All of the source (`REJECT_*`),
  a plain-text `redline` must read as its `modified` text under Accept All
  (`ACCEPT_INTENT`), and an untracked edit must neither change other paragraphs
  (`ACCEPT_NEIGHBORS`) nor create a revision (`UNTRACKED_CREATED_REVISION`). A
  violation restores the operation savepoint and returns
  `REVISION_FIDELITY_VIOLATION`; `warn` applies and warns, `off` is for
  diagnosis only.
- Allocate, continue, merge and validate numbering only through
  `services/numbering-allocation.js`. Never edit a shared `w:abstractNum` or
  `w:num` in place, never delete a definition Reject All may need, and keep
  `numPr` on tracked-deleted source paragraphs.
- Construct paragraph-mark `w:ins` / `w:del` only through
  `engine/paragraph-mark-revisions.js`, and read current paragraph/run
  properties only through `core/current-properties.js`. Two source-scanning
  tests enforce this: `tests/paragraph_mark_single_site_tests.mjs` rejects
  paragraph-mark revision construction (templates opening an `rPr` with
  `ins`/`del`, DOM creation attached to an `rPr`, retired helper names) outside
  the helper, and `tests/current_properties_usage_tests.mjs` rejects descendant
  lookups of property names (`numPr`, `pStyle`, `b`, ...) outside the accessor.
  Exceptions need an entry with a written reason in each test's allowlist.
- Do not write unknown `result.oxml` payloads directly into `word/document.xml`;
  normalize with `extractReplacementNodesFromOoxml(...)` or use
  `applyOperationToDocumentXml(...).documentXml` for full-document replacement.
- Run `validateRedlineOoxml(oxml)` on generated output before packaging it;
  it reports structural invariant violations as `{ valid, issues }`.

## Integration Contracts

### Targeting and text fidelity

- Paragraph targeting and replacement have different contracts. Targeting may
  use normalized text and fallback heuristics; replacement content is literal
  and must preserve tabs, line breaks, non-breaking spaces, repeated spaces,
  and boundary whitespace.
- Text extraction used for targeting represents the visible accepted view of
  tracked content: insertions are visible and deletions are excluded.
- Paragraph indexes are transient integration references, not user-facing
  document identifiers. Prefer paragraph IDs plus exact text where available.
- Ambiguous matches are unsafe for document mutation. New targeting surfaces
  should report candidate matches or `AMBIGUOUS_TARGET` rather than silently
  choosing the first paragraph.
- Do not remove operation-session savepoints merely to improve throughput.
  Any replacement must prove that a thrown error or false no-op cannot leak a
  partial DOM mutation or revision-ID allocation into later operations.

### Structural paragraph-to-list replacement

- In WordprocessingML, `w:numId w:val="0"` suppresses numbering. It represents
  no reusable list context and must never be returned by numbering allocation
  or assigned to a generated list item. Generated lists use a positive `numId`
  bound to a compatible `w:abstractNum` definition.
- Expanding one non-list paragraph into multiple list paragraphs is a block
  replacement, not an inline replacement. The original content remains in its
  own paragraph with a deleted paragraph mark; every replacement item occupies
  its own paragraph with an inserted paragraph mark and inserted text. This is
  required so Word can accept the edit as the new paragraphs or reject it as
  the exact original paragraph without joining the old heading to item one.
- Inserted list runs selectively inherit source typography (`w:rFonts`, size,
  language, and related script properties). Semantic emphasis such as bold,
  underline, italic, or strike is emitted only when requested by the replacement
  markup; heading emphasis must not leak into ordinary list body text.
- Low-level list output can carry a package-fragment sentinel needed by its
  standalone payload shape. Full-document mutation removes that sentinel while
  importing replacement nodes, so it cannot become an empty document paragraph.
- Tests for this route must verify positive numbering IDs, separate physical
  paragraphs, formatting inheritance/non-inheritance, valid revision markup,
  and exact accepted and rejected paragraph sequences.

### Atomic mixed-content replacement

- `planStructuredReplacement(...)` converts agent-authored Markdown into one
  replacement operation carrying `structuredContent: true`. The block plan is
  diagnostic metadata; execution remains one mutation so replacing the anchor
  cannot invalidate later blocks.
- Mixed-content parsing recognizes explicit Markdown headings, blank-line
  paragraph boundaries, adjacent list items, and contiguous table rows. Tables
  require a header separator and consistent column counts. Invalid input fails
  with `STRUCTURED_CONTENT_INVALID`; silently degrading table pipes into visible
  paragraph text is forbidden on this route.
- Newly inserted tables inside a mixed replacement are tracked once at block
  scope. Cell runs are not nested in additional `w:ins` wrappers, preserving the
  no-nested-revisions invariant while allowing Reject All to remove the table.
- Agents must use explicit heading markers and valid table grammar rather than
  relying on capitalization or layout inference. The engine preserves the
  declared block types and does not guess legal-document semantics.

### Package artifacts and transactions

- A real `.docx` stores document markup, comments, numbering, relationships,
  and content types in separate parts. APIs operating on one XML string cannot
  claim to update artifacts held in another part.
- Comment IDs must be allocated against both document anchors and the existing
  `word/comments.xml` part. Revision IDs must be allocated against the complete
  document revision scope.
- Numbering payloads are merged, never replaced, by the single merge in
  `services/numbering-allocation.js` (`mergeNumberingParts`), which both the
  runner's consumers and the `openDocx` facade call; conflicting ids fail with
  `NUMBERING_ID_CONFLICT`.
- Safe package mutation is transactional: retain the original package, apply
  operations, merge all artifacts, validate redline markup and package wiring,
  and commit only if every required check succeeds.
- `openDocx(buffer)` implements this transaction for Node without adding a ZIP
  dependency to browser or XML-only consumers. Unmodified part contents remain
  byte-identical after extraction, although the ZIP container is reserialized.

### Operation results

- `hasChanges: false` does not imply success. Callers must check `status`,
  `error`, and warnings to distinguish errors, missing anchors, and true no-ops.
- Batch results must preserve the caller's original operation indexes even when
  execution is reordered for stable anchors.
- Author attribution is externally visible document data. Agent-facing APIs
  should require an explicit author and report the author used for every
  operation rather than relying silently on a configured fallback.
- Operation-level authors override the batch author. Runtime results expose
  `authorUsed`, `authorsUsed`, `operationType`, `resolvedBy`, and resolved target
  metadata so integrations can audit what the engine actually selected.
- Target fingerprints are revision-view scoped. Inspection returns a
  `revisionView` beside every paragraph and computes its fingerprint from the
  same text/view pair. Restore normalization defaults omitted target and range
  endpoint views to `rejected`; all other operations default to `accepted`.
  Explicit views remain authoritative, and cross-view text/fingerprint mistakes
  return an actionable mismatch hint rather than silently weakening targeting.
- A normalized target match does not become an edit coordinate system for a
  text-bearing mutation. Mutation uses canonical accepted-view source text, and `resolvedTarget.targetTextMatch`
  records `exact`, `space_equivalent`, or `normalized` selection plus bounded
  invisible-character diagnostics.
- The CLI contract (version 9; `docx-redline version` lists the capabilities) is deliberately narrower than library result objects.
  `apply`, `accept`, `reject`, and `delete-comments` omit `documentXml`, package
  parts, inspection text, and full issue arrays. `completion` is true only when
  `written === true`, top-level status is neither error nor partial, and every
  operation result is non-error. The `validate` command remains the full issue
  reporting surface.
- `preflightOperations` is the read-only safety boundary for agent-generated
  batches. It uses strict targeting by default; mutation APIs retain permissive
  legacy targeting unless `strictTargets: true` is requested. In v1.0.0,
  application will default to strict targeting; in the current warning cycle,
  permissive resolution that chooses among multiple identical paragraphs emits
  `AMBIGUOUS_TARGET_HEURISTIC_USED` with candidate count and migration guidance.
  Both preflight and mutation runners share the same candidate resolver
  (`resolveTargetParagraph`).

### Mutation receipts and output reconciliation oracle

- `ReceiptCollector` tracks exact allocations (revision IDs with kind and target part,
  comment IDs, numbering IDs, relationship IDs, affected targets, and warnings)
  directly at the point of allocation/attachment during operation execution.
- Collector state is snapshotted within operation savepoints. Failed or rolled-back
  operations cleanly restore prior collector state without leaking orphaned allocations.
- Single operations expose `result.receipt`; batches expose per-item `results[i].receipt`
  and top-level `result.receipts`, with dispositions (`applied`, `refused`, `no_change`,
  `rolled_back`, `not_attempted`) and `committed: true/false`.
- The output reconciliation oracle (`reconcileReceiptsAgainstOutput`) performs an
  independent, non-negotiable verification: every durable ID reported as committed
  is parsed from the serialized output XML parts (`word/document.xml`, `word/comments.xml`,
  `word/numbering.xml`). Any discrepancy immediately fails the transaction and triggers
  atomic rollback.
- The oracle checks that durable IDs are committed; the revision fidelity
  verifier checks what Reject All would restore. A refused operation restores
  its savepoint and receipt state like any other refusal, so it contributes no
  committed IDs (`finalDisposition: 'refused'`).


## Build Output

`npm run build` generates CDN-ready ESM bundles under `dist/`:

- `dist/docx-redline-js.esm.js`
- `dist/docx-redline-js.esm.js.map`
- `dist/docx-redline-js.esm.min.js`
- `dist/docx-redline-js.esm.min.js.map`

The bundle inlines `diff-match-patch` and keeps `@xmldom/xmldom` external.

## Testing

- `npm test`
  - Runs all `tests/*.mjs` except setup helpers in separate Node processes with
    a conservative four-worker cap. Use `DOCX_TEST_CONCURRENCY=1` to reproduce
    the same sorted suite serially.
- `npm run benchmark:tests`
  - Compares paired serial and bounded-parallel suite runs and records JSON
    under ignored `tmp/benchmarks/`.
- `npm run test:isolation`
  - Runs boundary checks for Word API markers and dependency-graph isolation.
- `npm run check:types`
  - Smoke-checks `index.d.ts`.
- `node scripts/export-validation-fixtures.mjs`
  - Writes release-time validation fixtures to `tmp/validation-docx/` as
    `word/document.xml` parts, assembled `.docx` files, and expected-text sidecars.
- `tests/roundtrip_fuzz_tests.mjs` (part of `npm test`)
  - Seeded fuzz sweep of the accept/reject round-trip invariant; tune with
    `FUZZ_SEED` / `FUZZ_ITERATIONS`.
- `npm run smoke:word -- path/to/file.docx`
  - Optional Windows/Word COM smoke test for a completed `.docx`.
- `npm run smoke:word:diff`
  - Windows/Word COM differential test: Word itself accepts/rejects the
    generated fixtures and the resulting text is compared to expectations.
    Structural checks also compare paragraph counts, and list cases compare
    `ListFormat.ListString` labels after `Revisions.AcceptAll()` /
    `RejectAll()`.
- `bash scripts/validate-fixtures-xsd.sh`
  - Validates exported fixtures against the ECMA-376 transitional `wml.xsd`.
- `tests/reject_roundtrip_matrix_tests.mjs` (part of `npm test`)
  - Generated operation matrix over fixtures and operation kinds, asserting the
    revision fidelity invariants and receipt reconciliation (a
    `RECEIPT_RECONCILIATION_FAILED` fails the case). Replay with `FUZZ_SEED` /
    `MATRIX_ITERATIONS`.
- `.github/workflows/validation.yml`
  - Nightly independent-oracle validation: XSD schema check, LibreOffice
    conversion, an extended 20k-case fuzz sweep with a fresh seed, and the
    `matrix-extended` job (1000 reject-roundtrip matrix cases with a
    date-derived seed; a failure prints its seed and replay command).

Use these checks before publishing or tagging. See [docs/TESTING.md](./docs/TESTING.md#release-validation-and-independent-oracles).

## Fast Orientation For Contributors

Use this sequence to understand or modify behavior without reading everything:

1. Start at `index.js` to locate the exported API.
2. Follow exports into `engine/oxml-engine.js` or relevant `services/*` module.
3. For targeting bugs, inspect `core/paragraph-targeting.js`, `core/list-targeting.js`, and `core/table-targeting.js`.
4. For package wiring issues, inspect `services/standalone-docx-plumbing.js`.
5. For revision/comment cleanup behavior, inspect `services/revision-comment-management.js`.
6. For numbering/list issues, inspect `services/numbering-allocation.js`,
   `services/list-format*.js`, `services/list-conversion-diagnostics.js` and
   orchestration list-fallback modules; for batch composition,
   `services/operation-batch-compiler.js`; for Reject All refusals,
   `services/revision-fidelity-verifier.js`.
7. For reliability regressions, start with `tests/roundtrip_invariant_tests.mjs`,
   `tests/engine_reliability_tests.mjs`, `tests/paragraph_mark_revision_tests.mjs`,
   `tests/move_revision_tests.mjs`, and `tests/hardening_status_tests.mjs`.
