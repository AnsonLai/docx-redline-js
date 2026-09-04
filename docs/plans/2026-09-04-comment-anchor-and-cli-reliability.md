# Comment Anchor and CLI Reliability Plan

**Status:** Implementation complete — external release deployment pending
**Date:** 2026-09-04  

## Implementation progress

- **Phase 1 complete (2026-09-04):** The CLI now validates per-command
  options, rejects malformed or conflicting selectors, supports `--index`,
  accepts the documented and compatibility range spellings, and reports
  `indexBase: 1`.
- **Phase 2 complete (2026-09-04):** Operation normalization no longer derives
  `textToComment` from paragraph target text. An omitted anchor now uses exact
  canonical text from the resolved paragraph.
- **Phase 3 complete (2026-09-04):** Preflight and application now share a
  non-mutating resolver with exact-match priority, unique ASCII-space/NBSP
  matching, ambiguity detection, raw run offsets, and resolution diagnostics.
- **Phase 4 complete (2026-09-04):** Missing and ambiguous anchors now produce
  structured errors, failed resolution does not allocate comment IDs, and
  atomic package/CLI operations roll back without writing partial output.
- **Phase 5 complete (2026-09-04):** The package now publishes thin
  `extract_text.mjs` and `apply_changes.mjs` compatibility entrypoints that
  delegate to the supported CLI, accept legacy `changes` arrays, preserve the
  legacy positional invocation and author fallback, and have subprocess tests
  for atomic failure and output safety. No local `docx-redline` plugin or
  marketplace entry is installed in this environment, so rollout into the
  reported `/mnt/skills/plugins/docx-redline` installation remains a release
  deployment step rather than an additional implementation.
- **Previous-release deletion regression closed (2026-09-04):** A
  whole-paragraph delete now stops with `COMMENTED_CONTENT_DELETE` when the
  target carries an existing comment, returning comment IDs and package-level
  author/text metadata for follow-up. The transaction remains byte-identical,
  and package validation independently rejects partial ranges, dangling
  usages, orphan definitions, and duplicate definitions. Regression coverage
  includes preflight diagnostics, atomic rollback, metadata preservation,
  point-comment validity, and each malformed comment-linkage shape.

## Summary

A reported agent workflow exposed three related reliability problems:

1. Unsupported or malformed extraction filters can be silently ignored, so an
   agent may receive the wrong paragraph while believing that it selected a
   specific index or range.
2. Paragraph targeting and comment-anchor placement use different whitespace
   rules. A paragraph can resolve successfully through normalized matching but
   its comment anchor can subsequently fail on an ordinary-space versus
   non-breaking-space difference.
3. A missing comment anchor is returned as `no_change` plus a warning. Atomic
   batches only roll back explicit errors, so a mixed batch can write a
   partially modified document and still appear successful.

The report used the legacy `extract_text.mjs` and `apply_changes.mjs` plugin
wrappers, while this repository now provides the `docx-redline` CLI and Node
package facade. The library must be safe independently of those wrappers, and
the wrappers must be migrated so they cannot reintroduce the old behavior.

## Goals

- Make paragraph selection explicit, predictable, and consistently 1-based.
- Make preflight and application resolve comment anchors identically.
- Tolerate safe space-representation differences without changing document
  text or selecting an ambiguous anchor.
- Treat unresolved and ambiguous anchors as operation failures.
- Guarantee byte-exact rollback and no output write for failed atomic batches.
- Preserve intentional no-ops as `no_change` rather than treating every
  unchanged result as an error.
- Provide regression coverage at the unit, service, CLI, DOCX package, and
  agent-workflow levels.

## Non-goals

- Do not split a single OOXML `w:p` into multiple paragraphs because formatting
  makes part of it look like a heading. Inspection must reflect the source
  document structure.
- Do not normalize or rewrite the contents of `w:t` nodes while locating an
  anchor.
- Do not make approximate or fuzzy comment-anchor selection silently succeed.
- Do not change `atomic: false` into all-or-nothing behavior.
- Do not retain two independent batch engines for the package and plugin
  workflows.

## Confirmed failure paths

### Extraction filters

`node/cli.js` currently recognizes `--indexes 2,5,8` and `--range 2:8`.
Unrecognized flags and malformed values can pass through argument parsing
without an error. For example, `--index 7` or `--range 5-9` can result in an
unfiltered extraction rather than a selection failure.

### Anchor whitespace

Paragraph resolution uses `normalizeWhitespaceForTargeting`, but
`findTextInParagraphIndex` uses a literal `String#indexOf`. In addition,
`normalizeDocumentOperation` currently derives `textToComment` from target text
when the caller did not explicitly provide an anchor. This creates a split
contract: the supplied target can resolve a paragraph after whitespace
normalization, then fail when reused as an exact anchor.

The current runner can therefore produce an outcome shaped like:

```json
{
  "hasChanges": false,
  "results": [
    {
      "type": "comment",
      "status": "no_change",
      "warnings": ["Could not find ..."]
    }
  ]
}
```

even though the requested operation was not successfully completed.

### Atomic classification

`applyCommentToParagraphByExactText` returns `hasChanges: false` and warnings
when no comment was placed. `applyOperationsToDocumentXml` only marks a batch
operation as failed when `status === "error"` or an `error` object exists. If
another operation did make a change, the batch can serialize and commit that
partial result despite atomic mode.

Preflight already reports `ANCHOR_NOT_FOUND`, so preflight and apply disagree
about the same input.

## Proposed behavior contract

### CLI filters

- Paragraph indexes and `P<n>` references remain 1-based.
- `--index N` is accepted as the single-paragraph form of `--indexes N`.
- `--range START:END` remains canonical.
- `START-END` and `START,END` may be accepted as compatibility spellings if
  they can be parsed without ambiguity.
- Missing, non-numeric, zero, negative, reversed, or partially numeric values
  return `status: "error"` with `INVALID_FILTER`.
- Unknown flags return `UNKNOWN_OPTION` rather than being ignored.
- Inspection and extraction results state `indexBase: 1`.
- A selection filter that matches no paragraphs succeeds with an empty list;
  a malformed filter fails. These cases must remain distinguishable.

### Comment anchors

- `target` identifies the paragraph; `textToComment` identifies an optional
  sub-paragraph anchor.
- When `textToComment` is omitted, use the exact canonical text of the resolved
  paragraph. Do not copy caller-supplied target text into the anchor field.
- For an explicit anchor, try an exact match first.
- If exact matching fails, allow a mapped space-equivalent match that treats
  ordinary spaces and non-breaking spaces as equivalent. Keep tabs and line
  breaks structurally distinct unless a separate, explicit policy is added.
- Map normalized match offsets back to the original run and character offsets;
  never rewrite the source text to make a match possible.
- If no match exists, return `ANCHOR_NOT_FOUND`.
- If normalization produces more than one candidate, return
  `AMBIGUOUS_ANCHOR` with candidate offsets. Do not select the first candidate.
- Report how an anchor resolved, for example `exact_anchor`,
  `space_equivalent_anchor`, or `whole_paragraph`.

### Operation and batch status

- A comment request that places no comment is an error unless the public API
  explicitly defines an idempotent-comment mode in the future.
- Missing and ambiguous anchors return `status: "error"` and a structured
  error object, not only a warning.
- `no_change` remains valid for genuine successful no-ops such as an identical
  replacement. Do not make every unchanged operation fatal.
- An atomic batch containing an anchor error returns the original
  `documentXml`, `hasChanges: false`, `rolledBack: true`, no comments or
  numbering artifacts, and `BATCH_OPERATION_FAILED` at batch level.
- The package facade returns `written: false`, and the CLI does not create an
  output file.
- With `atomic: false`, successful operations may be retained, but failed items
  remain explicitly marked `error`.

## Implementation plan

### Phase 1: Harden CLI parsing and selection

Update `node/cli.js` to validate commands against per-command option sets and
parse selectors through one dedicated helper.

- Add `--index` as a supported alias.
- Parse and validate all supported range spellings in one place.
- Detect conflicting selectors such as `--index` together with `--range`.
- Reject unsupported flags and invalid selector values before opening or
  inspecting a DOCX.
- Add `indexBase: 1` to inspection/extraction CLI results and document the
  convention in `README.md` and `docs/AGENT-WORKFLOW.md`.

Keep `inspectDocumentParts` focused on structured filtering; user-facing string
syntax belongs at the CLI boundary.

### Phase 2: Separate target normalization from anchor defaults

Update `services/document-operation-contract.js` so a missing
`textToComment` remains missing after normalization. Let
`applyCommentToParagraphByExactText` derive a whole-paragraph anchor from the
resolved paragraph's canonical text.

Preserve compatibility aliases for operation types and target descriptors, but
do not infer a sub-paragraph anchor from `target.exactText`.

### Phase 3: Build a shared anchor resolver

Refactor `services/comment-locator.js` to expose a read-only anchor-resolution
helper used before marker insertion.

The resolver should:

1. Build a canonical paragraph index with source offsets and run ownership.
2. Search exact text and return the unique raw span when found.
3. Build a mapped representation for ordinary-space/NBSP equivalence.
4. Convert a normalized candidate back to raw start and end offsets.
5. Detect multiple candidates.
6. Return a structured result without modifying the DOM.

Marker injection should consume the resolved raw span rather than perform a
second search. Allocate a comment ID only after successful resolution so a
failed attempt cannot consume package-scoped IDs.

Preflight should call the same resolver instead of `paragraphText.includes`.
This makes diagnostics and application agree by construction.

### Phase 4: Propagate structured anchor failures

Update `services/comment-engine.js`,
`services/document-operation-mutations.js`, and
`services/document-operation-applier.js` to preserve structured locator
errors.

- Return `ANCHOR_NOT_FOUND` and `AMBIGUOUS_ANCHOR` with useful diagnostics.
- Restore the per-operation DOM savepoint on either error.
- Ensure failed comment definitions and markers are never added to output.
- Retain warnings for non-fatal information only.

The batch orchestrator should continue to roll back explicit errors. Add a
defensive invariant that a comment operation reporting zero applied comments
cannot be classified as a successful `no_change`.

### Phase 5: Migrate the external plugin wrappers

The reported `extract_text.mjs` and `apply_changes.mjs` files do not live in
this repository. Update their owning plugin after the library changes land.

- Delegate extraction and application to the supported Node facade or CLI.
- Use 1-based indexes and emit `P<n>` references.
- Default to strict targets, atomic batches, and package validation.
- Run preflight before apply or surface equivalent apply-time errors.
- Exit nonzero and omit the output path when an atomic transaction fails.
- Pin a package release containing these fixes and remove duplicated legacy
  targeting logic.

## Regression test plan

The tests below deliberately overlap at different boundaries. A locator unit
test alone will not catch a wrapper that discards its error, and a CLI test
alone will not isolate an offset-mapping regression.

### CLI argument and filter tests

Extend `tests/agent_cli_edge_tests.mjs` with:

- `--index 7` returns only paragraph 7.
- `--indexes 2,5,8` returns exactly those paragraphs in document order.
- Canonical `--range 5:9` includes both endpoints.
- Supported compatibility range spellings return the same indexes.
- `--range 7:7` returns one paragraph.
- Zero, negative, reversed, missing-end, non-numeric, and decimal selectors
  return `INVALID_FILTER`.
- `--index` combined with `--range` returns a clear conflict error.
- A valid range outside the document returns an empty paragraph list.
- A misspelled flag such as `--indxe` returns `UNKNOWN_OPTION` and never falls
  back to an unfiltered extraction.
- `runCli` returns a nonzero exit code for all invalid selector cases.
- Inspection/extraction JSON includes `indexBase: 1` and stable `P<n>` refs.

Add command-specific unknown-option tests so an `apply`-only flag cannot be
silently accepted by `extract`, and vice versa.

### Canonical extraction tests

Extend `tests/canonical_paragraph_text_tests.mjs` and
`tests/document_inspection_edge_tests.mjs` with paragraphs containing:

- ordinary spaces split across differently formatted runs;
- NBSPs before and after an underlined run;
- repeated spaces and leading/trailing spaces;
- tabs, line breaks, soft hyphens, and non-breaking hyphens;
- comments and tracked revisions crossing run boundaries.

Assert that extraction preserves the exact source characters and never
synthesizes NBSPs because formatting changes between runs. Verify accepted,
rejected, and current revision views independently.

### Anchor resolver unit tests

Extend `tests/comment_tests.mjs` or add a focused
`tests/comment_anchor_locator_tests.mjs` covering:

- exact anchor within one run;
- exact anchor spanning two or more formatted runs;
- anchor beginning or ending at a run boundary;
- ASCII spaces matching source NBSPs;
- source ASCII spaces matching anchor NBSPs;
- multiple adjacent space characters with correct raw offset mapping;
- punctuation or letter differences remaining unmatched;
- tabs and line breaks not being silently treated as ordinary spaces;
- duplicate exact anchors returning `AMBIGUOUS_ANCHOR` if uniqueness is
  required by the selected policy;
- duplicate space-equivalent anchors returning `AMBIGUOUS_ANCHOR`;
- missing anchors returning `ANCHOR_NOT_FOUND`;
- failed resolution leaving the paragraph DOM byte-equivalent after
  serialization;
- successful insertion preserving formatting, hyperlinks, bookmarks,
  revision wrappers, and existing comment markers around the anchor;
- run splitting continuing to allocate unique `w:rPrChange` IDs.

For every successful case, inspect the actual text enclosed by
`w:commentRangeStart` and `w:commentRangeEnd`, not merely the presence of marker
elements.

### Operation contract and preflight parity tests

Extend `tests/agent_operation_contract_tests.mjs` with:

- a comment target with no `textToComment` anchors the exact resolved
  paragraph, even when caller target whitespace differs;
- an explicit space-equivalent anchor resolves identically in preflight and
  apply;
- a missing anchor produces `ANCHOR_NOT_FOUND` in both paths;
- an ambiguous anchor produces `AMBIGUOUS_ANCHOR` in both paths;
- `resolvedBy`, resolved raw offsets, and candidate diagnostics agree between
  preflight and apply;
- a genuine identical replacement remains `no_change` and is not mislabeled
  as an error.

Use a table-driven parity test that runs each anchor scenario first through
`preflightOperations` and then through `applyOperationToDocumentXml`.

### Atomic batch tests

Extend `tests/standalone_operation_runner_tests.mjs` with:

- seven valid operations plus one missing comment anchor;
- one successful replacement followed by a failed comment;
- a failed comment scheduled before a replacement;
- failures at the first, middle, and last logical input indexes;
- `continueOnError: true` still reporting every attempted operation while the
  atomic transaction rolls back;
- `continueOnError: false` stopping after the first failure and rolling back;
- `atomic: false` retaining valid operations while reporting the comment as
  `error`;
- an all-no-op, genuinely successful batch not being treated as failed;
- rollback clearing comment and numbering artifacts and leaving runtime target
  snapshots uncommitted;
- input and returned rollback XML being byte-identical.

Assert both original result order and execution order because comments are
scheduled before replacements.

### Package facade and CLI transaction tests

Extend `tests/docx_package_facade_tests.mjs`,
`tests/docx_package_transaction_edge_tests.mjs`, and
`tests/agent_cli_tests.mjs` with:

- a mixed batch whose one missing anchor yields `written: false`;
- `toBuffer()` returning bytes identical to the original DOCX after rollback;
- no output file being created by CLI apply after an atomic failure;
- an existing requested output file remaining untouched after failure;
- no orphaned comment relationship, content-type entry, marker, or definition;
- retrying after a failed comment not skipping or colliding comment IDs;
- existing comments with high IDs remaining intact;
- unrelated ZIP entries remaining byte-identical;
- CLI JSON reporting batch error, failed item index, error code, and
  `outputPath: null`;
- `runCli` returning nonzero for the failed transaction.

### End-to-end reported-case fixture

Create a minimal, reviewable DOCX fixture modeled on the report:

- at least ten paragraphs so selector behavior is observable;
- a target paragraph containing multiple runs and underlined text;
- NBSPs around `Subscription Term`;
- a heading-looking run at the end of the same OOXML paragraph;
- one existing comment elsewhere in the document;
- an operation file with seven valid changes and one whitespace-mismatched
  comment anchor.

Exercise the complete workflow:

1. Extract the target using `--index` and `--range`.
2. Verify exact text and 1-based references.
3. Preflight the operation file.
4. Apply it atomically.
5. Validate the output package.
6. Reinspect comments and verify the intended anchored text.

Include two variants:

- a safely space-equivalent anchor that should apply all eight operations;
- a genuinely absent anchor that must roll back all eight operations and write
  no output.

### Property and fuzz tests

Add bounded randomized coverage for the shared anchor resolver:

- generate paragraph text split across random run boundaries;
- randomly replace ordinary spaces with NBSPs in either source or anchor;
- retain an oracle mapping from logical anchor characters to raw offsets;
- assert that successful resolution returns the oracle span;
- insert duplicate anchors and assert ambiguity rather than first-match
  selection;
- mutate one non-space character and assert no match;
- verify resolver calls do not mutate their input DOM.

Use fixed seeds and print the seed with failures so every case is reproducible.

### Type and schema tests

Update declarations and schemas if new anchor diagnostics are public:

- `ANCHOR_NOT_FOUND` and `AMBIGUOUS_ANCHOR` result shapes compile correctly;
- `resolvedBy` accepts the new anchor-resolution values;
- operation schemas continue to make `textToComment` optional;
- examples with descriptor-only whole-paragraph comments type-check;
- examples with explicit sub-paragraph anchors type-check.

### Performance tests

Extend the existing session benchmarks or performance boundary tests to verify:

- anchor indexing remains linear in paragraph size;
- exact matches do not build the fallback mapped representation unnecessarily;
- batches with many comments reuse paragraph indexes where safe;
- preflight and application do not repeatedly serialize the entire document;
- ambiguity detection remains bounded for long repeated paragraphs.

Set generous regression thresholds appropriate for CI; the purpose is to catch
accidental quadratic scans, not microbenchmark noise.

## Validation commands

During implementation, run the focused suites after each phase:

```bash
node tests/agent_cli_edge_tests.mjs
node tests/canonical_paragraph_text_tests.mjs
node tests/comment_tests.mjs
node tests/agent_operation_contract_tests.mjs
node tests/standalone_operation_runner_tests.mjs
node tests/docx_package_facade_tests.mjs
node tests/docx_package_transaction_edge_tests.mjs
```

Before release, run:

```bash
npm test
npm run test:isolation
npm run check:types
npm run lint
node scripts/export-validation-fixtures.mjs
```

If Word is available on Windows, also run the generated fixture through:

```bash
npm run smoke:word -- path/to/comment-anchor-regression.docx
```

Manually confirm that Word opens the output without repair, displays all
expected comments, selects the intended text for each comment, and accepts or
rejects unrelated tracked changes correctly.

## Acceptance criteria

- Invalid or unknown extraction filters can never silently return an
  unfiltered document.
- All documented paragraph indexes and references are consistently 1-based.
- Whole-paragraph comments use resolved document text, not reconstructed caller
  text.
- Safe ASCII-space/NBSP anchor differences resolve without modifying source
  text.
- Missing and ambiguous anchors are structured errors in both preflight and
  apply.
- No failed comment operation is reported only as `no_change` plus a warning.
- Atomic batches with any failed anchor return and preserve the exact original
  DOCX bytes and create no output file.
- Intentional no-ops remain supported.
- The legacy plugin wrappers use the supported facade/CLI and surface the same
  failure semantics.
- Focused, full, isolation, type, lint, package-validation, and optional Word
  smoke tests pass.

## Rollout

1. Land the library and CLI changes with the reported-case fixture.
2. Publish a new package version and document the status-contract change.
3. Update and republish the external docx-redline plugin/skill wrappers with the
   new package pinned.
4. Run the reported workflow against the installed plugin, not only the source
   checkout.
5. Keep compatibility selector spellings for at least one release cycle, while
   documenting `--index`, `--indexes`, and colon-delimited `--range` as the
   canonical interface.
