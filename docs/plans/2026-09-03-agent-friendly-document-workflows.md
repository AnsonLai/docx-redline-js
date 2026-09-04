# Agent-Friendly Document Workflow Plan

**Status:** Implementation complete — external release validation noted below  
**Date:** 2026-09-03  

This plan turns integration lessons from a real `docx-redline-js` skill into a
supported product surface. The skill successfully uses the engine, but it also
has to reimplement document inspection, comment reading, package transactions,
visible provision references, per-operation authorship, and safe command-line
workflows. Those are recurring integration responsibilities rather than
skill-specific business logic.

The core reconciliation engine should remain host-independent and focused on
OOXML. Agent-facing `.docx` file handling should be added as a separate adapter
or companion CLI so the core does not acquire a mandatory ZIP implementation.

## Baseline

As of 2026-09-03:

- Paragraph/range reconciliation is available through `applyRedlineToOxml`.
- Full `word/document.xml` operations are available through the
  `@ansonlai/docx-redline-js/standalone-runner` export.
- Atomic batches reorder comments ahead of text edits, but accept one author
  for the entire batch.
- The public declarations do not define discriminated operation types or the
  standalone runner functions, and comment request arrays are typed as
  `unknown[]`.
- Targeting supports transient paragraph references and text heuristics, but
  text-only duplicate matches can resolve to the first candidate.
- There is no supported structured document-inspection or comment-reading API.
- Package callers must merge comments and numbering, allocate identifiers, and
  invoke validation correctly themselves.
- `ensureNumberingArtifactsInZip` replaces existing numbering unless the caller
  supplies `mergeNumberingXmlBySchemaOrder`.
- Revision-safe cloning work for comment/highlight run splitting is in progress
  separately and is not re-scoped by this plan.

This plan overlaps with
`2026-09-01-performance-and-complexity-reduction.md`. Implement the operation
schema and batch facade against the modularized/in-memory runner from that plan
where practical; do not create a second batch engine.

## Compatibility strategy

| Phase | Expected impact |
|---|---|
| 1. Public contracts | Additive. Existing operation objects and runner signatures remain supported. |
| 2. Targeting and preflight | Additive by default. Strict ambiguity handling is opt-in until the next major version. |
| 3. Structured inspection | Additive. New read-only APIs. |
| 4. Transactional package facade | Additive companion surface. Core keeps no mandatory ZIP dependency. |
| 5. Safe defaults and CLI | Mostly additive. Changing numbering replacement behavior requires a deprecation period or major release. |

---

## Phase 1 — Typed operations and per-operation attribution

**Status:** Complete (2026-09-03)

### Problem

The runner's operation field names are discoverable mainly from implementation
code. Its batch author applies to every operation, forcing consumers that need
different redline and comment authors to rebuild scheduling and atomicity.

### Work

1. Define and export discriminated types for redline, comment, and highlight
   operations, runner options, per-operation results, and batch results.
2. Export runner declarations from the stable `./standalone-runner` package
   subpath and consider re-exporting the primary runner functions at the root.
3. Add optional `author` to each operation. Resolution order should be explicit:
   operation author, batch default author, configured legacy fallback.
4. Include `authorUsed`, `resolvedBy`, and resolved target metadata in every
   operation result.
5. Add runtime operation validation with stable `INVALID_OPERATION` diagnostics
   for missing or incompatible fields.
6. Keep the existing operation names (`modified`, `textToComment`, and
   `commentContent`) compatible. If ergonomic aliases are introduced, normalize
   them once at the API boundary and document one canonical representation.

### Acceptance

- TypeScript catches a comment using redline-only fields and vice versa.
- One atomic batch can emit revisions and comments under different authors.
- Existing JavaScript callers using the batch-level author continue to work.
- Every result identifies the author and target actually used.

### Implementation notes

- Added discriminated operation/result declarations at the stable
  `./standalone-runner` package subpath and exported the reusable types from the
  root declaration file.
- Added runtime `INVALID_OPERATION` validation while retaining `replace`,
  `format`, `list-change`, `table-reconciliation`, `insert`, and `delete` as
  compatibility labels for redline operations.
- Added per-operation author resolution, `authorUsed`, `authorsUsed`, canonical
  `operationType`, `resolvedBy`, and resolved target metadata.
- Preserved the batch-level author as the fallback for existing callers.
- Verified with `tests/agent_operation_contract_tests.mjs`, the existing
  standalone runner tests, and `npm run check:types`.

---

## Phase 2 — Safe target descriptors and preflight

**Status:** Complete (2026-09-03)

### Problem

Text-only targeting is convenient but unsafe when boilerplate paragraphs repeat.
Agents also need to discover existing-revision conflicts and invalid anchors
before mutating a document.

### Work

1. Introduce a target descriptor supporting exact text, paragraph ID, transient
   index, table/list context, occurrence, and an optional source fingerprint.
2. Add `AMBIGUOUS_TARGET` with candidate metadata. Provide a strict targeting
   option immediately and make strict behavior the default in the next major
   version.
3. Return normalization diagnostics when the matched text differs from the
   supplied target in whitespace or punctuation.
4. Add `preflightOperations(documentParts, operations, options)` that performs no
   mutation and reports:
   - missing and ambiguous targets;
   - anchor presence and candidate ranges;
   - existing revisions on affected paragraphs;
   - incompatible operations against the same target;
   - authors that will be written;
   - required comments and numbering artifacts.
5. Define ordering/conflict behavior for comment, highlight, and replacement
   operations aimed at the same paragraph. Do not rely only on global type
   priority when one operation changes another operation's target or revision
   policy.

### Acceptance

- Strict mode never silently edits the first of multiple exact matches.
- Preflight results are deterministic, serializable JSON and map one-to-one to
  original operation indexes.
- A consumer can resolve all target and author questions before applying edits.

### Implementation notes

- Added target descriptors for exact text, paragraph ID, index, occurrence,
  table context, and deterministic source fingerprint.
- Added opt-in `strictTargets` application behavior with `AMBIGUOUS_TARGET` and
  no fuzzy fallback. Legacy application remains permissive by default.
- Added read-only `preflightOperations`, strict by default, with target and
  anchor diagnostics, existing-revision policy checks, author reporting,
  artifact requirements, and same-paragraph operation conflicts.
- Added candidate metadata to ambiguity errors so callers can choose a stable
  disambiguator and retry.
- Added focused coverage for duplicate exact text, occurrence and fingerprint
  selection, stale fingerprints, paragraph IDs, missing anchors, existing
  revisions, artifact prediction, operation conflicts, per-operation authors,
  and runtime validation.
- Verified the completed phases with 47/47 test files, type checks, dependency
  isolation, changed-file lint, and `git diff --check`.

---

## Phase 3 — Structured document inspection

**Status:** Complete (2026-09-03)

### Problem

Consumers currently reconstruct paragraph inventories, accepted-view text,
comment anchors, headings, and visible list numbers themselves. This produces
multiple subtly different definitions of document text and target identity.

### Work

1. Add a read-only inspection API over document parts that returns paragraphs
   in document order with:
   - exact canonical visible text;
   - paragraph ID/path and transient index;
   - table-cell and list context;
   - heading/style context;
   - existing revision and comment authors;
   - whether revision markup intersects the paragraph.
2. Add a comment reader that joins `word/comments.xml` definitions to document
   anchors and reports exact anchored text in document order.
3. Centralize accepted/rejected/current-view text semantics, including moves,
   tabs, breaks, fields, notes, and other structural content. Reuse the canonical
   extractor in targeting and ingestion rather than maintaining parallel walkers.
4. Add optional visible numbering resolution from `numbering.xml`, including
   levels, overrides, restarts, and section-continuation cases covered by tests.
5. Produce human-facing references from provision number, enclosing heading,
   and a quotable excerpt. Treat computed numbering as advisory and always retain
   the excerpt.
6. Support filtering by index/range, text search, revision presence, table
   context, and non-empty content without requiring callers to post-process a
   full-document dump.

### Acceptance

- Inspection output can supply target and replacement source strings without
  losing tabs or boundary whitespace.
- Comment definitions, ranges, authors, and anchor text round-trip on existing
  commented documents.
- Targeting and inspection use the same canonical text representation.
- Provision references are verified against fixtures with numbering restarts
  and overrides.

### Implementation notes

- Added `inspectDocumentParts(...)` with exact text, target identity,
  heading/table/list context, revision authors, comment joins, and filters.
- Centralized accepted/rejected-view text and reused it in targeting and
  ingestion, including moves, tabs, breaks, and hyphens.
- Added advisory numbering resolution for level formats and start/level
  overrides while retaining exact excerpts.
- Added focused coverage in `tests/document_inspection_tests.mjs`.

---

## Phase 4 — Transactional package facade

**Status:** Complete (2026-09-03)

### Problem

The XML runner cannot safely own concerns that span `document.xml`,
`comments.xml`, `numbering.xml`, relationships, and content types. Callers must
currently compose those steps correctly and preserve the original package for
rollback.

### Work

1. Add a separate Node adapter or companion package with an API similar to:

   ```js
   const document = await openDocx(inputBuffer);
   const inspection = await document.inspect(options);
   const result = await document.applyOperations(operations, {
     author: 'Editor',
     atomic: true,
     validate: true
   });
   const outputBuffer = await result.toBuffer();
   ```

2. Seed comment IDs from both document anchors and the existing comments part.
   Keep comment IDs and revision IDs in explicit document/package-scoped
   allocators.
3. Merge numbering with schema-order-safe behavior by default and preserve all
   existing definitions.
4. Update comment definitions and anchors together for add/delete operations.
5. Treat apply, artifact merge, relationship/content-type wiring, and validation
   as one transaction. Atomic failure returns the untouched input buffer.
6. Validate both the original package and generated package so diagnostics can
   distinguish pre-existing defects from newly introduced ones.
7. Return a complete report containing operation results, authors used,
   artifacts changed, validation issues, execution order, and whether output was
   written.

### Acceptance

- Callers can safely edit a `.docx` without manually opening or rewriting ZIP
  parts.
- Existing comments and numbering survive new comments and lists.
- A generated validation failure cannot produce a writable partial result in
  atomic mode.
- No ZIP dependency becomes mandatory for browser or XML-only core consumers.

### Implementation notes

- Added the isolated `@ansonlai/docx-redline-js/node` entry point with
  `openDocx`, inspection, preflight, batch application, and serialization.
- Package mutation clones parts, allocates comment IDs above existing package
  IDs, merges numbering by schema order, updates OPC wiring, validates, and
  commits only on success.
- Results report `written`; atomic failures return the original input bytes.
  Unmodified entry contents remain byte-identical after extraction.
- Added coverage in `tests/docx_package_facade_tests.mjs`.

---

## Phase 5 — Supported agent CLI and safer defaults

**Status:** Complete (2026-09-03)

### Problem

Agents are more reliable with small JSON-producing commands than with ad hoc ZIP
scripts or large XML dumps. The real skill demonstrates a useful command set,
but every skill author should not need to maintain a private integration layer.

### Work

1. Build a CLI over the package facade with commands for `inspect`, `extract`,
   `preflight`, `apply`, `accept`, `reject`, `delete-comments`, and `validate`.
2. Make output structured JSON with stable error codes and meaningful process
   exit codes. Provide built-in range/search/revision/table filters.
3. Require explicit edit and comment authors for mutating agent workflows, while
   retaining legacy fallback behavior in low-level APIs.
4. Accept a documented JSON operation file and provide a published JSON Schema.
5. Preserve input files by default; require an explicit option for in-place
   mutation.
6. Deprecate silent numbering replacement. In the next major version, merge by
   default or throw if safe merging is unavailable.
7. Publish a compact agent workflow guide and use it as the source for future
   skills instead of vendoring integration logic independently.

### Acceptance

- A skill can consist mainly of workflow and review guidance, with no custom ZIP
  or OOXML manipulation code.
- Exact paragraph text can flow from `extract` to an operation file without
  whitespace normalization.
- CLI failures never report a mutation as successful and never overwrite the
  input by default.
- The CLI works on supported Node versions across Windows, macOS, and Linux.

### Implementation notes

- Added the `docx-redline` executable with JSON `inspect`, `extract`,
  `preflight`, `apply`, `accept`, `reject`, `delete-comments`, and `validate`.
- Added filters for ranges, indexes, search, revision presence, table/body
  context, empty paragraphs, and accepted/rejected views.
- Mutations require explicit attribution, are transactional, preserve the input
  by default, refuse existing destinations, and require `--in-place` or
  `--force` for the corresponding destructive intent.
- Added `docs/schemas/document-operations.schema.json` and the compact
  `docs/AGENT-WORKFLOW.md` source for future skills.
- Deprecated low-level silent numbering replacement with a runtime warning and
  documented the next-major error behavior. The facade and CLI merge today.
- Added `tests/agent_cli_tests.mjs` covering every command family and the
  recommended end-to-end workflow.

## Testing and release gates

Each phase must add focused unit tests plus package-level fixtures where relevant.
Before release:

- Run `npm test`, `npm run test:isolation`, `npm run check:types`, and
  `npm run lint`.
- Add duplicate-target, existing-comment-ID, multi-author, numbering-merge, and
  atomic-validation regression cases.
- Add Word differential cases for comments on previously revised runs, multiple
  existing comment authors, and lists added to documents with existing numbering.
- Confirm accept/reject round trips and Word-open-without-repair behavior.
- Document any default change in the changelog and provide a migration example.

### Phase 1–2 verification record (2026-09-03)

- `npm test`: 47/47 test files pass.
- `npm run check:types`: pass; the stable `./standalone-runner` self-import is
  exercised by the TypeScript usage fixture.
- `npm run test:isolation`: pass.
- ESLint on all Phase 1–2 implementation files: pass.
- `git diff --check`: pass.
- The repository-wide `npm run lint` remains blocked by pre-existing unused
  variables in `scripts/render-agenda-multilevel.mjs` and
  `scripts/render-multilevel-cases.mjs`; these files were outside this plan and
  were not changed.

### Phase 3–4 verification record (2026-09-03)

- `npm test`: 49/49 test files pass, including structured inspection and real
  DOCX-buffer transaction regressions.
- `npm run test:isolation`: pass; the Node facade is a separate export and the
  root/browser dependency graph remains host-independent.
- `npm run check:types`: pass, including declarations for both new surfaces.
- `git diff --check`: pass.

### Phase 5 and final plan audit (2026-09-03)

- `npm test`: 50/50 test files pass. The final suites cover duplicate targets,
  per-operation authors, existing high comment IDs, numbering preservation,
  exact comment anchors, coordinated comment deletion, atomic validation
  rollback, CLI safety, and accept/reject transforms.
- `npm run test:isolation`: pass. Node filesystem, ZIP, and DOM dependencies
  remain confined to the separate Node surface.
- `npm run check:types`: pass; 106 root runtime exports have declarations and
  the self-import fixtures cover the standalone and Node subpaths.
- `npm run lint`: pass. The five pre-existing unused-variable failures in two
  one-off rendering scripts were removed during the final release-gate audit.
- `npm run build`: pass.
- `npm pack --dry-run`: pass and includes the executable, Node facade, agent
  guide, and operation JSON Schema.
- `git diff --check`: pass; only line-ending conversion notices are emitted.

| Phase | Final acceptance audit |
|---|---|
| 1. Public contracts | Complete: discriminated types, validation, attribution, and result metadata are implemented and tested. |
| 2. Targeting/preflight | Complete: strict ambiguity, stable descriptors, diagnostics, conflicts, and artifact prediction are implemented and tested. |
| 3. Inspection | Complete: canonical text, exact comment anchors, styles/headings, table-cell coordinates, structural references, advisory numbering, human references, and filters are implemented and tested. |
| 4. Package facade | Complete: add/delete comment coordination, scoped IDs, numbering preservation, validation comparison, reports, and byte-safe rollback are implemented and tested. |
| 5. Agent CLI | Complete: all planned commands, JSON output, explicit attribution, schema, safe paths, deprecation, and workflow guidance are implemented and tested. |

The implementation plan is complete. Desktop Word differential/visual runs and
multi-OS Node execution remain release-environment checks rather than local
implementation blockers. The existing Word, corpus, XSD, LibreOffice, and CI
lanes documented in `docs/TESTING.md` and `docs/VALIDATION.md` should run before
publishing a release. No local Word or independent-office validation result is
claimed by this audit.

## Non-goals

- Moving ZIP or filesystem dependencies into the host-independent core.
- Editing headers, footers, text boxes, or notes before the corresponding core
  document-part support is designed and tested.
- Treating paragraph indexes or computed provision numbers as permanent document
  identifiers.
- Automatically accepting existing revisions merely to make an operation pass.
