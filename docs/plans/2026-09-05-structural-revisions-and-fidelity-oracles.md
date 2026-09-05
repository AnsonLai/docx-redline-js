# Structural Revisions, Stable Addressing, and Fidelity Oracles

**Status:** Proposed — implementation-ready specification  
**Date:** 2026-09-05  
**Target releases:** v0.5.0–v1.0.0  
**Priority:** Correct accepted/rejected results and document fidelity over throughput or review-pane cosmetics

## 1. Purpose

The package already performs surgical text redlining, paragraph-mark revision
generation, tracked run formatting, structured-content planning, strict
preflight, atomic live-DOM batches, and package validation. The next releases
should harden and extend those capabilities without creating a second mutation
model or weakening existing rollback and targeting guarantees.

This plan focuses on the remaining gaps:

1. Prove what is preserved before expanding mutation behavior.
2. Give callers precise revision-view addressing without pretending that All
   Markup is one unambiguous string.
3. Make replacement metadata and insertion affinity explicit while preserving
   structural containers.
4. Complete paragraph-boundary and formatting-revision semantics already
   partially implemented.
5. Add package concurrency guards, dependent batch steps, and commit-aware
   receipts.
6. Promote strict application targeting only through a documented major-version
   migration.

The plan does not introduce an editor tree, layout engine, HTML round-trip, or
second document authority. The OOXML package and the current live DOM remain
authoritative.

## 2. Current Baseline

Before implementation, tests and documentation must acknowledge what exists.
Treat these as foundations to extend, not greenfield work:

- `engine/run-builders.js` creates `w:ins`, `w:del`, paragraph-mark revisions,
  and `w:rPrChange` snapshots.
- `services/revision-comment-management.js` accepts and rejects text, move,
  paragraph-mark, and property-change revisions.
- `core/paragraph-text.js` already extracts accepted and rejected revision
  views.
- `services/operation-preflight.js` is strict by default and reports ambiguous
  targets and anchors.
- The Node facade and CLI default to strict target application; the lower-level
  application API remains permissive for compatibility.
- `DocumentOperationSession` owns one live DOM and one document-scoped
  `RevisionIdAllocator`, with per-operation DOM and allocator savepoints.
- Batch execution uses an immutable start-of-batch target-reference snapshot,
  comment-first scheduling, and one final serialization.
- `planStructuredReplacement` validates headings, paragraphs, lists, and tables
  before a large structural insertion is applied.
- Word differential, visual, package, round-trip, and real-document tests are
  already part of the verification stack.

### Required baseline audit

Create a short capability matrix before feature work begins. For each proposed
behavior, label it `implemented`, `partial`, `missing`, or `intentionally
unsupported`, and link the production code and regression test. At minimum the
matrix must cover:

- text insertion, deletion, and replacement;
- paragraph insertion, deletion, split, and boundary removal;
- run and paragraph property revisions;
- hyperlinks, bookmarks, comments, fields, SDTs, tabs, breaks, and notes;
- lists and table structural revisions;
- accepted/rejected text extraction;
- target resolution, batch savepoints, and package artifact wiring.

No phase may duplicate an implemented path under a new option without a written
migration and compatibility reason.

## 3. Non-Negotiable Invariants

### 3.1 Lifecycle correctness

For every generated revision:

- Accept All produces the requested final document.
- Reject All restores the intended source document.
- Selective author acceptance/rejection does not consume another author's work.
- Repeating acceptance or rejection is idempotent.
- Paragraph count, order, list membership, tables, comments, and section
  ownership are checked in addition to plain text.

### 3.2 Preservation envelope

“Untouched” has three distinct meanings and must not be conflated:

1. **Untouched package entry:** compare the uncompressed entry payload
   byte-for-byte. ZIP container bytes are not expected to match after repacking.
2. **Untouched subtree inside a modified XML part:** compare using a
   namespace-aware canonical structural digest.
3. **Intentionally modified region:** validate against an operation-specific
   mutation envelope and lifecycle oracle.

An operation must declare which XML nodes, ancestor containers, relationships,
content types, and companion parts it is permitted to change. Unexpected drift
outside that envelope fails the test.

### 3.3 Structural safety over presentation heuristics

Bookmarks, comment ranges, hyperlinks, fields, SDTs, proofing markers, drawings,
and revision boundaries must not be moved or deleted merely to influence how
Word groups a change in the Reviewing Pane.

### 3.4 Refusal over guessing

Ambiguous targets, anchors, capture selections, revision views, or structural
ownership must produce structured errors. Fuzzy or first-match fallback is a
legacy compatibility behavior, never an agent recommendation.

### 3.5 One mutation authority

All public package operations continue through the existing session, allocator,
savepoint, artifact, and validation paths. New functionality must not introduce
parallel ID allocation, string-only package mutation, or a second batch state.

## 4. Compatibility Matrix

| Initiative | Initial behavior | Classification | Default change |
|---|---|---|---|
| Fidelity and mutation-envelope oracles | Test-only | Non-breaking | None |
| Explicit accepted/rejected addressing | Additive option; omitted means current behavior | Non-breaking | None |
| Preserving existing revisions during addressed edits | New opt-in policy | Additive, high-risk | No automatic default |
| Paired replacement metadata | Opt-in serialization policy | Observable semantic change | Consider only after Word evidence |
| Explicit insertion affinity | Additive; omitted means legacy resolution | Non-breaking | No implicit `left` default |
| Completed paragraph-boundary handling | New support for previously refused/fallback cases | Internal semantic change | Route only proven cases |
| Author-aware formatting coalescing | Opt-in policy | Observable semantic change | Preserve current behavior initially |
| Package revision token | Additive precondition | Non-breaking | No check when omitted |
| Capture chaining | Additive batch feature | Non-breaking for independent batches | Dependency scheduling only when used |
| Mutation receipts | Additive result field | Non-breaking | Existing result fields remain authoritative |
| Strict lower-level application by default | Default behavior change | Breaking | v1.0.0 only |

“Non-breaking” means more than an unchanged TypeScript signature. Changes to
generated OOXML, formatting inheritance, accepted/rejected structure, revision
count, author attribution, warnings, or Word presentation are observable and
must be tested and documented.

## 5. Milestone 0 — Fidelity Oracle Foundation

**Target:** v0.5.0, before new mutation semantics  
**Classification:** Non-breaking, test and validation infrastructure

### 5.1 Part inventory and mutation envelopes

Add a fixture helper that inventories every package entry and records:

- entry URI and SHA-256 of its uncompressed bytes;
- content type and relationship ownership;
- XML root namespace and canonical digest for XML parts;
- known volatile metadata that a test intentionally ignores.

Each operation family declares allowed effects. Examples:

- Text-only body edit: `word/document.xml` only.
- List creation: document XML, numbering XML, numbering relationship, and
  content type when absent.
- Comment creation: document XML, comments XML, comments relationship, and
  content type when absent.
- Header edit: selected header part only, plus artifacts explicitly required by
  that edit.

The oracle must fail when an undeclared part changes, even if the resulting DOCX
opens successfully.

### 5.2 Canonical subtree comparison

Do not implement canonicalization by simple prefix replacement. Use a proven XML
canonicalization implementation where possible. If a local OOXML subset is
necessary, it must:

- compare expanded names, not literal prefixes;
- sort attributes by namespace URI and local name;
- retain child order;
- preserve text, tabs, breaks, and `xml:space` semantics;
- preserve or correctly rewrite QName-valued attributes such as
  `mc:Ignorable`;
- distinguish absent properties from explicit off-values when Word does;
- normalize only syntax that is semantically equivalent.

Untouched subtree identity cannot rely solely on `w14:paraId`: IDs may be absent
or duplicated in real documents. Match nodes using a combination of stable IDs,
canonical fingerprints, ancestor context, and ordered-neighbor evidence. Exclude
the mutation target, its structurally affected ancestors, and any explicitly
declared sibling boundary from the untouched set.

### 5.3 Required oracles

Add tests for:

- accepted and rejected text and structure;
- revision ID uniqueness across every revision element supported by validation;
- comment and numbering artifact integrity;
- untouched entry byte equality;
- untouched subtree canonical equality;
- save-and-reopen semantic equality;
- no-op and atomic rollback returning the exact original XML/package bytes;
- Word open, Accept All, Reject All, and save/reopen behavior.

Reviewing Pane grouping is a visual compatibility observation, not a schema or
lifecycle oracle.

### Acceptance gate

No later phase begins until representative paragraph, list, table, comment,
header/footer, field, hyperlink, bookmark, drawing, and SDT fixtures pass the
new preservation envelope.

## 6. Milestone 1 — Revision-View Addressing

**Target:** v0.5.0  
**Classification:** Additive, with mutation support gated separately

### 6.1 Use semantic view names

Expose revision addressing as:

```ts
export type RevisionView = 'accepted' | 'rejected';

export interface ParagraphTargetDescriptor {
  exactText?: string;
  paragraphId?: string;
  index?: number;
  occurrence?: number;
  fingerprint?: string;
  revisionView?: RevisionView; // Defaults to accepted/current behavior.
}
```

Do not call accepted text “No Markup or Original.” Original/rejected and
accepted/final are different views. Do not expose a single `tracked` string made
by concatenating deleted and inserted text; a replacement has two overlapping
histories, not one stable user-facing coordinate line.

For inspection that needs revision history, return segments instead:

```ts
interface RevisionTextSegment {
  text: string;
  kind: 'baseline' | 'insertion' | 'deletion' | 'move_from' | 'move_to';
  author?: string;
  revisionId?: string;
  acceptedStart: number | null;
  rejectedStart: number | null;
}
```

Offsets use JavaScript UTF-16 code units. DOM mapping must cover `w:t`,
`w:delText`, tabs, breaks, carriage returns, soft/no-break hyphens, entities,
and revision/move ancestry.

### 6.2 Separate inspection from mutation policy

Reading or locating text in a rejected view does not automatically authorize
editing historical deletion markup. Initial scope:

- Inspection and preflight support both views.
- Comments may target a supported revision range only after schema and Word
  round-trip fixtures prove the anchor representation.
- Text mutation of another pending deletion or insertion is refused until an
  explicit author-aware lifecycle policy exists.
- The existing `reject-input` and `accept-all-first` policies remain unchanged.
- If a future `preserve-input` policy is added, it must prohibit illegal nested
  revisions and define same-author coalescing separately.

### Acceptance criteria

- The same descriptor resolves deterministically in accepted and rejected
  views.
- Duplicate text remains ambiguous within the selected view.
- Emoji, entities, tabs, breaks, moves, and mixed insert/delete replacements map
  to the correct DOM boundaries.
- Omitting `revisionView` produces byte-for-byte current behavior.

## 7. Milestone 2 — Replacement Pairing and Insertion Affinity

**Target:** v0.5.0 experimental; stabilize in v0.6.0  
**Classification:** Opt-in observable serialization behavior

### 7.1 Paired replacement metadata

Model a delete-followed-by-insert diff as one internal replacement event when
both sides share a safe parent and no structural boundary must be crossed. The
event allocates distinct revision IDs but may share one author and timestamp.

```ts
interface ReplacementRevisionEvent {
  deletionId: string;
  insertionId: string;
  author: string;
  date: string;
}
```

Rules:

- `w:del` and `w:ins` IDs remain distinct and document-scoped.
- Shared metadata is allocated once for the event.
- Empty runs created solely by splitting are removed.
- `xml:space="preserve"` is based on emitted text boundaries; consecutive
  interior spaces are also preserved defensively.
- Pairing never crosses a hyperlink, field, SDT, comment, bookmark, move, or
  incompatible revision parent.
- Required structural markers may remain between changes. Fidelity wins over
  visual grouping.
- Word may present a paired event as a combined replacement, but the public API
  does not guarantee a particular balloon shape.

### 7.2 Separate formatting affinity from container affinity

One `left | right` flag is insufficient for every boundary. Define:

```ts
interface InsertionAffinity {
  formatting?: 'left' | 'right' | 'none';
  hyperlink?: 'inside' | 'outside' | 'preserve';
  revision?: 'coalesce_same_author' | 'separate';
}
```

Bookmarks and comments are ranges represented by markers, not formatting
wrappers. Their membership is determined by the insertion point relative to
start/end markers and must be tested independently.

When affinity is omitted, retain the existing boundary behavior. Do not claim
that a universal implicit `left` policy matches Word in every container.

### Acceptance criteria

- Accepted and rejected content is identical with pairing enabled or disabled.
- Pairing never removes or reorders structural markers.
- Every formatting/hyperlink/comment/bookmark boundary combination has an
  explicit expected DOM position.
- Word-version visual fixtures record presentation differences without making
  them correctness requirements.

## 8. Milestone 3 — Complete Paragraph-Boundary Semantics

**Target:** v0.6.0  
**Classification:** Internal semantic change for newly supported cases

The paragraph-mark primitives already exist. This phase completes the operation
matrix and removes unsafe reconstruction only where native boundary revisions
are proven.

### 8.1 Boundary removal versus paragraph deletion

These operations must remain distinct:

1. **Remove the boundary between A and B while retaining both texts**
   - Keep both physical paragraphs in the pending-revision document.
   - Mark the paragraph mark ending A with `w:pPr/w:rPr/w:del`.
   - Do not mark B's retained text as deleted.
   - Accept joins the contents; Reject restores the boundary.
2. **Delete Paragraph B and its text**
   - Track B's content deletion.
   - Track the relevant paragraph mark so Reject does not leave a ghost
     paragraph, bullet, or number.
   - Accept removes B; Reject restores B exactly.

For boundary acceptance, do not assume that A's paragraph properties win. The
remaining paragraph mark normally determines final paragraph formatting. Build
native Word fixtures for conflicting A/B styles, numbering, spacing, and
section properties, then encode the observed ownership rule before implementing
the merge algorithm.

### 8.2 Paragraph split

When splitting A into A and B:

- Split at a structurally safe run boundary.
- Create B with the correct inherited paragraph properties.
- Mark the newly inserted boundary—the paragraph mark ending A—with
  `w:pPr/w:rPr/w:ins`.
- Move retained trailing runs without incorrectly marking their text inserted.
- Accept keeps A and B; Reject moves B's retained content back into A and
  removes the inserted boundary.

### 8.3 Unsafe cases

Refuse or retain the current proven fallback when the boundary owns or crosses:

- `w:sectPr` with uncertain ownership;
- table-cell terminal paragraphs;
- incompatible list definitions;
- overlapping bookmarks/comments;
- fields, SDTs, drawings, notes, or unsupported revision nesting.

### Acceptance criteria

- A full matrix covers split, join, delete, list-item insertion/deletion, and
  boundaries adjacent to structural content.
- Word and internal Accept All/Reject All agree on paragraph order, text,
  paragraph properties, numbering, and section ownership.
- Existing paragraph-mark regression tests remain valid and are expanded rather
  than replaced.

## 9. Milestone 4 — Complete Tracked Formatting Semantics

**Target:** v0.6.0  
**Classification:** Existing behavior plus opt-in author-aware coalescing

`w:rPrChange` generation and property-change lifecycle handling already exist.
This phase closes coverage gaps and defines a public formatting operation model.

### 9.1 Operation contracts

Distinguish character and paragraph formatting:

```ts
interface CharacterFormatOperation {
  type: 'format';
  target: ParagraphTargetDescriptor;
  textToFormat: string;
  properties: {
    bold?: boolean;
    italic?: boolean;
    underline?: boolean;
    strike?: boolean;
    highlight?: string | null;
  };
}

interface ParagraphFormatOperation {
  type: 'paragraph-format';
  target: ParagraphTargetDescriptor;
  properties: {
    alignment?: 'left' | 'center' | 'right' | 'both';
    keepNext?: boolean;
    keepLines?: boolean;
    pageBreakBefore?: boolean;
  };
}
```

The property set may expand only with prior-state snapshot and lifecycle tests.
The previous-state child of `w:rPrChange` or `w:pPrChange` contains the complete
direct property state required to reverse the change, excluding nested change
records.

### 9.2 Author-aware policy

Preserve current formatting revision behavior by default. Add an experimental
policy only after selective-author tests:

```ts
formattingRevisionPolicy?: 'always' | 'coalesce-own-insertion';
```

- `always`: current behavior; create the appropriate property revision.
- `coalesce-own-insertion`: formatting inside the active author's pending
  insertion updates that insertion directly.
- Formatting baseline text or another author's insertion creates a separate
  property revision attributed to the active author.

### Acceptance criteria

- Bold/italic/underline/strike/highlight on/off transitions round-trip.
- Paragraph-property changes round-trip independently of run properties.
- Same-author and different-author pending insertions are tested separately.
- Run splitting never duplicates an existing `w:rPrChange` ID.
- Omitting the new policy retains current serialized behavior.

## 10. Milestone 5 — Versioned Concurrency Tokens

**Target:** v0.7.0  
**Classification:** Additive precondition

### 10.1 Two token scopes

Do not describe a three-part hash as a complete package revision.

```ts
interface RevisionToken {
  algorithm: 'sha256';
  version: 1;
  scope: 'document-parts' | 'package';
  value: string;
}
```

- `document-parts`: hash a documented, sorted list of provided part names and
  exact strings/bytes. Suitable for lower-level APIs.
- `package`: hash every relevant ZIP entry name plus its uncompressed bytes.
  Suitable for `openDocx` and CLI workflows.

Use length-prefixed binary framing so part boundaries cannot collide. Include
the token version and scope in the digest input. A package token necessarily
requires reading the ZIP directory and selected entries; it can still reject
before any mutation is staged.

### 10.2 Check location

- Inspection returns the token and the covered part list.
- Apply recomputes it from the exact incoming state before creating mutation
  savepoints or package artifacts.
- Mismatch returns `REVISION_MISMATCH`, `hasChanges: false`, no artifacts, and
  the original bytes.
- A token from one scope cannot satisfy another scope.

### Acceptance criteria

- Changes to document text, comments, numbering, relationships, headers,
  footers, styles, notes, or media are detected by the package token.
- ZIP recompression without uncompressed entry changes does not invalidate a
  package token.
- Omitted preconditions preserve current behavior.

## 11. Milestone 6 — Dependency-Aware Capture Chaining

**Target:** v0.7.0  
**Classification:** Additive batch behavior when captures are used

### 11.1 Contract

```json
[
  {
    "operationId": "insert-privacy-section",
    "type": "replace",
    "target": { "exactText": "11. Term and Termination" },
    "modified": "11. Term and Termination\n\n12. Data Privacy\nService Provider shall maintain the required controls.",
    "structuredContent": true,
    "captureKey": "privacy-section"
  },
  {
    "operationId": "comment-controls",
    "type": "comment",
    "target": {
      "captureRef": "privacy-section",
      "select": "required controls"
    },
    "commentContent": "Confirm the evidence and audit cadence."
  }
]
```

### 11.2 Dependency scheduling

Build a dependency graph before execution:

- A `captureRef` depends on the operation exporting its `captureKey`.
- Reject duplicate keys, missing references, forward references not resolvable
  by the graph, and cycles.
- Preserve source order among independent operations.
- Retain comment-first scheduling only when it does not violate dependencies.
- `executionOrder` reports actual topological order.

Static preflight reports capture-dependent targets as `deferred`, not missing.
An optional mutation preview may execute against an isolated clone to resolve
all dynamic selections without committing output.

### 11.3 Stable capture representation

Do not store raw DOM nodes as the durable capture value. Savepoint restoration
replaces the DOM tree and invalidates those references. A capture records a
serializable session-local identity:

```ts
interface CapturedEntity {
  captureKey: string;
  operationIndex: number;
  kind: 'paragraph' | 'range' | 'table' | 'list';
  generatedParagraphIds: string[];
  fingerprints: string[];
  expectedText: string[];
  structuralPathHints: string[];
}
```

Resolution reacquires nodes from the current session index and verifies text and
fingerprint before mutation. The capture table participates in every savepoint:
failed/no-op steps restore its exact prior state, and atomic rollback discards it.
If a captured entity is replaced later, the executor must either update the
capture deterministically or return `CAPTURE_STALE`.

### 11.4 Structured insertion guidance

Large generated content must be analyzed before execution:

1. Run `planStructuredReplacement`.
2. Refuse malformed tables or ambiguous block boundaries.
3. Emit separate heading, paragraph, list, and table nodes.
4. Capture block identities, not a flattened combined string.
5. Apply dependent comments or formatting only after those nodes exist.

### Acceptance criteria

- Insert-then-comment and insert-then-format work in one atomic batch.
- Independent comments retain current comment-first safety.
- Savepoint restoration never leaves stale node references or captures.
- Failure of a dependent step rolls back the entire atomic batch and all
  artifacts.

## 12. Milestone 7 — Commit-Aware Mutation Receipts

**Target:** v0.7.0  
**Classification:** Additive result metadata

Receipts describe both what was attempted and what survived the transaction:

```ts
interface MutationReceipt {
  operationIndex: number;
  operationId?: string;
  attemptedDisposition: 'applied' | 'no_change' | 'refused' | 'not_attempted';
  finalDisposition: 'applied' | 'no_change' | 'refused' | 'rolled_back' | 'not_attempted';
  committed: boolean;
  authorUsed?: string;
  revisionItems: Array<{
    id: string;
    kind: 'ins' | 'del' | 'move_from' | 'move_to' | 'rPrChange' | 'pPrChange' | 'structural';
    partName: string;
  }>;
  commentIds: string[];
  numberingIds: string[];
  relationshipIds: string[];
  affectedTargets: ResolvedDocumentTarget[];
  warnings: string[];
}
```

Rules:

- Existing result fields and status strings remain unchanged.
- Collectors are scoped to the operation savepoint.
- IDs consumed by a failed/no-op step are absent after restore.
- A later atomic failure changes earlier provisional receipts to
  `finalDisposition: 'rolled_back'` and `committed: false`.
- No IDs are presented as durable handles unless the transaction commits.
- Receipt generation must not add paragraph IDs or otherwise mutate the
  document solely for telemetry.

### Acceptance criteria

- Receipts enumerate exact committed revision, comment, numbering, and
  relationship allocations.
- No-op, refusal, stopped execution, per-step rollback, and whole-batch rollback
  are distinguishable.
- Receipt contents agree with a fresh parse of the committed package.

## 13. Milestone 8 — Strict Application Migration

**Target:** warning cycle in v0.7.x; default change in v1.0.0  
**Classification:** Breaking default change

Preflight is already strict by default, and the Node facade/CLI already prefer
strict application. The remaining migration concerns callers of permissive
lower-level application APIs.

### Migration

- v0.7.x: ambiguous permissive resolution succeeds only where legacy behavior
  currently permits it and emits `AMBIGUOUS_TARGET_HEURISTIC_USED`.
- Documentation and generated agent examples always use strict application.
- v1.0.0: lower-level application defaults to strict.
- An explicit `strictTargets: false` compatibility escape hatch may remain for
  one major release, with warnings.

Strict descriptors are conjunctive assertions, not a priority list: when both
`paragraphId` and `exactText` are supplied, both must describe the same current
target. An index or occurrence alone is not a concurrency guarantee.

### Acceptance criteria

- Zero, duplicate, descriptor-conflict, stale fingerprint, and anchor ambiguity
  have distinct errors and candidate diagnostics.
- Preflight and strict application resolve the same target or both refuse.
- CLI and Node package workflows never silently choose candidate one.

## 14. Verification Matrix

Every milestone runs:

```bash
npm test
npm run test:isolation
npm run check:types
node scripts/export-validation-fixtures.mjs
```

Risk-specific gates:

| Change | Required additional verification |
|---|---|
| Canonical fidelity | New fidelity-oracle suite plus real package corpus |
| Revision addressing | Accepted/rejected view, entity, Unicode, move, and mixed-revision fixtures |
| Pairing/affinity | Hyperlink, field, SDT, bookmark, comment, and multi-author fixtures |
| Paragraph boundaries | Internal round trip plus desktop Word Accept/Reject and visual review |
| Formatting revisions | Selective-author Accept/Reject and complete previous-state snapshots |
| Concurrency tokens | Entry mutation/recompression/token-version matrix |
| Capture chaining | Dependency graph, savepoint, stale capture, and atomic rollback matrix |
| Receipts | Fresh-package audit of every reported committed ID |

Desktop Word automation proves open/save and accepted/rejected outcomes. Visual
inspection remains required for layout, paragraph marks, change bars, and
Reviewing Pane presentation because those UI details are not fully exposed as a
stable COM assertion surface.

## 15. Release Sequence

### v0.5.0 — Evidence and addressing

- Current-capability audit.
- Mutation envelopes and fidelity oracles.
- Accepted/rejected revision-view inspection and targeting.
- Experimental paired replacement metadata and explicit insertion affinity.

### v0.6.0 — Structural lifecycle completion

- Proven paragraph boundary join/split/delete cases.
- Complete character and paragraph formatting revision contracts.
- Expanded selective-author and native Word lifecycle fixtures.

### v0.7.0 — Transactional agent workflows

- Versioned document/package concurrency tokens.
- Dependency-aware capture chaining.
- Commit-aware mutation receipts.
- Permissive-target deprecation warnings.

### v1.0.0 — Strict defaults

- Strict lower-level application by default.
- Compatibility flags and deprecated helpers reviewed under normal semver rules.
- Full fidelity, lifecycle, package, and Word validation gates required for
  release.

## 16. Exit Criteria

This program is complete when:

1. Every claimed existing and new capability has production-code and regression-
   test evidence.
2. Every operation declares and stays within its mutation envelope.
3. Accepted and rejected results match internal and desktop Word outcomes for
   text and structure.
4. Large structured insertions remain headings, paragraphs, lists, and tables
   rather than flattened text.
5. Dependent batch steps use stable captures and preserve atomic rollback.
6. Receipts never report rolled-back IDs as committed.
7. Stale package mutations fail before document state is changed.
8. Ambiguous agent operations fail closed under all recommended public workflows.

## 17. Coding-Agent Execution Protocol

This section is normative. An implementation agent should follow it even when a
work package appears simple.

### 17.1 Unit of work

Implement exactly one work package from Section 18 at a time. Do not combine
unrelated packages, opportunistically refactor adjacent modules, or change a
default before the package that explicitly authorizes that change.

For each work package:

1. Read `AGENTS.md`, `ARCHITECTURE.md`, `docs/TESTING.md`, and every production
   file listed under **Files to inspect**.
2. Run the listed baseline tests before editing. If they fail, record the
   pre-existing failure and stop unless the task explicitly includes repairing
   it.
3. Add the smallest failing regression test that expresses the requested
   invariant.
4. Implement through existing helpers and session state. Do not create a second
   parser, allocator, target resolver, or package writer.
5. Run the focused tests, then the shared gates in Section 17.5.
6. Update public declarations, schemas, README, architecture, testing guidance,
   and changelog only when the work package changes those contracts.
7. Report files changed, tests run, compatibility impact, and any deferred
   unsafe cases.

### 17.2 Mutation rules

- Create Word elements through `createWordElement`.
- Create revision metadata through the document-scoped allocator path already
  used by `createRevisionMetadata` and `RevisionIdAllocator`.
- Never allocate revision IDs with module globals, timestamps, random numbers,
  or a fresh per-operation allocator inside a batch.
- Mutate the live session DOM. String serialization is a boundary operation,
  not a mutation technique.
- Retain per-operation DOM, allocator, runtime-context, artifact, capture, and
  receipt savepoints.
- Never move or delete unknown OOXML merely because the current renderer ignores
  it.
- Preserve source whitespace exactly. Replacement text is caller data and must
  not be trimmed, collapsed, Markdown-normalized, or line-ending-normalized.
- Keep accepted/rejected text extraction centralized in
  `core/paragraph-text.js`.
- Keep target selection centralized in `core/paragraph-targeting.js` and
  `services/operation-preflight.js`.
- Keep package mutations in the Node/package layer. Browser-safe root imports
  must not acquire Node built-ins.

### 17.3 Error behavior

Expected refusal is not an exception. Return the established error shape:

```ts
{
  hasChanges: false,
  status: 'error',
  error: {
    code: 'STABLE_MACHINE_CODE',
    message: 'Actionable human-readable explanation.'
  }
}
```

An atomic batch failure additionally returns the original document, empty
uncommitted artifacts, `rolledBack: true`, and per-operation attempt results.
Never convert a structured refusal into `no_change`.

New error codes introduced by this plan:

| Code | Meaning |
|---|---|
| `UNSUPPORTED_REVISION_VIEW_MUTATION` | Requested view can be inspected but not safely mutated |
| `UNSAFE_REVISION_NESTING` | Mutation would create unsupported nested/overlapping revisions |
| `UNSUPPORTED_INSERTION_AFFINITY` | Requested container placement is not legal or deterministic |
| `UNSAFE_PARAGRAPH_BOUNDARY` | Split/join crosses unsupported structural ownership |
| `REVISION_MISMATCH` | Expected package/document token does not match current state |
| `REVISION_TOKEN_SCOPE_MISMATCH` | Token scope/version cannot guard the selected API |
| `DUPLICATE_CAPTURE_KEY` | More than one step exports the same capture key |
| `CAPTURE_NOT_FOUND` | No step exports the requested capture |
| `CAPTURE_DEPENDENCY_CYCLE` | Capture dependencies cannot be topologically ordered |
| `AMBIGUOUS_CAPTURE_SELECTION` | `select` matches multiple locations inside a capture |
| `CAPTURE_STALE` | Captured identity no longer resolves to the verified current entity |

Add codes to runtime validation, TypeScript declarations, JSON schema, README,
and tests in the same work package that first emits them.

### 17.4 Public-contract propagation checklist

For every public option, operation field, result field, or error payload, inspect
and update all applicable surfaces:

- `index.js` and `index.d.ts`;
- `services/document-operation-contract.js`;
- `services/standalone-operation-runner.d.ts`;
- `docs/schemas/document-operations.schema.json`;
- `node/docx-document.js`;
- CLI parsing/help/examples;
- `README.md`, `ARCHITECTURE.md`, `docs/AGENT-WORKFLOW.md`,
  `docs/TESTING.md`, and `CHANGELOG.md`;
- contract, schema, package facade, CLI, and type-export tests.

Do not update only the JavaScript implementation and leave declarations or the
operation schema behind.

### 17.5 Shared verification gates

Run these after every work package:

```bash
npm test
npm run test:isolation
npm run check:types
```

Run these when OOXML generation or package plumbing changes:

```bash
node scripts/export-validation-fixtures.mjs
docx-redline validate <representative-output.docx>
```

Run the relevant Word differential suite before completing any work package that
changes paragraph marks, revision wrappers, run properties, numbering, tables,
comments, relationships, or content types. Visual suites require human review of
the generated evidence; successful rendering alone is not visual approval.

## 18. Ordered Implementation Work Packages

### WP-00 — Capability and Gap Audit [COMPLETED 2026-09-05]

**Depends on:** None  
**Produces runtime changes:** No  
**Status:** Completed — audit matrix published at `docs/plans/structural-revision-capability-matrix.md` with all 67 baseline test suites verified passing.

**Files to inspect**

- `engine/run-builders.js`
- `engine/surgical-run-splitting.js`
- `engine/surgical-diff-application.js`
- `engine/reconstruction-mode.js`
- `pipeline/list-generation.js`
- `pipeline/structured-content.js`
- `core/paragraph-text.js`
- `core/paragraph-targeting.js`
- `services/document-operation-session.js`
- `services/batch-operation-orchestrator.js`
- `services/revision-comment-management.js`
- `services/operation-preflight.js`
- `node/docx-document.js`

**Steps**

1. Create `docs/plans/structural-revision-capability-matrix.md`.
2. Add one row for every behavior named in Section 2.
3. For each row, record status, production symbol, existing test, missing case,
   and proposed work package.
4. Mark a capability `implemented` only when generation, validation, Accept All,
   and Reject All are covered.
5. Mark behavior that exists only for lists or only for whole-paragraph deletion
   as `partial`, not implemented.
6. Record current serialized output for one canonical fixture per partial area.

**Tests**

No new test is required, but all tests linked from the matrix must be executed.

**Done when**

No later work package describes existing behavior as absent, and every claimed
gap has a reproducible fixture or test.

### WP-01 — Package Part Inventory Helper [COMPLETED 2026-09-05]

**Depends on:** WP-00  
**Produces runtime changes:** No; test helper only  
**Status:** Completed — helper implemented in `tests/helpers/package-fidelity.mjs`, verified by `tests/package_fidelity_inventory_tests.mjs`.

**Files to inspect**

- `node/zip-archive.js`
- `services/package-builder.js`
- `services/standalone-docx-plumbing.js`
- existing ZIP/package transaction tests

**New files**

- `tests/helpers/package-fidelity.mjs`
- `tests/package_fidelity_inventory_tests.mjs`

**Algorithm**

1. Read every ZIP entry as its uncompressed byte payload.
2. Normalize entry names to OPC forward-slash form; reject duplicate normalized
   names.
3. Sort by entry name for reporting only. Never treat archive order, compression
   method, CRC field position, or ZIP timestamps as content.
4. Return `{ entries: Map<name, { size, sha256, bytes }> }`.
5. Provide `comparePackageEntries(before, after, allowedChangedEntries)`.
6. Fail with a report containing unexpected changed, added, and removed entries.
7. Do not silently ignore `docProps`, relationships, media, content types, or
   custom XML.

**Required tests**

- Recompressing identical entries produces no content differences.
- One-byte payload changes are detected.
- Added and removed entries are reported separately.
- Duplicate normalized paths are refused.
- Binary media is compared without text decoding.

**Done when**

The helper can prove exact untouched entry payload preservation without claiming
the entire ZIP byte stream remains identical.

### WP-02 — OOXML Canonical Subtree Helper [COMPLETED 2026-09-05]

**Depends on:** WP-01  
**Produces runtime changes:** No; test helper only  
**Status:** Completed — canonical XML helper implemented in `tests/helpers/canonical-ooxml.mjs` supporting QName-valued attribute canonicalization, verified by `tests/canonical_ooxml_tests.mjs`.

**Files to inspect**

- `adapters/xml-adapter.js`
- `core/word-xml.js`
- `core/xml-query.js`
- fixtures containing `mc:Ignorable`, alternate prefixes, drawings, and SDTs

**New files**

- `tests/helpers/canonical-ooxml.mjs`
- `tests/canonical_ooxml_tests.mjs`

**Steps**

1. Evaluate an existing standards-compliant canonicalization dependency before
   writing custom logic. Do not add a dependency without documenting bundle and
   browser impact.
2. Represent element and attribute names as `{ namespaceURI, localName }`.
3. Preserve element child order and exact character data.
4. Canonicalize namespace declarations without breaking QName-valued attribute
   contents.
5. Treat comments and processing instructions consistently and document whether
   they participate in the digest.
6. Return both canonical bytes and SHA-256 so failures can print a readable
   canonical diff.

**Required tests**

- Different harmless prefixes compare equal.
- Different attribute orders compare equal.
- Different child orders compare unequal.
- `w:b` absent and `w:b w:val="0"` compare unequal.
- Leading/trailing spaces and `xml:space` differences compare unequal when they
  change Word text semantics.
- `mc:Ignorable` remains bound to the intended namespace prefixes.

**Stop condition**

If QName-valued attributes cannot be canonicalized safely, limit the first
oracle to exact serialized subtree bytes and document the limitation. Do not
ship a prefix-replacement approximation as canonical XML.

### WP-03 — Mutation Envelope Registry and Fidelity Suite [COMPLETED 2026-09-05]

**Depends on:** WP-01, WP-02  
**Produces runtime changes:** No unless a discovered fidelity bug is separately scoped  
**Status:** Completed — mutation envelopes registry implemented in `tests/helpers/mutation-envelopes.mjs`, verified against the multi-feature fixture matrix in `tests/fidelity_oracle_tests.mjs`.

**New files**

- `tests/helpers/mutation-envelopes.mjs`
- `tests/fidelity_oracle_tests.mjs`

**Steps**

1. Define envelopes by operation kind and structural route: surgical text,
   reconstruction, list, structured content, table reconciliation, comment,
   highlight, accept/reject, and package plumbing.
2. Each envelope lists allowed part changes and affected subtree classes.
3. Capture before-state entry hashes and candidate subtree identities.
4. Apply exactly one operation.
5. Reopen the output package before measuring after-state.
6. Compare untouched package entries exactly.
7. Reacquire untouched subtrees using IDs plus fingerprints and ordered context.
8. Compare their canonical digests.
9. Run internal Accept All and Reject All and compare expected structure.
10. Emit a diagnostic that names the first unexpected entry or subtree.

**Required fixture matrix**

- normal paragraph surrounded by untouched paragraphs;
- hyperlink and bookmark adjacent to an edit;
- comment range adjacent to an edit;
- field-code paragraph;
- drawing and image relationship;
- SDT containing text;
- numbered and bulleted lists;
- existing table and newly inserted Markdown table;
- header, footer, footnote, and endnote package parts;
- documents with and without comments/numbering parts.

**Done when**

The suite passes current behavior or produces separately tracked fidelity bugs.
Do not weaken an envelope merely to make a failure disappear.

### WP-04 — Revision-View Segment Extraction

**Depends on:** WP-03  
**Produces runtime changes:** Additive inspection API

**Files to edit**

- `core/paragraph-text.js`
- `services/document-inspection.js`
- public exports and declarations from Section 17.4

**New tests**

- `tests/revision_view_segment_tests.mjs`

**Algorithm**

1. Add one depth-first paragraph walker that emits structural text pieces.
2. Track ancestors for `ins`, `del`, `moveFrom`, and `moveTo`.
3. Emit text for `w:t`, `w:delText`, tabs, breaks, carriage returns,
   soft hyphens, and no-break hyphens.
4. Compute accepted and rejected visibility using the existing canonical rules.
5. Maintain accepted and rejected UTF-16 cursors independently.
6. For a hidden segment, store `null` for that view's start.
7. Merge adjacent pieces only when kind, author, revision ID, and DOM carrier are
   compatible.
8. Reimplement existing accepted/rejected string extraction as a projection of
   the walker, or prove parity before retaining both implementations.

**Required assertions**

- Concatenating accepted-visible segments equals current accepted extraction.
- Concatenating rejected-visible segments equals current rejected extraction.
- Entity decoding does not alter DOM offsets.
- An emoji counts as two UTF-16 code units.
- A replacement exposes deletion and insertion as separate segments, never as
  one concatenated “tracked” word.
- Move revisions map consistently with existing accept/reject behavior.

**Compatibility gate**

Existing extraction output must be byte-for-byte unchanged for all current
fixtures.

### WP-05 — Revision-View Target Resolution

**Depends on:** WP-04  
**Produces runtime changes:** Additive target option for inspection/preflight

**Files to edit**

- `core/paragraph-targeting.js`
- `services/operation-preflight.js`
- `services/document-operation-contract.js`
- public contract surfaces from Section 17.4

**Steps**

1. Add `revisionView` validation with only `accepted` and `rejected` values.
2. Build or cache paragraph metadata per revision view; never reuse accepted
   text for rejected matching.
3. Resolve all supplied descriptor fields conjunctively.
4. Calculate `occurrence` within the selected view.
5. Include the selected view in candidate diagnostics and fingerprints.
6. Restrict the initial implementation to inspection and preflight.
7. If apply receives `revisionView: 'rejected'`, return
   `UNSUPPORTED_REVISION_VIEW_MUTATION` until WP-06 explicitly supports that
   operation kind.

**Required tests**

- Same phrase at different accepted/rejected locations.
- Phrase exists only in an insertion.
- Phrase exists only in a deletion.
- Duplicate phrase in one view but unique in the other.
- Descriptor fields disagree despite one individually matching.
- Invalid view is rejected by runtime contract and JSON schema.

### WP-06 — Existing-Revision Mutation Policy

**Depends on:** WP-05  
**Produces runtime changes:** Opt-in only

**Goal**

Decide which mutations can safely preserve existing revisions. Do not add a
general `preserve-input` switch until the matrix below is implemented.

**Policy matrix to implement and test**

| Existing content | Same author | Different author | Initial action |
|---|---:|---:|---|
| Baseline text | n/a | n/a | Normal tracked mutation |
| Pending insertion, comment only | Yes | Yes | Allow if comment range is schema-valid |
| Pending insertion, text replacement | Yes | No | Refuse initially; later consider same-author coalescing |
| Pending deletion, comment only | Yes | Yes | Refuse until native Word fixture proves representation |
| Pending deletion, text replacement | Yes | No | Refuse with `UNSAFE_REVISION_NESTING` |
| Move source/destination | Any | Any | Refuse until move-specific lifecycle is designed |

**Steps**

1. Generate native Word fixtures for every proposed allowed case.
2. Inspect exact OOXML before copying the pattern.
3. Add schema, Word open/save, selective-author Accept/Reject, and internal
   lifecycle tests.
4. Add only narrowly named policies for proven cases.
5. Retain current `reject-input` default.

**Stop condition**

If internal and Word selective-author outcomes differ, keep the case refused.

### WP-07 — Replacement Event Metadata

**Depends on:** WP-03  
**Produces runtime changes:** Opt-in paired metadata

**Files to inspect/edit**

- `engine/surgical-diff-application.js`
- `engine/surgical-run-splitting.js`
- `engine/run-builders.js`
- `core/types.js`
- `core/redline-validation.js`

**Steps**

1. Identify delete/insert token pairs before either wrapper is created.
2. Confirm both sides can be emitted under the same legal parent.
3. Allocate one event author/date and two unique revision IDs.
4. Pass explicit metadata into builders; do not call the clock independently for
   each half.
5. Remove only empty runs introduced by the current split operation.
6. Preserve all pre-existing zero-width structural nodes in their original order.
7. Fall back to independent revisions when safe adjacency is impossible.
8. Return a warning such as `PAIRING_SKIPPED_STRUCTURAL_BOUNDARY` only when the
   caller explicitly requested pairing.

**Tests**

- Plain run replacement receives identical author/date and distinct IDs.
- Accepted/rejected output matches independent mode.
- Replacement across two formatting runs retains both formats.
- Hyperlink, comment, bookmark, field, SDT, and existing-revision boundaries
  trigger safe fallback without node movement.
- Multiple replacements in one operation receive separate event pairs.
- Validation rejects duplicate revision IDs.

**Word evidence**

Record Word version, build, platform, displayed revision count, and screenshots.
Do not encode “one balloon” as a cross-version API guarantee.

### WP-08 — Insertion Affinity

**Depends on:** WP-07  
**Produces runtime changes:** Additive explicit option

**Files to inspect/edit**

- `engine/surgical-spans.js`
- `engine/surgical-run-splitting.js`
- `engine/surgical-diff-application.js`
- hyperlink/bookmark/comment targeting helpers

**Steps**

1. Create an internal boundary descriptor containing left/right carrier runs and
   all open/closing structural boundaries at the insertion point.
2. Apply formatting affinity independently from hyperlink/revision affinity.
3. Treat bookmark and comment range markers as ordered boundaries, not wrappers.
4. Validate requested affinity before mutation.
5. If the requested placement is illegal or ambiguous, return
   `UNSUPPORTED_INSERTION_AFFINITY`.
6. When omitted, call the exact legacy path.

**Minimum test grid**

For insertion at the start and end of each structure, test left/right formatting
and legal inside/outside placement:

- normal to bold run;
- paragraph start/end with no carrier on one side;
- hyperlink;
- bookmark range;
- comment range;
- field begin/separate/end sequence;
- SDT content;
- same-author insertion;
- different-author insertion.

Assert exact parent, sibling order, inherited `w:rPr`, accepted/rejected text,
and preservation of every range marker.

### WP-09 — Paragraph Boundary Fixture Matrix

**Depends on:** WP-03  
**Produces runtime changes:** No; fixture/oracle work

**New files**

- `tests/fixtures/paragraph-boundaries/`
- `tests/paragraph_boundary_matrix_tests.mjs`

**Fixture creation**

Use desktop Word to create golden examples for:

- split one paragraph in the middle;
- delete only the boundary between two paragraphs;
- delete an entire middle paragraph;
- insert a blank paragraph;
- boundaries between different paragraph styles;
- boundaries with different list levels/numIds;
- boundary before a section break;
- first/last paragraph in a table cell;
- adjacent bookmark/comment ranges;
- same-author and multi-author revisions.

For every golden fixture, save pending, accepted, and rejected copies. Extract
the relevant XML and record which paragraph owns the inserted/deleted mark and
which `w:pPr` survives acceptance.

**Done when**

The ownership rules are evidence, not assumptions in prose.

### WP-10 — Paragraph Boundary Implementation

**Depends on:** WP-09  
**Produces runtime changes:** New proven structural routes

**Files to inspect/edit**

- `engine/reconstruction-mode.js`
- `engine/reconstruction-writer.js`
- `engine/run-builders.js`
- `services/revision-comment-management.js`
- `services/document-operation-mutations.js`

**Implementation order**

1. Boundary removal retaining both texts.
2. Paragraph split retaining all source text.
3. Entire paragraph deletion.
4. Blank paragraph insertion/deletion.
5. List paragraph cases.
6. Only then consider section/table-cell cases supported by WP-09 evidence.

**Boundary removal algorithm**

1. Resolve adjacent A/B paragraphs under the same legal parent.
2. Verify no unsupported boundary ownership.
3. Mark the paragraph mark ending A as deleted.
4. Leave B text active and in B while the revision is pending.
5. On acceptance, merge using the golden-fixture property ownership rule.
6. On rejection, remove only the boundary deletion marker.

**Split algorithm**

1. Resolve a safe UTF-16 split boundary using surgical span mapping.
2. Clone only the paragraph properties proven to transfer.
3. Move trailing retained nodes into B without wrapping them as inserted text.
4. Mark the paragraph mark ending A as inserted.
5. On rejection, move retained B content back into A in exact order and remove B.

**Required safeguards**

- Never create an empty numbered/bulleted ghost paragraph after Reject All.
- Never place `w:sectPr` on both paragraphs.
- Never split a field instruction or move only one bookmark/comment boundary.
- Preserve tabs, breaks, hyperlinks, drawings, and SDTs as nodes.
- Return `UNSAFE_PARAGRAPH_BOUNDARY` before mutation for unsupported cases.

### WP-11 — Formatting Contract Completion

**Depends on:** WP-03, WP-06  
**Produces runtime changes:** Additive operations; current default retained

**Files to inspect/edit**

- `engine/format-application.js`
- `engine/formatting-removal.js`
- `engine/run-builders.js`
- `engine/rpr-helpers.js`
- `services/document-operation-contract.js`
- `services/revision-comment-management.js`

**Steps**

1. Inventory current Markdown format hints and highlight behavior.
2. Define explicit character-format and paragraph-format operations in runtime,
   declarations, and schema.
3. Split only the requested character span.
4. Snapshot complete direct prior properties without nested change records.
5. Apply explicit on/off properties in schema order.
6. Generate one property change per resulting changed carrier.
7. Preserve fresh unique IDs when a run containing history is cloned.
8. Implement paragraph `w:pPrChange` generation separately from run formatting.
9. Add author-aware coalescing only after the WP-06 selective-author policy
   permits it.

**Tests**

- Every supported property: absent→on, on→off, and value→different value.
- Partial-run and multi-run spans.
- Character style plus direct formatting.
- Existing `w:rPrChange` and `w:pPrChange` history.
- Same/different author pending insertion.
- Accept/reject one author while retaining another's property change.
- No requested property change returns `no_change` without consuming an ID.

### WP-12 — Revision Token Core

**Depends on:** WP-01  
**Produces runtime changes:** Additive inspection utility

**New/edited files**

- `services/revision-token.js` or an equivalently focused leaf module
- `services/document-inspection.js`
- `node/docx-document.js`
- browser/root dependency isolation tests

**Binary framing**

Hash this conceptual sequence using UTF-8 names and raw payload bytes:

```text
magic = "docx-redline-revision-token\0"
version = uint32be(1)
scopeLength + scope
entryCount
for each entry sorted by normalized name:
  nameLength + name
  payloadLength + payload
```

Use fixed-width unsigned big-endian lengths or another unambiguous documented
encoding. Reject duplicate normalized entry names. The browser implementation
may use Web Crypto asynchronously; the Node package path may use `node:crypto`
behind a Node-only module. Do not import Node crypto from the browser-safe root
graph.

**Tests**

- Stable across ZIP compression/order/timestamp changes.
- Changes when any entry name or payload changes.
- Stable regardless of Map/object enumeration order.
- Different scope or version produces a different token.
- Binary zero bytes and non-ASCII entry names frame correctly.

### WP-13 — Revision Token Enforcement

**Depends on:** WP-12  
**Produces runtime changes:** Additive apply precondition

**Steps**

1. Add `expectedRevision` to package and applicable lower-level options.
2. Validate token syntax, version, algorithm, and scope before comparison.
3. Recompute from the exact input supplied to apply.
4. Perform comparison before creating session mutations or companion artifacts.
5. Use a timing-safe equality primitive where available; correctness must not
   depend on it.
6. Return `REVISION_TOKEN_SCOPE_MISMATCH` for incompatible scopes and
   `REVISION_MISMATCH` for unequal valid tokens.
7. Ensure atomic and non-atomic modes both leave input untouched on mismatch.

**Tests**

- Matching token applies.
- Stale token refuses.
- Token from another document refuses.
- Parts token cannot guard package API and vice versa.
- No operation result or artifact is partially produced.
- Omitting the option follows the byte-for-byte legacy path.

### WP-14 — Capture Contract and Dependency Graph

**Depends on:** WP-05  
**Produces runtime changes:** Validation and scheduling for capture batches

**Files to inspect/edit**

- `services/document-operation-contract.js`
- `services/operation-preflight.js`
- `services/batch-operation-orchestrator.js`
- public contract surfaces from Section 17.4

**Steps**

1. Validate `operationId`, `captureKey`, and `target.captureRef` as non-empty,
   bounded strings.
2. Build `captureKey -> source operation index`.
3. Reject duplicates and missing producers.
4. Add dependency edges from consumer to producer.
5. Perform a stable topological sort using original operation index as the tie
   breaker.
6. Apply existing comment-first priority only among currently ready independent
   nodes.
7. Detect cycles before constructing `DocumentOperationSession` mutations.
8. Preflight static targets immediately and mark capture consumers `deferred`.

**Tests**

- Linear dependency, diamond dependency, independent comments, duplicate key,
  missing producer, self-cycle, multi-node cycle, and stable tie order.
- `results` remain sorted by original index while `executionOrder` reports actual
  execution.
- Batches without capture fields preserve current scheduling exactly.

### WP-15 — Capture Storage, Resolution, and Savepoints

**Depends on:** WP-14  
**Produces runtime changes:** Capture execution

**Files to inspect/edit**

- `services/document-operation-session.js`
- `services/document-operation-applier.js`
- `services/document-operation-mutations.js`
- structured/list/table replacement paths

**Steps**

1. Add `captureTable` to the session.
2. Clone/restore it with every savepoint.
3. After a successful changed producer, derive captures from the actual imported
   live-DOM nodes—not from the requested text alone.
4. Store serializable identity evidence; do not store raw nodes across
   savepoints.
5. Reacquire candidates from the current paragraph/table index.
6. Verify fingerprints and expected text before resolving `select`.
7. Require a unique `select` match within the capture.
8. Invalidate or deterministically update captures affected by later structural
   replacement.
9. Discard all captures on atomic rollback.

**Required route tests**

- Plain paragraph insertion → comment.
- Structured heading/paragraph insertion → comment.
- List insertion → format one item.
- Table insertion → target one cell only if the capture schema supports cells.
- Producer no-op, producer failure, consumer failure, stale capture, ambiguous
  selection, and later replacement of captured content.

### WP-16 — Receipt Collector

**Depends on:** WP-03  
**Produces runtime changes:** Internal telemetry collector

**Files to inspect/edit**

- `core/types.js`
- revision builders
- comment engine
- numbering helpers
- package relationship/content-type helpers
- `services/document-operation-session.js`

**Steps**

1. Add a session-owned collector with `beginOperation(index)`, record methods,
   `commitOperation()`, and `restore(savepoint)`.
2. Record IDs at the successful allocation/attachment point, not by reparsing a
   partially serialized string.
3. Associate every item with kind and part name.
4. Snapshot collector state in operation savepoints.
5. Ensure failed/no-op allocations disappear on restore.
6. Keep the collector internal until WP-17 proves reconciliation with output.

**Tests**

- Text replacement records distinct deletion/insertion IDs.
- Formatting records property-change IDs.
- Comment/list/package plumbing records companion IDs.
- Failed and no-op operations leave collector state unchanged.

### WP-17 — Public Commit-Aware Receipts

**Depends on:** WP-16  
**Produces runtime changes:** Additive result field

**Steps**

1. Convert internal collector records to the receipt schema in Section 12.
2. Attach provisional receipts to item results.
3. At successful transaction commit, set applied receipts `committed: true`.
4. On atomic rollback, rewrite every provisionally applied receipt to
   `finalDisposition: 'rolled_back'` and `committed: false`.
5. Mark unexecuted steps `not_attempted` when `continueOnError: false`.
6. Reopen committed output and reconcile every reported durable ID.
7. If reconciliation fails, treat it as a transaction error rather than return
   inaccurate telemetry.

**Tests**

- Successful atomic and non-atomic batches.
- Early stop and continue-on-error.
- Later failure rolling back earlier success.
- Mixed text/comment/list/table/format operations.
- Existing IDs are never incorrectly reported as newly allocated.

### WP-18 — Strict Application Warning Cycle

**Depends on:** WP-05  
**Produces runtime changes:** Warning only before v1.0.0

**Steps**

1. Identify every public application entry point and its current strict default.
2. Add a warning only when permissive resolution actually chooses among multiple
   candidates—not merely when `strictTargets: false` is present.
3. Include candidate count and migration guidance without exposing document
   content beyond existing diagnostics policy.
4. Ensure preflight and strict application share the same candidate resolver.
5. Update recommended examples to strict descriptors.

**v1.0.0 task**

Change the remaining permissive defaults only in a dedicated major-version
work package. Update semver notes, migration examples, and compatibility tests at
that time; do not silently make the change during v0.x implementation.

## 19. Test Fixture Naming and Assertions

Use names that state the invariant, not the implementation. Examples:

- `paragraph-boundary-join-retains-second-text`
- `paragraph-split-reject-restores-source-properties`
- `replacement-pair-preserves-comment-boundary`
- `rejected-view-target-is-ambiguous`
- `capture-savepoint-reacquires-current-node`
- `atomic-receipt-marks-prior-success-rolled-back`

Every structural fixture should assert at least:

1. operation status and error code;
2. generated revision element location and unique IDs;
3. exact accepted text;
4. exact rejected text;
5. accepted paragraph/table/list structure;
6. rejected paragraph/table/list structure;
7. preservation of nearby unknown/structural nodes;
8. validation result;
9. package artifact presence/absence;
10. no unexpected mutation-envelope drift.

Avoid assertions that merely search the serialized XML for a tag. Parse the
output and assert parent, child, sibling order, namespace, attributes, and
affected text. Regex assertions may supplement but not replace DOM assertions.

## 20. Implementation Reporting Template

The implementing agent should finish each work package with this report:

```text
Work package: WP-XX — Name
Outcome: complete | partial | blocked

Behavior implemented:
- ...

Compatibility:
- Public API change: ...
- Default behavior change: none | ...
- Serialized OOXML change: none | ...

Files changed:
- path — purpose

Tests added:
- test — invariant

Verification run:
- command — result

Deferred or refused cases:
- case — reason and error code

Manual Word evidence:
- not required | artifact path and review status
```

Do not report a work package complete when focused tests pass but shared gates,
package validation, required Word evidence, declarations, schemas, or docs remain
unfinished.
