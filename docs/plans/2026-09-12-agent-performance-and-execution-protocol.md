# Agent Performance and Execution Protocol Plan

**Status:** In progress — WP-00 through WP-03 implemented as development-only benchmark/sample work; external agent observations and WP-04+ remain
**Date:** 2026-09-12  
**Priority:** Reduce AI-agent wall time, tool turns, generated tokens, and recovery
reasoning without weakening targeting, revision fidelity, validation, rollback, or
output-safety guarantees.

---

## 1. Executive Summary

`@ansonlai/docx-redline-js` is already fast and reliable at its core task. In a
small in-memory fixture benchmark, opening, inspecting, applying a one-character
redline, validating, and repacking completed in roughly 44–87 ms after process
startup. The user-visible delay in Claude, OpenCode, and similar harnesses is now
primarily an agent-protocol problem rather than an OOXML mutation problem.

**Implementation scope decision (2026-09-12):** The agent session is a
development-only test and demonstration. It is not a package export, supported
core API, or alternate mutation engine. The README may show it as an example of
how a custom harness can be built over the published Node facade.

The remaining latency is concentrated in five areas:

1. Agents ingest a large instruction surface that mixes ordinary document use,
   wrapper construction, contributor architecture, advanced revision cases,
   error recovery, and testing.
2. Agents copy exact targets and generate complete replacement paragraphs even
   when only a few words change.
3. Shell-based agents perform separate extract, operations-file creation, and
   apply turns.
4. Agents reason about batch order because targets are re-resolved against a
   changing document and can become falsely stale after unrelated structural
   edits.
5. Recovery policy is documented in prose, while runtime errors do not always
   preserve enough structured context to perform the next safe action directly.

This plan addresses those costs through one coherent architecture:

- a development-only, thin agent facade with stateful document sessions, narrow tools,
  opaque target handles, and explicit safe profiles;
- localized exact-span edits that are deterministically expanded into the
  library's canonical complete desired text;
- a compile-and-bind batch phase that resolves source targets against the
  immutable batch-start document and schedules only genuine dependencies;
- a versioned, machine-actionable recovery envelope and explicit retry plan;
- compact consumer instructions and modest shell ergonomics for environments
  that cannot use structured tools.

The existing reconciliation engine remains the authority. The new layers must
translate into the canonical document-operation schema and delegate to the Node
facade or full-document runner rather than reproducing targeting, diffing, ZIP,
comment, numbering, revision, or validation behavior.

## 2. Confirmed Findings

### 2.1 The common agent path is smaller than the current instruction surface

The ordinary editing workflow is essentially:

1. inspect the relevant provision;
2. compose the desired change;
3. apply one stable batch;
4. require complete success;
5. inspect changed structure only when necessary.

The current `AGENTS.md` also contains wrapper blueprints, contributor routing,
special restoration/list behavior, detailed failure rules, and testing guidance.
Those topics remain useful, but they should not all be loaded for every ordinary
contract edit.

### 2.2 Batch scheduling is only partially automatic

`services/batch-operation-orchestrator.js` currently:

- builds explicit dependencies from `captureKey` and `captureRef`;
- gives comments priority among otherwise-ready operations;
- executes operations progressively against one live document DOM;
- retains source indexes in results while separately reporting execution order;
- installs a batch-start numeric-reference snapshot to repair some index drift.

This does not bind ordinary source targets. Paragraph IDs, fingerprints, exact
text, and anchors are still evaluated against current document metadata during
execution.

Paragraph fingerprints currently include document index. A reproduced case
demonstrates the resulting false dependency:

1. Operation 1 splits an early paragraph.
2. Operation 2 targets a later, otherwise untouched paragraph using the exact
   descriptor returned by inspection.
3. Operation 2 fails with `TARGET_FINGERPRINT_MISMATCH` because the earlier split
   changed its document index.
4. Reversing the same two operations succeeds.

Same-source writes are a different class of problem. Two complete desired texts
for one paragraph cannot safely be merged by sorting. A deletion and later
formatting of the same source content may also be semantically incompatible.

### 2.3 Error data becomes lossy across layers

Core target resolution can produce useful candidates and code-specific details,
but `normalizeOperationError` and CLI compaction currently preserve only selected
fields. Some generated-output validation information, current target text,
revision mismatch data, and corrective context are consequently unavailable to
thin harnesses.

Progressive partial output creates additional reasoning work. The caller must
infer whether to retry against the original input or written output, which
operations committed, and whether replaying the whole batch would duplicate
work. A CLI result with `status: "partial"` also currently has
`completion: false` but can exit with code zero.

### 2.4 Most ordinary defaults already exist

The CLI already supplies or enforces defaults for author, strict targeting,
validation, tracked changes, existing-revision safety, and derived output names.
A mutable global configuration system would therefore save less work than a
session facade and would make runs harder to reproduce. Defaults should be
explicit properties of an agent session or named execution profile.

## 3. Goals

- Let an ordinary agent locate and revise a provision through two structured tool
  calls: inspect/find, then apply.
- Eliminate operations-file creation for structured-tool integrations.
- Avoid echoing exact target text back into an apply request when a session-bound
  target handle is available.
- Let agents express small wording changes with small, exact edit intents while
  retaining the canonical complete-desired-text reconciliation path internally.
- Make independent source-targeted batch operations insensitive to caller order.
- Detect true source overlap and created-content dependencies before mutation.
- Return enough bounded, structured error context for one-turn recovery whenever
  recovery is deterministic and safe.
- Make retry base, replay scope, authorization requirements, and unchanged-retry
  safety explicit.
- Preserve current package validation, revision validation, rollback, comments,
  numbering, receipts, and Accept/Reject behavior.
- Keep existing public APIs and CLI commands backward compatible through additive
  capabilities and contract-version negotiation.
- Measure agent wall time, tool calls, model output tokens, failed attempts, and
  document fidelity separately from native engine time.

## 4. Non-goals

- Do not add an LLM or legal-drafting model to the library.
- Do not make generic instructions such as "make this provision mutual" a
  deterministic library operation. The agent remains responsible for legal
  composition and intent.
- Do not replace the canonical rule that an ordinary redline's `modified` value
  is the complete desired accepted-view content.
- Do not silently choose among ambiguous targets or anchors.
- Do not automatically accept or reject another reviewer's revisions.
- Do not automatically remove comments to make an operation succeed.
- Do not automatically retry semantic failures.
- Do not auto-merge two incompatible complete desired states for the same source
  paragraph.
- Do not promise that every batch is order-independent. Operations targeting
  created content and operations with overlapping write scopes have real
  dependencies or conflicts.
- Do not use blanket bottom-up sorting as the batch architecture.
- Do not add mutable process-global library defaults.
- Do not weaken source-file overwrite rules or present partial output as complete.
- Do not optimize core reconciliation paths unless profiling shows a measurable
  bottleneck after agent-protocol overhead is removed.

## 5. Hard Invariants

### 5.1 Thin integration boundary

- The agent facade must translate ergonomic inputs into canonical operations and
  call `openDocx(...).applyOperations(...)` once per stable batch.
- It must not implement its own ZIP mutation, target matcher, paragraph merger,
  revision allocator, comment relationship manager, numbering merger, or OOXML
  validator.
- Localized edits must be expanded into complete desired content before entering
  the ordinary redline path.

### 5.2 Target and version safety

- Every target handle is scoped to one exact package revision and revision view.
- Applying a stale handle fails closed with structured current/expected revision
  data.
- Handles are opaque conveniences, not durable references across unrelated
  sessions or externally modified files.
- A successful mutation invalidates affected handles. The result may return new
  handles bound to the new package revision.
- The source document is never overwritten unless the caller explicitly requests
  in-place behavior.

### 5.3 Mutation and rollback safety

- Agent profile operations default to strict targeting, validation, tracked
  changes, and atomic batches.
- An atomic failure returns no completed output and leaves the retry base
  byte-identical to the input.
- Progressive output remains opt-in for the agent facade and must return an
  explicit retry plan.
- Existing savepoint, receipt reconciliation, package validation, and generated-
  output validation behavior remains authoritative.

### 5.4 No hidden semantic choices

- Ambiguous matches return candidates; they do not select the first candidate.
- Cross-author revision policy changes are explicit.
- Policies that normalize, accept, reject, or remove reviewer work identify that
  they require user authorization.
- Library/builder failures are distinguished from operation-input failures so an
  agent is not encouraged to alter legal wording to work around invalid OOXML.

## 6. Proposed Agent Protocol

### 6.1 Integration surfaces

The reference implementation is a Node-only development example under
`examples/`, built on `@ansonlai/docx-redline-js/node`. It is intentionally
excluded from package `files` and `exports`. Custom harnesses can copy or adapt
the pattern. A production MCP or other tool server would remain a separate thin
transport owned by its host rather than becoming part of the core package.

The common logical tool surface is:

| Tool | Side effect | Purpose |
|---|---|---|
| `inspect_docx` | None | Open or reuse a session, find relevant clauses, return context and target handles. |
| `apply_docx_edits` | Produces new bytes/file | Apply complete desired text, localized exact-span edits, deletes, comments, or supported formatting through canonical operations. |
| `resolve_docx_review` | Produces new bytes/file | Accept/reject revisions or remove comments through the existing package facade. |

Advanced restoration, rejected-view insertion, list, table, and comment-reply
features may be exposed through narrow convenience calls or an advanced operation
field. They should not expand the common tool schema unless usage data justifies
it.

### 6.2 Inspection response

Example:

```json
{
  "status": "ok",
  "sessionId": "D1",
  "packageRevision": {
    "algorithm": "sha256",
    "version": 1,
    "scope": "package",
    "value": "..."
  },
  "targets": [
    {
      "handle": "T7",
      "exactText": "The Company may terminate this Agreement on 30 days' notice.",
      "humanReference": "Section 8.2 — Termination",
      "nearestHeading": "Termination",
      "inTable": false,
      "list": null
    }
  ]
}
```

The visible response contains the text needed to draft the change. The session
registry retains the full strict descriptor and package token so the agent does
not repeat them in the apply request.

Inspection should support context windows:

```js
inspectDocx({
  input,
  search: 'termination',
  around: 2,
  revisionView: 'accepted'
});
```

Search results must distinguish direct hits from surrounding context and return
handles only for targetable paragraphs. This should replace the common
search-then-range two-call pattern.

### 6.3 Apply request

Example with complete desired text:

```json
{
  "sessionId": "D1",
  "edits": [
    {
      "operationId": "mutual-termination",
      "target": "T7",
      "desiredText": "Either party may terminate this Agreement on 30 days' notice."
    }
  ]
}
```

Example with a localized exact-span replacement:

```json
{
  "sessionId": "D1",
  "edits": [
    {
      "operationId": "mutual-termination",
      "target": "T7",
      "replacements": [
        {
          "find": "The Company may terminate",
          "replace": "Either party may terminate"
        }
      ]
    }
  ]
}
```

The response must preserve canonical status, results, receipts, warnings,
validation, output bytes/path, and recovery information. A facade-level `ok`
must be true only when the operation is fully complete.

### 6.4 Agent execution profile

The initial named profile should be explicit and inspectable:

```json
{
  "name": "legal-agent",
  "author": "AI Redliner",
  "atomic": true,
  "continueOnError": true,
  "strictTargets": true,
  "validate": true,
  "generateRedlines": true,
  "existingRevisions": "merge-same-author",
  "outputPolicy": "derived"
}
```

`continueOnError: true` in atomic mode allows the engine to diagnose the complete
batch while still rolling all mutations back when any operation fails.

Call-level values override session-profile values. The general library and
existing CLI retain their current defaults for compatibility. The response
records the effective profile and all overrides that affect semantics.

## 7. Localized Exact-Span Edit Contract

Localized edits are facade conveniences that compile into one canonical redline
per source target.

### 7.1 Matching rules

- All `find` values are evaluated against the immutable text associated with the
  target handle.
- Exact matching is the default. Whitespace-equivalent or case-insensitive modes
  require explicit, separately named policies if added later.
- A `find` value must be unique unless the caller supplies a positive occurrence.
- Zero matches return `PATCH_SOURCE_NOT_FOUND`.
- Multiple matches without an occurrence return `AMBIGUOUS_PATCH_SOURCE` with
  bounded offsets/excerpts.
- Overlapping replacement ranges return `OVERLAPPING_PATCHES`.
- Identical replacement ranges with incompatible values return
  `CONFLICTING_PATCHES`.
- Replacement positions are calculated before any replacement is applied, then
  materialized in a deterministic order.

### 7.2 Canonical delegation

After validation, the facade constructs the complete desired accepted-view text
and emits one canonical `redline` operation containing:

- the session-bound exact target descriptor;
- the complete constructed `modified` text;
- the operation ID, author, and explicit effective policies.

The normal reconciliation engine performs diffing and tracked-change generation.
The facade must not construct revision OOXML directly.

### 7.3 Round-trip assertion

The facade verifies that the requested replacement set reconstructs exactly to
the computed `modified` value before calling the engine. Existing output
validation and receipt reconciliation then verify the mutation. Any mismatch is
an operation error; the facade must not fall back to a fuzzy or broader edit.

### 7.4 Scope boundaries

Localized replacement is appropriate for text inside one source paragraph. It
must not silently span paragraphs, table cells, list items, fields, hyperlinks,
comments, bookmarks, or unsupported revision boundaries. Those cases use the
existing structured operations or fail with a specific boundary error.

## 8. Compile-and-Bind Batch Architecture

### 8.1 Compile against the immutable source state

Add an internal `compileOperationBatch` phase to the ordinary apply path. This is
not a new agent-visible preflight command and must not add another tool turn.

The compiler should:

1. normalize and validate every source operation;
2. build the explicit capture dependency graph;
3. create accepted and rejected metadata indexes for the initial document;
4. resolve every non-capture target, range endpoint, and anchor against that
   initial state;
5. validate source fingerprints, revision views, exact text, occurrences, and
   package/document revision tokens once;
6. assign session-local source identities and mutation footprints;
7. build ordering edges and conflict records;
8. return a compiled execution schedule or a complete set of refusal results.

### 8.2 Stable source identity

Compiled bindings must not depend on recomputing an index-sensitive fingerprint
after unrelated mutations.

Introduce a session-local `SourceTargetRegistry` that maps immutable source
identities to live DOM lineage. It must:

- avoid injecting temporary attributes into serialized OOXML;
- survive unrelated insertions and removals elsewhere in the document;
- track one-to-one replacement lineage where the original semantic target
  remains meaningful;
- mark identities consumed when an operation removes or structurally replaces
  the target;
- allow capture-produced identities for newly created content;
- rehydrate bindings after a savepoint restore or cloned-DOM rollback;
- expose enough metadata for receipts and causal errors without leaking DOM
  objects into public results.

The implementation may use node references during uninterrupted successful
execution, but it must define an explicit remapping strategy because savepoint
rollback replaces the live document with a clone.

### 8.3 Read/write footprints

Each compiled operation should declare its source footprint:

- target paragraph identity or contiguous range;
- optional exact anchor range;
- operation class such as read-anchor, character-format, paragraph-format,
  text-write, structural-write, comment-write, or created-content consumer;
- whether it preserves, replaces, splits, or consumes the source identity.

These footprints permit deterministic conflict and dependency analysis before
mutation.

### 8.4 Scheduling rules

- Preserve explicit capture producer-to-consumer edges.
- Apply source-anchored comments before an operation that would consume or alter
  their source anchor, when that combination is otherwise safe.
- Apply formatting/highlighting before text or structural writes only when the
  mutation contracts guarantee that the resulting revisions remain valid.
- Preserve source order when no rule requires reordering; caller order remains a
  useful deterministic tiebreaker, not a correctness requirement.
- Do not use global reverse-document order as a substitute for binding.
- Reconcile numbering and comment artifacts through the existing shared batch
  context.

### 8.5 Conflicts and consumed targets

Return causal, operation-indexed errors before mutation when possible:

| Code | Condition | Suggested action |
|---|---|---|
| `OVERLAPPING_SOURCE_TARGETS` | Multiple incompatible text/structural writes address the same source scope. | Consolidate into one desired state. |
| `TARGET_CONSUMED_BY_OPERATION` | A prior planned write deletes or structurally replaces a later source target. | Consolidate or target created output explicitly. |
| `CAPTURE_FANOUT_CONFLICT` | Multiple mutating consumers depend on one capture whose first mutation makes later consumers stale. | Chain captures or consolidate consumers. |
| `REVISION_ORDER_CONFLICT` | Formatting/highlighting and text mutation cannot be safely composed. | Consolidate or split into an explicit dependency. |
| `AMBIGUOUS_CREATED_TARGET` | A capture selector does not uniquely identify created content. | Supply a narrower selector. |

Operations intentionally targeting content created by another operation continue
to use `captureKey`/`captureRef` internally. The agent facade may expose this as a
simpler `resultOf` plus selector relationship rather than requiring models to
learn capture mechanics.

### 8.6 Apply integration

`applyOperationsToDocumentXml` should consume a compiled plan by default. It must
continue to:

- share one live DOM and revision allocator;
- create per-operation savepoints;
- preserve original operation indexes in results;
- report actual execution order;
- reconcile receipts against final output;
- roll back atomically on any operation or output-validation failure.

Preflight and apply should share the compiler so their validation and conflict
semantics cannot drift. External preflight remains optional and is not added to
the normal agent workflow.

## 9. Machine-Actionable Recovery Contract

### 9.1 Versioned envelope

Define one additive error normalizer and recovery registry used by preflight,
operation apply, batch apply, the package facade, agent facade, and CLI
compaction.

Example:

```json
{
  "code": "EXISTING_REVISIONS",
  "message": "Target contains pending revisions from Reviewer A.",
  "stage": "target-safety",
  "category": "operation_requires_change",
  "context": {
    "operationIndex": 3,
    "operationId": "mutuality-3",
    "revisionAuthors": ["Reviewer A"],
    "currentPolicy": "merge-same-author"
  },
  "recovery": {
    "sameArgumentsSafe": false,
    "action": "set_option",
    "field": "existingRevisions",
    "recommendedValue": "slice-cross-author",
    "requiresReinspection": false,
    "requiresUserAuthorization": false
  }
}
```

The envelope is additive: existing `code` and `message` remain stable.

### 9.2 Recovery categories

At minimum, support:

- `request_fixable`: malformed operation or missing required field;
- `target_refresh_required`: package or target state changed;
- `candidate_selection_required`: multiple safe candidates exist;
- `policy_choice_required`: a revision or formatting policy must be explicit;
- `user_authorization_required`: reviewer work would be accepted, rejected,
  removed, or otherwise normalized;
- `source_conflict`: multiple batch intents cannot be composed automatically;
- `library_or_builder_failure`: valid input produced invalid OOXML, serialization,
  package, or lifecycle output;
- `manual_document_resolution`: unsupported Word structure requires human work.

### 9.3 Bounded corrective context

- `TARGET_TEXT_MISMATCH` with a uniquely resolved paragraph ID may include the
  current bounded `exactText`, fingerprint, handle, and
  `requiresRecomposeModified: true`.
- `AMBIGUOUS_TARGET` includes bounded candidate handles, human references,
  paragraph IDs, fingerprints, revision views, and excerpts without selecting a
  candidate.
- `REVISION_MISMATCH` returns structured expected and current revision tokens,
  not only hashes embedded in prose.
- Comment-protection errors retain comment IDs, authors, and bounded text and are
  classified as requiring user action.
- `GENERATED_OOXML_INVALID`, serialization failures, and package validation
  failures retain summarized issue codes and stages and are classified as
  library/builder failures.
- `INVALID_OPERATION` includes a machine field path, expected shape, and actual
  shape when deterministically known.

### 9.4 Retry plan

Every failed or partial mutation result from the agent facade should include:

```json
{
  "retryPlan": {
    "base": "original",
    "committedIndexes": [],
    "failedIndexes": [3],
    "unattemptedIndexes": [],
    "replayWholeBatch": true,
    "sameArgumentsSafe": false
  }
}
```

Rules:

- Atomic failure uses the unchanged original input and replays the corrected
  complete batch.
- Progressive partial output uses the written output and retries only failed or
  unattempted operations.
- A successful no-op remains `no_change`, not a failure.
- Repeating unchanged failed arguments is never recommended.

### 9.5 CLI completion semantics

The agent CLI/profile must return a nonzero exit code for both `error` and
`partial`, with distinct codes if useful. Existing general CLI exit behavior may
remain backward compatible until a major version; an additive
`--require-complete` or agent-specific command/profile can enforce the stronger
contract immediately.

## 10. Shell-Only Ergonomics and Documentation

### 10.1 Operations from stdin

Support `--operations -` so a shell harness can provide serialized JSON without a
persistent `ops.json` file. Parsing, Unicode handling, schema validation, and
result behavior must be identical to file-based operations.

Do not encourage shell interpolation of legal text. Harnesses should pass stdin
through a structured execution API or JSON serializer.

### 10.2 Explicit configuration, not hidden globals

The session profile is the primary location for defaults. If a project config is
added for shell-only use:

- prefer explicit `--config <path>` before considering automatic discovery;
- define precedence as call flag > explicit config > environment > named profile;
- return the effective settings in mutation output;
- exclude authorization-sensitive policies from implicit configuration unless
  the user deliberately places them in the explicit file;
- never mutate global library state.

### 10.3 Tiered agent documentation

Create a compact consumer document, tentatively
`docs/AGENT_FAST_START.md`, limited to the ordinary path:

- inspect/find and use returned handles;
- submit complete desired text or exact-span replacements;
- treat only `ok: true`/`completion: true` as complete;
- follow structured recovery actions;
- never retry unchanged failed operations.

Target size: 40–60 lines and no more than roughly 600 words.

Refactor `AGENTS.md` into a short launch card that routes:

- ordinary DOCX edits to the consumer quick start or agent tools;
- wrapper work to the thin-facade contract;
- source changes to architecture and focused testing references;
- advanced restore/list/table/revision cases to topic-specific material.

Keep detailed examples, recovery rationale, contributor architecture, and testing
lanes in the knowledge base. Tool descriptions and runtime recovery envelopes
should carry the minimum machine-facing contract so ordinary agents do not need
to load the full knowledge base.

## 11. Work Packages

### WP-00: Establish agent-performance and fidelity baselines [IMPLEMENTED 2026-09-12 — external observations pending]

**Goal:** Measure the correct bottleneck before changing interfaces.

- Build a repeatable task corpus covering:
  1. add terminal punctuation;
  2. change a notice period;
  3. make a simple termination provision mutual;
  4. make a substantive full-clause rewrite;
  5. apply ten independent edits including an earlier paragraph split;
  6. mix comments with redlines;
  7. recover from stale target text;
  8. encounter ambiguous text, comments, and foreign revisions.
- Record native library time separately from model/tool wall time.
- Record prompt/input tokens, generated/output tokens, tool calls, retries, and
  failure causes for Claude skill and OpenCode/custom-harness paths.
- Preserve source, pending, accepted, and rejected outputs for fidelity checks.

**Acceptance:** Baseline results are reproducible and do not combine engine time
with model or transport latency.

**Delivered:** `scripts/benchmark-agent-workflow.mjs` and
`scripts/lib/agent-performance-cases.mjs` provide a repeatable observational
benchmark over five deterministic legal-edit shapes plus ordering, stale-handle,
ambiguity, and foreign-revision diagnostics. The report separates native timing,
heap, serialized request size, and lifecycle fidelity from explicitly unmeasured
LLM/provider/transport latency and writes
`tmp/benchmarks/agent-workflow-latest.json`. Real Claude and OpenCode observations
remain external harness runs rather than simulated repository measurements.
The checked seven-iteration/two-warmup run is refreshed as each protocol work
package lands. WP-03 measurements are recorded below. Native execution remains
in the same tens-of-milliseconds range for both shapes, confirming that the
sample's immediate benefit is protocol compactness rather than a claimed core
engine speedup.

### WP-01: Define and implement the agent session facade [COMPLETED 2026-09-12 AS DEVELOPMENT SAMPLE]

**Goal:** Provide the common inspect/apply/resolve protocol without duplicating
core behavior.

- Add a Node-only development session facade outside the published package.
- Do not add a package export, public declaration, or core capability requirement.
- Implement explicit profiles and effective-option reporting.
- Store package-scoped revision tokens and session-bound target descriptors.
- Return canonical results without collapsing receipts or validation.
- Ensure failed work never exposes original bytes as completed output.

**Acceptance:** In the sample harness, an ordinary full-text edit requires one
inspect call and one apply call, with no operations file, repeated exact target,
or repeated policy flags. The published Node facade remains unchanged.

**Delivered:** `examples/agent-session-wrapper.mjs` demonstrates explicit safe
profiles, stateful `DocxDocument` reuse, package-revision-bound opaque handles,
context-window inspection, narrow redline/delete/comment translation, atomic
retry semantics, handle invalidation, and review resolution. Focused tests in
`tests/agent_session_example_tests.mjs` verify delegation, accepted/rejected
text, rollback, stale handles, context, and absence from `node/index.js`.
`package.json` also excludes the development wrapper and its benchmark inputs
from the published package contents.

### WP-02: Add opaque handles and context-window inspection [COMPLETED 2026-09-12 AS DEVELOPMENT SAMPLE]

**Goal:** Remove target-copying and search-then-range turns.

- Add session-local opaque target handles.
- Add `around` context to inspection/search.
- Distinguish hits from context paragraphs.
- Bind handles to exact package revision and revision view.
- Invalidate or refresh handles after mutations.
- Return new revision/handle information after successful apply.

**Acceptance:** A search for a provision returns enough surrounding text to draft
the edit and enough hidden target state to apply it safely without a second
extraction.

**Delivered:** The example session registry retains strict descriptors and the
package revision while exposing only short handles and drafting context.
`inspect({ search, around })` labels direct matches separately from context,
supports accepted/rejected views, and returns the active package token and
effective profile. A mutation retires old handles and returns `refreshedTargets`
with new revision-bound handles for successful surviving targets. Focused tests
verify context-window classification, stale-handle refusal, handle reuse within
one revision, and a follow-up edit through a returned handle without another
inspection.

### WP-03: Add localized exact-span edit compilation [COMPLETED 2026-09-12 AS DEVELOPMENT SAMPLE]

**Goal:** Let small wording changes produce small agent requests.

- Define typed replacement inputs and deterministic matching.
- Detect zero, duplicate, overlapping, and conflicting spans.
- Construct complete desired text from the session snapshot.
- Emit one canonical redline per target.
- Preserve existing revision-policy and boundary failures.
- Return patch-specific structured errors and recovery actions.

**Acceptance:** Punctuation, notice-period, and simple mutuality cases generate the
same accepted text and valid tracked-change lifecycle as equivalent full-text
operations while materially reducing generated request tokens.

**Delivered:** `compileExactReplacements` resolves all exact matches against the
immutable text stored with a handle, requires `occurrence` for duplicates,
deduplicates identical patches, rejects missing/ambiguous/overlapping/conflicting
patches with structured recovery actions, and constructs one complete `modified`
value for the existing redline path. It does not write OOXML. Focused unit and
fixture tests cover simultaneous length-changing replacements, term changes,
simple mutuality, accepted/rejected lifecycle fidelity, failure immutability, and
the refreshed-handle chain.

In the checked seven-iteration/two-warmup benchmark, the development session
reduced serialized apply-request bytes by 86.11% for punctuation, 89.96% for a
term-duration phrase, and 82.22% for deterministic mutuality; the full rewrite
and mixed batch remained supported.
These are JSON byte measurements, not provider-token or end-to-end latency
claims. The benchmark continues to report native time separately and leaves the
WP-04 ordering diagnostic failing in the known split-first direction.

### WP-04: Compile and bind source-targeted batches

**Goal:** Make independent operations caller-order agnostic.

- Implement shared batch compilation for preflight and apply.
- Add `SourceTargetRegistry` and rollback rehydration.
- Resolve source targets against the immutable batch-start document.
- Add read/write footprint and conflict analysis.
- Integrate capture dependencies and created-target selectors.
- Replace false current-fingerprint checks with compiled source preconditions.
- Preserve execution order, source result indexes, savepoints, receipts, and
  atomic rollback.

**Acceptance:** The reproduced split-before-later-target case succeeds in either
input order, while same-source conflicts fail before mutation with causal errors.

### WP-05: Add the recovery registry and retry plan

**Goal:** Make the next safe action machine-readable.

- Centralize normalization and code-specific recovery metadata.
- Preserve bounded code-specific details through facade and CLI compaction.
- Add unique current-target context, ambiguity candidates, issue summaries,
  authorization flags, and structured revision tokens.
- Add atomic/progressive retry plans.
- Correct `EXISTING_REVISIONS` guidance so surgical slicing is distinguished from
  authorization-sensitive normalization.
- Add complete/partial exit behavior for agent CLI use.

**Acceptance:** Each common failure payload tells a thin harness whether to
reinspect, change one field, choose a candidate, request authorization, report a
library failure, or stop for manual resolution.

### WP-06: Add shell fallback and tiered documentation

**Goal:** Reduce overhead where structured tools are unavailable.

- Support operations JSON from stdin.
- Add an agent-oriented CLI/profile with complete-success exit semantics.
- Decide whether explicit config materially improves measured shell performance;
  implement only if it does.
- Publish the compact consumer quick start.
- Slim and route `AGENTS.md`; preserve advanced material in topic documentation.
- Update version capabilities and wrapper compatibility guidance.

**Acceptance:** A shell-only ordinary batch needs extract plus apply, without
operations-file management or repeated common flags, and the minimum required
instructions fit within the quick-start budget.

### WP-07: End-to-end rollout and regression audit

**Goal:** Demonstrate speed improvements without accuracy regression.

- Re-run the WP-00 corpus through the legacy CLI workflow, compact shell workflow,
  and structured session facade.
- Compare model/tool wall time, tokens, turns, retries, and native engine time.
- Verify package and redline validation.
- Verify Accept All and Reject All accepted/rejected text.
- Verify multi-author revision preservation and comment integrity.
- Confirm old CLI and Node consumers remain compatible.
- Publish measured results; do not substitute estimated latency claims.

**Acceptance:** The facade materially reduces ordinary tool calls and generated
tokens, independent batches no longer require manual reordering, common errors
provide deterministic next actions, and all fidelity gates remain green.

## 12. Focused Verification

Add focused suites near the owned subsystem:

- `tests/agent_session_facade_tests.mjs`
- `tests/agent_target_handle_tests.mjs`
- `tests/agent_context_inspection_tests.mjs`
- `tests/localized_edit_contract_tests.mjs`
- `tests/batch_source_binding_tests.mjs`
- `tests/batch_conflict_graph_tests.mjs`
- `tests/error_recovery_contract_tests.mjs`
- `tests/agent_cli_protocol_tests.mjs`

Extend existing suites where they already own behavior:

- `tests/capture_dependency_graph_tests.mjs`
- `tests/capture_resolution_tests.mjs`
- `tests/operation_preflight_tests.mjs`
- `tests/standalone_operation_runner_tests.mjs`
- `tests/agent_operation_contract_tests.mjs`
- `tests/agent_cli_tests.mjs`
- `tests/agent_cli_edge_tests.mjs`
- `tests/performance_phase1_session_tests.mjs`
- `tests/performance_phase2_boundary_tests.mjs`

Required cases include:

1. target handle succeeds only against its exact package revision;
2. successful mutation invalidates affected old handles;
3. context-window inspection returns stable hit/context classification;
4. localized replacements preserve all untouched text exactly;
5. duplicate and overlapping spans fail without mutation;
6. independent batch permutations produce equivalent results and receipts;
7. an early split does not invalidate a bound later target;
8. same-source writes fail before mutation;
9. created-content dependencies execute after their producer;
10. capture fan-out conflicts are diagnosed before execution;
11. savepoint restore rehydrates compiled bindings correctly;
12. atomic recovery identifies the original input as retry base;
13. progressive recovery identifies output and failed/unattempted operations;
14. ambiguous errors return candidates but never auto-select;
15. comment/revision recovery identifies authorization requirements;
16. generated OOXML failures remain library/builder failures;
17. agent CLI partial completion returns nonzero;
18. original source files are never overwritten implicitly.

Run the closest suites first. Because the completed work crosses session state,
targeting, orchestration, facade, CLI, and documentation contracts, run `npm test`
before release-level handoff. Use existing Word/lifecycle oracles when mutation
ordering could change emitted OOXML; do not create Word fixtures for facade-only
serialization behavior.

## 13. Performance and Quality Gates

Do not use estimated claims such as "95% faster" or "3–8 seconds" as acceptance
criteria. Report observed values by harness and model.

Track:

- time to first relevant document text;
- time from user instruction to completed output;
- native inspect/apply time;
- number of agent/tool turns;
- input and output tokens attributable to library instructions and operations;
- bytes of repeated unchanged clause text in apply requests;
- number of failed operations and recovery turns;
- rate of ambiguous or stale target errors;
- batch permutation success rate for independent operations;
- package/revision validation results;
- accepted/rejected lifecycle equivalence;
- comment and multi-author revision preservation.

Initial directional targets, subject to WP-00 baselines:

- two structured calls for the ordinary inspect-and-apply path;
- no repeated exact target text in handle-based apply requests;
- at least 80% fewer generated operation tokens for localized edits on long
  clauses;
- zero order-dependent failures across permutations of independent source
  targets;
- one-turn deterministic recovery for uniquely stale targets and schema errors;
- no regressions in existing validation, rollback, receipt, and lifecycle suites.

## 14. Compatibility and Rollout

- Keep the reference agent protocol in a non-published example rather than a new
  export or core capability.
- Let production harnesses version their request/result contracts independently
  from internal classes.
- Keep canonical document operations as the durable interoperability contract.
- Keep existing Node facade, standalone runner, and CLI invocation forms working.
- Do not change general progressive defaults in a minor release.
- Put atomic and complete-success behavior in the named agent profile.
- Add recovery fields without removing existing error codes/messages.
- Gate compiled batch execution behind focused parity tests before making it the
  sole apply path.
- During transition, compare compiled and legacy planning on the test corpus, but
  do not maintain two permanent mutation engines.
- Document the minimum existing Node/CLI contract needed by the example. Any
  production skill or wrapper owns its own version negotiation and must fail
  closed with a clear upgrade instruction when the underlying library lacks a
  capability it depends on.

## 15. Risks and Mitigations

### Stable bindings across cloned savepoints

**Risk:** DOM references become invalid when a failed operation restores a cloned
savepoint.  
**Mitigation:** Make rollback rehydration a first-class `SourceTargetRegistry`
contract and test continue-on-error after a restored savepoint.

### Hidden state in sessions

**Risk:** Handles obscure the exact descriptor or revision used.  
**Mitigation:** Scope handles to a package token, return effective profile and
revision data, expose optional diagnostics, and fail stale sessions closed.

### Localized edits across complex Word structures

**Risk:** A simple-looking substring crosses a field, hyperlink, comment,
bookmark, or revision boundary.  
**Mitigation:** Let the canonical engine and existing boundary policies remain
authoritative; reject unsupported spans rather than broadening or flattening the
edit.

### Over-aggressive automatic scheduling

**Risk:** A precedence rule changes comment, formatting, numbering, or revision
semantics.  
**Mitigation:** Bind targets first, schedule only proven-safe combinations, reject
uncertain overlaps, and verify lifecycle outputs. Do not depend on blanket
bottom-up order.

### Error payload growth

**Risk:** Rich diagnostics increase token and privacy costs.  
**Mitigation:** Bound excerpts, return full text only for a uniquely identified
target when it enables safe recovery, use handles for candidates, and summarize
validation issues by code/stage.

### Configuration drift

**Risk:** Ambient config makes the same command behave differently across
machines.  
**Mitigation:** Prefer explicit session profiles, report effective options, and
defer auto-discovered config unless measured shell benefits justify it.

## 16. Definition of Done

This plan is complete when:

1. A development-only agent facade demonstrates inspect, apply, and resolve
   capabilities through the existing published Node authority without becoming a
   package export.
2. Ordinary provision edits use session-bound handles and require no operations
   file or repeated target descriptor.
3. Context-window inspection removes the common search-then-range turn.
4. Localized exact-span edits compile into canonical complete desired text and
   pass all fidelity gates.
5. Source targets are compiled and bound against the batch-start document.
6. Independent batch permutations succeed equivalently, including the reproduced
   split-before-later-target case.
7. Genuine overlaps and created-content dependencies produce causal structured
   results before unsafe mutation.
8. Errors carry versioned recovery actions, authorization requirements, bounded
   corrective context, and explicit retry plans.
9. Agent-mode partial completion is unmistakably unsuccessful to both JSON
   consumers and process-level harnesses.
10. Consumer instructions are materially smaller and advanced guidance is loaded
    only when relevant.
11. End-to-end benchmarks show measured reductions in agent turns and generated
    tokens without package, revision, Accept/Reject, comment, or multi-author
    fidelity regressions.
12. Existing public APIs and supported CLI workflows remain compatible.
