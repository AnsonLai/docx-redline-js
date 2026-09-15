# Localized Sub-Paragraph Patching and Speculative Execution Plan

**Status:** Complete — WP-01 through WP-05 completed 2026-09-14

**Date:** 2026-09-14

**Target:** CLI Contract Version 8 / Next Release (0.7.0)

**Priority:** Eliminate remaining agent latency bottlenecks by removing the "paragraph output tax" and collapsing two-turn mechanical edits into single-turn executions, while preserving all underlying OOXML reconciliation and validation invariants.

---

## 1. Executive Summary

Empirical transcripts from production agent runs (Claude 3.5/3.7, GPT-4o) using `docx-redline` 0.6.2 demonstrate that the OOXML diffing, reconciliation, and packaging engines are no longer the bottleneck (~100–250ms total execution). Over 95% of user-perceived wall-clock time is spent in:

1. **Output Token Generation ("The Paragraph Tax"):** The model must emit the entire modified paragraph in `modified: "..."`. For long legal clauses (200–400 words), sequential token streaming costs 6 to 12 seconds per paragraph, even when changing only a single word or date.
2. **Tool Turn Multiplicity (2-Turn Floor):** Every edit requires at least two full LLM round-trips: `extract` followed by `apply`. With network latency and Time-to-First-Token (TTFT), a two-turn interaction imposes a hard floor of 15–25 seconds regardless of engine speed.
3. **Cognitive Overhead / Deliberation Inflation:** When forced to reproduce a 300-word paragraph verbatim to change one phrase, reasoning models spend extra thinking tokens verifying that they didn't drop commas, punctuation, or sentences from the rest of the clause.

This plan outlines three coordinated improvements to address these bottlenecks directly. Localized replacements are a small extension at the compilation boundary; speculative mutation adds a deliberately narrow, fail-closed convenience path:

1. **Sub-Paragraph Localized Replacements (`replacements: [{ find, replace }]`):** Allow agents to supply targeted `find`/`replace` patches against a resolved paragraph instead of generating the entire paragraph text.
2. **Speculative 1-Turn Apply (`apply --find ... --replace ...`):** Allow single-command search-and-replace for obvious, unambiguous edits without requiring a preliminary `extract` call. An optional text anchor and directional paragraph window can narrow a multi-paragraph provision without admitting text before a heading.
3. **Fast-Path Skill Directives:** Update skill-authoring guidance and fast-start templates with simple instructions directing models to bypass deliberative preludes on routine mechanical edits.

---

## 2. Empirical Latency Breakdown

| Phase | Current 0.6.2 Workflow | Proposed 0.7.0 Workflow | Latency Impact |
|---|---|---|---|
| **Turn 1: Discovery** | Model calls `extract --search ...` | *Skipped on simple edits* via speculative search | **-8 to -15s** (eliminates 1 tool turn) |
| **Model Drafting** | Model reconstructs 300-word paragraph in `modified` | Model emits `{ find: "30 days", replace: "60 days" }` | **-5 to -10s** (drastic token reduction) |
| **Turn 2: Apply** | Model calls `apply` with full text JSON | Model calls `apply` with micro-patch | **-5 to -10s** (smaller payload) |
| **Engine Execution** | Node parses, diffs, commits (~150ms) | Compiles patch to full text, diffs, commits (~155ms) | Neutral (<5ms diff) |
| **Total Turnaround** | **20–45 seconds** | **4–8 seconds** | **Target: ~60–75% reduction** |

These figures are working hypotheses drawn from observed agent transcripts, not release claims. WP-05 must reproduce them with the repository benchmark and report median/p95 wall time, tool turns, request bytes, and generated operation tokens. Engine time and model/network time must be reported separately.

---

## 3. Technical Design

### 3.1 Feature 1: Sub-Paragraph Localized Replacements

#### Concept & Architecture
The underlying OOXML reconciliation engine relies on having the complete target text and complete modified text to compute exact character- and word-level diffs. We do **not** need to alter the low-level OOXML reconciliation engine.

Instead, localized patching is handled at the operation normalization/compilation boundary (`services/operation-batch-compiler.js` and `services/document-operation-contract.js`):
1. Validate the localized-patch shape without requiring `modified` yet.
2. Resolve the target paragraph against the immutable batch-start accepted view via `paragraphId`, `fingerprint`, `index`, or `exactText`.
3. Pass the resolved paragraph text to `compileExactReplacements(sourceText, replacements)` (already prototyped and validated in `examples/agent-session-wrapper.mjs`).
4. Apply all non-overlapping replacements simultaneously to synthesize the canonical `modified` string.
5. Validate the synthesized canonical redline operation.
6. Send it through the unchanged OOXML diff and reconciliation pipeline with its existing transactionality, receipts, validation, and rollback.

This two-stage validation matters because the current public operation validator requires `modified` before the batch compiler resolves its target. The implementation must not relax canonical redline validation globally.

#### Operation Schema Extension
In `docs/schemas/document-operations.schema.json` and `document-operation-contract.js`:

Allow `replacements` on `redline`. Do not add a separate `patch` operation alias in v1; keeping one canonical operation type makes the agent surface smaller:
```json
{
  "type": "redline",
  "target": { "paragraphId": "1A2B3C4D" },
  "replacements": [
    {
      "find": "thirty (30) days",
      "replace": "sixty (60) days",
      "occurrence": 1
    }
  ]
}
```

V1 scope and validation rules:
- Accepted-view, single-paragraph redline operations only.
- Exact, case-sensitive matching against the accepted-view paragraph text.
- Captures, created-content dependencies, rejected-view mutation, and cross-paragraph localized patches are deferred.
- An operation must provide either `modified` (string) OR `replacements` (non-empty array of `{ find, replace, occurrence? }`). Providing both is an error (`INVALID_OPERATION`).
- Each replacement must have non-empty string `find` and string `replace`.
- `replace` may be the empty string so an exact span can be deleted.
- Without `occurrence`, `find` must be unique within the resolved paragraph. With it, the positive 1-based occurrence must exist.
- A replacement where `find === replace` fails as a no-op.
- Identical duplicate replacements may be deduplicated; different replacements for the same range fail with `CONFLICTING_PATCHES`.
- Overlapping replacement ranges fail closed before mutation with the existing `OVERLAPPING_PATCHES` code.

#### CLI Flags
Extend `apply` with inline patch arguments:
```bash
# Target by ID with inline patch:
docx-redline apply contract.docx --target-id 1A2B3C4D --find "thirty (30) days" --replace "sixty (60) days" --output out.docx

# Target by exact text with inline patch:
docx-redline apply contract.docx --target "The Company may terminate upon thirty (30) days..." --find "thirty (30) days" --replace "sixty (60) days" --output out.docx
```

`--target-ref <N>` may also be used when a caller already has a fresh machine paragraph index. `--target-id`/`paragraphId` is preferred when a prior inspection or stateful wrapper already supplied it. Neither is the primary one-turn path because an agent cannot know those values without an earlier inspection, and a numeric paragraph index is more vulnerable to document drift.

Inline flags are convenience syntax for short literals. Wrappers and shell hosts handling quotes, dollar signs, newlines, or other shell-sensitive content should use a structured request or serializer-backed stdin rather than constructing a quoted command string.

---

### 3.2 Feature 2: Speculative 1-Turn Execution

#### Concept
For simple edits (e.g. "Change the termination notice in Section 4.1 from 30 to 60 days"), agents currently run:
1. `docx-redline extract ...`
2. Inspect JSON
3. `docx-redline apply ...`

With speculative execution, the agent can attempt mutation directly in **Turn 1**. If the exact source phrase is globally unique, no contextual anchor is necessary:
```bash
docx-redline apply contract.docx --find "thirty (30) days" --replace "sixty (60) days" --profile agent --output out.docx
```

For a provision spanning multiple paragraphs, `--search` locates contextual text anchors and `--context-range <START:END>` selects physical paragraphs at signed offsets from every matching anchor. The agent does not need to guess the exact machine paragraph number. For a heading whose operative text must appear within the next three paragraphs, use `1:3`: this excludes the heading paragraph and every preceding paragraph.
```bash
docx-redline apply contract.docx --search "Section 4.1" --context-range 1:3 --find "thirty (30) days" --replace "sixty (60) days" --profile agent --output out.docx
```

Here, `--search` does not itself select the paragraph to mutate. This avoids the common case where `Section 4.1` is a heading or lead-in paragraph and the editable language is in a following paragraph. The exact `find` span always selects the mutation location. `--around 3` remains convenient shorthand for a symmetric `--context-range -3:3`, but directional ranges should be preferred for known headings.

Relative-range semantics are explicit:

- `0:0` searches only the anchor paragraph.
- `0:3` searches the anchor and the next three physical paragraphs.
- `1:3` searches strictly within the next three physical paragraphs.
- `-3:-1` searches strictly within the three preceding physical paragraphs.
- `-3:3` searches the anchor plus three paragraphs in both directions and is equivalent to `--around 3`.
- `START` must be less than or equal to `END`; both offsets are bounded to `-20..20` and clamped at document boundaries.

#### Paragraph Number vs. Text Anchor
A visible legal identifier such as "Section 4.1" is useful as a text/context anchor and may be supplied directly from the user's instruction. A machine paragraph reference such as `P42` or `--target-ref 42` is different: the agent normally cannot know it without extraction, and insertions or document drift can invalidate it. Therefore:

- Do not ask for a machine paragraph number as the default speculative workflow; doing so restores the discovery turn this feature is intended to remove.
- If a fresh `paragraphId` is already held by a wrapper/session, use it as the strongest selector.
- If only a fresh machine paragraph index is already available, accept `--target-ref` but retain the ordinary strict-target and revision checks.
- If the user names a visible section or paragraph number, treat that value as `--search` context and let exact `find` text identify the mutation paragraph within an explicit directional window.

The agent chooses only a conservative scope such as `--context-range 1:3`; it is not trusted to guess which machine paragraph inside that scope is correct. The library performs the exact match and refuses the edit unless one paragraph and one permitted range remain.

#### Resolution & Fail-Closed Guardrails
1. **Build the candidate scope:**
   - Default to accepted-view text.
   - With a strong target (`--target-id`, fresh `--target-ref`, or exact `--target`), resolve that target normally and search for `find` only inside it.
   - Without a strong target or `--search`, inspect all accepted-view paragraphs containing the exact, case-sensitive `find` string.
   - With `--search`, find every anchor paragraph using the existing case-insensitive inspection semantics. Repeated header/cross-reference text is expected and does not by itself select or reject a target.
   - Apply `--context-range START:END` independently to every anchor paragraph, then take the union of the resulting physical paragraphs. Only paragraphs in that union are eligible to contain `find`. Overlapping windows and duplicate candidate paragraphs are deduplicated.
   - With no context option, use `0:0`, so only each anchor paragraph is eligible. `--around N` is shorthand for the symmetric range `-N:N`; `--around` and `--context-range` are mutually exclusive.
2. **Require a unique mutation location:**
   - Exactly one distinct eligible paragraph must contain `find`; otherwise fail closed. Multiple search anchors are acceptable only when their combined directional scope still yields one unique mutation paragraph.
   - Within that paragraph, `find` must occur exactly once unless a positive `--occurrence` is supplied.
   - `--occurrence` may disambiguate ranges only within a uniquely resolved paragraph. It must never choose among multiple eligible paragraphs.
   - The replacement must not cross a paragraph boundary.
3. **Never write on failed speculation:**
   - If **0 eligible matches** are found, fail with `PATCH_SOURCE_NOT_FOUND` (or `TARGET_NOT_FOUND` when the contextual/strong target itself is absent). Do not create, truncate, or overwrite the output path.
   - If **multiple eligible paragraphs** remain, fail with `AMBIGUOUS_TARGET` and compact paragraph candidates.
   - If the paragraph is unique but the patch span is ambiguous, fail with `AMBIGUOUS_PATCH_SOURCE` and compact offset candidates.
   - Failed recovery envelopes include compact anchor and patch candidates and direct the agent to narrow the anchor/range, select a strong returned target, or run `extract`; they never recommend retrying unchanged arguments.
4. **Preserve the ordinary apply guarantees:**
   - Once resolved, compile the patch to a canonical `modified` string and run one ordinary redline operation.
   - A successful command must still require `completion: true`, `written: true`, a non-null output path, valid receipts, and no operation errors.
   - Return a compact change summary containing the selected `paragraphId` when present, human-facing document reference, exact applied replacement(s), bounded before/after context, contextual anchor match(es), and resolved directional range. This provides post-call evidence without returning the full paragraph or requiring a verification extraction.
   - The existing revision policy and `--profile agent` semantics remain unchanged.

#### Explicit Success Evidence
The response must make successful mutation distinguishable from merely writing a valid package. Each localized-patch operation result receives one `change` object; do not duplicate it inside the structural mutation receipt:

```json
{
  "status": "ok",
  "written": true,
  "completion": true,
  "results": [
    {
      "index": 1,
      "status": "ok",
      "change": {
        "kind": "localized_replacement",
        "target": {
          "paragraphId": "1A2B3C4D",
          "humanReference": "Section 4.1 — Termination"
        },
        "context": {
          "search": "Section 4.1",
          "range": "1:3",
          "anchorMatchCount": 2
        },
        "replacements": [
          {
            "find": "thirty (30) days",
            "replace": "sixty (60) days",
            "occurrence": 1,
            "beforeExcerpt": "...terminate upon thirty (30) days' written notice...",
            "afterExcerpt": "...terminate upon sixty (60) days' written notice..."
          }
        ],
        "verification": {
          "acceptedViewMatchesCompiledText": true
        }
      }
    }
  ]
}
```

The change summary must be derived from the resolved batch-start text and the actual post-mutation accepted view, not merely echo the request. Before/after excerpts should be centered on the changed range, whitespace-preserving, deterministically bounded (target: at most 120 characters each), and omitted only when unavailable. The complete paragraph remains internal: compare it during verification but do not return it by default. For multiple localized replacements, return one concise entry per applied range in source order.

`acceptedViewMatchesCompiledText: true` confirms the resulting accepted-view paragraph equals the complete text synthesized by the localized replacement compiler. It complements rather than replaces `written`, `completion`, output-path checks, validation, and committed mutation receipts. A missing or false verification flag makes the agent profile incomplete and nonzero.

For example, an absent contextual anchor returns a recovery envelope such as:
```json
{
  "error": {
    "code": "TARGET_NOT_FOUND",
    "message": "No paragraph matched search query: 'Section 4.1'.",
    "recovery": {
      "action": "reinspect",
      "requiresReinspection": true,
      "sameArgumentsSafe": false
    }
  }
}
```
When speculative apply succeeds, the mutation is completed in **one tool call without a discovery call**. When it fails, it provides immediate recovery diagnostics pointing to `extract`.

This is a deterministic literal-edit fast path, not semantic clause selection. A unique string match can still be the wrong legal provision if the caller supplies a poor anchor. Skills should use it only when the user's request supplies or strongly implies the existing literal text and the contextual anchor is adequate. Requests such as "make this provision mutual" still require inspection and drafting.

---

### 3.3 Feature 3: Fast-Path Skill Directives (Keep It Simple)

We avoid complex prompt state machines or heavy prompt engineering. Instead, we add simple, crisp behavioral directives in `SKILL_AUTHORING.md`, `AGENT_FAST_START.md`, and the user's `SKILL.md`:

```markdown
### Fast-Path for Simple Edits
- For literal mechanical edits (names, numbers, dates, defined terms, or an exact isolated phrase replacement):
  - Prefer localized `find`/`replace` patches over generating entire paragraphs.
  - If `find` is expected to be globally unique, you may run `apply --find "..." --replace "..."` directly.
  - For a heading followed by a multi-paragraph provision, use a meaningful text anchor and a directional window, for example `--search "Section 4.1" --context-range 1:3`.
  - Use symmetric `--around N` only when text on either side of the anchor is genuinely eligible.
  - If a fresh `paragraphId` is already available, prefer it over a numeric paragraph index.
  - If a speculative apply reports `requiresReinspection: true`, fall back to `extract`.
- Do not use speculative execution for semantic drafting requests such as making a provision mutual. Inspect enough context first.
- For a literal edit, proceed directly to the tool call without a conversational preamble unless the user requested alternatives or explanation.
```

---

## 4. Implementation Work Packages

### WP-01: Localized Replacements Compiler Integration — Complete
* **File:** `services/operation-batch-compiler.js`, `services/document-operation-contract.js`
* **Task:**
  * Port `compileExactReplacements` from `examples/agent-session-wrapper.mjs` into a reusable service module; keep the example wrapper consuming that shared helper rather than retaining a second implementation.
  * Support `replacements` array on operation objects.
  * Implement the two-stage patch-shape validation, batch-start target binding, compilation to `modified`, and canonical redline validation described above.
  * Restrict v1 to accepted-view, single-paragraph source targets.
* **Verification:** Unit tests verifying single, multi, and conflicting replacements against a paragraph.

### WP-02: CLI Flags & Schema Updates — Complete
* **Files:** `node/cli.js`, `node/cli-help.js`, `docs/schemas/document-operations.schema.json`
* **Task:**
  * Add `--find`, `--replace`, `--occurrence`, and `--target-id` flags to strong-target inline `apply`; retain the existing `--target-ref` path for callers with a fresh paragraph index.
  * Update `document-operations.schema.json` to formally include `replacements` schema.
  * Update `docx-redline apply --help` with machine-readable option metadata and compact examples.
* **Verification:** CLI tests validating inline `--find`/`--replace` execution and schema validation.

### WP-03: Speculative Search-and-Apply — Complete
* **Files:** `node/cli.js`, `services/standalone-operation-runner.js`
* **Task:**
  * Allow exact global `--find` resolution on `apply` when a strong target is omitted.
  * Allow `--search <query> [--context-range START:END]` to bound candidate paragraphs without treating the anchor paragraph as the mutation target.
  * Add `--context-range <START:END>` and define `--around N` on speculative apply as shorthand for `--context-range -N:N`; reject combining both options.
  * Evaluate every search-anchor match, union and deduplicate its directional windows, then require one unique patch location across the combined scope.
  * Implement unique-paragraph and unique-range resolution with safe fail-closed ambiguous/miss envelopes and no output write on failure.
  * Return the compact, post-mutation-verified `change` object on success so the agent can confirm both the selected legal location and the exact before/after language from the apply result.
  * Ensure `--profile agent` sets exit code 2 on ambiguous/missing speculative targets.
* **Verification:** CLI tests verifying successful 1-turn apply, 0-match fail-closed, and multi-match ambiguous fail-closed.

**Progress log (2026-09-14):**

- Implementation boundary selected: resolve speculative candidates in the CLI
  against the already-open accepted-view inspection, translate the unique result
  into one ordinary strong-target localized redline, and delegate to the existing
  facade. No OOXML mutation, validation, receipt, or rollback logic will be
  duplicated.
- Implemented targetless exact global patch resolution and case-insensitive
  contextual anchors with bounded signed `--context-range` windows. Search-anchor
  windows are unioned and deduplicated before the exact patch source is required
  to resolve to one paragraph.
- Preserved the strong-target apply pipeline: the CLI converts a speculative
  match to an accepted-view descriptor and delegates compilation, mutation,
  package validation, receipts, and rollback to the existing facade.
- Added fail-closed `TARGET_NOT_FOUND`, `PATCH_SOURCE_NOT_FOUND`, and
  `AMBIGUOUS_TARGET` recovery envelopes. Candidate descriptors survive CLI error
  normalization, and failed speculation never reaches output writing.
- Added post-mutation `change` evidence with resolved location, bounded
  before/after excerpts, anchor count/range, commit disposition, and actual
  accepted-view verification. Atomic rollback now marks this evidence
  `committed: false` and `finalDisposition: rolled_back`.
- Added `tests/cli_speculative_apply_tests.mjs`, covering global resolution,
  repeated section anchors, directional exclusion of preceding paragraphs,
  overlapping-window deduplication, ambiguity despite `--occurrence`, protected
  output on refusal, missing anchors, and invalid ranges. Focused speculative,
  localized replacement, agent CLI, and help suites pass.
- WP-04 documentation and contract-version work is now in progress.

### WP-04: Skill Authoring & Fast-Start Documentation — Complete
* **Files:** `docs/SKILL_AUTHORING.md`, `docs/AGENT_FAST_START.md`, `docs/AGENT_KNOWLEDGE_BASE.md`, `README.md`
* **Task:**
  * Document `replacements` and speculative 1-turn `apply`.
  * Add concise fast-path instructions for AI agents.
  * Warn generated skills and harnesses that the library remains before 1.0: record and pin the exact tested package release, declare the required CLI contract/capabilities, and treat a different release or contract as unverified until compatibility tests pass.
  * Increment `contractVersion` to 8 and declare capabilities: `localized-replacements-v1`, `speculative-search-apply-v1`, and `localized-change-summary-v1`.
* **Verification:** Run `node tests/agent_documentation_contract_tests.mjs`.

**Progress log (2026-09-14):**

- Bumped the CLI contract to version 8 and advertised
  `localized-replacements-v1`, `speculative-search-apply-v1`, and
  `localized-change-summary-v1`. Protocol and recovery contract tests now assert
  all three capabilities.
- Added public `LocalizedReplacementChange` declarations to the standalone and
  root type surfaces, including commit disposition, contextual scope, bounded
  excerpts, and accepted-view verification.
- Updated Agent Fast Start, Skill Authoring, README, CLI help, and the advanced
  knowledge base. Literal mechanical edits lead with targetless one-turn apply;
  semantic requests such as “make this provision mutual” still lead with
  focused extraction and drafting.
- Documented directional heading scope (`1:3`) separately from symmetric
  `--around`, repeated-anchor behavior, intra-paragraph-only `--occurrence`, and
  the localized change evidence required before reporting completion.
- Kept the authoring documents narrow after adding the feature: Agent Fast Start
  is 58 lines/380 words and Skill Authoring is 166 lines/980 words. Its pre-1.0
  warning now demonstrates a pinned planned `0.7.0` release and contract 8.
- Verification completed: 113/113 tests pass sequentially, type declarations
  pass, dependency/isolation checks pass, the focused changed-file lint passes,
  and `npm pack --dry-run --json` includes the compiler and narrowed published
  documentation. The repository-wide lint command still reports two unrelated
  pre-existing unused constants in
  `tests/cross_author_slicing_advanced_synthetic_tests.mjs`; neither file nor
  warning was changed in WP-03/WP-04.

### WP-05: Test Suite & Verification — Complete
* **Files:** `tests/localized_replacement_tests.mjs`, `tests/cli_speculative_apply_tests.mjs`
* **Task:**
  * Automated regression tests covering all single-turn and multi-patch scenarios.
  * Benchmark token and turn reduction against baseline 0.6.2 workflows, separating engine time from model/network time and reporting median/p95 measurements.

**Progress log (2026-09-14):**

- WP-05 started. The existing benchmark infrastructure measures current-process
  native/CLI execution, serialized requests, and protocol-call counts, but does
  not yet compare contract-8 speculative apply against the 0.6.2 two-turn
  extract/full-paragraph workflow. A dedicated reproducible benchmark and
  checked validation report will be added without estimating provider latency.
- The initial WP-03 suite covers the primary success/refusal paths. WP-05 will
  fill the remaining matrix: every directional range form, same-paragraph
  occurrence ambiguity, strong-ID mismatch, excerpt bounds/order, and rollback
  disposition evidence.
- Expanded both focused suites with that remaining matrix; all focused tests and
  scoped lint pass. Atomic rollback now has a direct regression assertion that
  localized evidence becomes non-committed and `rolled_back`.
- Added `npm run benchmark:localized` and ran 3 warmups plus 11 measured
  iterations for short punctuation, a longer provision, and directional heading
  scope. The one-turn path cut tool turns by 50% and the deterministic request
  token proxy by 64.77%–78.98%; median native engine/I/O time remained neutral
  (-0.58% to 5.41% improvement). Provider/model time remains explicitly
  unmeasured.
- Checked the methodology and results into
  `docs/validation-reports/2026-09-14-localized-patching-rollout.md`. The report
  is published with the package; the fixture-dependent benchmark script remains
  development-only and is excluded from package files.
- Final verification passed: 113/113 test suites, type declarations (123 runtime
  exports), dependency/Word-API isolation, focused lint, documentation contract,
  and diff whitespace checks. Package dry-run contains 154 entries, includes
  the compiler and rollout audit, and excludes the fixture-dependent benchmark.
- No WP-05 work remains. External Claude/OpenCode transcript instrumentation is
  deliberately outside this repository benchmark and is the only way to measure
  actual provider reasoning, streaming, network, or billed-token savings.

---

## 5. Verification Plan

### Automated Test Matrix
1. **Core Patching:** `node tests/localized_replacement_tests.mjs`
   - Single replacement compiles to expected paragraph text.
   - Multiple non-overlapping replacements within one paragraph compile in one simultaneous pass.
   - Ambiguous matches without occurrence fail closed with candidate offsets.
   - Missing find text fails closed with `PATCH_SOURCE_NOT_FOUND`.
   - Empty replacement deletes the exact span; empty `find` and no-op replacements fail.
   - Captures, rejected-view targets, and cross-paragraph patches fail as unsupported in v1.
   - Compiled patch output produces the same canonical mutation behavior as supplying the synthesized `modified` text.
2. **CLI Speculative Apply:** `node tests/cli_speculative_apply_tests.mjs`
   - A globally unique exact `find` + `replace` succeeds and commits without prior extraction.
   - A repeated heading/cross-reference anchor plus `--context-range 1:3` can select a unique `find` in a following paragraph without admitting the anchor or preceding paragraphs.
   - `0:0`, `0:3`, `1:3`, `-3:-1`, and `-3:3` obey the documented inclusive offset semantics and document-boundary clamping.
   - Overlapping windows deduplicate the same paragraph and patch range.
   - Multiple anchor matches are allowed when their combined scope yields one unique patch location; multiple resulting patch paragraphs fail closed.
   - A zero-match context anchor returns `TARGET_NOT_FOUND`; a missing patch span returns `PATCH_SOURCE_NOT_FOUND`.
   - Multiple eligible paragraphs return `AMBIGUOUS_TARGET`; multiple spans in one unique paragraph return `AMBIGUOUS_PATCH_SOURCE`.
   - `--occurrence` never chooses among multiple eligible paragraphs.
   - Every failed speculative case leaves the output path untouched.
   - Known `paragraphId` succeeds; a stale or mismatched ID fails through the ordinary target recovery contract.
   - Success includes the compact resolved paragraph/human reference, applied patch occurrence, bounded before/after excerpts, anchor count/range, and positive accepted-view verification without echoing the complete paragraph.
   - The reported excerpts are derived from actual source/output text, preserve meaningful whitespace, are deterministically bounded, and list multiple patches in source order.
   - A missing or failed accepted-view comparison prevents agent-profile completion.
3. **Full Regression:**
   - `npm test`
   - `npm run check:types`
   - `npm run test:isolation`
