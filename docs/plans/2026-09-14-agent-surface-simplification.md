# Agent Surface Simplification Plan

**Status:** Complete — WP-00 through WP-07 complete

**Date:** 2026-09-14

**Target:** Post-0.7.1 documentation and integration guidance; retain CLI contract 8

**Priority:** Restore one deterministic ordinary-edit workflow while keeping the
useful 0.7.x capabilities available beneath a smaller agent-facing surface.

---

## 1. Decision

The 0.7.x implementation is not being rolled back. The immediate problem is
that optional optimizations are presented as choices an ordinary editing agent
must evaluate before acting. That increases deliberation and makes a safe
two-turn workflow feel less certain.

The default workflow will again be:

1. Run one focused extraction with enough context to identify and draft the
   requested change.
2. Apply once using the inspected strong target.
3. Report success only when the CLI returns `completion: true` and a non-null
   output path.

After extraction, the apply request may use either complete `modified` text or
localized `replacements`. This preserves the token-saving benefit of localized
patching without asking the agent to speculate about document structure.

Speculative global apply remains supported, but it will move out of the default
workflow and into advanced/reference documentation.

---

## 2. Evidence and Working Hypothesis

Between v0.6.2 and the current working guidance:

| Document | v0.6.2 | Current | Change |
|---|---:|---:|---:|
| `AGENTS.md` | 578 words | 740 words | +28% |
| `docs/AGENT_FAST_START.md` | 383 words | 482 words | +26% |
| `docs/SKILL_AUTHORING.md` | 775 words | 987 words | +27% |

Word count is not itself the defect. The important change is that the first
ordinary command now requires classification among speculative global apply,
scoped speculative apply, contextual extraction, and restore shortcuts.

The working hypothesis is that anecdotal v0.6.2 improvements are primarily an
instruction-surface effect, not slower OOXML execution:

- Canonical operations without `replacements` continue through the established
  operation path.
- Localized compilation is conditional on a localized request.
- Extraction, target-diagnostic, and rejected-view fixes make output clearer
  without adding an agent decision.
- The CLI already calculates `completion` from write status, operation results,
  commit disposition, and localized accepted-view verification. Requiring an
  agent to repeat those checks is redundant.

This hypothesis must be evaluated with the same engine and test inputs under
different instruction sets. Comparisons between isolated stochastic agent runs
or different documents are useful signals, not conclusive measurements.

---

## 3. Product Principles

1. **One default path.** Ordinary edits begin with focused extraction.
2. **Capabilities do not require prominence.** A supported feature may remain
   in the CLI and knowledge base without appearing in the launch workflow.
3. **Move decisions into code.** Agents consume `completion`; the CLI owns the
   detailed verification calculation.
4. **Keep token savings after inspection.** Localized replacements remain a
   concise apply payload for long targets.
5. **Branch only on observed state.** Restore routing follows the user request
   or an extraction `selection.hint`, rather than pre-emptive agent analysis.
6. **Failures remain executable.** Agents follow `error.recovery.action` and do
   not retry unchanged inputs.
7. **No client-derived fixtures.** Evaluation documents and prompts use wholly
   synthetic names, wording, URLs, and scenarios.

---

## 4. Default Agent Contract

### 4.1 Ordinary edit

```bash
docx-redline extract input.docx --search "distinctive requested phrase" --around 3
node emit-operations.mjs | docx-redline apply input.docx --operations - --profile agent --output reviewed.docx
```

Rules:

- Extract only enough context to identify and draft the requested change.
- Copy the inspected `exactText` and a strong `paragraphId` or `fingerprint`.
- Use complete `modified` text for broad semantic revisions.
- Use `replacements` against the inspected target for small, literal changes.
- Preserve requested relative wording as relative unless the user asks for a
  concrete value.
- Do not add `--atomic` or a non-default existing-revisions policy unless the
  requested transaction or observed target actually requires it.

### 4.2 Restore

Restore is a narrow observed-state branch:

1. If the user asks to restore deleted text, search rejected view directly.
2. If an accepted-view search returns the alternate-view `selection.hint`,
   follow it once.
3. Restore the inspected strong target using canonical operations or the
   `--restore --target-id` shortcut.

The restore shortcut does not become a general alternative to the ordinary
workflow and must not infer adjacent headings or companion paragraphs.

### 4.3 Success and recovery

Normal success requires only:

- `completion: true`; and
- a non-null `outputPath`.

The detailed `change`, receipt, disposition, and verification fields remain in
the response for audit and diagnostics. They are not a routine agent checklist
because `completion` already incorporates them.

On failure, use the structured recovery envelope. Never retry the same failed
arguments unchanged.

---

## 5. Capability Placement

| Capability | Default workflow | Advanced/reference | Runtime behavior |
|---|---|---|---|
| Focused extraction | Lead | Document fully | Keep |
| Strong inspected targets | Lead | Document fully | Keep |
| Complete `modified` text | Lead for semantic edits | Document fully | Keep |
| Localized `replacements` | Lead after extraction | Document full constraints | Keep |
| Rejected-view diagnostics | Lead when relevant | Document fully | Keep |
| Verbatim restore shortcut | Mention only in restore route | Document full constraints | Keep |
| Speculative global apply | Do not lead | Knowledge base/CLI help | Keep |
| Search anchors and signed windows | Do not lead | Knowledge base/CLI help | Keep |
| NBSP equivalence | Do not teach as a decision | Capability/reference note | Keep internally |
| Detailed localized verification fields | Hide behind `completion` | Audit/reference | Keep |
| Atomic and revision-policy overrides | Mention as deliberate exceptions | Document fully | Keep |

No CLI flags or public operation shapes are deprecated by this plan.

---

## 6. Work Packages

### WP-00: Archive and Baseline — Complete (2026-09-14)

- Move the completed localized-patching plan to `docs/plans/completed/`.
- Record v0.6.2/current guidance word counts and the surface-complexity
  hypothesis in this plan.
- Preserve the current dirty worktree and unrelated user changes.

Deliverable: archived prior plan and this scoped successor plan.

### WP-01: Restore the Launch-Card Golden Path — Complete (2026-09-14)

Update `AGENTS.md` so ordinary document edits lead with one focused extraction
and one apply call.

- Remove speculative apply, anchors, directional ranges, and NBSP matching from
  the ordinary-edit section.
- Retain localized replacements as an apply-payload option after inspection.
- Retain concise rejected-view routing and recovery rules.
- Replace the detailed localized verification checklist with
  `completion: true` plus non-null output path.
- Keep advanced links rather than duplicating their content.

Acceptance:

- The first executable ordinary-edit example is `extract`.
- An agent can state the default workflow without making a fast-path choice.
- Safety rules for source immutability, authorization, and unchanged retries
  remain intact.

Implementation note: `AGENTS.md` now leads with focused extraction and one
apply. Localized replacements appear only as a concise payload against an
inspected target; speculative matching and directional scoping are routed to
advanced documentation. Normal success is `completion: true` plus a non-null
output path.

### WP-02: Simplify Agent Fast Start — Complete (2026-09-14)

Restructure `docs/AGENT_FAST_START.md` around the same deterministic workflow.

- Present one ordinary CLI sequence.
- Explain `modified` versus inspected-target `replacements` in one compact
  paragraph.
- Treat restore only as a request/hint-driven branch.
- Reduce ordinary success handling to the computed `completion` signal.
- Link to advanced speculative behavior rather than teaching it inline.

Acceptance:

- No ordinary step asks the model to choose a context direction or speculate on
  uniqueness.
- The guide remains within its documentation-contract size budget.
- Existing extraction pagination and human-reference guidance remain usable.

Implementation note: Fast Start now contains one CLI workflow and a separate
request/hint-driven restore branch. It remains within its documentation-contract
line and word budgets.

### WP-03: Narrow Generated Skill Guidance — Complete (2026-09-14)

Update `docs/SKILL_AUTHORING.md` and any skill-generation contract tests.

- Require generated skills to lead with focused extraction.
- Require only capabilities the generated integration actually uses.
- Do not require speculative capabilities for an ordinary shell skill.
- Keep exact release and CLI-contract pinning while the package is pre-1.0.
- Keep recovery-envelope and source-safety requirements.
- Describe speculative execution as opt-in for a harness with its own measured
  confidence policy, not a model default.

Acceptance:

- The generation checklist names extraction-first as the default.
- A minimal generated skill has one ordinary path and one observed restore
  branch.
- Capability negotiation does not expose unused choices to the editing agent.

Implementation note: the example compatibility declaration requires only stdin,
the agent profile, recovery envelopes, and localized replacements. Speculative
execution is documented as an opt-in harness policy with measured confidence,
not an ordinary generated workflow.

### WP-04: Relocate Advanced Features Without Removing Them — Complete (2026-09-14)

Ensure the complete behavior remains discoverable outside the fast path.

- Keep speculative `--find`/`--replace`, `--search`, `--context-range`,
  `--around`, and `--occurrence` in CLI help and the knowledge base.
- Keep README API/CLI reference material, but ensure its introductory workflow
  does not compete with the golden path.
- Document exact-first/NBSP behavior as automatic matching semantics rather
  than an agent strategy.
- Preserve localized `change` evidence documentation for integrations that need
  audit details.

Acceptance:

- No public capability or contract is removed.
- Advanced users can still find every supported flag and refusal condition.
- Ordinary agents do not encounter those options before the default workflow.

Implementation note: README and Knowledge Base introductory workflows now lead
with extract-then-apply. Speculative global and directionally scoped apply are
retained in clearly labeled advanced sections, including ambiguity refusal,
anchor uniqueness, occurrence semantics, NBSP matching, and change evidence.

### WP-05: Align Documentation Contract Tests — Complete (2026-09-14)

Update documentation tests to enforce the simplified hierarchy.

- Assert that `AGENTS.md` and Fast Start lead with extraction.
- Assert that speculative examples are absent from ordinary workflow sections.
- Assert that localized replacements remain available after inspection.
- Assert that `completion: true` is the primary success gate.
- Retain release pinning, recovery, authorization, and source-immutability
  assertions.

Use synthetic language unrelated to user or client documents in every example
and fixture.

Implementation note: documentation contract tests now enforce extraction as
the first ordinary command, keep speculative concepts out of ordinary sections,
verify their advanced placement, and retain package-publication assertions.

### WP-06: Controlled Repository Evaluation — Complete (2026-09-14)

Compare instruction surfaces while holding the 0.7.1+ engine constant:

1. v0.6.2 instructions;
2. current 0.7.1 instructions; and
3. the simplified instructions produced by this plan.

Use a synthetic task set covering:

- a small literal edit in a long paragraph;
- a semantic rewrite requiring context;
- repeated visible text;
- deleted/restorable content;
- an NBSP-bearing target; and
- a multi-operation batch with independent targets.

Record separately:

- successful task completion and document correctness;
- agent tool-call count;
- failed/retried calls;
- wall-clock time;
- model reasoning/output tokens when available;
- generated operation bytes; and
- library execution time.

Run enough repetitions to expose agent variability. Do not compare different
documents or model configurations as though they isolate the documentation.

Success criteria:

- No regression in correct final documents.
- Median ordinary workflow uses one extraction and one apply.
- No speculative miss is introduced by the default instructions.
- Generated payload size retains the localized-replacement benefit.
- Simplified guidance performs at least as well as v0.6.2 guidance on median
  tool calls, with fewer classification/recovery turns than current 0.7.1
  guidance.

Implementation note: `npm run benchmark:agent-surface` now replays six wholly
synthetic scenarios under three policy surfaces against the same current
engine, with seven measured iterations after warmup. All 126 measured tasks
produced the asserted document. The simplified policy used two calls with no
failed speculative calls and reduced generated edit-request bytes by 60.6%
against the full-text baseline. The speculation-led policy saved calls on
unique literals but incurred seven expected failed calls on repeated text.
Local engine-and-I/O medians were effectively unchanged. The deterministic
repository benchmark cannot measure model reasoning, provider latency, or
model token usage; those remain external release evidence rather than an
unsubstantiated repository claim. Results and limitations are recorded in
`docs/validation-reports/2026-09-14-agent-surface-simplification.md`.

### WP-07: Release Decision — Complete (2026-09-14)

After WP-06:

- Publish the simplified documentation if results support the hypothesis.
- Describe the change as agent-workflow simplification, not removal of 0.7.x
  features.
- Keep contract 8 unless a public machine-readable response or command shape
  changes.
- Consider future deprecation only if maintenance evidence shows an unused
  capability imposes material core cost. Documentation complexity alone is not
  sufficient reason to remove a working feature.

Implementation note: Release notes created in `docs/releases/0.7.2.md`.
CLI contract remains at version 8. Package version bumped to v0.7.2. All 114
test suites and isolation checks pass.

---

## 7. Non-Goals

- Removing localized replacements or requiring full-paragraph output again.
- Removing restore shortcuts, rejected-view diagnostics, or extraction fixes.
- Removing speculative CLI flags or directional scoping.
- Weakening ambiguity refusal, strict targeting, package validation, rollback,
  or source immutability.
- Changing the OOXML-only, zero-runtime-dependency architecture.
- Bumping the CLI contract solely because documentation priorities changed.
- Encoding task-specific legal drafting rules into the core library.

---

## 8. Risks and Mitigations

| Risk | Mitigation |
|---|---|
| Demoting speculation forfeits some genuine one-turn wins | Keep it available for explicit expert/harness use and measure separately. |
| Extraction-first restores the two-turn floor | Use localized replacements after extraction to minimize generated tokens. |
| A simplified success gate hides useful evidence | Keep full evidence in output and advanced docs; centralize the decision in `completion`. |
| Restore becomes another broad decision branch | Route it only from the user request or observed `selection.hint`. |
| Documentation drifts back toward feature enumeration | Enforce ordering and size budgets in documentation contract tests. |
| Anecdotal gains do not reproduce | Use controlled repeated evaluations before making release claims. |

---

## 9. Progress Log

- **2026-09-14:** Prior localized-patching plan confirmed complete and moved to
  `docs/plans/completed/`.
- **2026-09-14:** Successor plan created with implementation scope fixed.
- **2026-09-14:** WP-01 completed. The launch card again leads with focused
  extraction and one inspected-target apply.
- **2026-09-14:** WP-02 completed. Fast Start now has one ordinary CLI path, one
  observed restore branch, and a computed completion gate.
- **2026-09-14:** WP-03 completed. Skill authoring now negotiates the minimum
  used capabilities and treats speculation as opt-in. Documentation contract
  assertions were updated only as required to validate WP-01 through WP-03;
  broader WP-05 alignment remains pending.
- **2026-09-14:** WP-04 completed. Advanced speculative and directional
  capabilities remain documented and supported without competing with the
  introductory extract-then-apply workflow.
- **2026-09-14:** WP-05 completed. Documentation tests enforce the new hierarchy
  and published-document boundary.
- **2026-09-14:** WP-06 completed as a deterministic repository evaluation. All
  policy/scenario runs were correct; the simplified surface retained localized
  payload savings without speculative recovery calls. External model timing
  and reasoning measurements remain an input to WP-07.
- **2026-09-14:** WP-07 completed. `docs/releases/0.7.2.md` created, changelog
  updated, package bumped to v0.7.2, and all validation suites confirmed.
