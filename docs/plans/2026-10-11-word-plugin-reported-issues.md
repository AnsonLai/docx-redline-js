# Word Add-in Reported Issues (2026-10-11)

Date: 2026-10-11
Status: IN PROGRESS. Baseline: `master` after the 2026-09-30 fidelity and list plans (unreleased 0.9.0, contract 9).
Source: `AIWordPlugin/docs/library-issues/` (open reports as of 2026-10-02, against installed 0.8.4). Each report has a
portable reproducer that exits 0 while the defect reproduces and 1 once fixed. To run one against this working tree, copy
it into `tmp/library-issues/` (so `../../tests/setup-xml-provider.mjs` resolves to this repo's tests) and run it with
node; the package self-reference resolves `@ansonlai/docx-redline-js` to the working tree.

Related plans: `2026-09-30-revision-fidelity-invariants.md` (verifier, Word lanes), `2026-09-30-canonical-list-operations.md`
(list-format). Update them in the same change when an item here moves an item there.

## Progress and hand-off (keep current; resume here)

| # | Report | State on 2026-10-11 (working tree) | Next |
|---|---|---|---|
| R1 | `2026-09-30-canonical-list-operations.md` | **Fixed by design**: `list-format` convert-plain / change-format / join (list plan WP-4/5) and redline + list-format composition (WP-6). The reproducer now gets an explicit refusal instead of a silent `ok` (atomic batch -> `UNSUPPORTED_LIST_CONVERSION` with `suggestedOperation`). | Plugin moves it to `completed/` after upgrading |
| R2 | `2026-10-02-separate-list-operations-restart-numbering.md` (0.8.4 regression) | **Partly fixed**: labels are now `A.`, `B.` (start values), but the two headers are still separate `w:num` instances (`numId 7`, `11`). | I2 |
| R3 | `2026-10-02-weak-target-soft-break-deletion.md` | **Reproduces**: weak target (index + exactText) on a `<w:br/>` paragraph tracks the whole paragraph as deleted, `ok`. | I1 |
| R4 | `2026-10-02-generate-redlines-false-cross-author-insertion.md` | **Reproduces**: `generateRedlines: false` + `slice-cross-author` still creates a tracked `w:ins`. | I1 |
| R5 | `2026-10-01-word-reject-table-append-paragraph.md` | Unverified on this tree; Word reported an extra empty paragraph and one leftover revision after `Revisions.RejectAll()` for an end-of-document table append. Our simulator and library reject say 7 paragraphs, so if Word still disagrees our Word-semantics model has a gap. | I3 |
| R6 | `2026-10-02-word-existing-comments-delta-transport.md` | Consumer transport incident, resolved in the plugin; no library defect. | Plugin moves it to `completed/` |

## Why two of these escaped the new safety net

The Reject All verifier only proves "Reject All restores the source". R3 (a tracked whole-paragraph deletion) and R4 (an
unwanted tracked insertion) both reject back to the source, so they pass. Two missing invariants would have caught them:

- **`UNTRACKED_CREATED_REVISION`**: an operation with `generateRedlines: false` must not create any new revision (no new
  `w:ins`/`w:del`/`*PrChange`/paragraph-mark revision ids). Cheap: compare revision ids before/after.
- **`ACCEPT_INTENT`** (deferred in the fidelity plan): when an operation declares its accepted result (`modified`, or a
  localized `replacements` result), Accept All of the target must equal it. A whole-paragraph deletion when the caller asked
  for a one-word replacement fails this. Localized replacements already have `acceptedViewMatchesCompiledText`; R3 shows it is
  not enforced on every path (weak targets).

## Work items (easiest first)

### I1. Untracked edits stay untracked; weak targets never delete (R4, R3) — easy/medium

1. R4: find where the surgical (`slice-cross-author`) path ignores `generateRedlines: false` (start at
   `engine/surgical-mode.js` / `surgical-*` and how `options.generateRedlines` reaches them from
   `engine/oxml-engine.js`). Untracked surgical edits must write plain runs, matching the `merge-same-author` control.
2. Add the `UNTRACKED_CREATED_REVISION` invariant to `services/revision-fidelity-verifier.js` (refuse by default, same
   payload conventions, explanation + retry text in `REVISION_FIDELITY_INVARIANTS`).
3. R3: find why a weak target on a soft-break paragraph routes to a whole-paragraph delete (compare the strong-target path:
   `services/localized-replacement-compiler.js`, `services/operation-batch-compiler.js` source binding,
   `applyToParagraphByExactText`). Fix the routing; and enforce Accept All intent for localized replacements on every path
   (refuse with `PATCH_ROUNDTRIP_MISMATCH` or a new `ACCEPT_INTENT` violation instead of `ok`).
4. Regression tests from both reproducers (bare + `openDocx`), plus matrix generator coverage: weak-target variants of the
   `replacements` kind on the soft-break fixture, and `generateRedlines: false` variants on the multi-author fixture.

### I2. One list across separate operations in a batch (R2) — medium

Separate `redline` operations that convert continuing markers (`A.` then `B.`) in one batch must share one `w:num` (as
0.8.3 did and as `list-format` convert-plain now does via `listFallbackSequenceState`), still allocated through
`services/numbering-allocation.js` so ids cannot collide. Route the text-equal/explicit header conversion paths through the
same sequence state as `list-format`. Test with the reproducer's Word-authored fixture and a Word ListString check
(Accept All shows `A.`, `B.`; inserting an item between would renumber).

### I3. Word Reject All after an end-of-document table append (R5) — investigate

Re-run the plugin's exporter (`2026-10-01-word-reject-table-append-paragraph-export.mjs`) against this tree and check in
Word with `Document.Revisions.RejectAll()`: paragraph count and `Revisions.Count`. If Word still leaves an extra paragraph or
a revision, record the shape in the WP-2 COM matrix (`scripts/generate-word-revision-semantics-matrix.mjs`), fix the
simulator so the verifier sees it, then fix the generator (likely the paragraph mark that hosts or follows the appended
table: a table must be followed by a paragraph, so the append probably needs the *existing* final paragraph's mark tracked
differently, as Word itself writes it). Add the case to the Word differential (paragraph-count check).

### I4. Make the escapes impossible to reintroduce — small

- Add both invariants to the KB failure table, `docs/TESTING.md`, CHANGELOG, and the plugin hand-off brief
  (`docs/integrations/2026-10-11-word-plugin-handoff.md`).
- Update the fidelity plan's "Open questions"/hand-off: `ACCEPT_INTENT` is no longer optional.

## Not library work

- R6 and the plugin's own `TRACKING_POLICY_MISMATCH` guard: the guard can stay as defense in depth after I1 ships.
- Moving R1/R6 to the plugin's `completed/` and updating its README is the plugin's job after it upgrades (its README says
  fixed reports move only after the reproducer exits 1 against the installed package).

## Definition of done

- R2, R3, R4 reproducers exit 1 against this tree; R5 verified in Word (fixed, or documented as a Word limitation with
  evidence); R1 documented as superseded by `list-format`.
- `UNTRACKED_CREATED_REVISION` and Accept All intent for localized replacements enforced by default.
- `npm test`, `check:types`, `test:isolation`, matrix (default + one extended seed), Word synthetic and corpus lanes green.
