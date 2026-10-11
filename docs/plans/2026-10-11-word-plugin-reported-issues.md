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
| R2 | `2026-10-02-separate-list-operations-restart-numbering.md` (0.8.4 regression) | **Fixed (I2)**: the text-changed header conversion (`trySingleParagraphListStructuralFallback`) allocated a `w:num` per operation, so `A.`/`B.` read right (start values) but were two lists (`numId 7`, `11`). It now goes through the batch sequence state `list-format convert-plain` uses (`continueBatchListSequence` in `services/numbering-allocation.js`): same kind+format+level and a continuing value reuses the live `numId`; a gap or another format allocates. Reproducer exits 1; Word differential (`*-list-redline-separate-*`) shows one list via `ListParagraphs.Count`. | Plugin moves it to `completed/` after upgrading |
| R3 | `2026-10-02-weak-target-soft-break-deletion.md` | **Fixed (I1.3-4)**: `compositeTextTarget` in `services/operation-batch-compiler.js` made a weak target whose `exactText` has `
` (a `w:br`) dynamic, so replacements were never compiled and an empty `modified` was applied. Replacements now always bind at batch start; the applier refuses any localized result whose accepted text differs from the compiled text (`PATCH_ROUNDTRIP_MISMATCH`). Reproducer exits 1. | Plugin moves it to `completed/` after upgrading |
| R4 | `2026-10-02-generate-redlines-false-cross-author-insertion.md` | **Fixed (I1.1-2)**: `engine/surgical-diff-application.js` `processInsert` passed a flipped `generateNestedRevision` flag; the new `UNTRACKED_CREATED_REVISION` invariant refuses any untracked output with new revision ids. Reproducer exits 1. | Plugin moves it to `completed/` after upgrading |
| R5 | `2026-10-01-word-reject-table-append-paragraph.md` | Verified in Word 16.0.20430 on this tree: the plugin exporter's package leaves 7 paragraphs and 0 revisions under `Revisions.RejectAll()` and the ribbon command; the reported 8 paragraphs + 1 revision reproduce only with `Document.RejectAllRevisions()` (Word API quirk with a tracked-inserted last mark; Word's own tracked end-table append survives both). A real adjacent defect was found and fixed: pretty-printed bodies (whitespace before `w:sectPr`) got no trailing paragraph (differential 5 vs 4). | **Done (I3)**. Plugin: use `Document.Revisions.RejectAll()` / the ribbon command in its Word oracle, not `RejectAllRevisions()`. |
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

Status: items 1-2 (R4) done (`processInsert` flag fix in `engine/surgical-diff-application.js`; `UNTRACKED_CREATED_REVISION` in the
verifier, wired in the applier; tests in `tests/revision_fidelity_verifier_tests.mjs`, untracked variants in the reject-roundtrip
matrix). Items 3-4 (R3) done: root-cause routing fix plus a pre-commit accepted-text intent check for localized replacements
(chosen over a verifier `ACCEPT_INTENT` invariant: the compiled text already exists on the operation, so no new invariant or code is needed);
tests in `tests/soft_break_localized_replacements_tests.mjs` and the `replacements-weak` matrix generator kind.

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

Status: DONE (`be539d5`, Word case `03fdca5`, counts `332d964`). Root cause: the reproducer's text-changed conversion (`1.` -> `A.`) runs the structural fallback, not the text-equal in-place core, and that path only shared numbering for explicit decimal starts; letters/roman/bullets allocated per operation (`B.` got a `startOverride`, so labels looked right). Fix: `continueBatchListSequence` (same keys as `convert-plain`, entry reused only while its numId is live in the DOM). Tests: `tests/list_format_operation_tests.mjs` section 12 (openDocx + bare, gap, other format, mixed with convert-plain), matrix kind `list-sequence-redline` (asserts one numId), Word differential cases `legal-list-redline-separate-continuing-headers` / `administrative-list-redline-separate-gap-headers` with the new `expectedAcceptedListGroupSizes` check (pre-fix Word reports groups 1|1 for the continuing case). Known limit (not R2): a text-changed conversion of a header with direct paragraph formatting (`keepNext`, spacing) is refused by the verifier (`REJECT_PARAGRAPH_PROPS`) because the fallback's deleted paragraph drops its `pPr`; use `list-format` convert-plain + a `redline` of the text after the marker (WP-6 recipe) for those.

### I3. Word Reject All after an end-of-document table append (R5) — done

**Result (2026-10-11):** Word 16.0.20430 agrees with the library on this tree for the plugin's package: `Revisions.RejectAll()` and the
ribbon give 7 paragraphs, 0 revisions, 0 tables; `Document.RejectAllRevisions()` gives 8 paragraphs + 1 revision (the documented API
quirk for a tracked-inserted last mark; Word's own COM-authored tracked end-table append, whose *final paragraph's* mark is the
inserted one, survives both). Found and fixed a real neighbour: a pretty-printed body got no trailing paragraph (differential
5 vs 4). Matrix cases `table-append-end-*` (4), simulator rule "body ends with a table => implicit untracked paragraph", regression in
`tests/table_append_fidelity_tests.mjs`, Word differential case `legal-table-append-after-final-paragraph` (65/65). Commits: matrix shapes (`test: pin Word Reject/Accept All for end-of-document table append shapes`), fix `0835f58`
(the generator hunk in `services/document-operation-mutations.js` landed inside R2's `be539d5`), then the differential case.

Original task:

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
