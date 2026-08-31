# Reliability and Testing Improvement Plan — Round 3

**Status:** In progress — Phases 1, 2, and 4 complete

This plan follows the completed
`completed/2026-08-02-reliability-improvements.md` plan. It selects four of the
next-step opportunities identified while completing that work:

1. Preserve and target structural tabs and complex fields correctly.
2. Extend the synthetic DOCX packager for related Word parts.
3. Add behavior-focused tests in the least-covered production paths.
4. Turn the Word catalogue into an explicit task/structure coverage matrix
   (selected recommendation 6 from the review).

The published library remains host-independent JavaScript. Microsoft Word COM,
ZIP assembly, downloaded corpus documents, and other platform-specific tooling
remain development-only.

## Baseline

As of 2026-08-30:

- `npm test`: 30/30 test files pass.
- Synthetic Microsoft Word differential: 25/25 cases pass.
- Reviewed SuperDoc Microsoft Word differential: 20/20 documents pass.
- Coverage: 79.96% lines/statements, 80.95% functions, 69.52% branches.
- The synthetic packager supports `word/document.xml` and optional numbering,
  but not the complete related-part graph needed for comments, notes,
  headers/footers, or relationship-backed hyperlinks.
- A paragraph containing `w:tab` cannot currently be targeted consistently
  through reconstruction.
- Reconstruction adjacent to a complex field can discard `w:fldChar` and
  `w:instrText` scaffolding.

Coverage percentages are diagnostic baselines, not correctness scores or
release thresholds. Each phase below is accepted through observable behavior.

## Production API compatibility

| Phase | Expected impact |
|---|---|
| 1 | **Conditional production behavior.** Correct output changes around tabs and fields; public signatures should remain unchanged. |
| 2 | **Not breaking.** Development-only fixture packaging and Word automation. |
| 3 | **Not breaking unless a test exposes a production defect.** Test additions alone are development-only; any resulting runtime fix must document its own compatibility impact before implementation. |
| 4 | **Not breaking.** Test metadata, reporting, and contributor workflow only. |

If implementation reveals that an established public text representation for
tabs or fields must change, stop and update this compatibility section and the
changelog before landing that behavior.

---

## Phase 1 — Structural tabs and complex fields

**Status:** Complete (2026-08-30)

### Problem

Word structures such as tabs and fields occupy positions in visible document
content but are not ordinary `w:t` text:

- Paragraph targeting currently omits `w:tab`, while ingestion/reconstruction
  expects a complete contiguous text range. A target containing the visible tab
  therefore fails before the Word oracle can run.
- A complex field consists of `w:fldChar` begin/separate/end nodes,
  `w:instrText`, and visible result runs. Reconstruction adjacent to its result
  can retain the visible text while losing the field itself.

Both failures are dangerous in legal and administrative documents: the visible
text may appear plausible even though alignment or an automatically generated
reference has been destroyed.

### Work

1. Define one canonical visible-text representation for `w:tab` across
   ingestion, targeting, offsets, exact-text assertions, and reconstruction.
   Prefer `\t`; do not collapse it to a space.
2. Represent complex field scaffolding as inert structural spans. Field
   instructions must not become editable visible text, and edits adjacent to a
   field result must not split, duplicate, reorder, or discard the field nodes.
3. Preserve the distinction among a field instruction, its displayed result,
   and surrounding editable text. Do not update or evaluate fields in the JS
   engine.
4. Add focused automated tests for tabs at the beginning, middle, and end of a
   run; edits on either side of a field; multiple fields; and field results
   split across runs.
5. Add synthetic Word cases for at least one tab-bearing administrative layout
   and one locked complex field. Require the original structural elements and
   verify exact Accept All and Reject All behavior in Word.
6. Add both shapes to deterministic fuzz/invariant testing once their text
   representation is stable.

### Acceptance

- Tab-bearing paragraph targets resolve without normalization or whitespace
  loss, and accept/reject round trips preserve exact tabs.
- Adjacent edits preserve `w:fldChar` begin/separate/end ordering and
  `w:instrText` byte content.
- `validateRedlineOoxml` reports no structural errors.
- The new synthetic packages open in Word without repair; Word sees tracked
  revisions; Accept All matches intent; Reject All restores the source.
- Existing 25 synthetic and 20 corpus Word cases remain green.

### Completion record

- Paragraph targeting and reconstruction now use literal `\t` for `w:tab`,
  including leading and trailing tabs; the standalone runner no longer trims
  those target boundaries.
- Complex-field begin/instruction/separate/end nodes are retained as inert,
  zero-width reconstruction sentinels. Their cached display runs remain inside
  the field and are not revised when only adjacent text changes.
- Added `tests/structural_tab_field_tests.mjs` plus deterministic tab/field fuzz
  shapes. Focused assertions cover field ordering, exact instruction bytes,
  valid run parents, split cached results, and exact Accept/Reject text.
- Added three synthetic Word cases: tab-aligned administrative text, boundary
  tabs, and a locked PAGE field. The catalogue is now 28 cases; Word 16.0 build
  16.0.20326 passed 28/28, and the reviewed SuperDoc lane passed 20/20.
- AI visual preflight inspected Word-produced final and markup renderings for
  all three new cases. It initially exposed a duplicated visible PAGE result in
  markup; the sentinel mapping was corrected and the repeat preflight passed.
  See `docs/validation-reports/2026-08-30-phase-1-word-visual-preflight.md`.
- Final verification: 31/31 JavaScript test files, isolation, declarations,
  lint, build, 100 deterministic fuzz cases, and 80.01% statement/line,
  80.95% function, and 69.71% branch coverage all pass.

---

## Phase 2 — Richer synthetic DOCX package fixtures

**Status:** Complete (2026-08-30)

### Problem

The current synthetic fixture builder creates a deliberately minimal package.
That is sufficient for body paragraphs, numbering, bookmarks, internal links,
content controls, and tables, but it cannot independently test structures that
require additional package parts and relationships.

### Work

1. Extend the script-local package builder with explicit, optional support for:

   - `word/comments.xml` and comment anchors/references;
   - `word/footnotes.xml` and `word/endnotes.xml`;
   - header and footer parts, including section references;
   - document relationships for external hyperlinks;
   - the required `[Content_Types].xml` overrides and relationship entries.

2. Keep every addition opt-in so the smallest fixtures remain easy to inspect.
   Do not add a ZIP dependency to runtime library code.
3. Validate relationship targets, content-type declarations, referenced IDs,
   and required separator entries for footnotes/endnotes before writing a
   package.
4. Add package-level assertions that distinguish edited parts from untouched
   parts and hash untouched entries where appropriate.
5. Add at least one English legal or administrative Word differential case for
   each supported related-part family. Expected results must come from edit
   intent or Word's original source text, never from this library's own
   accept/reject helpers.
6. Document the fixture schema and small reusable constructors in
   `docs/TESTING.md`.

### Acceptance

- Synthetic fixtures containing comments, notes, headers/footers, and external
  hyperlinks open in desktop Word without repair.
- Word Accept All and Reject All pass for every new case.
- Required parts, content types, and relationships are validated before Word
  runs.
- Untouched package parts remain byte-identical.
- Runtime dependency and isolation checks prove that packaging/COM code has not
  entered `index.js`, `core/`, `engine/`, `pipeline/`, or runtime services.

### Completion record

- `scripts/lib/minimal-zip.mjs` now accepts opt-in comments, footnotes,
  endnotes, header/footer parts, and external hyperlink relationships in
  addition to numbering. It generates relationships and content-type overrides
  without adding a runtime ZIP dependency.
- Pre-emission validation covers well-formed roots, unique relationship IDs and
  part names, external targets, document relationship references, comment/note
  IDs, and required note separator IDs `-1` and `0`.
- Added reusable escaped constructors in
  `tests/fixtures/word-package-parts.mjs` and comprehensive positive/negative
  package tests in `tests/minimal_docx_package_tests.mjs`.
- Related parts are byte-compared after packaging, recorded by SHA-256 in each
  fixture sidecar, and rechecked by the Word differential before Word opens the
  package.
- Added five English legal/administrative cases. Word 16.0 build 16.0.20326
  passed 33/33 synthetic cases; the reviewed SuperDoc lane remains 20/20.
- AI visual preflight inspected final and markup renderings for all five new
  cases. It prompted self-contained superscript formatting for note references;
  the repeated preflight passed. See
  `docs/validation-reports/2026-08-30-phase-2-word-visual-preflight.md`.
- Final verification: 32/32 JavaScript test files, isolation, declarations,
  lint, build, and 80.35% statement/line, 81.29% function, and 69.60% branch
  coverage all pass.

---

## Phase 3 — Behavior-focused coverage in thin production paths

**Status:** Planned

### Targets

The current report covers 442 of 546 functions, leaving 104 functions
unexecuted. Module loading can execute declarations and top-level statements
without invoking a function, so line coverage alone hides the sharpest gaps:

| Priority | Area | Current signal | Why it is a problem area |
|---|---|---:|---|
| P0 | `services/numbering-helpers.js` | 23.55% lines, **0/18 functions** | Allocates and merges package identifiers; a collision can corrupt lists or relationships. |
| P0 | `orchestration/route-plan.js` | 23.75% lines, **0/6 functions** | Chooses the editing path; a wrong route can bypass the safeguards of the intended mode. |
| P0 | `orchestration/list-markdown.js` | 24.11% lines, **0/6 functions** | Converts user intent into list structure and is currently visible only through module loading. |
| P0 | `pipeline/patching.js` | 52.08% lines, 50% functions | Directly mutates reconstructed content; missed boundaries can lose or duplicate text. |
| P0 | `engine/format-span-application.js` | 62.56% lines, 40% functions | Applies formatting over offsets; split-run and boundary mistakes are visually subtle. |
| P1 | `orchestration/list-structural-fallback.js` | 57.19% lines, 47.05% functions, 55% branches | Large fallback path with list-level, adjacency, and failure decisions. |
| P1 | `engine/table-mode.js` | 75.93% lines, **26.08% branches** | Most table decisions are not exercised despite acceptable line coverage. |
| P1 | `core/table-targeting.js` | 74.67% lines, 45.83% branches | Ambiguous cells and partial matches can target the wrong administrative data. |
| P1 | `services/standalone-operation-runner.js` | 74.05% lines, 59.57% branches | The largest module coordinates targeting, artifacts, ordering, and rollback. |
| P1 | `pipeline/pipeline.js` | 79.94% lines, 58.33% functions, 56% branches | Central orchestration still has untested modes and error paths. |
| P2 | adapters and defensive package validation | mixed | Lower-risk fallbacks and host-specific diagnostics; cover after reachable content-changing paths. |

The goal is not to inflate a global percentage. It is to exercise meaningful
decisions, error contracts, and collision/rollback behavior that are currently
easy to change unnoticed.

### Coverage method

1. Add a detailed JSON coverage artifact and a small reporting script, exposed
   as `npm run coverage:gaps`, that lists every uncovered production function
   with file, name, and source line. Keep the existing summary for trend
   history.
2. Classify each uncovered function before writing a test:

   - **reachable behavior** — cover through a public or supported deep-import
     entry point;
   - **pure helper with meaningful boundaries** — direct unit tests are
     appropriate;
   - **defensive/environment-specific** — exercise with a realistic injected
     failure where possible;
   - **unreachable or obsolete** — delete it, or document why it remains.

3. Rank gaps using four questions: can failure lose/corrupt content, can it edit
   the wrong target, is it on a public/common path, and has the code recently
   changed? Content loss/corruption and wrong-target risks are always P0.
4. For each selected function, use a behavior set rather than a single smoke
   call: normal case, boundary case, malformed/unsupported input, no-op, and
   collision or rollback case where applicable.
5. Prefer tests that enter through the public API and naturally traverse
   several internal functions. Use direct helper tests to cover algorithms and
   boundary values, not to manufacture execution without observable behavior.
6. Pair function coverage with branch decision tables. A function is not
   considered adequately tested merely because its first line executed.
7. Do not add `c8 ignore` annotations to improve the score. An exclusion needs
   a written environment or reachability reason and review alongside the code.
8. Track both the number of newly executed functions and the behaviors added.
   A working milestone is at least 40 additional previously-unexecuted
   production functions, moving global function coverage toward 88% and branch
   coverage toward 75%. These are navigation targets, not sufficient acceptance
   criteria by themselves.

### Work

1. Generate the uncovered-function inventory and attach the P0/P1/P2 ranking.
   Record which branches are public behavior, defensive impossibilities, or
   obsolete code. Update the inventory after each test group so effort moves to
   the next actual gap.
2. Add numbering tests for missing parts, existing abstract/instance ID
   collisions, malformed numbering XML, multiple list styles, and deterministic
   allocation across independent documents. Exercise every reachable function
   in `numbering-helpers.js`, including relationship/content-type merging and
   idempotent repeated calls.
3. Add route-plan tests covering every supported operation route, ambiguous
   inputs, table/list precedence, unsupported/native fallbacks, and stable error
   results. Express the routes as a decision table and require one positive and
   one rejection/fallback case per route.
4. Add list-markdown and structural-fallback tests for ordered/unordered lists,
   nesting and level changes, adjacent insertion/deletion, formatting, existing
   numbering, and failure without partial mutation. Ensure all public and
   reachable internal functions in `list-markdown.js` execute through real list
   scenarios rather than import-only tests.
5. Add standalone runner tests for mixed-operation ordering, package artifacts,
   `continueOnError`, atomic/non-atomic results, and target snapshot invalidation.
6. Add patching/format-span boundary tests at offset zero, run boundaries,
   whitespace, tabs/breaks, overlapping formatting, empty spans, and the final
   character. Assert exact XML plus accept/reject text.
7. Add table decision tests for duplicate cell text, merged cells, nested
   tables, row/column additions, partial table targets, format-only edits, and
   clean failure without modifying untargeted cells.
8. Add pipeline tests for each routing mode, parse and diff failures, no-op
   behavior, existing revisions, formatting removal, and generated package
   artifacts.
9. Delete unreachable branches when evidence shows they are obsolete rather
   than writing artificial tests solely to execute them.
10. Record before/after function and branch counts for every targeted file, plus
    the global snapshot, after the behavioral cases land.

### Acceptance

- Every P0 module has a documented function inventory and focused regression
  tests for all reachable content-changing functions.
- No reachable function in `numbering-helpers.js`, `route-plan.js`, or
  `list-markdown.js` remains at zero hits; any retained uncovered function has a
  reviewed reachability/environment explanation.
- At least 40 previously-unexecuted production functions are exercised, unless
  the inventory proves that fewer than 40 are reachable—in which case obsolete
  functions are removed and the remainder are explicitly justified.
- P0 and P1 branch decisions have named behavior cases even where coverage-tool
  instrumentation groups multiple decisions onto one line.
- New tests assert outputs and error contracts, not merely that functions ran.
- The report fails if a targeted file loses covered functions relative to the
  new Phase 3 baseline. No global percentage is used as the sole gate; per-file
  changes are reviewed alongside the added behavior matrix.
- `npm test`, isolation, types, lint, and deterministic fuzz checks pass.
- Any production defect discovered by the new tests receives its own changelog
  and compatibility assessment before being fixed.

---

## Phase 4 — Explicit task/structure coverage matrix

**Status:** Complete (2026-08-30)

### Problem

The Word catalogue currently has useful `category` and `task` labels, while the
SuperDoc scenarios record `shape` and free-form coverage labels. These show what
individual cases do, but they do not reveal which combinations are covered,
planned, intentionally excluded, or impossible in the current harness.

### Work

1. Define a shared vocabulary for:

   - task: replace, insert, delete, format, comment, accept/reject, list change,
     table reconciliation, and mixed batch;
   - structure: plain paragraph, multi-paragraph, formatted runs, list, table,
     bookmark, hyperlink, content control, tab/break, field, comment, note,
     header/footer, section boundary, and prior revisions;
   - oracle: JS exact round trip, runtime validator, XSD, LibreOffice, synthetic
     Word, real-document Word, AI Word visual preflight, and human Word visual
     review.

2. Normalize synthetic and corpus metadata to this vocabulary without removing
   human-readable review notes.
3. Add a deterministic report command that prints the task-by-structure matrix
   and identifies uncovered combinations. Store machine-readable exclusions
   with reasons such as “requires related-part packager” rather than silently
   treating absence as coverage.
4. Make catalogue validation reject unknown labels, missing oracle metadata,
   duplicate scenario identities, and claims unsupported by the actual fixture.
5. Use the report to choose new tests based on risk and structural diversity,
   not simply to increase the case count.
6. Add a planned `npm run review:word:prepare` helper that selects all new or
   changed Word cases plus a rotating release sample and writes a review
   manifest under ignored `tmp/` storage. The helper may prepare disposable
   accepted/rejected copies, but it must never mark a visual check as passed.
7. Add periodic AI visual preflights using the local Word UI and screenshots:
   after layout-sensitive phases, after every five to ten new Word cases, for
   every new structure family, and before the human release sample. Record the
   exact cases/views, Word build, screenshot evidence, observed failures, and
   uncertainty. AI review remains advisory and must not open sensitive/private
   documents without explicit authorization.
8. Require human review for every new structure/operation, every affected case
   after layout-sensitive engine or package changes, all cases using normalized
   text comparison, and any unexplained automated Word failure.
9. Before each release, visually inspect all new/changed cases, at least 20% of
   unchanged synthetic cases on a rotating basis, and at least one legal and one
   administrative SuperDoc result. Record Word version/build and the All Markup,
   Accept All, and Reject All result for each selection.
10. Document the process in `docs/TESTING.md`, use
   `docs/WORD-MANUAL-REVIEW.md` as the checklist/template, and retain the signed
   report with release-validation artifacts.

### Acceptance

- One command produces a stable, reviewable matrix for synthetic and SuperDoc
  cases.
- Every case declares task, structure, and oracle metadata from validated
  vocabularies.
- Every empty high-priority matrix cell is either backed by a planned test or an
  explicit exclusion with a reason and dependency.
- Adding a new case updates the matrix automatically.
- The test guide explains how contributors use the report to select the next
  case.
- The matrix distinguishes automated Word semantics from human Word visual
  review, records AI visual preflights separately, and identifies cases whose
  manual review is missing or stale.
- Each layout-sensitive phase and each group of five to ten new Word fixtures
  receives an AI visual preflight unless no authorized local documents are
  available; skipped preflights have a recorded reason.
- Release validation includes a recorded human sign-off with the reviewer,
  Word build, selected cases, and results for all three Word views.

### Completion record

- Added a shared, validated vocabulary covering nine tasks, fifteen structures,
  eight independent oracles, and explicit missing/current/stale human-review
  state. All 33 synthetic and 20 reviewed SuperDoc cases now resolve to declared
  canonical metadata without removing their descriptive labels.
- Added `npm run report:word:coverage`, which deterministically reports the live
  53-case task/structure matrix, individual case identities in JSON mode,
  missing/stale human review, and every uncovered high-priority cell.
- Added machine-readable high-priority cells and eight explicit planned gaps,
  each with a reason and dependency. Catalogue validation rejects unknown
  labels, missing oracle/review metadata, duplicate identities, stale gap
  dispositions, and structure/task claims unsupported by the fixture.
- Added `npm run review:word:prepare -- --cycle=N`. It conservatively selects
  changed catalogue families, a rotating 20% synthetic sample, and one legal
  plus one administrative SuperDoc case. Every view and human-sign-off field is
  emitted as `pending`; the helper cannot certify a review.
- The matrix records eight synthetic AI visual preflights from Phases 1 and 2
  separately from automated Word semantics. No human sign-off is fabricated:
  the current report correctly lists all 53 cases as missing human review, and
  the signed three-view review remains an operational release gate.
- Added focused matrix tests and contributor/reviewer instructions in
  `docs/TESTING.md` and `docs/WORD-MANUAL-REVIEW.md`.
- Validation passed with 33/33 JavaScript suites, 33/33 synthetic Word cases,
  20/20 real-document Word cases, isolation, declarations, lint, and build.
  Coverage after the Phase 4 tooling tests is 80.60% statements/lines, 70.13%
  branches, and 82.02% functions.

---

## Execution order

```text
Phase 1 (tabs and fields) ───────────────┐
                                         ├─> Phase 4 (coverage matrix)
Phase 2 (related-part package fixtures) ─┘

Phase 3 (thin-path behavioral tests) can proceed independently, but any newly
discovered runtime defect must be scoped and documented before implementation.
```

Phase 1 supplies two important structural shapes. Phase 2 unlocks the package
families currently missing from synthetic Word validation. Phase 4 should use
those stable capabilities rather than encoding temporary harness limitations as
the final matrix. Phase 3 can run alongside them because it mostly targets
lower-level routing and numbering behavior.

## Verification commands

```bash
npm test
npm run test:isolation
npm run check:types
npm run lint
npm run test:coverage
npm run coverage:gaps              # planned in Phase 3
npm run test:word                 # Windows + desktop Microsoft Word
npm run test:corpus:word          # pinned local corpus + Word
npm run review:word:prepare -- --cycle=0  # pending human-review set
node scripts/export-validation-fixtures.mjs
FUZZ_SEED=1 FUZZ_ITERATIONS=12000 node tests/roundtrip_fuzz_tests.mjs
```

Release validation continues to include the ECMA-376 transitional XSD and
LibreOffice lanes described in `docs/VALIDATION.md`.
