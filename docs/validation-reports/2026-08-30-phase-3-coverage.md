# Phase 3 behavior-focused coverage report — 2026-08-30

Phase 3 used `coverage/coverage-final.json` and `npm run coverage:gaps` to count
production functions and branches. The checked baseline is
`tests/coverage-data/phase3-baseline.json`; retained P0/P1 gaps are classified in
`tests/coverage-data/phase3-reviewed-gaps.json`.

## Outcome

- Production functions: **437/540 → 496/542**. At least 59 previously
  unexecuted functions now have behavior-driven hits, exceeding the 40-function
  milestone. V8 discovered two additional nested callback functions when their
  containing paths first executed.
- Production branches: **2333/3351 → 2919/3956**. Covered branches increased by
  586. V8 exposes additional branch sites when previously cold functions run,
  so both numerator and denominator grow; the final production ratio is 73.79%.
- Full c8 snapshot: **87.92% statements/lines, 73.90% branches, 91.80%
  functions**.
- No target file lost covered functions or branches relative to the baseline.
- No production defect was discovered, so Phase 3 required no compatibility or
  behavior change.

## Target files

| Priority | File | Functions before → after | Covered branches before → after |
|---|---|---:|---:|
| P0 | `services/numbering-helpers.js` | 0/18 → 18/18 | 1 → 109 |
| P0 | `orchestration/route-plan.js` | 0/6 → 6/6 | 1 → 37 |
| P0 | `orchestration/list-markdown.js` | 0/6 → 6/6 | 1 → 39 |
| P0 | `pipeline/patching.js` | 6/12 → 12/12 | 36 → 121 |
| P0 | `engine/format-span-application.js` | 2/5 → 5/5 | 10 → 40 |
| P1 | `orchestration/list-structural-fallback.js` | 8/17 → 17/17 | 44 → 162 |
| P1 | `engine/table-mode.js` | 2/3 → 3/3 | 6 → 13 |
| P1 | `core/table-targeting.js` | 7/11 → 11/11 | 22 → 99 |
| P1 | `services/standalone-operation-runner.js` | 41/50 → 42/52 | 255 → 284 |
| P1 | `pipeline/pipeline.js` | 7/12 → 12/12 | 28 → 57 |

Branch totals are intentionally omitted from the last column because V8 did
not expose all cold-function branch sites at baseline. The checked regression
gate compares covered counts, while the JSON report retains current totals.

## Behavior matrices added

- Numbering: missing/malformed parts, identifier collisions, preferred-range
  overflow, independent-document determinism, pair reservation, paragraph
  reference overwrite/extraction, explicit starts, payload remapping, schema
  ordering, duplicate idempotence, and malformed merge fallback.
- Routing and list markdown: every route kind, list/table precedence, literal
  escapes, empty and block inputs, ordered/unordered lists, nesting/outdenting,
  marker stripping, decimal/alpha/Roman styles, and alpha rollover past Z.
- Patching and format spans: offset zero/end, run boundaries, whitespace,
  hyperlink/structural runs, insert/delete/equal, containers, multiline list
  insertion, style selection, overlapping hints, empty spans, and final-character
  formatting with tracked/untracked output assertions.
- Structural list fallback: detection/rejection, existing numbering, explicit
  sequence reuse/reset, binding cleanup/preservation, injected generation
  failure, start overrides, and trailing paragraph cleanup.
- Tables and pipeline: paragraph-block inference, symmetric row insertion,
  ambiguity/no-mutation, nested tables, table no-ops, validation modes, web
  yielding, wrapping, indentation, and valid/invalid table generation.
- Standalone runner: highlighting, localized preservation, stop-on-error,
  atomic artifact rollback, non-atomic retention, and malformed-document errors.
  Existing tests continue to cover mixed ordering, comments, numbering artifacts,
  target snapshot invalidation, and full continue-on-error batches.

## Reviewed retained gaps

All P0 functions are covered. The ten remaining P1 function entries are nine
empty default logging callbacks and one last-resort paragraph constructor whose
failure prerequisite cannot be injected through a supported API. They contain
no untested routing or identifier behavior and are retained for diagnostics and
content-preserving defense. Exact locations and classifications are stored in
`tests/coverage-data/phase3-reviewed-gaps.json`.
