# Agent CLI Efficiency Rollout Audit

**Date:** 2026-09-13  
**Scope:** WP-03 through WP-05 profile composition, compact mutation output,
skill-authoring guidance, and package parity

## Outcome

CLI contract version 7 separates complete-success behavior from workflow
policy. `--profile agent` now retains progressive execution and the ordinary
`merge-same-author` policy; `--atomic` and explicit revision policies compose
with it and appear in `effectiveOptions`. Reviewer precedence and the visible
`AI Redliner` fallback are unchanged.

Compact CLI mutation output now contains one authoritative top-level receipt per
operation and no nested receipt copy. Successful exact-match diagnostics are
omitted; successful space-equivalent matches retain `{ mode, differenceCount }`.
Errors keep their bounded recovery evidence. The Node facade and standalone
runner retain their complete per-result receipt and target-match structures.

## Measured output

`npm run benchmark:agent-cli` uses the same one-operation sample established in
the WP-00 baseline:

| Mutation response | Bytes | Change from pre-WP-04 |
|---|---:|---:|
| Pre-WP-04 pretty JSON | 2,968 | — |
| Contract-7 pretty JSON | 2,216 | 25.34% smaller |
| Contract-7 `--compact` JSON | 1,557 | 47.54% smaller |

One-line serialization is 29.74% smaller than the structurally deduplicated
pretty response, so contract version 7 exposes it explicitly through `--compact`
and `compact-cli-json-v1`. It changes whitespace only.

## Representative workflow audit

The checked repository run records:

- one relevant command-help lookup;
- zero source, distribution-bundle, or ZIP-part inspection calls;
- one contextual clause search and one apply attempt;
- progressive agent-profile execution with complete-success exits;
- exact accepted text and exact source restoration after rejection;
- unchanged comment count and byte-identical source input; and
- one top-level receipt with no nested copy.

The existing agent workflow benchmark separately covers mixed comment/redline
work, independent batch permutations, stale and ambiguous targeting recovery,
and cross-author attribution. File-based and serializer-backed stdin operations
are both exercised with the same Unicode payload by the CLI protocol suite.

## Documentation and package parity

The shipped documentation now includes `docs/SKILL_AUTHORING.md`. It separates
safety invariants, workflow policy, transport choices, four-field recovery, and
presentation rules. Executable documentation tests require each ordinary-use
section to lead with focused contextual extraction and require skills to cite
`humanReference`, provisions, or headings rather than machine paragraph
ordinals.

The npm allowlist includes the launch card, fast start, knowledge base,
skill-authoring contract, schema, testing guide, and checked rollout reports.
Development-only session examples and benchmark scripts remain excluded.

The checked `npm pack --dry-run` contains 152 files (approximately 1.07 MB).
It includes `node/cli-help.js`, the fast start, skill-authoring contract, and this
audit; it excludes both agent benchmark scripts and the development session
example.
An unpacked tarball invocation reported contract version 7 and the same 16
capabilities as the source CLI.

## Automated verification

- `npm test`: 110 passed, 0 failed.
- `npm run check:types`: 123 runtime exports have declarations.
- `npm run test:isolation`: dependency and Word-API isolation passed.
- Focused ESLint for every WP-03 through WP-05 source, test, and benchmark file
  passed.
- `npm run build` and `npm pack --dry-run`: passed.

## Measurement boundary

These are native runtime, serialized-byte, command-count, and correctness
measurements. They do not estimate Claude/OpenCode tokens, model reasoning time,
provider transport latency, or end-to-end harness wall time.
