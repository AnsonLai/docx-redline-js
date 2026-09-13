# Agent Protocol Rollout Audit

**Date:** 2026-09-12  
**Scope:** WP-06 shell/documentation work and WP-07 repository-level rollout  
**Benchmark:** `npm run benchmark:agent` on Node v24.11.1 / Windows x64,
7 measured iterations after 2 warmups

## Outcome

The compact shell path removes the operations-file tool turn, and the
development session example removes repeated target/policy payload. All measured
paths produced the requested accepted view, restored the exact source text in
the rejected view, passed package/revision validation, and preserved comment
content. These are protocol and native-runtime measurements; no LLM/provider
latency or token count was estimated.

| Measure | Legacy | WP-06/07 result |
|---|---:|---:|
| Ordinary agent instructions | 2,023 words | 349 words (82.75% reduction) |
| Repository launch card | 2,023 words | 527 words |
| Shell workflow calls | extract + file write + apply (3) | extract + stdin apply (2; 33.33% reduction) |
| Independent early-split permutations | one previously order-sensitive direction | both directions succeed |

## Measured workflow results

Median wall time includes the in-process CLI file/stdin transport for CLI rows.
It does not include process startup, a model, provider tokenization, or tool
transport.

| Case | Legacy CLI | Compact stdin CLI | Session example | Session request bytes vs legacy |
|---|---:|---:|---:|---:|
| Terminal punctuation | 49.63 ms | 46.05 ms | 49.50 ms | 125 vs 255 (50.98% fewer) |
| Term-duration phrase | 55.50 ms | 51.30 ms | 56.05 ms | 128 vs 630 (79.68% fewer) |
| Simple mutuality | 62.82 ms | 50.22 ms | 58.45 ms | 281 vs 935 (69.95% fewer) |
| Full-clause rewrite | 50.55 ms | 50.31 ms | 53.81 ms | 223 vs 407 (45.21% fewer) |
| Comment + redline | 71.94 ms | 62.11 ms | 63.81 ms | 290 vs 724 (59.94% fewer) |

Compared with the canonical stateless Node request envelope, localized session
requests remained 86.11% smaller for punctuation, 89.96% smaller for a duration
change, and 82.22% smaller for deterministic mutuality. A complete clause
rewrite is naturally less compressible and is not presented as a localized-edit
win.

## Fidelity and recovery audit

- The benchmark validates accepted and rejected paragraph text for every text
  edit and verifies the mixed-batch comment definition remains present.
- Both ten-edit permutations around an early paragraph split complete without
  target errors.
- An ambiguous target returns two candidates without auto-selection.
- A stale session handle returns `STALE_TARGET_HANDLE` with recovery action
  `reinspect`.
- The default cross-author policy refuses the edit with `EXISTING_REVISIONS`.
  `slice-cross-author` succeeds while retaining both original authors (Anson
  Lai and John Doe), and rejecting the benchmark author's work still retains
  those authors.
- The CLI protocol regression proves UTF-8 stdin transport, package-revision
  enforcement, source immutability, agent-profile defaults, atomic rollback,
  progressive exit code 3, and rejected-view restoration.

## Configuration decision

No project or auto-discovered configuration file was added. The CLI already
defaults author, strict targeting, validation, tracked changes, revision safety,
and output naming. The explicit `--profile agent` flag captures the remaining
atomic/complete-success policy and reports `effectiveOptions`. Removing that one
flag would not justify hidden configuration discovery or precedence reasoning.

## Remaining external measurement

Claude/OpenCode wall time, prompt tokens, generated tokens, and provider tool
latency must be captured by those calling harnesses. The repository benchmark
records bytes, words, calls, native time, failure behavior, and document
fidelity so those external observations can be compared without conflating the
library with model reasoning.

## Automated verification

- `npm test`: 107 passed, 0 failed.
- `npm run check:types`: 123 runtime exports have declarations.
- `npm run test:isolation`: core dependency and Word-API isolation passed.
- Focused ESLint for all WP-06/07 source, benchmark, and test files passed.
