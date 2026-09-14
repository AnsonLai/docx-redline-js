# Localized Patching and Turn-Reduction Rollout Audit

**Date:** 2026-09-14
**Scope:** WP-05 correctness matrix and controlled workflow benchmark for CLI
contract 8

## Outcome

The contract-8 literal fast path removes one of two tool turns and reduces the
generated mutation request by roughly two-thirds to four-fifths in the measured
cases. Native DOCX execution remains essentially the same speed: median changes
ranged from 0.58% slower to 5.41% faster. This is the intended result—the
feature attacks agent turn and output-generation latency without replacing the
OOXML reconciliation or validation pipeline.

All successful candidates produced the expected accepted-view text, positive
localized change verification, and a byte-identical source document. Refused
speculation remained non-writing in the automated suite.

## Controlled comparison

Run:

```bash
npm run benchmark:localized
```

The benchmark uses `tests/fixtures/sample_doc_test.docx`, 3 warmups, and 11
measured iterations per workflow. Both paths run on the same current source
tree to isolate workflow shape from runtime or machine drift:

- **0.6.2-style baseline:** focused `extract`, followed by a canonical
  full-paragraph operation.
- **Contract-8 candidate:** targetless localized `apply`, either globally or
  within a directional heading scope.

| Case | Tool turns | Request bytes | Token proxy | Median local ms | p95 local ms |
|---|---:|---:|---:|---:|---:|
| Short punctuation, baseline → candidate | 2 → 1 | 244 → 79 | 61 → 20 | 55.34 → 53.37 | 83.89 → 62.11 |
| Long provision, baseline → candidate | 2 → 1 | 352 → 123 | 88 → 31 | 58.02 → 54.88 | 70.04 → 59.44 |
| Directional heading, baseline → candidate | 2 → 1 | 626 → 129 | 157 → 33 | 56.46 → 56.79 | 77.35 → 63.92 |

Observed reductions:

- Tool turns: **50%** in every case.
- Generated request bytes: **65.06%–79.39%**.
- Deterministic token proxy: **64.77%–78.98%**.
- Median local engine/I/O time: **-0.58%–5.41%** (effectively neutral).
- p95 local engine/I/O time: **15.13%–25.96% lower** in this run; timing is
  observational and not a correctness gate.

The directional candidate returned 20 more response bytes than its two-turn
baseline because its one response includes explicit location and before/after
verification. That small cost is intentional: it removes the verification turn
while making the applied change auditable.

## Measurement boundary

`generatedTokenProxy` is `ceil(UTF-8 request bytes / 4)`. It is a deterministic
comparison aid, not a provider tokenizer or billed-token count. Local timing
uses `performance.now()` and includes DOCX reads, inspection, mutation,
validation, and output writing.

The benchmark does **not** measure model reasoning, token streaming, tool
transport, network latency, or end-to-end Claude/OpenCode wall time. Those
components require external harness transcript instrumentation and must not be
inferred from native timing. The JSON artifact is written to
`tmp/benchmarks/localized-turn-reduction-latest.json`.

## Correctness coverage

The expanded suites cover:

- simultaneous localized replacements, ambiguity, deletion, no-op and overlap
  rejection, unsupported v1 shapes, summary ordering, excerpt bounds, and
  atomic rollback disposition;
- global speculative success; `0:0`, `0:3`, `1:3`, `-3:-1`, and `-3:3`
  directional windows; boundary clamping; repeated-anchor window deduplication;
  missing anchors and patch sources; multi-paragraph and same-paragraph
  ambiguity; occurrence scoping; stale IDs; protected outputs; response bounds;
  and agent-profile exit behavior.

## Verification

- `npm test`: 113 suites passed, 0 failed with the expanded assertions.
- `npm run check:types`: 123 runtime exports have declarations.
- `npm run test:isolation`: dependency and Word-API isolation passed.
- Focused ESLint for the localized compiler, CLI, benchmark, and tests passed.
- `npm pack --dry-run --json`: 154 entries; the compiler and this report are
  included, while the fixture-dependent benchmark script is excluded.
