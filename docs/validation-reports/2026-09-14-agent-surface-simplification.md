# Agent Surface Simplification Replay

**Date:** 2026-09-14  
**Runtime:** Node.js 24.11.1 on Windows x64  
**Engine:** Current post-0.7.1 source for every policy  
**Command:** `npm run benchmark:agent-surface`  
**Configuration:** 7 measured iterations and 1 warmup per scenario/policy

## Scope

This repository-controlled replay compares three workflow policies without
changing the underlying engine:

1. `v0.6.2-extract-full`: extraction followed by complete modified text.
2. `v0.7.1-speculation-led`: speculative literal edits with extraction/apply
   recovery, and extraction for semantic or batch work.
3. `simplified-extract-localized`: extraction followed by a compact localized
   operation against `paragraphId` plus `fingerprint`.

The six wholly synthetic scenarios cover a long literal edit, semantic drafting,
repeated visible text, rejected-view restoration, an NBSP-bearing target, and an
independent two-target batch. Every output is re-extracted and compared with its
predetermined accepted-view text. Benchmark-only verification calls are not
counted as workflow calls.

## Aggregate Results

| Policy | Correct runs | Total calls | Calls/run | Failed calls | Edit request bytes | Median local ms |
|---|---:|---:|---:|---:|---:|---:|
| v0.6.2 extract/full | 42/42 | 84 | 2.00 | 0 | 27,034 | 3.86 |
| v0.7.1 speculation-led | 42/42 | 77 | 1.83 | 7 | 8,988 | 3.86 |
| simplified extract/localized | 42/42 | 84 | 2.00 | 0 | 10,661 | 3.77 |

The simplified policy reduced generated edit-request bytes by 60.6% relative to
the complete-paragraph baseline while preserving a fixed two-call workflow and
zero failed calls. On the long-paragraph case, its median edit request was 193
bytes rather than 2,104 bytes, a 90.8% reduction.

The speculation-led policy saved seven total calls because two unique literal
scenarios completed in one call. The repeated-text scenario cost three calls and
one failure on every iteration. This demonstrates both the legitimate one-turn
benefit and the recovery cost that the simplified default is intended to avoid.

The localized form is not always smaller for a short paragraph. In the NBSP
scenario it used 344 bytes versus 343 bytes for complete modified text. The
benefit is strongest when the inspected source is long relative to the changed
span.

## Interpretation

The repository evidence supports the surface decision:

- keep speculative apply as an available expert/harness capability;
- do not make ordinary agents classify whether to use it;
- lead with deterministic extraction and one apply;
- retain localized replacements after inspection to avoid long output payloads;
- use compact `paragraphId` plus `fingerprint` targets for localized requests,
  while retaining extracted text locally for drafting; and
- keep restoration as a request/hint-driven branch.

Local execution time is effectively neutral at this scale and remains far below
agent deliberation or provider latency. No performance claim should be based on
the small local timing differences in this report.

## Limitations

This is a deterministic workflow replay, not a model evaluation. It cannot
measure reasoning tokens, provider tokenization, transport latency, stochastic
instruction following, or subjective drafting quality. Repeated external agent
runs using the same model, prompt, document, and engine are still required before
claiming end-to-end latency improvement.

The machine-readable output is written to
`tmp/benchmarks/agent-surface-simplification-latest.json` and is intentionally
not committed.
