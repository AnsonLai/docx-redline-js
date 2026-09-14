# AGENTS.md — Launch Card

Use this file to route work on `@ansonlai/docx-redline-js`. Do not explore the
whole repository before acting, and never inspect `dist/`, a vendored CLI bundle,
or an installed plugin bundle to infer public behavior.

## Pick the route

| Task | Start here |
|---|---|
| Edit or review a complete `.docx` | [Agent Fast Start](docs/AGENT_FAST_START.md) and the `docx-redline` CLI |
| Build or update an agent skill/tool wrapper | [Skill Authoring Contract](docs/SKILL_AUTHORING.md), then [README wrapper example](README.md#example-agent-session-wrapper-development-only) |
| Change paragraph/range reconciliation | `index.js` → `engine/oxml-engine.js` → selected `engine/*-mode.js` |
| Change complete-document operations | `services/standalone-operation-runner.js` → `services/document-operation-*.js` |
| Change DOCX ZIP or CLI behavior | `node/index.js`, `node/docx-document.js`, `node/cli.js` |
| Choose or add tests | Closest `tests/*.mjs`, then [Testing Guide](docs/TESTING.md) |
| Understand ownership/dependencies | [Architecture](ARCHITECTURE.md) |

Open the full [Agent Knowledge Base](docs/AGENT_KNOWLEDGE_BASE.md) only for the
specific advanced operation, API, or recovery topic you need.

## Ordinary document edits

For a literal mechanical edit with known old and new text, try one fail-closed
apply call. When global matching is too broad, use a longer, fairly unique
nearby phrase plus a directional range:

```bash
docx-redline apply contract.docx --search "distinctive nearby heading or phrase" --context-range 1:3 --find "thirty (30) days" --replace "sixty (60) days" --profile agent --output reviewed.docx
```

Generic or repeated anchors widen the unioned scope and may return
`AMBIGUOUS_TARGET`, costing a recovery turn. On success, treat
`results[i].change.context.anchorMatchCount > 1` as a signal to confirm the
returned location and excerpts.

For semantic drafting, inspect once and apply once. With a structured wrapper,
use revision-bound handles. For shell-only work, use serializer-backed stdin:

```bash
docx-redline extract contract.docx --search "termination" --around 3
node emit-operations.mjs | docx-redline apply contract.docx --operations - --profile agent --compact --output reviewed.docx
```

For every ordinary text operation, `modified` is the complete desired
accepted-view content. Copy inspected `exactText` verbatim and include
`paragraphId` or `fingerprint`. Independent strong targets are bound against the
batch-start document, so do not manually sort around structural edits.
Consolidate incompatible writes to the same source; use captures for intentional
created-content dependencies.

Require `completion: true`, `written: true`, a non-null output path, and no
per-operation errors. Localized edits also require
`results[i].change.committed: true`, `finalDisposition: "applied"`, and
`verification.acceptedViewMatchesCompiledText: true`. Follow
`error.recovery.action` and `retryPlan`; never retry unchanged failed arguments.
Follow `selection.hint` for restorable text. Do not accept/reject another
reviewer's work or remove comments without explicit user authorization. The
source is never overwritten unless `--in-place` is explicit.

The agent profile preserves progressive execution and the ordinary revision
policy. Add `--atomic` or `--existing-revisions slice-cross-author` only when
that policy is intended, and confirm the resolved `effectiveOptions`.

Advanced restore, rejected-view insertion, list, table, formatting, comments,
revision policies, and failure examples live in the
[knowledge base](docs/AGENT_KNOWLEDGE_BASE.md#agent-document-workflow-cli) and
[operation schema](docs/schemas/document-operations.schema.json).

## Thin wrappers

Wrap the CLI for shell hosts or `openDocx` from
`@ansonlai/docx-redline-js/node` for byte-oriented Node hosts. A wrapper should
inspect, translate narrow ergonomic inputs into canonical operations, call the
facade once, and return its structured result. Do not reproduce ZIP handling,
targeting, revision allocation, comments, numbering, validation, or rollback.

The stateful wrapper in `examples/agent-session-wrapper.mjs` is a testable
development demonstration only. It is excluded from package files and exports.
Production harnesses own their transport and negotiate the minimum CLI
`contractVersion`/capabilities they use. The 0.7.0 localized fast path requires
contract 8 and its localized replacement, speculative apply, and change-summary
capabilities.

## Code map

```text
index.js       host-independent public API
core/          OOXML primitives, text views, targeting, validation
pipeline/      ingestion, diffing, Markdown, lists, serialization
engine/        reconciliation modes and run-level mutation
orchestration/ route planning and structural conversion
services/      document operations, comments, receipts, artifacts
node/          Node-only ZIP and whole-DOCX facade/CLI
tests/*.mjs    directly runnable suites
```

Keep shared modules from importing `index.js`; keep host-independent code from
importing `node/`. Use `rg` to follow only the symbol being changed. Preserve
unrelated worktree changes and edit source files, never generated `dist/` files.

## Verification

Run the closest test first: `node tests/<focused-suite>.mjs`. Use `npm test` for
cross-subsystem or release handoff, plus `npm run check:types` and
`npm run test:isolation` when boundaries or declarations change. Word COM,
visual, corpus, coverage, and fixture-generation lanes are separate; select them
from the [Testing Guide](docs/TESTING.md).
