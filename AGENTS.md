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
| Change complete-document operations | `document/docx-document.js` → `services/standalone-operation-runner.js` → `services/document-operation-*.js` |
| Change DOCX ZIP or CLI behavior | `document/zip-archive.js`, `node/index.js`, `node/docx-document.js`, `node/cli.js` |
| Choose or add tests | Closest `tests/*.mjs`, then [Testing Guide](docs/TESTING.md) |
| Understand ownership/dependencies | [Architecture](ARCHITECTURE.md) |

Open the full [Agent Knowledge Base](docs/AGENT_KNOWLEDGE_BASE.md) only for the
specific advanced operation, API, or recovery topic you need.

## Ordinary document edits

Use one focused extraction and one apply call. With a structured wrapper, use
its revision-bound target handles. For shell-only work, use serializer-backed
stdin and the explicit agent profile:

```bash
docx-redline extract contract.docx --search "termination" --around 3
node emit-operations.mjs | docx-redline apply contract.docx --operations - --profile agent --compact --output reviewed.docx
```

Retain inspected `exactText` for drafting and target with its compact
`paragraphId` plus `fingerprint`. Use complete `modified` text for a broad
semantic revision; for a small literal change, use `replacements` so the apply
payload need not repeat the whole paragraph. Independent strong targets bind to the batch start, so do
not manually sort around structural edits. Consolidate incompatible writes to
one source; use captures only for intentional created-content dependencies.

If the user requests restoration, search `--view rejected` directly. Otherwise
follow `selection.hint` if extraction identifies an alternate revision view.
Restore only the inspected strong target; do not infer adjacent content.

Require `completion: true` and a non-null output path. Follow
`error.recovery.action` and `retryPlan`; never retry unchanged failed arguments.
Do not accept/reject another reviewer's work or remove comments without explicit
user authorization. The source is never overwritten unless `--in-place` is
explicit.

The agent profile preserves progressive execution and the ordinary revision
policy. Add `--atomic` or `--existing-revisions slice-cross-author` only when
that policy is intended, and confirm the resolved `effectiveOptions`.

Speculative apply, directional search ranges, advanced restore, rejected-view
insertion, list, table, formatting, comments, revision policies, and failure
examples live in the
[knowledge base](docs/AGENT_KNOWLEDGE_BASE.md#agent-document-workflow-cli) and
[operation schema](docs/schemas/document-operations.schema.json).

## Thin wrappers

Wrap the CLI for shell hosts or `openDocx` from `@ansonlai/docx-redline-js`
(or `@ansonlai/docx-redline-js/bundle` for zero-dependency sandboxes, or
`@ansonlai/docx-redline-js/node` for backward-compatible Node hosts) for
byte-oriented hosts. A wrapper should inspect, translate narrow ergonomic inputs
into canonical operations, call the facade once, and return its structured result.
Do not reproduce ZIP handling, targeting, revision allocation, comments, numbering,
validation, or rollback.

The stateful wrapper in `examples/agent-session-wrapper.mjs` is a testable
development demonstration only. It is excluded from package files and exports.
Production harnesses own their transport and negotiate the minimum CLI
`contractVersion` and capabilities they actually use. Optional capabilities do
not belong in an ordinary editing prompt merely because the CLI supports them.

## Code map

```text
index.js       host-independent public API & universal document facade
adapters/      DOM/XML parser provider with automatic pure-JS fallback
core/          OOXML primitives, text views, targeting, SHA-256, validation
document/      universal Uint8Array DOCX facade & fflate ZIP container
pipeline/      ingestion, diffing, Markdown, lists, serialization
engine/        reconciliation modes and run-level mutation
orchestration/ route planning and structural conversion
services/      document operations, comments, receipts, artifacts
node/          Node CLI and backward-compatibility re-exports
dist/          universal ESM and zero-dependency sandbox bundles
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
