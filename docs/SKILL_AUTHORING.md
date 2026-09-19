# Skill and Harness Authoring Contract

Use this page to create or update an agent skill, MCP server, or custom harness
for `@ansonlai/docx-redline-js`. Ordinary editing agents should load
[Agent Fast Start](AGENT_FAST_START.md), not this design contract.

> [!WARNING]
> Before 1.0, every integration must identify and pin its exact tested library
> release and required CLI contract/capabilities. Treat any different release or
> contract as unverified until compatibility tests pass.

Example declaration for a contract-8 shell integration that uses localized
operations over stdin:

```json
{
  "testedLibraryRelease": "0.7.1",
  "requiredCliContractVersion": 8,
  "requiredCapabilities": [
    "operations-stdin",
    "agent-safety-profile-v2",
    "recovery-envelope-v1",
    "localized-replacements-v1"
  ]
}
```

## Runtime negotiation

Run `docx-redline version` before the first document operation. Require CLI
contract version 8 and only the capabilities the integration actually uses. Check once per stable
runtime, fail closed when a requirement is absent, and never inspect bundles or
ZIP parts as a fallback.

Do not expose negotiated but unused capabilities as choices in the editing
prompt. `command-help-v1`, restore shortcuts, speculative apply, and localized
change summaries are optional requirements only for integrations that use them.

## Safety invariants

Every generated integration must:

- inspect narrowly, retain `exactText`, and target with `paragraphId` plus
  `fingerprint`;
- use strict targeting, package validation, and a derived output path;
- never overwrite the source unless `--in-place` is explicitly requested;
- require `completion: true` and a non-null `outputPath` before reporting
  success;
- follow `error.recovery.action` and never retry unchanged failed arguments; and
- require explicit authorization before accepting/rejecting another reviewer's
  work or deleting comments.

The CLI computes `completion` from write status, operation outcomes, commit
disposition, and localized accepted-view verification. Preserve the detailed
results for audit and recovery, but do not make the editing agent re-evaluate
those fields during ordinary success handling.

## Workflow policy choices

Keep these exceptions separate from the default workflow:

| Choice | Default | Override only when |
|---|---|---|
| Transaction | progressive | Partial output is unacceptable (`--atomic`). |
| Existing revisions | `merge-same-author` | Editing inside another reviewer's pending insertion is intended (`slice-cross-author`). |
| Reviewer identity | `AI Redliner` is a valid visible fallback | The host supplies a preferred flag, operation author, or environment value. |
| Output | derived sibling | The caller deliberately selects another destination. |

Check `effectiveOptions` when an integration deliberately sets a policy.

## Transport choices

A UTF-8 operations file and serializer-backed stdin are peer transports:

```bash
docx-redline apply input.docx --operations operations.json --profile agent --output reviewed.docx
node emit-operations.mjs | docx-redline apply input.docx --operations - --profile agent --compact --output reviewed.docx
```

Use `JSON.stringify` or a structured tool API. Never demonstrate `echo`, a
heredoc, or shell-interpolated document text. In JavaScript/TypeScript runtimes
(Node.js, browsers, Cloudflare Workers, Deno, Bun, or sandboxed environments like
n8n), an in-process wrapper should call `openDocx(uint8Array)` directly from
`@ansonlai/docx-redline-js` (or `@ansonlai/docx-redline-js/bundle` for zero-dependency
standalone sandboxes) and return the structured result instead of recreating CLI
serialization, ZIP handling, targeting, validation, or rollback.

## Ordinary generated workflow

The first executable editing example in a generated skill must be focused
extraction:

```bash
docx-redline extract input.docx --search "renewal notice" --around 3
```

Then apply once using the inspected strong target. Use complete `modified` text
for a broad semantic revision or `replacements` for a small literal change in
that target. Keep requested relative wording relative unless the user asks for
a concrete value.

Follow `selection.nextAfter` when extraction is truncated. Output projections must retain
`selection`, target descriptors, `humanReference`, and revision
cues.

### Observed restore branch

If the user requests restoration, search `--view rejected` first. Otherwise
follow an alternate-view `selection.hint` when returned. Restore only the
inspected strong target. Require `restore-shortcuts-v1` only if the integration
uses inline `--restore`; canonical operations remain available.

### Optional speculative execution

Do not make speculative apply part of an ordinary generated workflow. A harness
may opt into `speculative-search-apply-v1` only when it owns a measured confidence
policy and accepts the recovery cost of a miss. Keep anchor selection,
directional windows, occurrence handling, and exact matching semantics out of
the ordinary editing prompt; link to the knowledge base or CLI help instead.

## Recovery contract

Lead failure handling with these runtime fields:

1. `error.recovery.action`
2. `error.recovery.requiresReinspection`
3. `error.recovery.requiresUserAuthorization`
4. `error.recovery.sameArgumentsSafe`

Use `retryPlan.base` and its operation indexes to select the original or partial
output package. Do not reproduce an error-code decision tree in the skill.

## Presentation rules

Use `humanReference`, `provision`, or `nearestHeading` in completion reports.
Paragraph indexes and `P55`-style references are machine coordinates, not
locations a document user can follow. Keep the exact machine target internally.

## Generation checklist

- Exact tested release, contract 8, and the minimum capability set are pinned.
- Runtime compatibility is checked once.
- The first ordinary command is focused extraction.
- Apply uses the inspected target with `modified` or localized `replacements`.
- Restore appears only as a user-requested or observed-view branch.
- Speculative execution is absent unless the harness owns a measured policy.
- Ordinary success requires `completion: true` and a non-null `outputPath`.
- Recovery, authorization, source immutability, and structured serialization
  invariants are preserved.
- Wrappers delegate mutation and OOXML packaging to the library.
