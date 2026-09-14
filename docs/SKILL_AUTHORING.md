# Skill and Harness Authoring Contract

Use this page when creating or updating an agent skill, MCP server, or custom
harness for `@ansonlai/docx-redline-js`. Ordinary document-editing agents should
load [Agent Fast Start](AGENT_FAST_START.md), not this design contract.

## Runtime negotiation

Run this before the first document operation when the skill vendors or invokes
the CLI:

```bash
docx-redline version
```

Require contract version 7 and only the capabilities the integration uses. A
typical shell skill should require `command-help-v1`, `inspection-context-v1`,
`bounded-inspection-v1`, `human-document-references-v1`,
`agent-safety-profile-v2`, `deduplicated-cli-receipts`, and
`recovery-envelope-v1`. Require `operations-stdin` or `compact-cli-json-v1` only
when the host uses those features. Fail closed with an upgrade instruction when
a required capability is absent. Never inspect `dist/`, minified source, or ZIP
parts to compensate for a stale runtime.

## Safety invariants

Every generated integration must preserve these rules:

- Inspect narrowly and copy `exactText` plus `paragraphId` or `fingerprint`.
- Treat `modified` as the complete desired accepted-view target content.
- Use strict targeting, package validation, and a derived output path; never
  overwrite the source unless the caller explicitly requests `--in-place`.
- Require `completion: true`, `written: true`, a non-null `outputPath`, and no
  per-operation error before reporting completion.
- Never accept or reject another reviewer's work or delete comments without
  explicit user authorization.
- Never retry unchanged failed arguments.

The `agent` profile enables complete-success exit behavior and retains the
ordinary progressive default. It does not select atomic execution or a special
existing-revision policy. A skill must state those choices explicitly when its
workflow needs them.

## Workflow policy choices

Choose and document these separately from safety invariants:

| Choice | Options | Guidance |
|---|---|---|
| Transaction | progressive default or `--atomic` | Use progressive for independent edits; use atomic when partial output is unacceptable. |
| Existing revisions | `merge-same-author` default or explicit policy | Use `slice-cross-author` only when editing inside another reviewer's pending insertion is intended. |
| Reviewer identity | flag, operation author, environment, or fallback | `AI Redliner` is a valid visible fallback. `DOCX_REDLINE_AUTHOR` is an optional harness preference, not a guard. |
| Output naming | derived sibling or explicit destination | Keep source immutability unless in-place mutation is deliberately authorized. |

Examples of valid profile composition:

```bash
docx-redline apply input.docx --operations operations.json --profile agent --output reviewed.docx
docx-redline apply input.docx --operations operations.json --profile agent --atomic --output reviewed.docx
docx-redline apply input.docx --operations operations.json --profile agent --existing-revisions slice-cross-author --output reviewed.docx
```

Check `effectiveOptions` rather than assuming the resolved transaction,
revision, author, or redline policy.

## Transport choices

A UTF-8 operations file and serializer-backed stdin are peer transports. Choose
the form the host can construct safely:

```bash
docx-redline apply input.docx --operations operations.json --profile agent --output reviewed.docx
node emit-operations.mjs | docx-redline apply input.docx --operations - --profile agent --compact --output reviewed.docx
```

Use `JSON.stringify` or a structured tool API. Never demonstrate `echo`, a
heredoc, or shell-interpolated legal text. `--compact` changes stdout formatting,
not result semantics. A Node byte-oriented wrapper should call `openDocx` and
return the full structured result rather than recreating the CLI serializer.

## Ordinary generated workflow

The first executable document-editing example in a generated skill should be a
focused contextual extraction:

```bash
docx-redline extract input.docx --search "force majeure" --around 3
```

Use `selection.nextAfter` with `--after` when truncated. When a search for
deleted or restorable content yields 0 matches in accepted view, follow
`selection.hint` and re-run with `--view rejected`. Apply once per stable batch.
Run `docx-redline apply --help` for canonical redline, whole-paragraph comment,
and rejected-view restore shapes; CLI help documentation fields return canonical
GitHub URLs. The operation schema remains
[document-operations.schema.json](schemas/document-operations.schema.json).

## Recovery contract

Lead error handling with exactly these runtime fields:

1. `error.recovery.action`
2. `error.recovery.requiresReinspection`
3. `error.recovery.requiresUserAuthorization`
4. `error.recovery.sameArgumentsSafe`

Use `retryPlan.base` and its operation indexes to select the original or partial
output package. Do not generate a second prose decision tree for every error
code; error messages and bounded diagnostics explain the cause, while the four
fields above determine the next action.

## Presentation rules

Use `humanReference`, `provision`, or `nearestHeading` in user-facing completion
reports. `P55`, `index: 55`, and phrases such as “the 55th paragraph” are machine
targeting and pagination coordinates, not locations a Word user can follow.
When no formal heading exists, use the returned short text lead. Never replace
an exact machine target with a human-facing reference.

## Generation checklist

- Runtime version and required capabilities are checked once.
- The first ordinary edit command is focused `extract --search ... --around`.
- Safety invariants are not mixed with transaction or revision-policy choices.
- Operations-file and stdin examples use structured serialization.
- The four-field recovery contract precedes any diagnostic commentary.
- Completion prose uses legal references rather than machine ordinals.
- The wrapper delegates mutation, validation, rollback, comments, numbering,
  receipts, and OOXML packaging to the library.
