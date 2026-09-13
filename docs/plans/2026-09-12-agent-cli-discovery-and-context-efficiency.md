# Agent CLI Discovery and Context Efficiency Plan

**Status:** Completed

**Date:** 2026-09-12

**Target:** Next CLI contract after version 5; suitable for inclusion in 0.6.0
if completed before release

**Priority:** Remove avoidable agent tool turns and stdout/context waste without
weakening exact targeting, revision attribution, recovery diagnostics, or
transactional guarantees.

---

## 1. Executive Summary

Real Claude usage after WP-00 through WP-07 confirms that the core DOCX path is
not the remaining bottleneck. The observed waste occurred before and after the
mutation itself:

1. generic CLI help did not answer operation-shape questions, causing the agent
   to inspect a vendored/minified bundle;
2. search and context retrieval required separate commands;
3. undocumented case-insensitive search caused duplicate queries;
4. unscoped detailed inspection consumed thousands of unused tokens;
5. the agent discarded a valid fallback-authored document and applied the same
   operations again with a preferred reviewer identity;
6. broad searches exceeded the harness output limit; and
7. compact CLI results repeated every receipt and retained more successful-match
   detail than the agent needed.

The proposed fixes make sense and are consistent with the existing architecture.
They should be implemented at the CLI/inspection presentation boundary. The
OOXML engine, Node transaction facade, canonical operation schema, receipts, and
validation model remain authoritative and unchanged.

Three recommendations require narrower treatment than the raw report suggests:

- The fallback author worked as designed and is reported in `authorUsed`. A skill
  may optionally set `DOCX_REDLINE_AUTHOR` when personalized attribution matters,
  but the library should not reject a valid operation, add an author guard, or
  introduce author-provenance/version-handshake machinery.
- Receipt duplication should be removed from compact CLI JSON only. Full Node and
  standalone-runner results should retain their current receipt contract. In the
  CLI, keep the authoritative top-level `receipts` array because it also accounts
  for refused and unattempted operations; remove duplicate nested receipt bodies
  from `results[i]`.
- Stdin is a speed optimization only when a harness can pipe output from a real
  JSON serializer. Structured operation files remain first-class and preferable
  to raw shell interpolation for Unicode legal text.

## 2. Confirmed Current Behavior

| Observation | Current implementation | Assessment |
|---|---|---|
| `apply --help` is not useful | `node/cli.js` returns the same one-line usage object for global and command help | Confirmed; highest-priority discovery fix |
| Search requires a second range call for context | CLI inspection accepts search/index/indexes/range/view but no context option | Confirmed; add context-window selection to the shared inspection path |
| Search case behavior is unclear | `inspectDocumentParts` lowercases both the query and paragraph text | Behavior is already correct; expose it in help and compact docs |
| `inspect --non-empty` can dump the document | `inspect` returns the full detailed paragraph model when no positional scope is supplied | Confirmed; bound unscoped CLI discovery while retaining explicit full inspection |
| Agent repeated a fallback-authored apply | CLI intentionally resolves flag, then `DOCX_REDLINE_AUTHOR`, then visible fallback `AI Redliner`, and reports `authorUsed` | Not a library defect; retain the fallback and document optional harness configuration only |
| Broad searches can overflow tool transport | There is no result limit, cursor, or output-budget metadata | Confirmed; add stable result pagination and an intentional unbounded escape hatch |
| Receipts are duplicated | Compact CLI returns `results[i].receipt` and the same data in top-level `receipts` | Confirmed; deduplicate only the compact serialization contract |
| Successful whitespace matches are verbose | `targetTextMatch` can retain up to eight code-point differences and two excerpts | Confirmed; summarize successful equivalence while preserving full error recovery detail |
| The generated skill led with `inspect --non-empty` | The first example contradicted its “inspect narrowly” heading and the model copied the example | Confirmed; treat example order as behavioral policy and never lead with an unscoped inspection |
| The generated skill prohibited `--profile agent` | The profile combines complete-success safety with atomic transaction policy, conflicting with the skill’s progressive house rule | Confirmed; separate invariants from policy choices instead of asking every skill to undo a monolithic preset |
| Stdin had low adoption | The skill correctly distrusted raw shell construction of Unicode legal JSON and preferred a structured operations file | Expected; retain both transports and recommend stdin only when the harness owns a real serializer/pipe |
| Recovery prose became much smaller | The generated skill replaced many hand-written code branches with the recovery envelope’s four generic fields | Confirmed success; preserve the envelope-first rule and avoid rebuilding a duplicate prose decision tree |
| Machine paragraph ordinals leaked into user prose | Compact extraction emphasizes `index`/`ref` and omits the inspection model’s `humanReference` | Confirmed; return and prioritize a human-facing legal location while retaining ordinals for machine pagination only |

The agent’s decision to inspect a vendored bundle was contrary to repository
guidance, but better guidance alone is insufficient. A shipped CLI must be
self-describing enough that an agent never needs source access to construct a
supported operation.

### 2.1 Implications for generated skills

The comparison of the generated 0.5.4 and 0.6.0 skills shows that documentation
is not passive reference material. Models frequently treat the first command and
first policy statement as the default algorithm, even when the surrounding prose
says otherwise. The package therefore needs an explicit skill-authoring contract,
not just correct facts spread across README, AGENTS, and the knowledge base.

That contract should distinguish:

- **invariants**: strict/exact targeting, validation, source immutability,
  complete-success detection, machine-actionable recovery, and explicit
  authorization for review normalization or comment removal;
- **integration choices**: progressive versus atomic execution, default
  existing-revision policy, reviewer identity, output naming, and operations
  transport; and
- **presentation rules**: search narrowly first, prefer legal provisions and
  headings in user prose, and never expose `P55`, “paragraph 55,” or “the 55th
  paragraph” as though it were a location a Word user can follow.

The recovery-envelope simplification is evidence that runtime contracts can
shrink generated skills. New runtime metadata should follow that pattern rather
than adding another large error-code matrix to the prompt.

## 3. Goals

- Make global and command-specific help sufficient to construct ordinary
  `redline`, `comment`, and `restore` operations.
- Return a search hit and its drafting context in one inspection command.
- State search matching semantics in runtime help, not only prose documentation.
- Keep every ordinary discovery response within a predictable transport budget.
- Preserve complete `exactText` for every returned target; never truncate target
  strings merely to meet a byte limit.
- Separate agent safety guarantees from atomic/progressive and revision-policy
  choices so skills can adopt the safety layer without blacklisting the profile.
- Make focused contextual extraction the first executable example in every
  agent-facing document; move unscoped detailed inspection to advanced guidance.
- Treat operations files and stdin as equally valid transports selected by host
  capability, not as a safety hierarchy imposed by the library.
- Give agents a human-facing provision/heading reference alongside machine
  indexes so completion reports are actionable in Microsoft Word.
- Remove duplicate receipt payloads from CLI output without weakening the
  library-level receipt/reconciliation contract.
- Preserve actionable mismatch detail on errors while minimizing diagnostics on
  successful operations.
- Measure actual command count, response bytes, and provider-visible output
  separately from native library time.

## 4. Non-goals

- Do not change reconciliation, revision allocation, ZIP handling, package
  validation, or Accept/Reject behavior.
- Do not make the development agent-session example a published API or package
  dependency.
- Do not expose internal engine helpers or vendored implementation symbols in
  CLI help.
- Do not embed the entire JSON Schema in every help response.
- Do not silently choose one target from a broad or ambiguous result set.
- Do not truncate `exactText`, paragraph IDs, fingerprints, revision views, or
  recovery fields required for a safe apply.
- Do not add `.docx-redline.json` discovery, home-directory configuration, or
  other hidden author precedence.
- Do not add an author guard, new author-configuration error, author-provenance
  capability, or version-handshake requirement. `AI Redliner` remains a valid,
  visible fallback and `authorUsed` remains the audit field.
- Do not remove or reshape receipts from the Node facade or standalone runner.
- Do not remove detailed `targetTextMatch` diagnostics from errors where they are
  needed to correct exact whitespace or Unicode mismatches.
- Do not force stdin when the host cannot pipe serializer output without raw shell
  interpolation; a structured temporary JSON file remains a safe supported path.
- Do not declare atomic, progressive, `merge-same-author`, or
  `slice-cross-author` universally superior. They encode workflow intent, not
  generic safety.
- Do not remove machine paragraph indexes; they remain useful for descriptors,
  pagination, diagnostics, and internal execution.

## 5. Contract Principles

### 5.1 Help is a public machine-readable surface

Help responses remain valid JSON so shell agents can parse them. Global help
should list commands and route to command help. Command help should contain only
the relevant synopsis, option records, behavioral notes, exit codes, and a small
set of canonical examples.

`apply --help` must explain that `modified` is the complete desired accepted-view
content and provide minimal `redline`, `comment`, and rejected-view `restore`
operation objects. It must state that source is not overwritten without
`--in-place`, the agent profile enforces complete-success behavior, explicit
flags control transaction/revision policy, and authorization is required before
accepting/rejecting foreign work or deleting comments.

Help must link to packaged files (`AGENTS.md`, `docs/AGENT_FAST_START.md`, and the
operation schema) rather than internal source. A help response should be bounded;
advanced examples stay in the knowledge base.

### 5.2 Search context is selected from one immutable inspection

Context expansion must not trigger a second parse or package read. The shared
inspection implementation should identify matches against the complete selected
revision view, select/paginate direct matches, and then expand surrounding
paragraphs from that same snapshot.

Direct hits and context must be distinguishable. Context paragraphs must not be
mistaken for additional search hits or consume the direct-hit limit.

### 5.3 Bounded output must remain valid and lossless per item

Never allow the shell or harness to truncate JSON. Limit output by returning
fewer complete paragraph records, accompanied by explicit pagination metadata.
If one paragraph alone exceeds the soft byte budget, return that whole paragraph
and set an oversize indicator rather than corrupting `exactText`.

### 5.4 The fallback author is valid behavior

The CLI’s `AI Redliner` fallback is intentional, visible in tracked-change
metadata, and reported through `authorUsed`. It is not unsafe OOXML and does not
justify rejecting or delaying an otherwise valid edit. A skill or harness may
set `DOCX_REDLINE_AUTHOR`, pass `--author`, or use operation-level authors when a
particular reviewer identity matters. That is an optional integration policy,
not a universal library safety requirement.

### 5.5 Safety guarantees and workflow policy are different axes

An agent preset should contain only behavior broadly required for safe machine
automation. `requireComplete`, strict targeting, validation, and source
protection qualify. Atomic versus progressive execution and
`merge-same-author` versus `slice-cross-author` do not; those are deliberate
workflow choices.

Before the 0.6.0 contract is finalized, revise `--profile agent` so it does not
force atomic execution. It should inherit the ordinary CLI transaction and
existing-revision settings unless flags override them. Skills that require
progressive, complete-success execution can use the profile without negating it;
skills requiring atomic rollback pass `--atomic`; skills intending cross-author
slicing pass `--existing-revisions slice-cross-author`.

If contract version 5 is already relied upon outside the repository, preserve
its old atomic preset under an explicitly named compatibility profile or require
contract negotiation. Do not silently reinterpret a published contract.

### 5.6 Examples are executable policy

The first ordinary workflow example must be a focused command such as:

```bash
docx-redline extract input.docx --search "force majeure" --around 3
```

Do not place `inspect input.docx --non-empty`, `inspect input.docx`, or `--all`
in the first workflow block. These remain advanced diagnostics after scoped
search/range examples. Code-block order is part of the agent contract and should
be tested like prose requirements.

## 6. Proposed CLI Surface

### 6.1 Command-specific help

Support both:

```bash
docx-redline --help
docx-redline apply --help
docx-redline extract --help
```

Return a stable structure resembling:

```json
{
  "status": "ok",
  "command": "help",
  "forCommand": "extract",
  "usage": "docx-redline extract <file.docx> [options]",
  "options": [
    { "name": "--search <text>", "description": "Case-insensitive substring search." },
    { "name": "--around <N>, --context <N>, -C <N>", "description": "Return N surrounding paragraphs." },
    { "name": "--limit <N>", "description": "Limit direct matches; context does not count." }
  ],
  "notes": [],
  "examples": [],
  "documentation": []
}
```

Use data tables in an unbundled module owned by `node/`; do not hand-maintain a
second parser or infer flags by reading the bundle. Tests must ensure every
accepted command option appears in that command’s help and every documented
option is accepted by the parser.

Add a capability such as `command-help-v1` and advance the CLI contract version.

### 6.2 Context-window extraction

Add:

```bash
docx-redline extract contract.docx --search "Prohibited Use" --view rejected --around 4
docx-redline inspect contract.docx --search "Notices" -C 2
```

Rules:

- `--around`, `--context`, and `-C` are aliases.
- Accept integer values from 0 through 20.
- Initially require `--search`; index/range context can be added later only if a
  measured workflow needs it.
- Search is Unicode-preserving, case-insensitive substring matching, consistent
  with current behavior.
- Match filters (`revised`, `table`, `body`, `non-empty`, and explicit range)
  select direct hits. Context comes from adjacent physical paragraphs in the
  same revision view and remains clamped to an explicit range.
- Overlapping context windows are deduplicated and returned in document order.
- Each returned paragraph has `selectionRole: "match" | "context"`; optionally
  include `contextFor` indexes when one paragraph supports multiple hits.
- The response distinguishes `matchCount`, `returnedMatchCount`, and total
  returned paragraphs.

Implement the selection semantics once in `services/document-inspection.js` so
the Node facade and CLI agree. The development example may delegate to that
behavior but remains unpublished.

Add an `inspection-context-v1` capability.

### 6.3 Limits, pagination, and deliberate full output

Add:

```bash
docx-redline extract contract.docx --search "email" --limit 10
docx-redline extract contract.docx --search "email" --limit 10 --after 84
docx-redline inspect contract.docx --all
```

Rules:

- `--limit` counts direct matches before context expansion.
- `--after` is an exclusive, stable 1-based source paragraph index, not an
  opaque offset into the filtered array.
- Unbounded CLI inspection/search gets a conservative default direct-match cap.
  Start with 20 and tune against the transcript-derived benchmark corpus.
- Apply a soft serialized-response budget below the observed 64 KiB harness
  ceiling (target 48 KiB). Reduce the number of whole returned records if needed;
  never cut a record or emit invalid JSON.
- Return `{ totalMatches, returnedMatches, truncated, nextAfter,
  softByteLimit, oversizeItem }` metadata.
- `--all` explicitly removes the default result and soft-byte caps. It is
  incompatible with `--limit` and must be an intentional choice.
- Explicit `--index`, `--indexes`, or `--range` requests are already scoped and
  should not receive the default match cap, although an explicitly supplied
  `--limit` remains valid.
- `inspect --non-empty` without another scope returns a bounded page and a note
  recommending `extract --search`, rather than dumping the complete detailed
  document model.

Add a `bounded-inspection-v1` capability. Keep the underlying programmatic
inspection API capable of returning a complete document when explicitly asked;
the default cap is a CLI transport policy.

### 6.4 Composable agent safety profile

Revise the agent profile before release so it represents the automation safety
contract without selecting a transaction strategy:

- `--profile agent` enables complete-success exit behavior.
- It does not force `atomic: true`; absent an explicit flag, the ordinary CLI
  progressive default remains in effect.
- It does not select a special existing-revision policy. The ordinary default
  remains in effect unless the caller deliberately passes
  `--existing-revisions`.
- `--profile agent --atomic` and
  `--profile agent --existing-revisions slice-cross-author` are supported,
  documented compositions.
- `effectiveOptions` reports the final resolved values so a skill can verify its
  house rules rather than banning the profile.

Reviewer identity remains unchanged. When a skill or harness wants a personalized
reviewer rather than the valid `AI Redliner` fallback, it may configure the
existing environment variable once:

```bash
export DOCX_REDLINE_AUTHOR="Lai, Anson"
```

On PowerShell:

```powershell
$env:DOCX_REDLINE_AUTHOR = 'Lai, Anson'
```

This is optional integration guidance, not a new CLI requirement. Do not add an
author guard, configuration error, source/provenance field, capability, or
version response. The existing `authorUsed` field is sufficient to report the
identity actually stamped into the document.

### 6.5 Compact mutation output

For CLI serialization only:

- Keep top-level `receipts`, ordered by `operationIndex`, as the authoritative
  receipt collection.
- Remove the duplicate `receipt` body from each compact `results[i]` item.
  `results[i].index` maps directly to `receipts[i].operationIndex`; do not add a
  verbose replacement reference unless measurement shows it is necessary.
- Retain all receipts for applied, no-change, refused, rolled-back, and
  unattempted operations.
- Do not change Node facade or standalone runner results.
- For successful exact matches, omit `targetTextMatch` entirely.
- For successful space-equivalent/normalized matches, retain a compact summary
  such as `{ mode, differenceCount }`. Preserve the first difference only if it
  materially helps auditability.
- On errors, retain bounded code-point differences and excerpts required by the
  recovery contract.
- Measure pretty-printed versus minified JSON after structural deduplication.
  Add compact serialization for the agent path only if it produces a material
  additional provider-visible reduction; do not mix that optional decision with
  the receipt contract.

Advance the CLI contract and advertise `deduplicated-cli-receipts` (and a
separate compact-serialization capability only if implemented).

### 6.6 Human-facing location metadata

Compact `extract` currently retains `index`, `ref`, and `nearestHeading` but
drops `humanReference` and `provision` from the richer inspection record. Reverse
that presentation imbalance:

- Include a concise `humanReference` (or a clearly named `userReference`) in
  every extracted match when a provision, list label, or nearest heading is
  available.
- Keep `index` and `ref` for machine targeting, context windows, and pagination,
  but add top-level response guidance such as
  `machineReferencesAreNotUserLocations: true`.
- Prefer provision plus heading, for example `Section 4.2 — Prohibited Use`.
  Avoid appending a long paragraph excerpt when a short legal reference exists.
- When no formal reference exists, use the nearest heading and a short unique
  text lead rather than a whole-document ordinal.
- Put the human reference before `index`/`ref` in compact serialized paragraph
  records. JSON property order is not semantic to parsers, but it affects what a
  model sees first.
- State in help and skill-authoring guidance that `P55`, “paragraph 55,” and
  “the 55th paragraph” are machine ordinals and must not be used in user-facing
  completion reports.

This is additive output metadata. It does not change target descriptors or
permit a human label to replace exact targeting.

Add a `human-document-references-v1` capability.

## 7. Work Packages

### WP-00: Reproduce and baseline the observed waste

**Status:** Completed 2026-09-12

**Goal:** Turn the real transcript into stable measurements before changing the
surface.

- Add synthetic documents that reproduce heading search, rejected-view context,
  common-keyword fan-out, unscoped detailed inspection, NBSP-equivalent targeting,
  and multi-operation receipts.
- Record command count, serialized stdout bytes, paragraph records returned,
  native execution time, and whether output exceeds 48/64 KiB.
- Capture current global/apply/extract help payloads and confirm they do not
  contain the needed operation and search semantics.
- Diff the generated 0.5.4/0.6.0 skill behavior: first executable inspection,
  profile policy, operations transport, recovery-rule word count, and use of
  machine ordinals in user-facing text.
- Do not estimate model tokens or invent Claude wall time beyond the supplied
  observation.

**Acceptance:** Every proposed optimization has a reproducible before-state and
the benchmark separates CLI bytes/tool turns from native DOCX time.

**Delivered:** Added `npm run benchmark:agent-cli` and a checked validation
report. The before-state records identical 171-byte global/apply help, a
72,387-byte unscoped inspection, a two-command context lookup, broad-search
fan-out, and duplicated receipt bytes. Active measurements record native time
separately and make no provider-token or model-latency claim.

### WP-01: Ship self-describing command help

**Status:** Completed 2026-09-12

**Goal:** Eliminate source/bundle inspection as a discovery strategy.

- Add the help data module and command-specific response builder.
- Cover every accepted flag and alias.
- Include compact redline, comment, and restore examples for apply.
- State complete-`modified`, rejected-view restore, case-insensitive search,
  output safety, revision authorization, profile override/composition, safe
  file-versus-stdin transport, and exit semantics.
- Add the contract capability and update typings if help types are exported.

**Acceptance:** An agent can construct each common operation and a focused
search using only `version` plus the appropriate command help. Help stays
bounded and contains no internal implementation symbol.

**Delivered:** Added a shared CLI help catalog that also defines accepted option
keys, global command-index help, and per-command JSON help. Contract version 6
advertises `command-help-v1`; `apply` help contains redline, comment, and
rejected-view restore shapes, while inspection help documents case-insensitive
search, context, limits, pagination, and machine-reference rules.

### WP-02: Add context windows and bounded discovery

**Status:** Completed 2026-09-12

**Goal:** Make one search sufficient while preventing broad-output floods.

- Implement shared direct-hit selection and context expansion.
- Add `--around`/`--context`/`-C`, `--limit`, `--after`, and `--all` parsing and
  validation.
- Return match/context roles and stable pagination metadata.
- Add CLI soft-byte budgeting without truncating exact target records.
- Preserve revision-view, table/list, heading, paragraph ID, and fingerprint
  accuracy.
- Return concise human-facing provision/heading references while retaining
  machine ordinals for targeting and pagination.
- Add context and bounded-inspection capabilities.

**Acceptance:** The transcript’s search-then-range sequences become one command;
case variants return the same hits; broad `email`/`Notices` queries return valid,
bounded JSON with a continuation; and explicit `--all` retains deliberate full
inspection. Extracted results give the agent a usable legal location, and the
completion report does not describe that location as a document-wide paragraph
ordinal.

**Delivered:** Added immutable-snapshot context expansion, match/context roles,
`--around`/`--context`/`-C`, `--limit`, `--after`, and `--all`. CLI discovery
defaults to 20 direct hits and a 48 KiB soft budget, drops only whole records,
and returns an oversized target intact. Contract version 6 advertises
`inspection-context-v1`, `bounded-inspection-v1`, and
`human-document-references-v1`. The checked broad inspection fell from 72,387
to 25,093 bytes; contextual search now returns one match plus two surrounding
paragraphs in one command.

### WP-03: Decompose the agent profile

**Status:** Completed 2026-09-13

**Goal:** Let skills adopt machine-safety guarantees without inheriting a
conflicting transaction or revision policy.

- Remove the agent profile’s implicit atomic selection before the next contract
  is published, or preserve it under an explicit compatibility profile if
  external version-5 reliance is discovered.
- Prove explicit atomic/progressive and existing-revision flags override profile
  defaults and are reflected by `effectiveOptions`.
- Keep `AI Redliner`, `--author`, operation-level authors, and
  `DOCX_REDLINE_AUTHOR` behavior unchanged.
- Mention the environment variable only as an optional skill/harness convenience
  when personalized attribution matters.

**Acceptance:** Both progressive and atomic skills can use the safety profile
without contradicting their house rules. Profile and non-profile calls retain
the existing author fallback and report the stamped identity through
`authorUsed`; no new author-related failure path is introduced.

**Delivered:** CLI contract version 7 advertises `agent-safety-profile-v2`.
`--profile agent` now enables complete-success exits while retaining progressive
execution and `merge-same-author`; `--atomic`, `--atomic=false`, and explicit
existing-revision policies compose and are reported in `effectiveOptions`.
Author precedence and the visible `AI Redliner` fallback are unchanged.

### WP-04: Deduplicate compact CLI mutation results

**Status:** Completed 2026-09-13

**Goal:** Cut successful apply output without losing audit or recovery data.

- Strip nested receipt bodies from compact CLI results and retain the canonical
  top-level collection.
- Compact successful target-match metadata.
- Preserve full library-layer results and error diagnostics.
- Measure optional minified agent JSON separately.
- Add a contract capability and migration note.

**Acceptance:** No receipt body is repeated in CLI JSON; every operation still
has an authoritative receipt/disposition; atomic/progressive retry plans remain
complete; and NBSP mismatch errors retain corrective evidence.

**Delivered:** Contract version 7 advertises `deduplicated-cli-receipts` and
`compact-cli-json-v1`. Compact CLI results remove nested receipt bodies, retain
the complete ordered top-level receipt collection, omit successful exact-match
metadata, and reduce successful equivalent-whitespace metadata to mode/count.
Error evidence and full Node/standalone results are preserved. The measured
one-operation response is 2,216 bytes pretty-printed or 1,557 bytes with
`--compact`, versus the 2,968-byte pre-WP-04 response.

### WP-05: Documentation, bundle parity, and real-world audit

**Status:** Completed 2026-09-13

**Goal:** Ensure the shipped skill actually receives the improvements.

- Update `AGENTS.md`, `docs/AGENT_FAST_START.md`, the knowledge base, README,
  schema references, changelog, and 0.6.0 release notes.
- Add a compact, packaged skill-authoring contract that separates invariants,
  policy choices, transport choices, and presentation rules. Ordinary editing
  agents should not load it; it is input specifically for creating/updating a
  skill or harness. Use `docs/SKILL_AUTHORING.md` unless implementation reveals a
  clearer consumer-facing name, and add it to the curated npm documentation
  allowlist.
- Make focused contextual `extract --search` the first executable ordinary-use
  example everywhere. Move `inspect --non-empty` and `--all` to deliberate
  advanced-inspection sections.
- Give the fast start three compact, copyable operation shapes—ordinary redline,
  whole-paragraph comment, and rejected-view restore—or a single clearly
  executable route to the equivalent command help. Preserve the fast-start
  budget by removing superseded profile/transport prose rather than appending.
- Present operations files and serializer-backed stdin as peer transports. Never
  demonstrate raw `echo`, heredoc, or shell-interpolated legal JSON.
- Preserve the recovery-envelope master rule as the primary skill behavior and
  keep code-specific prose diagnostic rather than duplicating next-action logic.
- Add a human-reference rule that prohibits paraphrased whole-document ordinals
  in user-facing prose.
- Keep the fast start within its current line/word budget by replacing obsolete
  guidance rather than continually appending text.
- Rebuild distribution artifacts from unbundled source; never patch the vendored
  or minified bundle directly.
- Require the vendored skill to run `version` and verify the new contract and
  capabilities before use.
- Re-run the original Claude task shape and record help/source-inspection turns,
  search/context turns, apply attempts, output bytes, and fidelity.

**Acceptance:** A regenerated skill leads with focused contextual extraction,
states its explicit progressive/atomic and revision-policy choices, selects a
transport supported safely by its host, follows recovery metadata instead of a
duplicated error decision tree, and reports legal provisions/headings rather
than machine ordinals. The vendored runtime reports the same contract as source;
no bundle-source inspection occurs; the task uses one contextual search and one
apply; output remains below the harness limit; accepted and rejected views,
comments, and foreign author attribution remain correct. A skill that elects to
use a personalized reviewer configures it through the existing author inputs.

**Delivered:** Added the packaged `docs/SKILL_AUTHORING.md` contract, updated all
agent-facing routes and release documentation, and added executable documentation
guards for scoped-first examples, profile composition, transport parity,
four-field recovery, capability negotiation, and human references. Package
dry-run verifies the source CLI/help and curated documentation are present while
development benchmarks remain excluded. Repository benchmarks and lifecycle
tests cover one contextual search, one apply, output size, accepted/rejected
text, comments, source immutability, receipts, and foreign attribution without
claiming provider/model timings.

## 8. Focused Verification

Add or extend focused suites:

- `tests/agent_cli_help_tests.mjs`
- `tests/agent_cli_inspection_context_tests.mjs`
- `tests/agent_cli_output_budget_tests.mjs`
- `tests/agent_cli_protocol_tests.mjs`
- `tests/agent_cli_tests.mjs`
- `tests/document_inspection_tests.mjs`
- `tests/document_inspection_edge_tests.mjs`
- `tests/error_recovery_contract_tests.mjs`
- `tests/agent_documentation_contract_tests.mjs`

Required cases:

1. global help lists all commands;
2. every command flag is represented in its help;
3. apply help includes valid redline, comment, and restore operation objects;
4. extract help states case-insensitive search and context/limit semantics;
5. upper/lower-case queries return identical indexes;
6. rejected-view search plus around returns one hit and ordered context;
7. overlapping windows deduplicate paragraphs;
8. limit counts hits rather than context records;
9. after-cursor pagination has no missing or duplicate direct hits;
10. broad extract and inspect remain valid JSON below the soft budget;
11. an oversize single paragraph is returned whole with an explicit indicator;
12. `--all` deliberately bypasses bounds;
13. invalid around/limit/after combinations fail before opening output;
14. compact CLI results contain one copy of each receipt;
15. rolled-back and unattempted receipt dispositions remain available;
16. successful target-match metadata is compact while mismatch recovery retains
    code-point evidence;
17. package revision enforcement, source immutability, validation, Accept All,
    Reject All, comments, and cross-author attribution remain unchanged.
18. the agent profile inherits progressive execution unless `--atomic` is
    explicitly supplied;
19. explicit atomic/progressive and existing-revision flags override the profile
    and appear accurately in `effectiveOptions`;
20. profile and non-profile calls preserve the existing author fallback,
    explicit author inputs, and `authorUsed` reporting without a new failure;
21. operations-file and serializer-backed stdin paths apply identical Unicode
    operation payloads and return equivalent semantic results;
22. compact extraction includes a concise legal `humanReference` while retaining
    machine targeting fields;
23. the first ordinary workflow code block in every agent-facing document is a
    scoped `extract --search` or `extract --range`, never broad `inspect`;
24. skill-authoring guidance labels safety invariants separately from workflow
    policy and transport choices;
25. skill guidance leads with the four-field recovery-envelope rule and does not
    duplicate next-action prose for every error code;
26. generated completion examples cite a provision/heading or short text lead,
    never a document-wide paragraph ordinal.

Run the focused tests first, followed by:

```bash
npm test
npm run check:types
npm run test:isolation
npm run build
npm pack --dry-run
```

Word COM and visual lanes are required only if shared inspection changes alter
document mutation or emitted OOXML. Pure CLI selection/serialization changes
must prove lifecycle fidelity through existing accepted/rejected fixture tests
without creating new Word-generated fixtures.

## 9. Success Metrics

For the reproduced real-world workflow:

- zero reads of `dist/`, installed plugin bundles, or vendored/minified source;
- command discovery requires at most one relevant help call;
- the first generated-skill inspection is focused and contextual, not
  `inspect --non-empty`;
- each clause search returns its requested context in one call;
- no duplicate case-variant search;
- the generated skill can adopt the agent safety profile while declaring either
  progressive or atomic execution explicitly;
- operations-file and stdin adoption follow host serialization capability; no
  unsafe raw-shell JSON is introduced merely to reduce a tool turn;
- no unscoped response exceeds the configured soft transport budget unless one
  indivisible exact paragraph is explicitly marked oversize;
- existing fallback and explicit author behavior remain unchanged; personalized
  attribution is an optional harness choice, not a mutation gate;
- successful apply JSON contains exactly one receipt body per operation;
- user-facing completion text cites legal headings/provisions rather than
  `P<index>` or paraphrased whole-document ordinals;
- recovery-envelope guidance remains a short generic rule rather than expanding
  back into a per-code action manual;
- apply stdout bytes decrease materially without removing actionable errors;
- native mutation time does not regress by more than benchmark noise;
- all accepted/rejected text, package validation, comments, receipts, and
  cross-author attribution gates remain green.

Provider wall time and token changes should be reported only from the actual
Claude/OpenCode harness. Repository benchmarks should continue to report bytes,
commands, and native time without converting them into estimated model tokens.

## 10. Risks and Mitigations

| Risk | Mitigation |
|---|---|
| Default limits hide additional matches | Return total/truncated/continuation metadata and require explicit `--all` for unbounded output |
| Context records are mistaken for targets | Label match versus context and count only direct hits |
| Byte budgeting corrupts exact targeting | Remove whole records only; never truncate `exactText`; flag an indivisible oversize item |
| Profile decomposition changes version-5 atomic behavior | Make the change before 0.6.0 publication, advance the contract, test explicit overrides, and preserve a named compatibility profile if external reliance is found |
| Receipt deduplication breaks Node callers | Apply it only in CLI compaction; leave facade/runner contracts unchanged |
| Help becomes another large prompt | Keep command help scoped, use three minimal apply examples, and link to packaged advanced docs |
| A generated skill copies a misleading example despite correct prose | Test code-block order and keep broad inspection examples out of the ordinary-use section |
| Stdin optimization creates Unicode/quoting failures | Recommend it only for serializer-backed pipes; retain structured operation files as a first-class path |
| Human-reference metadata is treated as a target | Keep exact target descriptors separate and state that human references are presentation-only |
| Source and vendored bundle diverge | Regenerate from source and verify contract/capabilities during skill startup |
| Shared inspection changes affect existing APIs | Keep new fields/options additive and run existing inspection, targeting, revision-view, and facade suites |

## 11. Recommended Sequence

Implement WP-00 first. WP-01 and WP-02 can then proceed independently against
the captured baseline. WP-03 should land with the next negotiated CLI contract
and its updated help. WP-04 should land after the response-size baseline so its
reduction can be measured accurately. Finish with WP-05 and one real provider
run.

The highest-value first slice is WP-01 plus the case-sensitivity note: it removes
the most pathological behavior—reverse-engineering a minified bundle—with low
risk. The next slice is WP-02 because it removes repeated turns and prevents
transport truncation. Profile composition and CLI output deduplication follow as
small, separately testable contract changes.
