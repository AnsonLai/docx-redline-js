# Canonical List Operations: Header-to-List Conversion and List Format Change

Date: 2026-09-30
Status: DONE (2026-10-11): WP-0 through WP-7 done; `detach`/`materializeMarker` deferred (see WP-7 notes). Package
baseline 0.8.4+ (not bumped; release notes draft `docs/releases/0.9.0.md`), CLI `contractVersion` 9 with
`list-format-v1` and `list-format-change-v1`. Remaining work is listed under "Open after this plan".

## Progress and hand-off (keep current; resume here)

| WP | State | Commits | Next |
|---|---|---|---|
| 0 | Done | `73c67e8` (+ earlier item c fix) | - |
| 1 | Done: list diagnostics warn by default, error under `strictListConversion` (CLI agent profile / `--atomic`) or `atomic`; `listMarkers: 'literal'` opt-out; `UNKNOWN_PARAGRAPH_FORMAT_PROPERTY` warning; capability `list-conversion-diagnostics-v1` | `db10d39`, `7ab6490`; tests `tests/list_conversion_diagnostics_tests.mjs` | Promote unknown `paragraph-format` keys to an error in the next minor; one-line adjacency insertion form not covered; consumer impact: atomic batches that type `1.`/`A.` onto plain paragraphs are now refused (tell AIWordPlugin) |
| 2 | Done: `services/numbering-allocation.js` is the single allocator and the single numbering merge for runner and facade (fidelity plan WP-5a R1-R3); `reuseSourceListNumbering` removed; facade refuses dangling numbering references it adds; `NUMBERING_ID_CONFLICT` on conflicting ids | `4718c36`, `f41bd3b`; tests `tests/numbering_allocation_tests.mjs`, matrix numbering parity in `tests/reject_roundtrip_matrix_tests.mjs` | Residual: separate bare calls without shared `numberingXml`/`numberingIdState` can still allocate the same ids (now detected, not auto-remapped) |
| 3 | Done: `paragraph-format.numbering` binds or removes existing numbering with a tracked `pPrChange`; validates `numId`/level; preserves the earliest same-author snapshot; refuses cross-author snapshot replacement; letter/roman marker starts are retained | `dc3e2d0` plus Word differential fixture follow-up; tests `tests/paragraph_format_numbering_tests.mjs`, Word task cases for bind/remove/change-level | WP-4 `list-format convert-plain` |
| 4 | Done: `list-format` `convert-plain` (`strip: 'marker'` default, `'none'` with `format`/`start`), decimal/letters/roman/bullets, in place (tracked `w:del` of exactly the marker + tracked `pPrChange` via the WP-3 path); values from the markers (`startOverride`, typed `.`/`)` kept in `lvlText`); consecutive values share one numId, gaps restart at the typed value, later ops continue through `listFallbackSequenceState` (entry reused only while its numId is live in the DOM, so refused ops cannot poison it); refusals `LIST_MARKER_NOT_FOUND`/`_AMBIGUOUS`/`_UNSUPPORTED`, `LIST_FORMAT_TARGET_IS_LIST`, `LIST_FORMAT_STYLE_NUMBERED`, `PARAGRAPH_FORMAT_REVISION_CONFLICT`, `TARGET_RANGE_INVALID`; text-equal redline letter/roman/bullet conversions use the same core, and text-changed redline header conversions in one batch continue the same sequence (`continueBatchListSequence`, 10-11 plan I2); WP-1 recovery is `use_list_format_convert_plain` + `suggestedOperation`; contract 9, `list-format-v1` | `1240842`, `1cd174a`; tests `tests/list_format_operation_tests.mjs`, matrix kind `list-format-convert-plain` + `typed-headers` fixture, Word cases `*-list-format-convert-*` (synthetic 56/56, corpus 60/60) | WP-5; see "WP-4 notes for WP-5/6" |
| 5 | Done: `list-format` `change-format` (`scope: 'list'` default; `'range'` = target..targetEnd running to the end of the list) and `join` (`continueFrom`, compiled at batch start); a new `abstractNum`+`num` per change (source levels/geometry copied, one level redefined, `isLgl`/picture bullet/bullet symbol font dropped on a format change), every item rebound with one tracked `pPrChange`; ALL start values (carried source `startOverride`s, split continuation, `list.start`) are `w:start` on the new abstract, never `startOverride` (Word in-session Accept All quirk, see WP-5 notes); refusals `LIST_FORMAT_TARGET_NOT_LIST`, `_MIXED_RANGE`, `_STYLE_LINKED`, `_SPLIT_UNSUPPORTED`, `_DEFINITION_UNSUPPORTED`, `_JOIN_INVALID` (+ `_STYLE_NUMBERED`, `PARAGRAPH_FORMAT_REVISION_CONFLICT`), `LIST_FORMAT_UNCHANGED` warning on a no-op; `LIST_FORMAT_NOT_APPLIED`/`UNSUPPORTED_LIST_FORMAT_CHANGE` recover with `use_list_format_change_format` + `suggestedOperation`; capability `list-format-change-v1` (contract stays 9) | `564ad0e`, `579efae`, `616d317`; tests `tests/list_format_operation_tests.mjs` section 10, matrix kinds `list-format-change-format`/`list-format-join` + `adjacent-lists` fixture, Word cases `*-list-format-change-*`/`*-list-format-join-*` (synthetic 62/62, corpus 60/60) | WP-6; see "WP-5 notes for WP-6/7" |
| 6 | Done: batch-compiler rule for one `redline` + one `list-format` writing one batch-start paragraph (`checkListFormatComposition` in `services/operation-batch-compiler.js`): compatible when the redline is a single-paragraph `redline`/`replace` (no `targetEnd`/anchor/part/structuredContent, no line break, accepted view) that changes the text and keeps the typed-marker prefix byte-identical (`parseTypedListMarker` prefix, `null` = no marker); the list-format gets `_compiledRunAfter` and `buildOperationDependencyPlan` runs it after the redline whatever the batch order; otherwise `LIST_FORMAT_COMPOSITION_CONFLICT` (`marker-changed` + `sourceMarker`/`modifiedMarker`, `redline-unchanged`, `redline-shape`; action `split_text_and_list_changes` + `hint`). `scope: 'list'` change-format/join registers every live item of the target's list as a write (`rebindSourceIds`): another write there fails `list-rebind` unless it is a compatible redline (decision 2; this refuses the WP-5 "two change-formats of one list in one batch" case, now two applies). No `sameRawText`/`sameListText` in new code. `list-change` still does not auto-apply format changes (decision 3). | `6ace2c7`, `9c3e0e2`; tests `tests/list_format_operation_tests.mjs` section 11 (compiler rule, bare both orders, openDocx labels), matrix kind `list-format-redline-compose`, Word cases `*-list-format-redline-changed-header*` | - |
| 7 | Done except `detach`/`materializeMarker` (deferred, see WP-7 notes): schema composition note, `standalone-operation-runner.d.ts` (`OperationConflict.reason` etc.), README "List operations (`list-format`)", CHANGELOG, `docs/releases/0.9.0.md` draft, `docs/SKILL_AUTHORING.md` list negotiation. convert-plain start values: Word probe says keep `startOverride` (WP-7 notes) | `56cc200` (+ this plan update) | Release: bump version, finalize notes |

Gates every WP must keep green (verified 2026-10-11 after WP-7): `npm test` (146 suites), `check:types`,
`test:isolation`, `node tests/reject_roundtrip_matrix_tests.mjs` (+ `FUZZ_SEED=8 MATRIX_ITERATIONS=500`), Word
synthetic lane 64/64 and corpus lane 60/60 (`pwsh`, redirect output to a log file; see docs/TESTING.md). The fidelity
verifier refuses by default, so a new list operation that breaks Reject All cannot pass these gates.
Source: consumer report `AIWordPlugin/docs/library-issues/2026-09-30-canonical-list-operations.md`
(with `reproduce-plain-header-conversion.mjs`). Reproductions below were re-run against this tree.

## Why this plan exists

The add-in cannot move list work onto the canonical operation path because three things are
missing or misleading.

| Item | Symptom | Class |
|---|---|---|
| (a) Plain paragraph to list item | `redline` "Header" -> "A. Header" returns `ok`, `hasChanges: true`, but the marker stays literal text and `list` is `null`. | Missing capability plus silent success |
| (b) Change list kind or numbering style | `list-change` to a different kind returns `ok` but the source `numId` keeps its old definition. | Missing capability plus silent success |
| (c) Parser diagnostics while generating list OOXML | A strict parser logs a fatal namespace error during the list fallback although the call succeeds. | Bug (fixed in this change set, see below) |

Items (a) and (b) share one root cause: there is no operation whose job is "bind these paragraphs to
this numbering definition". Lists today are produced only as a side effect of text reconciliation
(`engine/*-mode.js` -> `pipeline/list-generation.js`), so anything that is not a text diff
cannot reach the numbering machinery.

## Current code paths

**Historical (pre-WP-2): describes the tree before the numbering allocator and `list-format`.** As built: `mergeNumberingXmlBySchemaOrder` lives in `services/numbering-allocation.js` (call sites in `document/docx-document.js`); the `mergeNumberingXml` callback is removed/ignored; a conflicting id is `NUMBERING_ID_CONFLICT`, not first-writer-wins.

- `services/document-operation-contract.js` accepts `list-change` but
  `getCanonicalOperationType` maps it to `redline`. There is no list-specific operation kind.
- `services/document-operation-mutations.js`, `applyToParagraphByExactText` (around line 2038), tries in order:
  1. `tryExplicitDecimalHeaderListConversion` (line ~1345): only when the source text already
     starts with a decimal marker (`7. Heading`) and the modified text is the same; strips the
     marker with a tracked deletion and binds `w:numPr` with a start override.
  2. `trySingleParagraphListStructuralFallback` (line ~1474): any other marker; converts a
     one-line marker paragraph.
  Both call `buildSingleLineListStructuralFallbackPlan`
  (`orchestration/list-structural-fallback.js:~395`), which returns `null` unless
  `sameRawText || sameListText` (original and modified text equal, or equal marker type and
  normalized content). That gate is exactly why "Header" -> "A. Header" and
  "1. Old heading" -> "1. New heading" fall through to ordinary text reconstruction.
- `mergeNumberingXmlBySchemaOrder` (`services/numbering-helpers.js:366`) merges by id and
  skips any incoming `abstractNum` or `num` whose id already exists. Callers pass it as
  `mergeNumberingXml` to `ensureNumberingArtifactsInZip` (`document/docx-document.js:429,626`).
- `paragraph-format` (`services/document-operation-mutations.js:~2540`) already has the machinery
  we need for tracking: `snapshotAndAttachPPrChange` (`engine/run-builders.js:361`) plus
  `applyParagraphPropertiesToPPr` / `checkParagraphPropertiesChanged` (`engine/rpr-helpers.js`).
  It supports `alignment`, `keepNext`, `keepLines`, `pageBreakBefore`, `style` and nothing about
  `w:numPr`. The schema (`docs/schemas/document-operations.schema.json:163`) leaves `properties`
  an open `object`, so unsupported keys are not rejected.
- `core/list-targeting.js` (`getParagraphListInfo`) and `services/document-inspection.js` expose
  `list.{numId, level, format}` on inspected paragraphs. Those files are under concurrent edit by
  the list-reject work and are not changed here.

## Reproductions

Scratch scripts (not committed): `scratchpad/repro-a.mjs`, `repro-a3.mjs`, `repro-b.mjs`, `repro-c2.mjs`.

### (a) Plain header to list item

```
== plain "Header" -> "A. Header"
 status ok  changes true numParts 0
 text "A. Header" list null warnings []
== plain "Header" -> "1. Header"
 status ok  changes true numParts 0
 text "1. Header" list null warnings []
== "1. Old heading" -> "1. New heading"
 status ok  changes true numParts 0
 text "1. New heading" list null warnings []
```

All three succeed with no numbering part and no list binding. The only conversions that bind a list
are text-equal ones (`"1. Header"` -> `"1. Header"`, `"A. Header"` -> `"A. Header"`).

Two further observations from the same run:

1. For an upper-letter marker the fallback does not convert the paragraph in place. Output is a
   whole-paragraph tracked delete of `A. Header` plus a new inserted paragraph `Header` bound to
   the new `numId`. The decimal explicit path instead edits the same paragraph and strips the
   marker. Reject All on the letter case therefore restores the original, but the two
   representations differ and the paragraph identity (`w14:paraId`, bookmarks, direct formatting)
   is not preserved for letters. Word verification: done in WP-4 (`1240842`, `1cd174a`; convert-plain edits in place).
2. Decimal text-equal conversion on a bare `<w:document>` with `generateRedlines: true` fails with
   `RECEIPT_RECONCILIATION_FAILED: Committed revision id '1000' (kind: del) was not found in
   word/document.xml` (both with and without `atomic`). With `generateRedlines: false` it works and
   binds `numId=2`. Reproduced only on a minimal document without a package; confirm against a
   packaged `.docx` before treating it as a defect. Existing suite coverage passes, so the
   explicit-path receipt wiring should be checked first (WP-0).

   **FIXED in `73c67e8` (2026-10-10).** Root cause: the bare runner did not seed `numberingIdState`, and
   `tryExplicitDecimalHeaderListConversion` recorded receipt revisions before an aborted conversion, so the receipt named
   ids absent from the output. The runner now always seeds the state (`services/batch-operation-orchestrator.js`) and the
   explicit path reserves numbering before redlining (`services/document-operation-mutations.js`). Verified by re-running
   `node tests/phase3_list_structural_fallback_tests.mjs` (PASS, includes the bare `1. Header` regression at line ~155:
   `status: ok`, `numId` bound, every receipt revision id present in the output). Whether it ever reproduced on a packaged
   `.docx` was never separately answered; the fidelity plan's WP-3 matrix will run both paths.

### (b) Format change keeps the old definition

Fixture `tests/fixtures/agentic-lists/nested-lists-source.docx` (Word-authored, untracked at time of
writing, owned by the list-reject work). Paragraphs 3-6 are a bullet list on `numId=1`
(`abstractNum 1`, `numFmt=bullet`). Applying

```json
{ "type": "list-change", "target": {"exactText": "Bullet Root A"},
  "modified": "1. Bullet Root A\n2. Bullet Insertion Anchor" }
```

returns `status: ok`, `applied`. Output: every bullet paragraph is still `<w:numId w:val="1"/>`,
numbering part still has `num 1 -> abstractNum 1` and abstract formats `0:decimal 1:bullet 2:decimal`.
The generated numbering is dropped by the id-keyed merge and nothing signals the no-op.

### (c) Parser diagnostics

```
ERR [XmlAdapter] XML fatal parse error: ... NamespaceError: prefix is non-null and namespace is null
  at extractFirstParagraphNumIdFromOxml (orchestration/list-structural-fallback.js)
  at executeSingleLineListStructuralFallback ...
```

Cause: `pipeline/list-generation.js` returns bare `<w:p>...</w:p>` fragments with no `xmlns:w`.
`extractFirstParagraphNumIdFromOxml` parsed that string as a standalone document. A strict
parser (xmldom) rejects the undeclared prefix, `parseOoxmlSafe` returns `doc: null`, and the
function returns `null`. The generated XML was never malformed; the helper was.

Consequence beyond log noise: with `numId` unresolved, `applyStartOverrideToNumberingXml` was a
no-op, so a marker start value (for example `7.`) was silently dropped on this path. A second
latent defect compounded it: the helper read only the first paragraph, and in tracked mode that is
the deleted-original paragraph, which has no `numPr`.

Status: FIXED (see Changes made). The consumer's "hierarchy" wording is likely their xmldom build's
message for the same malformed-parse class; confirm it is gone against their parser during WP-0.

Not addressed: start values are only derived for decimal markers (`parseMarkerStart` returns `null`
for letters and roman numerals), so `C. Header` is converted but numbers from `A`. Include in WP-3.

## Changes made with this plan (item c)

- `orchestration/list-structural-fallback.js`: new `parseOxmlFragment` wraps namespace-less fragments
  in a `w:root` declaring `xmlns:w` before parsing; used by `getFirstParagraphFromOxml` and
  `extractFirstParagraphNumIdFromOxml`. `extractFirstParagraphNumIdFromOxml` now returns the first
  `numId` found in any paragraph of the fragment.
- `tests/phase3_list_structural_fallback_tests.mjs`: regression that runs the real list pipeline for
  `7. Header`, captures `console.warn`/`console.error`, asserts none were emitted, and asserts
  `w:startOverride w:val="7"` is produced.

## Interim safety fix (historical: shipped in WP-1, `db10d39`)

**Status: DONE in `db10d39` (WP-1; capability `list-conversion-diagnostics-v1`, contract version stays 8).** Implementation notes and deviations
from the text below: strict policy means `strictListConversion: true` (CLI agent profile and `--atomic` set it) or `atomic`; plain
`strictTargets` is always on, so it cannot select the policy. Diagnostics live in `services/list-conversion-diagnostics.js`, appear as
`diagnostics[]` on the operation result plus a `CODE: message` string in `warnings`, and detect only a plain target whose original lacks an
equivalent marker (numbered markers need a `.`/`)` terminator, so `7 days` is not a marker). Format-change detection compares only retained
lines that carry a marker (inserted items are not compared: nested bullets under numbered items are a supported pattern) and covers the
insertion-only and explicit-range heuristics, not the one-line adjacency form (no marker is requested there). `UNSUPPORTED_LIST_FORMAT_CHANGE`
is the strict form of `LIST_FORMAT_NOT_APPLIED`. As built, the recovery actions are `use_list_format_convert_plain` / `use_list_format_change_format`, each with a `suggestedOperation` (not the "apply as plain text or convert in the host" text below). Unknown `paragraph-format` keys ship as the warning `UNKNOWN_PARAGRAPH_FORMAT_PROPERTY`
(error planned for the next minor); the schema's `paragraphFormatProperties` is closed.

Silent success is the worst property of both (a) and (b). Recommend a patch release that adds
detection only, no new conversion behavior:

1. List-marker detection. In `applyToParagraphByExactText`, after both fallbacks decline, if the
   target paragraph is not a list item and the modified text parses as a single-line list
   candidate (`parseSingleLineListCandidate`) while the original text does not carry an equivalent
   marker, do one of:
   - Strict profile (`--profile agent`, `strictTargets`, or `atomic`): fail with
     `UNSUPPORTED_LIST_CONVERSION` and `recovery.action` pointing at the canonical operation once it
     exists (until then, "apply as plain text or convert in the host").
   - Otherwise: succeed but add a structured warning `LIST_MARKER_NOT_BOUND`
     ("modified text begins with a list marker, but the marker was written as literal text; no list
     binding was created"). Surface it in `result.warnings` and the compact CLI receipt.
2. Format-change detection. For `list-change` (and the one-line adjacency form) whose target is
   already a list item, compare the requested kind and format (bullet vs numbered, decimal vs
   upperLetter and so on) with the source `numId` definition. If they differ and the source `num`
   is retained, return `UNSUPPORTED_LIST_FORMAT_CHANGE` (strict) or `LIST_FORMAT_NOT_APPLIED`
   (warning) instead of `ok`.
3. Escape hatch for callers that intentionally want literal text: `listMarkers: 'literal'` on the
   operation suppresses the error and warning. Keep it narrow; do not default it on.

Why error under strict policy: the consumer explicitly wants an "explicit unsupported-operation
error" so it can fall back to its host path. A warning alone requires every caller to parse warnings
to avoid corrupt-looking output. Behavior change: ordinary redlines whose new text happens to start
with `1.` or `A.` (legal headings typed as text) would now warn; hence warn by default and error only
under the strict profile, and the opt-out above.

No `contractVersion` bump for the interim fix; add capability `list-conversion-diagnostics-v1` so the
consumer can negotiate it.

## Contract as built (originally "Proposed operation contract")

Deviations from the proposal below: change-format and range splits write `w:start` on the NEW abstractNum (Word in-session Accept All `startOverride` quirk, `579efae`), while convert-plain keeps `startOverride` (Word probe evidence in `257fa7b`); style-linked lists are refused in both scopes; `detach`/`materializeMarker` are not implemented (use `paragraph-format` `numbering: null`); `scope: 'range'` must run to the end of the list (`LIST_FORMAT_SPLIT_UNSUPPORTED`); `paragraph-format.numbering` validates the numId (`PARAGRAPH_NUMBERING_NOT_FOUND`); sequence continuity is solved by `listFallbackSequenceState`; integrity checks are `validateNumberingIntegrity`, wired into `openDocx`. The text below is the original design and is partly superseded.

### Option chosen: a dedicated `list-format` operation, with a thin `paragraph-format` binding

Extending `paragraph-format` alone is not enough, because numbering needs definition allocation,
sibling continuity, level mapping and (for conversions) marker-text removal. Putting that in an
untyped `properties` bag would hide a cross-part operation behind a paragraph-property setter.

New operation kind `list-format` (canonical name; `type: "list-format"`):

```json
{
  "type": "list-format",
  "target": { "paragraphId": "1A2B3C4D", "fingerprint": "...", "exactText": "Header" },
  "targetEnd": { "paragraphId": "...", "exactText": "..." },
  "list": {
    "mode": "convert-plain | change-format | detach | join",
    "kind": "bullet | numbered",
    "format": "decimal | upperLetter | lowerLetter | upperRoman | lowerRoman | bullet",
    "lvlText": "%1.",
    "start": 1,
    "level": 0,
    "strip": "marker | none",
    "continueFrom": { "paragraphId": "..." },
    "scope": "range | list"
  }
}
```

Semantics:

- `convert-plain`: bind each paragraph in the target range as a list item. `strip: "marker"`
  (default when the text begins with a marker matching `list.format`) removes the literal marker
  with a tracked deletion in the same paragraph, preserving paragraph identity. This is the
  general form of today's decimal-only explicit path and also replaces the whole-paragraph
  delete/insert used for letters. Text changes are not part of this operation: "1. Old heading" ->
  "1. New heading" is expressed as a `redline` (text) plus a `list-format` (structure) in one batch
  on the same target, resolved by the existing batch-start binding (they touch disjoint concerns:
  runs vs `pPr`; the batch compiler must treat `list-format` as a `pPr` writer, like
  `paragraph-format` in `FORMAT_WRITE_KINDS`, `services/operation-batch-compiler.js:15`).
- `change-format`: change kind/format/`lvlText`/`start` of the list the target belongs to.
  `scope: "list"` re-points every paragraph sharing the source `numId`; `scope: "range"` splits the
  list at the range (see allocation).
- `detach`: remove `numPr` (list item to plain paragraph), optionally materializing the visible
  label as literal text (`materializeMarker: true`). Out of the first release.
- `join`: bind to the same `numId` as `continueFrom`. Used to get sibling continuity right.

`paragraph-format` gains a documented, validated `numbering` property that is a pure binding
(no allocation): `{ "numbering": { "numId": 3, "level": 1 } }` or `{ "numbering": null }`. It exists
for callers that already hold a `numId` from `inspect`. Unknown `properties` keys must start to be
rejected (today they are ignored, which is the same silent-success defect class) as part of WP-1.

`list-change` stays as the Markdown-structure operation (insert/remove items, nesting). Its
format-change case routes to the same allocator as `list-format` once available, so the two cannot
diverge.

### Schema and CLI contract

- **Done** (WP-1..WP-7; capabilities `list-format-v1`, `list-format-change-v1`, `paragraph-numbering-binding-v1`, `list-conversion-diagnostics-v1`, contract 9). Original bullets:
- `docs/schemas/document-operations.schema.json`: add a `list-format` branch to the operation union;
  add `numbering` to a closed `paragraph-format.properties` schema (also adds the missing
  `additionalProperties: false` that exposes silently ignored keys).
- `node/cli.js`: bump `CLI_CONTRACT_VERSION` 8 -> 9 when `list-format` ships, and add capabilities
  `list-format-v1` and `paragraph-numbering-binding-v1`. The interim diagnostics ship as
  `list-conversion-diagnostics-v1` at version 8 (additive, no new operation type).
  Per `AGENTS.md`, production harnesses negotiate only the capabilities they use; a wrapper that
  needs `list-format` must check `capabilities.includes('list-format-v1')`, and fall back to its
  host path otherwise. Old CLIs reject the unknown `type` with `INVALID_OPERATION`, which is
  already a safe failure.
- `index.d.ts`, `docs/AGENT_KNOWLEDGE_BASE.md` (new "Lists: choose by intent" table row set under
  "Operation Model"; remove the guidance that suggests `redline` for markers), `README.md`, and
  `CHANGELOG.md` (the plans dir convention). `AGENT_FAST_START.md` should stay unchanged: this is an
  advanced operation (see `AGENTS.md`).

## Numbering definition allocation and remapping

Same-kind edits (landed in `c44f227`, 0.8.3; since `4718c36` this is allocator policy, `allocateNumberingForReplacement` in `services/numbering-allocation.js`, with identical behavior): a `redline` over a list range whose generated list has the same level
formats (bullet-like vs bullet-like, otherwise equal `numFmt`) as the single existing list the
range uses now binds the new paragraphs to that existing `numId` and drops the generated numbering
payload (formerly `reuseSourceListNumbering`). Tracked-deleted source
paragraphs are never remapped, so Reject All restores the source `numPr`. The existing definitions
are read from `numberingIdState.sourceNumberingXml`. Regression:
`tests/facade_list_numbering_tests.mjs`. Format-changing and multi-list ranges still take the
allocation path below and are unchanged; `change-format` (WP-5) should reuse the same
"tracked-deleted source paragraphs keep their numPr" rule when it allocates.

Current limitation: `mergeNumberingXmlBySchemaOrder` is "first writer wins by id", so a generated
definition can never replace one the document already has. This is correct for the common case (do
not overwrite Word's definitions) and wrong for change-format. Do not make the merge overwrite.
Instead never produce colliding ids for a different definition:

1. Allocate, do not replace. `change-format` always creates a new `abstractNum` (copying level
   geometry `w:ind`, `w:rPr`, `w:lvlJc` from the source abstract for levels it does not redefine, so
   indentation and fonts survive) with a fresh `w:abstractNumId`, and a new `w:num`
   with a fresh `w:numId`. Use the document-scoped allocators already introduced for
   `document-scoped-list-revision-ids` (`createDynamicNumberingIdState`,
   `reserveNextNumberingIdPair`; `runtimeContext.numberingIdState`). The incoming ids then never
   collide, and the existing merge inserts them unchanged.
2. Rebind paragraphs. `scope: "list"` rewrites `w:numId` on every paragraph of the source `num` to
   the new one (a `pPrChange` per paragraph, see below). Continuity is by construction: they all
   share the one new `numId`. `scope: "range"` rewrites only the target range; this splits the list
   and must set `w:lvlOverride/w:startOverride` on the new `num` for each level that should
   continue the old count, computed from the count of preceding items of the old `num` at that
   level. Today only level 0 is overridden (`applyStartOverrideToNumberingXml`); generalize to all
   levels present in the range.
3. `start` (historical; as built, change-format/range write `w:start` on the new abstractNum and convert-plain keeps `startOverride`, see "Contract as built"): express with `w:lvlOverride/w:startOverride` on the new `num` (per-instance), never by
   mutating the shared abstract `w:start`. `setAbstractStartOverride` already defaults to
   false on the paragraph path; make that the only behavior for the new operation.
4. Do not delete the old `num`/`abstractNum`. They may still be referenced by other
   paragraphs, styles (`w:style/w:pPr/w:numPr`), headers/footers, or the "before" side of a
   `pPrChange` (Reject All must be able to bind back). Orphaned definitions are harmless to Word.
   If the source `numId` is style-linked (`w:pStyle` in the `abstractNum`), refuse
   `scope: "list"` with a precise error; direct per-paragraph `numPr` still works.
5. Mixed-state ranges: if the target range spans more than one source `numId` or includes
   non-list paragraphs, `change-format` refuses with the paragraph indexes that violate the rule.
   `convert-plain` accepts plain paragraphs only and refuses existing list items (as the plan builder
   does with `allowExistingList: false`).
6. Sequence interplay. `resolveSingleLineListFallbackNumberingAction` and
   `recordSingleLineListFallbackExplicitSequence` carry cross-operation state for "7." then "8."
   header runs. `list-format` ranges with several plain headers must go through the same state so
   two adjacent converts share one `numId` and continue the count; otherwise each header becomes its
   own list restarting at 1. This is the largest behavioral risk.
7. Package plumbing. `ensureNumberingArtifactsInZip` already creates the part, content-type and
   relationship when missing; confirm it for documents that have no `numbering.xml` (WP-2 test).
   `applyOperationsToDocumentXml` consumers receive `numberingXmlParts` and must merge with the same
   function; document that remapped ids are guaranteed not to collide, so a first-writer-wins merge is
   safe.
8. Invariant checks added to the validator (`core/redline-validation.js` or a numbering check):
   every `w:numId` in the body resolves to a `w:num`; every `w:num` resolves to an `abstractNum`;
   `w:lvl` count/levels in the abstract cover every `w:ilvl` used. Run before writing (atomic
   rollback already exists).

## Tracked-change representation

Structure changes must be reviewable and reversible in Word:

- Binding or rebinding `numPr` on an existing paragraph: attach `w:pPrChange` holding the previous
  `w:pPr` (including the previous `numPr`, or none) using `snapshotAndAttachPPrChange`. Reject All
  restores the old `pPr`; Accept All keeps the new one. Revision id from the shared allocator.
  If a paragraph already has a pending `pPrChange` by the same author, keep the original
  snapshot (the helper replaces an existing one; check that it keeps the earliest "before" state;
  add a test). Cross-author pending `pPrChange` follows the existing
  `existingRevisions` policy; default to refuse rather than overwrite another author's snapshot.
- Marker stripping: tracked `w:del` of the literal marker run text (`A. `, `1. `, tab/space),
  exactly like the explicit decimal path. Reject restores the text and, via `pPrChange`, removes the
  numbering; the two reject together because both are authored in one operation by one author.
  Because both Word and this library resolve revisions independently, test that rejecting only one of
  the two (user rejects the text deletion but not the style change) leaves a coherent paragraph
  (literal marker plus bullet is ugly but not corrupt). This caveat belongs in the docs.
- `generateRedlines: false`: write directly, no `pPrChange`, and no marker tracked deletion.
- Paragraph-mark side: do not use `w:rPr/w:ins` paragraph-mark revisions for a pure binding; that
  represents a paragraph insertion, not a property change. Another workstream is changing
  paragraph-mark reject handling; coordinate on `rejectPropertyChangeNode`
  (`services/revision-comment-management.js`) so `numPr` inside `pPrChange` round-trips (the list
  reject fidelity suite `tests/list_reject_fidelity_tests.mjs` is the place; do not duplicate it).
- Whole-paragraph delete/insert (today's letter fallback) must stop being used for conversions.
  It loses `w14:paraId`, bookmarks and direct formatting, and shows up in review as one deleted
  and one inserted line instead of a formatting change.

## Word verification requirements

The report states plainly that offline inspection is not sufficient, and this repo's own history
(`content-types` incident in the 2026-09-29 plan) agrees. Required before shipping WP-4 and later:

1. Author fixtures in desktop Word (not by this library): plain header, header with literal marker
   `A.`/`1.`/`I.`, bullet list, decimal list with a nested level, list with a Word-applied start
   override, style-linked list, and a document with no `numbering.xml`.
   **Status:** all but the last two are covered by Word differential cases. The style-linked and no-`numbering.xml` fixtures are NOT Word-authored: style-linked refusal is covered by unit tests (`tests/list_format_operation_tests.mjs`) and the matrix; the no-numbering-part path by `tests/numbering_allocation_tests.mjs` and package tests. Word-authored versions remain open.
2. Produce the "expected" side in Word by performing the same action by hand with Track Changes on
   (apply numbering to the paragraph; change bullet to decimal via the Define New Multilevel List
   dialog; change numbering format). Commit the resulting `document.xml` `pPr`/`pPrChange`/`numPr`
   structure and the new `numbering.xml` entries as the oracle, like the existing structural-revision
   goldens in `docs/TESTING.md`.
3. Word COM lane (`scripts/word-com-differential.ps1`, `scripts/word-com-visual-suite.ps1`): open each
   output, confirm no repair prompt, render list labels (A., B., C. / 1. 2. 3. / bullets), then
   Reject All (labels and literal text return to the source), Accept All (final labels correct), and
   re-save. Compare label text, not just XML.
4. Confirm list label continuity across a non-list paragraph in the middle of a converted run.
5. Strict-parse lane: run the generator under xmldom (no browser DOM) and assert no parser
   diagnostics (the regression added with item c, extended to `list-format`).

## Test plan

Run focused suites first (`node tests/<suite>.mjs`), then `npm test`, `npm run check:types`
(declarations change), `npm run test:isolation` (new module boundaries).

| Concern | Suite |
|---|---|
| Fallback planner, helper parsing, start values | `tests/phase3_list_structural_fallback_tests.mjs` (extend; c regression already here) |
| Numbering routing and sequence state | `tests/phase3_numbering_routing_tests.mjs` |
| Existing list behavior must not change | `tests/list_tests.mjs`, `tests/list_replacement_structure_tests.mjs`, `tests/multilevel_bullet_tests.mjs`, `tests/performance_phase4_list_and_text_parity_tests.mjs`, `tests/cross_author_slicing_list_tests.mjs` |
| Reject/accept lifecycle of numPr `pPrChange` | `tests/list_reject_fidelity_tests.mjs` (owned by the reject work) and `tests/revision_comment_management_tests.mjs` |
| Operation contract/schema, unknown keys | `tests/agent_operation_contract_tests.mjs` and the schema test, plus `tests/agent_cli_tests.mjs` for `version` capabilities and exit/error shape |
| Package: merge, missing numbering part, part order | `tests/docx_package_facade_tests.mjs`, `tests/docx_package_transaction_edge_tests.mjs` |
| New | `tests/list_format_operation_tests.mjs` (convert-plain, change-format list/range scope, join, refusal cases, id non-collision, per-level startOverride) |
| Interim diagnostics | `tests/list_conversion_diagnostics_tests.mjs` (`UNSUPPORTED_LIST_CONVERSION`, warnings; not `tests/error_contract_tests.mjs`) |
| Allocator, numbering binding | `tests/numbering_allocation_tests.mjs`, `tests/paragraph_format_numbering_tests.mjs` |
| List operations | `tests/list_format_operation_tests.mjs` (convert-plain, change-format, join, composition) |
| Word oracle | `tests/fixtures/agentic-lists/*` Word-authored additions plus the COM lane |

Fixtures: `tests/fixtures/agentic-lists/nested-lists-source.docx` already contains bullet, decimal,
nested, continuation, restart, plain-header candidates and `1. Supported Header ...` paragraphs;
extend it (or add a sibling) with lettered/roman literal headers, a style-linked list, and a doc
without numbering.

Property checks worth adding: after any `list-format` op, (1) every paragraph outside the range is
byte-identical (existing subtree fidelity oracle), (2) Reject All yields a `document.xml` canonical-equal
to the source, (3) `inspect` reports the requested `list.format` and a continuous `list.numId`
for the range.

## Work packages (in order)

| WP | Work | Risk |
|---|---|---|
| 0 | **Done.** Strict-parse and start-value fix (item c); bare-XML `RECEIPT_RECONCILIATION_FAILED` fixed in `73c67e8` (see section (a) observation 2). Remaining: re-check xmldom hierarchy warnings in the consumer's environment; bare-vs-packaged parity is covered by the fidelity plan's WP-3 matrix. | Low |
| 1 | **Done (`db10d39`).** Interim safety diagnostics: `UNSUPPORTED_LIST_CONVERSION`, `LIST_MARKER_NOT_BOUND`, `LIST_FORMAT_NOT_APPLIED`, `listMarkers: 'literal'` opt-out; closed `paragraph-format.properties` schema with unknown-key rejection; capability `list-conversion-diagnostics-v1`. Patch release. | Low-medium: unknown-key rejection could break callers that pass ignored keys; ship as warning first, error in the next minor |
| 2 | **Done (`4718c36`).** Numbering allocator module `services/numbering-allocation.js`: `allocateNumberingDefinition` (create new abstractNum + num, copy level geometry, per-level `startOverride`, ids via `reserveNextNumberingIdPair`), `allocateNumberingForReplacement` (replacement policy, replaces `reuseSourceListNumbering`), the single merge `mergeNumberingParts`, `validateNumberingIntegrity` (numId/abstractNum/ilvl; wired into the `openDocx` validation as introduced-issues-only). No operation exposed yet; `tests/numbering_allocation_tests.mjs`, parity in the WP-3 matrix. Not yet used by an operation: `allocateNumberingDefinition` (WP-4/5 callers), batch continuity (`listFallbackSequenceState`) stays with WP-4. | Medium: continuity across a batch |
| 3 | **Done (`dc3e2d0`).** `paragraph-format.numbering` binding (`numId`, `level`, null) with tracked `pPrChange`; schema-ordered `w:numPr`; existing-definition validation; same-author earliest-snapshot preservation and cross-author conflict refusal; roman/letter start values. Word differential cases verify bind/remove/change-level through Accept All and Reject All list labels. | Low-medium: pPr child order is schema-sensitive and Word repairs misordered children |
| 4 | **Done (`1240842`, `1cd174a`).** `list-format` `convert-plain` with `strip: marker` (decimal, letters, roman, bullets) and `strip: none`, replacing the whole-paragraph letter fallback for text-equal conversions; Word-authored-style fixtures and COM lane for A./1./I. and a C. start value (`ListString` after Accept/Reject); header runs share one `numId` via the sequence state. Contract version 9 and `list-format-v1`. | High: Word fidelity, partial reject coherence, sequence continuity |
| 5 | **Done (`564ad0e`, `579efae`, `616d317`).** `list-format` `change-format` (`scope: list`, `range`), `join`; the `list-change` format-change case is rerouted through recovery (`use_list_format_change_format` + `suggestedOperation`), not auto-applied (see WP-5 notes); refusal rules (style-linked, mixed ranges, mid-sequence splits, nested restarts, shared abstracts, instance level overrides). Word COM lane: bullet -> decimal, decimal -> upper letter, nested level, two range splits, join. Capability `list-format-change-v1`. | High: splitting a list mid-sequence, nested level remap, style-linked numbering |
| 6 | **Done (`6ace2c7`, `9c3e0e2`).** Changed-text header conversion ("1. Old" -> "1. New") = `redline` + `list-format` in one batch; batch-compiler composition rule, fixed order (redline first), `LIST_FORMAT_COMPOSITION_CONFLICT` for incompatible pairs and for other writes to `scope: 'list'` rebound items; Word cases (Accept All label `1.` + new text, Reject All literal `1. Old`, plus a `3.` start value with list-format first in the batch). | Medium. `sameRawText`/`sameListText` are not used by the new code. |
| 7 | **Done (`56cc200`), `detach`/`materializeMarker` deferred.** Docs, schema, d.ts, README, CHANGELOG, release notes draft, wrapper guidance in `docs/SKILL_AUTHORING.md`. | Low |

Ship WP-1 alone as the first patch; it fixes the dishonest success reports without waiting on the
Word-verification work. WP-4 and WP-5 should not merge without the COM lane results attached.

### WP-4 notes for WP-5/6

- Code: `services/list-format.js` (contract shape, marker grammar, level spec, sequence key) and
  `services/list-format-mutation.js` (`convertPlainParagraphsToList`: plan every refusal first, then mutate). Dispatch is
  `applyListFormatToParagraphByExactText` in `document-operation-mutations.js`; its `resolveOperationParagraphRange` is the
  restore-style range resolver (compiled source ids, refs, end descriptor) and should be reused by `change-format`.
- `validateListFormatShape` rejects `scope`, `continueFrom`, `lvlText`, `materializeMarker` and modes other than `convert-plain`
  with `INVALID_OPERATION` naming the field; WP-5 relaxes it per mode. `list.start` is refused with `strip: 'marker'`.
- Allocation: one `allocateNumberingDefinition` call per list group with only the target level redefined (other levels default
  decimal, indent 720*(n+1)/hanging 360; no source geometry because convert-plain has no source list). `change-format` must pass
  `sourceNumberingXml`/`sourceNumId` to keep geometry.
- Letter vs roman: a lone `I.`/`C.`/`D.` is refused `LIST_MARKER_AMBIGUOUS` unless `format` is given or the run's values are
  consecutive under exactly one reading (`C., D.` letters; `I., II.` roman). The text-equal redline route passes the format the
  legacy parser inferred (single letters are letters there), so it never hits the ambiguity refusal.
- Reject coherence: rejecting only the marker deletion leaves literal marker plus label (valid OOXML, tested); rejecting only the
  `pPrChange` leaves the stripped text unnumbered. Documented, not tied together.
- Finding: the legacy whole-paragraph fallback was reject-exact only for plain runs. With the marker in a hyperlink, smart tag,
  custom XML or the author's own pending insertion its Reject All drops the wrapper and the verifier refuses it; convert-plain
  refuses those shapes up front (`reason: marker-shape`). Converting a marker inside the author's own `w:ins` (insert then convert
  in one session) is a candidate improvement (delete-own-insertion semantics), not done.
- Batch compiler (superseded by `6ace2c7`: the composition rule is `checkListFormatComposition`, `list-rebind`, `_compiledRunAfter`, `LIST_FORMAT_COMPOSITION_CONFLICT`; see WP-6 row): `list-format` is a TEXT_WRITE kind, so a `redline` and a `list-format` on one batch-start target failed
  `OVERLAPPING_SOURCE_TARGETS` before WP-6. WP-6 had to add the explicit rule (redline first, list-format strips the marker of the
  redlined text) instead of moving it to FORMAT_WRITE_KINDS, which would allow a text+marker write race.
- Style-linked numbering detection needs `stylesXml` (`runtimeContext.stylesXml` or `options.stylesXml` on the bare runner; the
  facade now passes `word/styles.xml`). Without it a style-numbered paragraph is not detected on the bare path.
- Sequence state lives in `runtimeContext.listFallbackSequenceState.explicitByNumberingKey` under `list-format:<fmt>:<lvl>:<punct>`
  keys; `commitBatchRuntimeContext` copies it to the caller's runtime context, so a bare caller that reuses one context object
  continues lists across calls (same as `numberingIdState`).

### WP-5 notes for WP-6/7

- Code: `services/list-format-change.js` (`changeListParagraphsFormat`, `joinListParagraphs`); dispatch `applyListRebinding` in
  `document-operation-mutations.js` reads definitions from the source part plus the batch's pending `session.numberingXmlParts`,
  so a later operation can change a list an earlier one allocated or rebound. `continueFrom` is compiled in
  `operation-batch-compiler.js` (`_compiledContinueFromId`; a read, not registered as a write source).
- Word numbering model (COM probes; module header): `w:num` instances sharing one `w:abstractNum` count as ONE list (refused
  `shared-definition`); a level start (`w:start` or `w:startOverride`) re-applies on every restart of that level (a split that
  must change the start of a level which restarts inside the range is refused `nested-restart`, found by simulating old vs new
  counters); an item below an unused level implicitly starts that level at its start value (a split that starts at a nested item
  sets the ancestors' start to their current value; Word-verified).
- Word quirk found by the lane: after `Revisions.AcceptAll()` in a session, a level carrying a `w:startOverride` that resumes
  after deeper items is renumbered from the override (`B., B.` instead of `C., D.`); reopening shows `C., D.`. A `w:start` on an
  abstract used by one instance only renders correctly in both, so change-format writes every start there. convert-plain (WP-4)
  still uses `startOverride` on single-level lists, where the quirk does not arise; WP-7 could switch it for consistency.
- Batch compiler: `scope: 'list'` writes `pPr` on paragraphs outside the target that the compiler does not register, so another
  operation on such a paragraph in the same batch is not reported as a conflict; it runs in order and sees the rebound paragraph
  (a same-author `pPrChange` keeps the earliest snapshot, so Reject All stays exact). Superseded by `6ace2c7`: WP-6 registers these
  out-of-target writes (`rebindSourceIds`); another write there fails `list-rebind` unless it is a compatible redline.
- Untracked (`generateRedlines: false`) rebinding outside the target is refused (`INVALID_OPERATION`, field `generateRedlines`);
  the verifier's untracked locality check would refuse it anyway. A `target`..`targetEnd` range covering the list works.
- Not done: the `list-change` path does not apply a format change itself (a Markdown edit mixes text and structure; WP-6's
  redline + list-format composition is the place to offer it). Style-linked lists (heading numbering) are refused in both scopes,
  including direct-`numPr` ranges the plan allowed, because style-bound members are invisible to the split/continuation analysis.
  Text-box list members are refused. `join` of plain paragraphs is refused (`LIST_FORMAT_TARGET_NOT_LIST`; use
  `paragraph-format.numbering`). `inspect` labels for the implicit-level case come from the library's own label computation and
  were compared with Word only in the COM lane cases.

### WP-6 notes

- Decisions implemented (parent, 2026-10-10): (1) the redline + list-format pair is allowed under the rule above, bound at
  batch start, executed redline first; (2) `scope: 'list'` writes are registered, fail closed; (3) `list-change` keeps
  pointing at `change-format` through recovery (KB, README, CHANGELOG, SKILL_AUTHORING say so); (4) start values, see WP-7
  notes.
- Order is enforced through the dependency plan (`_compiledRunAfter` edges), so results and receipts keep the batch
  order while `executionOrder` shows the redline first. If the redline fails at run time the list-format still runs on
  the original paragraph (valid alone); under `atomic` the batch rolls back.
- The redline's reconstruction leaves `"1. "` as a plain run before the changed text, so convert-plain's marker locator
  finds it; if a future diff change put the marker inside the `w:ins`/`w:del`, convert-plain refuses `marker-shape`
  rather than guessing.
- Finding (pre-existing, not list work): a plain `redline` reconstruction rewrites the `w:p` and drops `w14:paraId`; the
  composed pair therefore does not keep paragraph identity while a lone convert-plain does.
- Comments are not writes and never conflict with rebinds; paragraphs outside the list are free.

### WP-7 notes

- convert-plain start values (decision 4), Word 16 probe 2026-10-11 (scratch, not committed): typed headers `3./4./5.`
  and `C./D./E.` converted, then nested items bound to the SAME new numId at level 1 by a second tracked apply
  (`paragraph-format.numbering`), and a variant with items of another list between headers. Labels read with
  `ListFormat.ListString` on open, after `Revisions.AcceptAll()` in the session, after save + reopen, and after
  `RejectAll()`: `3., 1., 2., 4., 1., 5.` / `C., 1., 2., D., 1., E.` in every phase with `w:startOverride`, identical to a
  `w:start` rewrite of the same files; Reject All shows no labels. Not affected, so convert-plain keeps `startOverride`
  (single-level instance; the WP-5 quirk needs a split instance whose level resumes after deeper items). The Word lane
  case `administrative-list-format-redline-changed-header-start-value` covers a `3.` start in-session.
- `detach` / `materializeMarker` deferred: (a) removing a direct binding is already `paragraph-format`
  `properties.numbering: null` (WP-3, Word-verified), so a mode adds contract surface, not capability; (b) a correct
  detach must also handle style-numbered items (direct `numPr` removal falls back to style numbering; needs `numId 0`)
  and the renumbering of later items, which is a policy choice; (c) `materializeMarker` needs Word-exact label text
  (multi-level `lvlText`, `isLgl`, bullet glyphs in Symbol/Wingdings private-use code points, picture bullets,
  `w:suff` tab/space/nothing), which the library computes itself and has compared with Word only in lane cases. Not small;
  recommend a later release limited to single-level decimal/letter/roman labels, with Word lane cases.

### Open after this plan

- `detach`/`materializeMarker` (above); style-linked and text-box lists stay refused.
- Release: bump `package.json`, finalize `docs/releases/0.9.0.md`, tell AIWordPlugin about refuse-by-default, atomic
  typed-marker refusal, renamed recovery actions and `list-rebind` refusals.
- Promote `UNKNOWN_PARAGRAPH_FORMAT_PROPERTY` to an error in the next minor (WP-1).

## Risks and open questions

Status 2026-10-11: items marked **Resolved** are settled; the rest are open.

- Splitting a list (`scope: range`) changes numbering for paragraphs after the range. Decide whether
  following siblings keep their old `num` (visible renumbering restart) or are also rebound
  (touches paragraphs the caller did not name). Recommendation: keep them on the old `num` and
  report it in the receipt; the caller can extend the range. **Resolved in WP-5:** refused instead
  (`LIST_FORMAT_SPLIT_UNSUPPORTED` `mid-sequence`, naming `lastListParagraphIndex`); a range must run to the end of the list.
- Reject coherence when a reviewer rejects only the marker deletion or only the `pPrChange`.
  Document; do not attempt to tie them with a shared revision id (Word does not support that). **Resolved (WP-4):** documented and tested, not tied.
- Style-linked numbering (`w:numPr` in `w:style`) cannot be rebound per paragraph without
  overriding the style; first release refuses it. **Resolved:** refused in both scopes (`LIST_FORMAT_STYLE_LINKED`, `LIST_FORMAT_STYLE_NUMBERED`); no Word-authored style-linked fixture yet (open).
- Headers/footers and comments parts have their own `document.xml`-adjacent content (see the
  2026-09-29 plan). First release is body-only; reject `part` targets with an explicit error. **Resolved:** body-only; `part` targets are refused.
- Shared in-flight work: `services/document-inspection.js` and the paragraph-mark/list reject paths
  are being edited elsewhere. WP-3 onward touches `rejectPropertyChangeNode` expectations; sequence
  after that work lands. **Resolved:** that work landed (fidelity plan); no longer a constraint.
- Whether `list-format` should auto-detect `strip` from the text or require it. Recommendation: auto
  for exact `^\s*marker[.)]?\s` matches of the declared `format`, otherwise leave text alone
  and say so in the receipt. **Resolved (WP-4):** `strip: 'marker'` is the default and auto-detects the marker; `strip: 'none'` opts out.
- Still open: unknown `paragraph-format` keys become an error in the next minor; converting a marker inside the author's own `w:ins` is refused (`marker-shape`), not handled; xmldom parse noise in the consumer environment is unconfirmed; Word-authored style-linked and no-numbering-part fixtures do not exist.

## Cross-plan dependencies (added 2026-10-10, updated 2026-10-11)

Companion: `2026-09-30-revision-fidelity-invariants.md` (the fidelity plan). Landed there so far: WP-1 verifier
(`services/revision-fidelity-verifier.js`; `05dcf4c`, `c933dcd`, `bbb23cb`), WP-2 Word-semantics alignment
(`1594add`, `584c83a`, `cedc68a`), WP-4 `core/current-properties.js` (`1a3eeb6`..`1f89ae0`), WP-5b (`73c67e8`), and the
numPr-without-`pPrChange` fix for explicit-start conversion (`f72f60a`). Consequences for the work packages here:

1. **Acceptance gate.** The post-apply verifier compares Word Reject All of the output with Word Reject All of the source;
   a new `list-format` / `change-format` / `paragraph-format.numbering` operation must leave them equal. The verifier refuses by
   default (`enforce`; refusal `REVISION_FIDELITY_VIOLATION`, disposition `refused`, per-invariant `violations[]` with
   `explanation`, source paragraph index/excerpt and `recovery.retryGuidance`), so a new operation that breaks Reject All
   cannot ship silently. This replaces "Reject All yields
   canonical-equal `document.xml`" as an ad hoc property check in the Test plan. Untracked (`generateRedlines: false`)
   operations are skipped by the verifier.
2. **Allocator requirements (WP-2 here; R1-R3 satisfied by `4718c36`; `reuseSourceListNumbering` replaced by allocator policy).** The allocator must satisfy: R1, tracked-deleted source paragraphs never change
   `numPr` (verifier invariant `REJECT_NUMBERING_REFS`, not a convention; already true of same-kind reuse); R2, the runner
   and the `openDocx` facade call one allocator function with the same arguments and neither carries its own merge; R3, the
   same-kind reuse heuristic `reuseSourceListNumbering` (`services/numbering-helpers.js`, call site
   `services/document-operation-mutations.js`) was replaced by allocator policy in `4718c36`; `tests/facade_list_numbering_tests.mjs` still guards the behavior.
3. **Reading properties.** Any new code that reads `numPr`, `pStyle`, `outlineLvl` or other paragraph/run properties must use
   `core/current-properties.js` (`getCurrentNumIdElements` etc.), never raw `w:pPr` descendants, which include `w:pPrChange`
   snapshots. `tests/current_properties_usage_tests.mjs` lints this. This closes the historical-`numPr` inspection
   concern; do not re-fix it here.
4. **Matrix.** The fidelity plan's WP-3 (landed, `f5b485d`: `tests/reject_roundtrip_matrix_tests.mjs`,
   `tests/helpers/operation-generators.mjs`) runs each operation on bare `document.xml` and packaged `.docx` paths and
   asserts matching paragraph texts and per-paragraph `numPr` presence in both views. Add `list-format` /
   `change-format` operation kinds to the generator when WP-4/5 here land.
5. **Paragraph marks.** Paragraph-mark revisions (`w:pPr/w:rPr/w:ins|del`) must be built only through
   `engine/paragraph-mark-revisions.js` (fidelity plan WP-6, `c356c5c`..`172fd56`; enforced by
   `tests/paragraph_mark_single_site_tests.mjs`). `pPrChange` snapshots for pure numbering binding stay with
   `snapshotAndAttachPPrChange`, and the binding helper `enforceListBindingOnParagraphNodes` now takes
   `trackPropertyChange` (`f72f60a`).
6. **Word checks.** Use `Document.Revisions.AcceptAll()/RejectAll()` in any new COM script, never
   `Document.AcceptAllRevisions()/RejectAllRevisions()`, which can leave the final paragraph mark's revision behind
   (fidelity plan, "Word verification"); `scripts/word-com-differential.ps1` also compares Reject All paragraph counts.

7. **Status 2026-10-11.** Every list operation here passes the enforce verifier, the matrix (kinds
   `list-format-convert-plain`, `-change-format`, `-join`, `-redline-compose`) and the Word lanes; nothing in this plan
   is blocked on the fidelity plan. Its open items (`ACCEPT_INTENT`, numbering-part existence check, idempotence
   property) would also cover list operations when they land.

### Conflicts and stale statements to resolve

- (Resolved in WP-3.) `numPr` is schema-ordered in the current `w:pPr`, while the original binding is retained in the
  `w:pPrChange` snapshot. Reject/accept behavior is covered by the Word-semantics simulator and Word differential list-label
  assertions.
- "Whole-paragraph delete/insert must stop being used for conversions" (letter markers): still a valid design goal, but the
  current letter form is reject-exact, so the verifier will not force the change. **Resolved by `1240842` (WP-4):** convert-plain edits in place, justified on identity and review quality, with Word lane cases.
- (Resolved) The fidelity plan's bare `1. Header` row now records the `73c67e8` fix.
- The "Interim safety fix" diagnostics are independent of the verifier modes; do not conflate `LIST_MARKER_NOT_BOUND` and
  friends with `REVISION_FIDELITY_VIOLATION`.
