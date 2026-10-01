# Canonical List Operations: Header-to-List Conversion and List Format Change

Date: 2026-09-30
Status: PROPOSED. Package baseline 0.8.2, CLI `contractVersion` 8.
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
   is not preserved for letters. Needs Word verification (WP-6).
2. Decimal text-equal conversion on a bare `<w:document>` with `generateRedlines: true` fails with
   `RECEIPT_RECONCILIATION_FAILED: Committed revision id '1000' (kind: del) was not found in
   word/document.xml` (both with and without `atomic`). With `generateRedlines: false` it works and
   binds `numId=2`. Reproduced only on a minimal document without a package; confirm against a
   packaged `.docx` before treating it as a defect. Existing suite coverage passes, so the
   explicit-path receipt wiring should be checked first (WP-0).

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

## Interim safety fix (recommended, small, ship first)

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

## Proposed operation contract

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

Same-kind edits (landed): a `redline` over a list range whose generated list has the same level
formats (bullet-like vs bullet-like, otherwise equal `numFmt`) as the single existing list the
range uses now binds the new paragraphs to that existing `numId` and drops the generated numbering
payload (`reuseSourceListNumbering` in `services/numbering-helpers.js`). Tracked-deleted source
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
3. `start`: express with `w:lvlOverride/w:startOverride` on the new `num` (per-instance), never by
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
| Interim diagnostics | new cases in `tests/error_contract_tests.mjs` for `UNSUPPORTED_LIST_CONVERSION` and the warning |
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
| 0 | Done in this change: strict-parse fix and start-value fix (item c). Still to do: confirm the decimal explicit-path `RECEIPT_RECONCILIATION_FAILED` with tracked changes on a packaged `.docx`; if real, fix before anything else because it is in the "supported" path. Re-check xmldom hierarchy warnings in the consumer's environment. | Low |
| 1 | Interim safety diagnostics: `UNSUPPORTED_LIST_CONVERSION`, `LIST_MARKER_NOT_BOUND`, `LIST_FORMAT_NOT_APPLIED`, `listMarkers: 'literal'` opt-out; closed `paragraph-format.properties` schema with unknown-key rejection; capability `list-conversion-diagnostics-v1`. Patch release. | Low-medium: unknown-key rejection could break callers that pass ignored keys; ship as warning first, error in the next minor |
| 2 | Numbering allocator module (new `services/numbering-allocation.js`): create-new-abstract-and-num, copy level geometry, per-level `startOverride`, id reservation through the document-scoped state; validator checks for numId/abstractNum/ilvl integrity. No operation exposed yet; unit tests only. | Medium: continuity across a batch |
| 3 | `paragraph-format.numbering` binding (`numId`, `level`, null) with `pPrChange`; extend `applyParagraphPropertiesToPPr`, `checkParagraphPropertiesChanged`, and `insertPPrChildInOrder` for `w:numPr` position (before `w:ind`, after `w:keepLines`/`w:pageBreakBefore` per the CT_PPr sequence); roman/letter start values (`parseMarkerStart`). | Low-medium: pPr child order is schema-sensitive and Word repairs misordered children |
| 4 | `list-format` `convert-plain` with `strip: marker` (decimal, letters, roman, bullets), replacing the whole-paragraph letter fallback; Word-authored fixtures and COM lane for A./1./I.; header runs share one `numId` via the sequence state. Contract version 9 and `list-format-v1`. | High: Word fidelity, partial reject coherence, sequence continuity |
| 5 | `list-format` `change-format` (`scope: list` then `range`), `join`; reroute the `list-change` format-change case; refusal rules (style-linked, mixed ranges). | High: splitting a list mid-sequence, nested level remap, style-linked numbering |
| 6 | Changed-text header conversion ("1. Old" -> "1. New"): documented as `redline` + `list-format` in one batch; add the batch-compiler rule that these two kinds on one target do not conflict; remove reliance on `sameRawText`/`sameListText` for new code. | Medium |
| 7 | `detach` and `materializeMarker` (optional), docs, schema, `index.d.ts`, README, CHANGELOG, release notes; wrapper guidance in `docs/SKILL_AUTHORING.md` (negotiate `list-format-v1`; otherwise keep the host path). | Low |

Ship WP-1 alone as the first patch; it fixes the dishonest success reports without waiting on the
Word-verification work. WP-4 and WP-5 should not merge without the COM lane results attached.

## Risks and open questions

- Splitting a list (`scope: range`) changes numbering for paragraphs after the range. Decide whether
  following siblings keep their old `num` (visible renumbering restart) or are also rebound
  (touches paragraphs the caller did not name). Recommendation: keep them on the old `num` and
  report it in the receipt; the caller can extend the range.
- Reject coherence when a reviewer rejects only the marker deletion or only the `pPrChange`.
  Document; do not attempt to tie them with a shared revision id (Word does not support that).
- Style-linked numbering (`w:numPr` in `w:style`) cannot be rebound per paragraph without
  overriding the style; first release refuses it.
- Headers/footers and comments parts have their own `document.xml`-adjacent content (see the
  2026-09-29 plan). First release is body-only; reject `part` targets with an explicit error.
- Shared in-flight work: `services/document-inspection.js` and the paragraph-mark/list reject paths
  are being edited elsewhere. WP-3 onward touches `rejectPropertyChangeNode` expectations; sequence
  after that work lands.
- Whether `list-format` should auto-detect `strip` from the text or require it. Recommendation: auto
  for exact `^\s*marker[.)]?\s` matches of the declared `format`, otherwise leave text alone
  and say so in the receipt.
