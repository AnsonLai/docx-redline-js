# Changelog

## Unreleased

### Changed

- **Unknown `paragraph-format.properties` keys are rejected:** `INVALID_OPERATION` with `field: properties.<key>`, `properties[]`
  and `recovery.action: change_request`; nothing is applied. The 0.8.x `UNKNOWN_PARAGRAPH_FORMAT_PROPERTY` warning no longer fires.
- **CLI contract version 9** (`list-format` operation type, capability `list-format-v1`). Integrations that do not send
  `list-format` can keep requiring version 8.
- **`UNSUPPORTED_LIST_CONVERSION` / `LIST_MARKER_NOT_BOUND` now point at `list-format`:** `recovery.action` is
  `use_list_format_convert_plain` (was `apply_plain_text_or_convert_in_host`) and, when the original paragraph has no marker,
  `recovery.suggestedOperation` carries the `list-format` request to send with the same target (for "Header" -> "C. Header":
  `{ mode: 'convert-plain', format: 'upperLetter', strip: 'none', start: 3 }`). `listMarkers: 'literal'` is unchanged.
- **`UNSUPPORTED_LIST_FORMAT_CHANGE` / `LIST_FORMAT_NOT_APPLIED` now point at `list-format` `change-format`:** `recovery.action`
  is `use_list_format_change_format` (was `change_list_format_in_host`) and `recovery.suggestedOperation` carries
  `{ type: 'list-format', list: { mode: 'change-format', format, level, scope: 'list' } }`; add a target on one item of the list.
  `list-change` still never applies a format change itself.
- **A `list-format` with `scope: 'list'` counts as a write to every item it rebinds:** another write to one of those items in the
  same batch (`paragraph-format`, `highlight`, `format`, a second `list-format` of the same list) now fails before mutation with
  `LIST_FORMAT_COMPOSITION_CONFLICT` (`reason: list-rebind`); in WP-5 it ran in batch order. Send it in a separate apply or use
  `scope: 'range'`. A single-paragraph `redline` of an item still composes (see Added).
- **Text-equal letter, roman and bullet header conversions are made in place:** a `redline` whose `modified` equals the plain
  paragraph's typed-marker text (`B. Header` -> `B. Header`) now strips the marker as a tracked deletion and binds the list with a
  tracked `w:pPrChange` in the same paragraph, like the decimal path. The previous whole-paragraph delete plus insert lost
  `w14:paraId`, bookmarks and direct formatting and reviewed as a deleted and an inserted line; it remains only for shapes the
  in-place conversion refuses. (It was reject-exact for plain runs but not for markers inside hyperlinks or smart tags, where the
  Reject All verifier refuses it.)
- **Operations that would break Reject All are refused by default:** the revision fidelity verifier now runs in
  `enforce` mode. A refused operation leaves the document unchanged and returns `REVISION_FIDELITY_VIOLATION` with a
  one-line message naming the invariant, the source paragraph (extraction index and excerpt), expected vs actual, an
  explanation and a retry instruction; `violations[]`, `invariants` and `recovery.retryGuidance` carry the same detail for
  agents, the message is also sent to `onWarn`, and compact CLI output keeps it. Opt out per call with
  `verifyRevisionFidelity: 'warn'` (apply and warn) or `'off'` (diagnosis). CLI: `--revision-fidelity <enforce|warn|off>`,
  shown in `effectiveOptions`; the agent profile rejects `off`. Capability `revision-fidelity-v1`. A single-target untracked
  operation (`generateRedlines: false`) is also refused when it changes a paragraph outside its target in the Accept All or
  Reject All view (invariant `ACCEPT_NEIGHBORS`).
- **`ensureNumberingArtifactsInZip` always merges numbering, and refuses id conflicts:** the optional `mergeNumberingXml`
  callback is ignored (the single merge is `mergeNumberingParts` in `services/numbering-allocation.js`), an existing
  `numbering.xml` is no longer replaced when no callback is given, and an incoming `w:abstractNum` / `w:num` whose id exists with
  different content throws `NUMBERING_ID_CONFLICT` instead of silently binding new paragraphs to another list. `openDocx` also
  rejects output that introduces a dangling `numId`, `abstractNumId` or `ilvl` (`NUMBERING_*` issues).

### Added

- **`redline` + `list-format` on one paragraph in one batch:** a changed-text header (`1. Old` -> `1. New`) is a `redline` that keeps
  the typed marker exactly (`modified: '1. New'` or `replacements`) plus a `list-format` `convert-plain` of the same paragraph. The
  batch compiler allows exactly this pair (both bind at batch start) and fixes the order whatever the batch order: the redline runs
  first, then the list-format strips the unchanged marker, so Accept All shows label `1.` with `New` and Reject All the literal
  `1. Old` (verified in Word with `ListFormat.ListString`, including a `3.` start value). A single-paragraph `redline` of a list item
  also composes with a `change-format`/`join` that rebinds it. Incompatible pairs are refused before mutation with
  `LIST_FORMAT_COMPOSITION_CONFLICT` (`reason`: `marker-changed` with `sourceMarker`/`modifiedMarker`, `redline-unchanged`,
  `redline-shape`; `recovery.action: split_text_and_list_changes`); other kinds keep `OVERLAPPING_SOURCE_TARGETS` /
  `REVISION_ORDER_CONFLICT`. No new capability: it is part of `list-format-v1` behavior.
- **`list-format` `change-format` and `join` modes (`list-format-change-v1`, contract version stays 9):** `change-format` changes
  the kind or numbering format of an existing list (`format`, `lvlText`, `start`, `level`, default the target's level). It never
  edits the shared definition: a NEW `w:abstractNum` + `w:num` is allocated with the source levels, indentation and fonts copied
  and one level redefined (`w:isLgl`, picture bullets and a bullet's symbol font are dropped on a format change), and each item is
  rebound with a tracked `w:pPrChange`, so Reject All restores the original labels and Accept All shows the new ones (verified in
  Word with `ListFormat.ListString` after `Revisions.AcceptAll()`/`RejectAll()`: bullets to `1.`, `1.` to `A.`, a nested level to
  `(i)`, and two range splits). `scope: 'list'` (default) rebinds every live item of the list (`items[].inTarget` marks those outside
  the target); `scope: 'range'` rebinds a `target`..`targetEnd` range that runs to the end of the list and continues its count
  (start values on the new definition, including a split that starts at a nested item). `join` binds the target list (or a range
  running to its end) to the `numId` of `list.continueFrom`, an item of an earlier list, so it continues that count; `continueFrom` is
  bound at batch start. Tracked-deleted paragraphs keep their `w:numPr`; a same-author earlier `w:pPrChange` snapshot is kept;
  later operations in a batch see definitions earlier ones allocated. Refused before any write: `LIST_FORMAT_TARGET_NOT_LIST`,
  `LIST_FORMAT_MIXED_RANGE` (plain paragraphs, several lists, or several levels without `list.level`), `LIST_FORMAT_STYLE_LINKED`
  (heading/list-style numbering), `LIST_FORMAT_STYLE_NUMBERED`, `LIST_FORMAT_SPLIT_UNSUPPORTED` (`mid-sequence` range,
  `nested-restart`, `shared-definition` abstract used by another list, `interleaved` join), `LIST_FORMAT_DEFINITION_UNSUPPORTED`
  (missing definition, `w:lvlOverride/w:lvl`, `w:lvlRestart` in a split), `LIST_FORMAT_JOIN_INVALID` and
  `PARAGRAPH_FORMAT_REVISION_CONFLICT`; an untracked rebinding outside the target is refused; a no-op is `ok` with a
  `LIST_FORMAT_UNCHANGED` warning. Results carry `list` (`sourceNumId`, `numId`, `items[]`, `levelStarts` or `continueFrom`).
  An older CLI refuses the new modes with `INVALID_OPERATION` (`field: list.mode`), so negotiate the capability.
- **`list-format` operation, `convert-plain` mode (`list-format-v1`):** binds plain paragraphs (a target, or `target`..`targetEnd`
  in one container) whose text starts with a typed marker (`1.`/`1)`, `A.`/`a)`, `IV.`/`iv.`, `-`/`•`) to a NEW list definition in
  place. The marker is removed with a tracked `w:del` of exactly the marker text (`strip: 'marker'`, default) and the `w:numPr`
  binding is a tracked `w:pPrChange`, so Reject All restores the literal marker and the unnumbered paragraph and Accept All shows
  the list label (verified in Word: `ListFormat.ListString` after `Revisions.AcceptAll()`/`RejectAll()` for `A.`, `1.`, `I.` and a
  `C.` start value). Each marker's value is the item's value (`startOverride` on the new `w:num`; the label keeps the typed `.`/`)`);
  consecutive values share one list, a gap starts a new list at the typed value, body paragraphs in the range stay plain, and a
  later `list-format` in the same batch whose first value continues the previous list joins it. `strip: 'none'` with `format`
  (and optional `start`) binds without touching the text. Ambiguous or unsupported shapes are refused with recovery envelopes:
  `LIST_MARKER_NOT_FOUND`, `LIST_MARKER_AMBIGUOUS` (a lone `I.`/`C.` reads as letter or roman; set `list.format`),
  `LIST_MARKER_UNSUPPORTED` (mixed markers, `2.1.`, `(a)`, a marker inside a hyperlink/field/revision, format mismatch),
  `LIST_FORMAT_TARGET_IS_LIST`, `LIST_FORMAT_STYLE_NUMBERED` (style-linked numbering; needs `stylesXml` on the bare runner, read from
  the package by `openDocx`) and `PARAGRAPH_FORMAT_REVISION_CONFLICT`. Body only. Results carry `list` (`format`, `items[]` with
  `numId`/`value`, `skipped[]`). Schema `$defs.listFormat`, `ListFormatDocumentOperation` in `index.d.ts`.
- **`paragraph-format.numbering` binding (`paragraph-numbering-binding-v1`, no contract bump):** `properties.numbering` is
  `{ numId, level? }` (bind the paragraph to an existing list) or `null` (remove its direct list binding), recorded as a tracked
  `w:pPrChange` so Reject All restores the previous `w:numPr` exactly and Accept All shows the new one (verified in Word with
  `Document.Revisions.RejectAll()/AcceptAll()` and `ListFormat.ListString`). `w:numPr` is inserted at its CT_PPr position
  (after `widowControl`, before `suppressLineNumbers`/`spacing`/`ind`; `ilvl` before `numId`). A binding never allocates: an unknown
  `numId` or undefined level is refused `PARAGRAPH_NUMBERING_NOT_FOUND`. `paragraph-format` now keeps the earliest `w:pPrChange`
  snapshot for repeated same-author edits (it used to replace it, so Reject All restored the intermediate state) and refuses to
  overwrite another author's pending snapshot (`PARAGRAPH_FORMAT_REVISION_CONFLICT`). Letter and roman literal-marker headers
  (`C.`, `iv.`) now convert with their start value instead of restarting at the first item.
- **List conversion diagnostics (`list-conversion-diagnostics-v1`, no contract bump):** text edits no longer report a silent `ok`
  when they cannot create list structure. A plain paragraph whose new text starts a list marker returns the warning
  `LIST_MARKER_NOT_BOUND` (`diagnostics[]` on the operation result, `CODE: message` in `warnings`), or `UNSUPPORTED_LIST_CONVERSION`
  under the strict policy (`--profile agent`, `--atomic`, or `strictListConversion: true`). A list item whose retained marker asks for a
  different kind or format than its numbering definition returns `LIST_FORMAT_NOT_APPLIED` / `UNSUPPORTED_LIST_FORMAT_CHANGE`.
  `listMarkers: 'literal'` on an operation opts out. `paragraph-format.properties` is now a closed schema; unknown keys (such as
  `spacingAfter`) first produced an `UNKNOWN_PARAGRAPH_FORMAT_PROPERTY` warning (superseded in 0.9.0, see Changed).
  Independent of `REVISION_FIDELITY_VIOLATION`; compact CLI output keeps `diagnostics` and the error fields.

- **Numbering allocator:** new `services/numbering-allocation.js` (exported from the package root): `allocateNumberingDefinition`
  (new abstractNum + num with fresh ids, copied level geometry, per-level `startOverride`), `allocateNumberingForReplacement`
  (same-kind source-list continuation or fresh allocation; tracked-deleted source paragraphs never change `numPr`),
  `mergeNumberingParts` / `mergeNumberingXmlBySchemaOrder` (now with an `onConflict` option), `canonicalNumberingDefinition` and
  `validateNumberingIntegrity`. The standalone runner's consumers and the `openDocx` facade merge through the same function;
  `reuseSourceListNumbering` is replaced by allocator policy.
- **Revision fidelity verifier:** every tracked operation now checks, on the live DOM and without extra
  serialization, that Word's Reject All of its output equals Word's Reject All of the source (paragraph count and order,
  canonical text including breaks and tabs, paragraph properties, `numPr`, hyperlink/field/content-control membership).
  A mismatch refuses the operation (see Changed); in `warn` mode it is applied with the message in `warnings` and
  `revisionFidelityViolations`. The Word-semantics simulator is `core/word-revision-semantics.js`, validated against
  recorded Word 16 Reject All / Accept All results (`tests/fixtures/word-revision-semantics/expectations.json`).

- **`UNTRACKED_CREATED_REVISION` invariant:** an operation with `generateRedlines: false` is refused (default `enforce`) when its
  output contains a revision id (`w:ins`, `w:del`, moves, `*PrChange`, paragraph-mark and row revisions) that the source did not
  have. A foreign revision sliced around the edit keeps its element, author and date and is not counted. It reports as
  `REVISION_FIDELITY_VIOLATION` with `violations[].invariant: 'UNTRACKED_CREATED_REVISION'`, also for range targets.

### Internal

- **One paragraph-mark revision helper:** `engine/paragraph-mark-revisions.js` is now the only module that constructs a tracked paragraph mark (`w:ins`/`w:del` under `w:pPr/w:rPr`): `markParagraphInserted`, `markParagraphDeleted`, `setParagraphPropertiesMarkRevision`, `restoreParagraphMarkInserted`, string builders `paragraphMarkRevisionXml` / `paragraphMarkPPrXml`, and `planParagraphMarks` (asserts the Reject All paragraph-count invariants). Document-level shortcuts, list generation, reconstruction, text-to-table and generated tables use it, and `tests/paragraph_mark_single_site_tests.mjs` fails on any other construction site. Output is unchanged except that a new mark `rPr` is placed before `sectPr`/`pPrChange`, an existing mark is replaced rather than duplicated, and the empty-modified shortcut's deleted-mark receipt kind is `del` (was `structural`).

### Fixed

- **Plain `redline` keeps `w14:paraId` / `w14:textId`:** reconstruction used to rebuild the paragraph without its identity 
  attributes; the rebuilt paragraph that survives for a source paragraph (not tracked-deleted) now keeps them, once (an inserted
  split gives them to the paragraph whose mark Reject All keeps). A composed `redline` + `list-format` pair keeps identity.
- **Text-changed header conversion with direct `keepNext` / spacing is no longer refused `REJECT_PARAGRAPH_PROPS`:** the list
  fallback's tracked-deleted source paragraph keeps its original `pPr`; the new list paragraph carries the direct properties and
  the source `w14:paraId`.
- **`ReconciliationPipeline.executeTableGeneration` inserts a tracked table block:** the public pipeline method generated
  tracked cell text in untracked rows (the defect `engine/table-mode.js` had); rows now carry `w:trPr/w:ins` so Reject All
  removes the table.
- **Separate list conversions in one batch continue one list:** two `redline` operations that turned non-adjacent headers into continuing markers (`A. Header`, `B. Header`) each allocated their own numbering instance, so Word showed `A.`, `B.` (start values) but as two lists, and a later item inserted between them would not renumber. The structural list fallback now uses the batch sequence state `list-format convert-plain` uses: same kind, format and level with a continuing value reuses the previous operation's live `numId` (a `redline` and a `convert-plain` of the same format join one list too); a gap (`A.` then `C.`) or a different format allocates a new instance. Ids still come from `services/numbering-allocation.js`. (Reported against 0.8.4.)
- **Weak-target localized replacements no longer delete soft-break paragraphs:** a `replacements` edit targeted only by `index` + `exactText` (no `paragraphId`/`fingerprint`) in a paragraph containing `<w:br/>` returned `ok` but tracked the whole paragraph as deleted and inserted nothing, because the batch compiler treated the line break in `exactText` as a multi-paragraph target and skipped replacement compilation. Localized replacements now always bind at batch start, so weak and strong targets behave identically. Every localized replacement also has its accepted result checked against the compiled text before commit (`PATCH_ROUNDTRIP_MISMATCH`, nothing applied), and uncompiled replacements are refused rather than applied.
- **Reconstruction keeps wrappers and deleted placeholders:** a text edit in a paragraph with a `w:fldSimple` no longer drops the field (Reject All lost its result text); a deleted manual break, tab or inline object is a tracked deletion instead of disappearing; and converting a paragraph to a list keeps hyperlink and content-control containers around the deleted runs, so Reject All keeps the text inside its link. Found by the reject-roundtrip matrix.
- **Punctuation edits stay surgical:** leading and trailing punctuation are separate diff tokens and unchanged substantive words anchor the re-diff, so "agents and attorneys" -> "agents, attorneys, accountants and other professional advisors" marks only " and" -> "," and the appended items.
- **Table reconciliation keeps untouched cell paragraphs:** editing one cell of a Markdown table reconciliation collapsed a multi-paragraph cell (for example `LS` plus a trailing empty paragraph) to one paragraph, untracked, so Reject All lost a paragraph. Cells whose only difference is trailing empty paragraphs are now unchanged, and a modified multi-paragraph cell is reconciled paragraph by paragraph (inserted paragraphs carry tracked marks).
- **Text-to-table redlines insert tracked table rows:** a paragraph range converted to a Markdown table carried tracked cell text but untracked rows, so Reject All left an empty table behind. Rows now carry `w:trPr/w:ins` and tracked cell paragraph marks like other inserted tables.
- **Explicit-start list conversion is reversible with Reject All:** converting a `1. Header` paragraph to a list bound `numPr` without recording a `w:pPrChange`, so rejecting the tracked deletion of the marker left a numbered item. The binding now records the original paragraph properties (empty `<w:pPr/>` when none) in a receipted `w:pPrChange`.
- **Reject/Accept All match Word for table cells and `pPrChange`:** rejecting an inserted paragraph mark keeps the cell's last paragraph (empty) instead of dropping it, accepting a fully deleted sole-paragraph table removes the table, and rejecting a `pPrChange` keeps the paragraph-mark `rPr` and `sectPr` instead of discarding them.
- **Property readers ignore `pPrChange`/`rPrChange` snapshots:** a new `core/current-properties.js` accessor reads only direct `w:pPr`/`w:rPr` children, so a tracked property change no longer makes its historical `numPr`, `pStyle`, `outlineLvl` or `b`/`i` look current. Markdown export, numbering-context detection, list-numbering reuse/overwrite, list-shape checks and captured-entity kind now use it; a usage test fails on new descendant property lookups.
- **Explicit `1. Header` list conversion works on a bare `document.xml`:** with no numbering part the batch had no numbering id state, so the conversion aborted after recording tracked deletions in the receipt and failed with `RECEIPT_RECONCILIATION_FAILED`. The bare runner now always seeds numbering ids as `openDocx` does, and an aborted conversion no longer records receipt revisions.
- **`generateRedlines: false` stays untracked under `slice-cross-author`:** the surgical insert computed its nested-revision flag as
  `!(generateRedlines && ...)`, which is true whenever `generateRedlines` is false, so an inserted word was wrapped in a new tracked
  `w:ins` (`applyRedlineToOxml` and `applyOperationsToDocumentXml`; `redline`/`replace`, `modified`/`replacements`). It now writes
  plain runs, like `merge-same-author`, and leaves other authors' revisions intact (a foreign insertion is sliced around the edit).
- **Table appended after the final paragraph of a pretty-printed document keeps its tracked trailing paragraph:** a whitespace
  text node between the last paragraph and `w:sectPr` made the retained-paragraph table append believe the table did not end
  the body, so the body ended with a table and Word added an untracked empty paragraph that Reject All left behind (plugin
  report R5; Word differential 5 vs 4 paragraphs). The append now skips non-element nodes when deciding. The Word-semantics
  simulator (and therefore the post-apply verifier) counts the untracked paragraph Word supplies after a body-final table, so
  the old output is refused with `REVISION_FIDELITY_VIOLATION`. Documents without whitespace between blocks (the plugin's
  own) were already correct on `Revisions.RejectAll()` and the ribbon; `Document.RejectAllRevisions()` still leaves the
  tracked-inserted last mark behind (a Word API quirk, see the fidelity plan).

## 0.8.4

### Added

- **`textOccurrence` on `format` / `character-format`:** a 1-based match of `textToFormat` inside the resolved paragraph,
  independent of `target.occurrence` (which selects the paragraph). Out of range fails with `PATCH_SOURCE_NOT_FOUND`;
  calls without it still read `target.occurrence`. CLI capability `format-text-occurrence-v1`; contract stays version 8.
- **`applyRedlineToOxml` numbering options:** `numberingXml` and `numberingIdState` remap generated list definitions to
  IDs that do not collide with the source numbering.

### Fixed

- **Soft-break text can be converted to a list:** a paragraph whose `A. …`, `B. …` items are separated by `<w:br/>` was
  refused with `TARGET_NOT_FOUND`, and an explicit list request identical to typed markers was a no-op. In-run `\n`/`\t`
  now serialize back to `<w:br/>`/`<w:tab/>`, including inside `w:delText`, so Reject All restores the breaks.
- **Localized `replacements` work in paragraphs with soft line breaks:** they were refused with `INVALID_OPERATION`.
  A `find` within one line now applies and keeps the breaks; a `find` or `replace` that spans or adds a break is refused.
- **Generated bullets no longer reuse existing numbering IDs:** bullets generated into a document that already defined
  `numId 1` continued an existing decimal list in Word. The bare runner and `applyRedlineToOxml` now allocate new IDs
  from the supplied `numberingXml`, as `openDocx` already did.
- **Appending a table after a paragraph keeps its underline:** the paragraph was tracked as deleted and re-inserted as
  plain text, so a prior same-author underline, or a `++underline++` sent with the table, was lost on Accept All. The
  source paragraph is now kept in place.
- **Word's Reject All removes generated tables:** tracked tables were wrapped in one `w:ins` around `w:tbl`, which Word
  does not recognize. Reject All left an empty table, and a table at the end of a document could hang Word. Tables now
  use row-level `w:trPr/w:ins` with inserted cell content and marks, and `rejectTrackedChangesInOoxml` removes a table
  once all of its rows are rejected. Confirmed in Word.

### Tests

- New `tests/soft_break_list_conversion_tests.mjs`, `tests/soft_break_localized_replacements_tests.mjs`,
  `tests/generated_bullet_numbering_collision_tests.mjs`, `tests/format_text_occurrence_tests.mjs` and
  `tests/table_append_fidelity_tests.mjs`.

### Known limitations

- Localized restore in the rejected view still refuses paragraphs with line breaks.
- The table-append path covers one paragraph followed by one table.
- Under `merge-same-author`, editing the text of a paragraph that carries your own pending formatting change still drops
  that formatting on Accept All.

## 0.8.3

### Fixed

- **Reject All no longer leaves an empty paragraph after an inserted paragraph:** `modified: "<original>\n<new>"` on a plain
  paragraph did not track the new paragraph's mark as inserted. It now does.
- **Reject All no longer merges source paragraphs after a list range replacement:** replacing a multi-paragraph list range
  (`target` + `targetEnd`) collapsed the sources into one deleted paragraph, so Reject All gave results such as
  `Bullet Root ABullet Insertion Anchor`, `AlphaGamma` for an empty middle paragraph, and lost both paragraphs of an
  all-empty range. Each source paragraph, empty ones included, now keeps its own tracked-deleted paragraph. Confirmed in Word.
- **Inspection no longer reports historical paragraph properties as current:** `list`, `styleId` and `headingLevel` (and list
  targeting) could read `numPr`, `pStyle` or `outlineLvl` stored under `w:pPrChange`. Only the paragraph's direct `w:pPr`
  is read now.
- **`openDocx` list range replacement no longer renumbers or re-points lists:** deleted source paragraphs were remapped to a
  new generic `numId` (Reject All restored them with the wrong list type), new items started a separate list (untouched
  following items restarted at 1), and unused numbering definitions were added. Deleted source paragraphs now keep their
  `numPr`, and a same-kind replacement reuses the source `numId`, leaving `numbering.xml` unchanged. Reuse applies when the
  range uses one `numId` and the generated level formats match; format-changing edits still allocate new definitions, and
  definitions created earlier in the same batch are not reused. Confirmed in Word.
- **Explicit list start values are applied:** generated list fragments without `xmlns:w` failed to parse in
  `orchestration/list-structural-fallback.js`, so `7. Header` lost its `startOverride` and xmldom logged parse warnings.

### Tests

- New `tests/list_reject_fidelity_tests.mjs` and `tests/facade_list_numbering_tests.mjs`, with the Word-authored fixture
  `tests/fixtures/agentic-lists/nested-lists-source.docx`; extended inspection and list structural fallback suites.

### Known limitations

- Plain paragraph to list conversion and list format changes are not yet supported canonically; see
  `docs/plans/2026-09-30-canonical-list-operations.md`.
- `1. Header` to `1. Header` on a bare `document.xml` (not a packaged `.docx`) can fail with `RECEIPT_RECONCILIATION_FAILED`.

## 0.8.2

### Fixed

- **Replacement next to a hyperlink no longer pulls adjacent plain text into the link** ([#3](https://github.com/AnsonLai/docx-redline-js/issues/3)):
  replacing `example.org` with `example.net` (localized `replacements` or full `modified`) in a hyperlink followed by a
  plain `.` run produced a link showing `example.net.` after Accept All. The reconstruction-mode word diff treated
  `example.org.` as one token spanning the hyperlink edge. Where a replacement's shared prefix or suffix reaches a wrapper
  boundary inside the deleted text, it is now kept unchanged.
- **Reject All no longer reorders text around a manual line break** ([#4](https://github.com/AnsonLai/docx-redline-js/issues/4)):
  replacing `tabbed` with `aligned` in `<tab>tabbed<w:br/>Line with ` rejected to `<tab>tabbedLine<break> with `, moving
  deleted text across the break. Break and note-reference placeholders are now standalone word-diff tokens (new internal
  `atomicChars` option), and change groups are split at references present on both sides. Covered by `tests/reconstruction_boundary_fidelity_tests.mjs`.

## 0.8.1

### Added

- **Header and footer editing:** add `part` to an operation (a selector `{ kind, type?, section? }` or a part path) to edit a
  header or footer; `inspect().headersFooters` lists the parts with their text, the sections they apply to (including
  inherited ones) and whether they contain fields. Body edits are unchanged and headers/footers are opt-in. Fields such as
  PAGE are atomic (`FIELD_EDIT_REFUSED`), comments are refused (`COMMENT_IN_HEADER_FOOTER`), bad selectors fail closed
  (`PART_NOT_FOUND`, `PART_AMBIGUOUS`), atomic batches roll back across body and parts, revision ids stay unique across the
  package, receipts name the part, `accept`/`reject` also resolve revisions inside headers and footers, and
  `validateDocxPackage` checks header/footer references, relationships and content types.
- **Resolve and reopen comment threads:** `doc.resolveComment(id, { resolved })` and the batch operation
  `{ "type": "comment_resolve", "commentId": "0", "resolved": true }`, so the CLI can resolve through `apply --operations`.
  Word resolves whole threads, so the root and every reply change together, whichever id is given. Only
  `word/commentsExtended.xml` changes; `document.xml` stays byte-identical. Unknown ids fail with `COMMENT_NOT_FOUND`
  (also in `preflight`).
- **Delete comments by id:** `doc.deleteComments({ ids })` and `docx-redline delete-comments --comment-id 3,4`. Deleting a
  thread root also deletes its replies; unknown ids fail closed with `COMMENT_NOT_FOUND`.
- CLI capabilities `comment-resolve-operation-v1`, `delete-comments-by-id-v1` and `header-footer-parts-v1` (additive; the
  contract version is unchanged).
- `services/package-parts.js`, a registry of package part paths, content types and relationship types checked against
  Word-saved fixtures (`tests/fixtures/word-authored/`), and `npm run test:word:package`, which opens generated outputs in
  real Word with repair disabled and checks comment counts, thread parents, resolved state and header/footer revisions
  through Word's object model.

### Fixed

- **Word "unreadable content" prompt on commented documents:** `word/commentsExtended.xml` was relabeled
  `application/vnd.ms-word.commentsExtended+xml` on every save of any document that already had comments.
  Word expects `application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml` and
  offered to repair the file. The constant is corrected, an edit that does not change comment threading no
  longer rewrites the extended part or its content type/relationship, and files already damaged by earlier
  versions are repaired on the next save (`repairKnownContentTypes`, also exported). `validateDocxPackage`
  now rejects the wrong type instead of accepting it. `tests/package_content_type_contract_tests.mjs` checks this
  against a Word-authored content-type table on every OS.
- **Comment threads keyed on the last paragraph:** Word joins `comments.xml` to `commentsExtended.xml` and
  `commentsIds.xml` through the `w14:paraId` of a comment's LAST paragraph. Inspection, package validation,
  replies and delete-cascade used the first, so Word's own multi-paragraph threaded documents failed
  validation and replies to them attached to the wrong key.
- **Edits no longer drop comment anchors:** a mid-paragraph edit of a commented paragraph dropped the trailing
  `commentRangeEnd` and `commentReference` (markers at the very end of the paragraph text were only emitted when the
  last diff part was an insert), leaving the comment unanchored and failing package validation. Found by
  opening our output in real Word; covered by `tests/comment_marker_preservation_tests.mjs`.
- **Replies are now visible in Word:** `comment_reply` only wrote `comments.xml`/`commentsExtended.xml`, and Word
  does not show a reply that has no body markers. Replies now get their own `commentRangeStart`/`commentRangeEnd`/
  `commentReference` next to the thread's, laid out like Word writes them. A reply to a reply joins the root thread.
  A parent with no body anchor fails with `PARENT_ANCHOR_NOT_FOUND`.
- **Replies to a resolved thread are resolved too**, as Word writes them (`reply-to-resolved.docx` fixture). They were
  written open, leaving a thread that was half resolved.
- `commentsIds.xml` and `commentsExtensible.xml` are kept consistent with `comments.xml` on add/delete when present.
- `deleteComments` and `resolveRevisions` report `artifactsChanged` from a real diff of the package (they listed a fixed
  set of parts and missed `commentsIds.xml`, `commentsExtensible.xml` and `[Content_Types].xml`).

## 0.8.0

### Universal Document Facade & Zero-Node Architecture

- **Universal `DocxDocument` & `openDocx`:** Implemented complete-DOCX lifecycle operations
  (`openDocx`, `doc.inspect()`, `doc.applyOperations()`, `doc.resolveRevisions()`, `doc.deleteComments()`,
  `doc.toUint8Array()`, `doc.getRevisionToken()`) on standard `Uint8Array` in `document/docx-document.js`,
  re-exported directly from `index.js`. Enables complete-document redline workflows in any JavaScript
  runtime without Node.js dependencies.
- **Pure-JS ZIP container (`fflate`):** Replaced `node:zlib` with lightweight, zero-dependency `fflate`
  in `document/zip-archive.js`, providing cross-runtime raw DEFLATE / INFLATE compression and decompression
  with full OOXML archive compatibility.
- **Pure-JS FIPS 180-4 SHA-256:** Created zero-dependency, synchronous SHA-256 implementation in
  `core/sha256.js`, removing `node:crypto` from package revision token calculations while retaining
  exact deterministic hash parity with standard NIST vectors.
- **Transparent XML DOM fallback:** Added `resolveDomParserConstructor()` and
  `resolveXmlSerializerConstructor()` in `adapters/xml-adapter.js`, automatically utilizing native
  `globalThis.DOMParser`/`globalThis.XMLSerializer` when present and falling back transparently to
  `@xmldom/xmldom` in headless, server, edge, and sandbox environments without requiring manual provider setup.
- **Zero-dependency sandbox bundles:** Configured `scripts/build.mjs` with `esbuild` (`platform: 'neutral'`)
  to produce single-file, zero-dependency standalone bundles in both ESM (`dist/docx-redline.bundle.js`)
  and CommonJS (`dist/docx-redline.bundle.cjs`), exported as `@ansonlai/docx-redline-js/bundle` for direct
  use in workflow engines (e.g., n8n Code nodes), Cloudflare Workers, AWS Lambda, Deno, and QuickJS sandboxes.
- **Strict runtime boundary enforcement:** Removed `node:zlib` and `node:crypto` from allowed external
  imports in `tests/no_word_api_index_check.mjs`. Automated sandbox isolation tests
  (`tests/sandbox_bundle_isolation_tests.mjs`) verify zero Node built-in imports and full document
  reconciliation in restricted VM contexts without `process`, `Buffer`, or `require`.
- **Backward compatibility:** Preserved full API and `Buffer` compatibility for `@ansonlai/docx-redline-js/node`
  consumers and `doc.toBuffer()`.

## 0.7.2

### Agent Surface Simplification

- **Extraction-first golden path:** Re-established the single deterministic two-turn
  "extract-then-apply" sequence as the default agent workflow across `AGENTS.md`,
  `docs/AGENT_FAST_START.md`, and `docs/SKILL_AUTHORING.md`. Eliminates the
  upfront "choice tax" and model deliberation overhead.
- **Localized replacements after inspection:** Decoupled localized patching from
  speculation. Agents inspect the target paragraph first, then supply compact
  `replacements` instructions, achieving 60.6%–90.8% payload reductions without
  speculative ambiguity risk.
- **Computed success gate:** Consolidated standard success verification onto
  `completion: true` and non-null `outputPath`, removing multi-field verification
  checklists from primary prompts while retaining full `change` audit evidence.
- **Advanced features relocated:** Speculative 1-turn apply, directional context
  windows (`--context-range`), and anchor diagnostics are relocated to
  `docs/AGENT_KNOWLEDGE_BASE.md` and CLI help as opt-in harness policies.
- **Observed-state restoration:** Restructured restoration as a clean request-
  or hint-driven branch (`--view rejected` or `selection.hint`) rather than an
  upfront classification decision.
- **Empirical surface evaluation:** Added synthetic replay benchmark
  (`npm run benchmark:agent-surface`) and validation report documenting 100%
  first-attempt task correctness and payload savings across three workflow policies.

## 0.7.1

### Reliability and Discovery Refinements

- Localized replacement matching now tries exact case-sensitive text first,
  then a narrow ASCII-space/non-breaking-space equivalent match. Ambiguous
  equivalent matches still fail closed, and unchanged boundary NBSPs are
  retained to avoid incidental whitespace redlines.
- Added a narrow single-paragraph restore shortcut. A restore with a strong
  rejected-view target may omit `modified` for verbatim restoration or provide
  exact `replacements`; the CLI exposes the same path through `apply --restore
  --target-id ID`. Ranges, captures, weak targets, and inferred companion
  content remain excluded.
- Fixed rejected-view restore receipts so `resolvedTarget` text, fingerprint,
  revision view, and match diagnostics describe the view actually used to
  resolve the target.
- Fixed alternate-view search hints to honor `--revised`, `--table`, and
  `--body`, and retained compact revision authors plus `deleted`/`inserted` cues
  in `extract` output so empty revision-view paragraphs are self-explanatory.

## 0.7.0

### Agent Performance and Localized Editing

- **Localized exact replacements:** `redline`/`replace` operations may provide
  `replacements: [{ find, replace, occurrence? }]` instead of a complete
  `modified` paragraph. The compiler resolves the accepted-view source once,
  applies non-overlapping exact replacements simultaneously, synthesizes the
  canonical desired paragraph, and delegates to the unchanged OOXML engine.
- **One-turn speculative apply:** `docx-redline apply --find ... --replace ...`
  can resolve a globally unique source literal without preliminary extraction.
  `--search` plus signed `--context-range START:END` scopes visible headings and
  nearby content; `--around N` is symmetric shorthand. Longer, distinctive
  anchors are recommended because generic or repeated phrases widen the unioned
  scope and may cost an ambiguity-recovery turn.
- **Fail-closed selection:** Missing anchors and patch sources, multiple eligible
  paragraphs, ambiguous occurrences, overlaps, conflicts, and unsupported
  localized shapes refuse mutation. A failed speculative command never writes
  or truncates its output path.
- **Verified change evidence:** Localized operation results include a compact
  `change` object with target location, contextual anchor/range, exact
  replacement entries, bounded before/after excerpts, commit disposition, and
  post-mutation accepted-view verification. Atomic rollback marks the evidence
  non-committed and `rolled_back`.
- **CLI contract version 8:** Adds `localized-replacements-v1`,
  `speculative-search-apply-v1`, and `localized-change-summary-v1`. Skills and
  harnesses should pin version 0.7.0 and negotiate only the capabilities used.
- **Measured workflow reduction:** The checked benchmark records 50% fewer tool
  turns and a 64.77%–78.98% smaller deterministic request-token proxy across
  three literal-edit cases. Median native engine/I/O time remained effectively
  neutral; provider reasoning and network latency were not estimated.

### Compatibility

- Existing `modified` operations, strong targets, operations files/stdin,
  tracked-change generation, package validation, rollback, comments,
  Accept/Reject, and revision policies remain supported.
- Localized replacement fields and successful `change` evidence are additive.
  Wrappers that adopt them must require CLI contract 8; integrations using only
  earlier capabilities can continue using their existing operation shapes.
- The implementation remains OOXML-only and adds no runtime dependency or Word,
  Office automation, COM, or proprietary document API requirement.

## 0.6.0–0.6.2

### Safety Fixes

- **Foreign deleted-paragraph resurrection guard**: Refuses non-empty same-paragraph edits when another author owns the paragraph-mark deletion and all existing paragraph content is deleted. The operation now returns `FOREIGN_PARAGRAPH_MARK_DELETION` with the owning author instead of emitting lifecycle-unsafe OOXML; atomic document operations roll back byte-for-byte. `validateRedlineOoxml` reports already-authored instances as warnings.
- **Source-truth whitespace replacement alignment (WP09a)**: Strict target resolution may equate ordinary spaces with NBSPs, but mutation offsets now always come from the resolved paragraph's exact accepted view. Space-equivalent word replacement hunks are refined to character-local edits, and multiple plain-text edits are applied right-to-left against refreshed live spans. This tracks NBSP-to-space substitutions exactly while preserving unchanged hyperlinks and prevents earlier run splits from invalidating later anchors.
- **Exact mismatch diagnostics**: `PATCH_ROUNDTRIP_MISMATCH` now includes the expected and actual code points at the first mismatch. Document-operation results also report bounded `targetTextMatch` diagnostics when target selection used equivalent whitespace.
- **Baseline-delta validation (WP09c)**: Restoration, operation, and package validation now compare issue multisets against the source. Unchanged legacy defects remain visible without blocking safe work, while any added occurrence or mutation-envelope error fails closed with `GENERATED_OOXML_INVALID` before an operation is reported as applied.
- **Revision identity sanitation (WP09d)**: New paragraph/list builders inherit effective `pPr`/`rPr` formatting without cloning historical revision descendants. Operation-level whole-document validation catches duplicate revision IDs and rolls back the operation savepoint before committing its receipt.
- **View-consistent restore targeting**: `restore` targets now default to the rejected view, where a wholly deleted paragraph's source text exists. Rejected-view inspection now computes fingerprints from rejected-view text and reports the view beside each paragraph. Explicit cross-view text or fingerprint mistakes retain strict refusal while returning an actionable view hint.

### New Features

- **Self-describing and bounded agent discovery (WP-00–WP-02)**: CLI contract version 6 adds command-specific JSON help, canonical redline/comment/restore examples, case-insensitive contextual search with `--around`/`--context`/`-C`, direct-hit pagination through `--limit` and `--after`, and deliberate `--all` opt-out. Unscoped inspection is capped at 20 direct records and a 48 KiB soft budget without truncating any returned `exactText`; extraction now distinguishes matches from context and prioritizes human-facing provision/heading references over machine paragraph ordinals.
- **Composable agent safety and compact output (WP-03–WP-04)**: CLI contract version 7 replaces the atomic-coupled agent profile capability with `agent-safety-profile-v2`. `--profile agent` now enforces complete-success exits while preserving progressive execution and the ordinary revision policy; `--atomic` and `--existing-revisions` compose explicitly. Compact CLI JSON keeps receipts only in the authoritative top-level array, omits exact-match diagnostics, summarizes successful equivalent-whitespace matches, and supports one-line `--compact` output without changing Node or standalone-runner results.
- **Packaged skill-authoring contract (WP-05)**: `docs/SKILL_AUTHORING.md` separates automation invariants from transaction/revision policy, file-versus-stdin transport, reviewer preferences, recovery-envelope behavior, and human-facing presentation. It requires contract/capability negotiation, focused contextual extraction, and library delegation while retaining the visible `AI Redliner` fallback without a new author guard.
- **Batch-start source binding (WP-04)**: Strong document-operation targets are compiled once against the immutable source DOM and carried through structural index/fingerprint drift with rollback-aware session identities. Independent edits no longer need manual bottom-up sorting; same-source writes, incompatible revision/format combinations, and unsafe capture fan-out fail before mutation with causal conflict codes. Unique references to paragraph text created elsewhere in the same batch become explicit internal capture dependencies.
- **Machine-actionable recovery (WP-05)**: Operation, package-facade, example-session, and compact CLI failures now carry recovery envelope version 1 with stable stage/category/action metadata, bounded corrective context, and authorization flags. Failed and partial mutation results include an explicit original-vs-output retry plan and operation indexes. `apply --require-complete` exits with code 3 for partial work while preserving legacy behavior without the flag.
- **Compact shell-agent protocol (WP-06)**: CLI contract version 5 adds `operations-stdin` and `agent-profile-v1`. `--operations -` accepts UTF-8 operation arrays/envelopes without a temporary operations file, while explicit `--profile agent` combines atomic rollback and complete-success exit behavior and reports its effective settings. `AGENTS.md` is now a 527-word routing card, and the 349-word `docs/AGENT_FAST_START.md` carries the ordinary edit contract. Ambient project configuration was deliberately deferred because it would save only the profile flag while adding hidden state.
- **Agent rollout audit (WP-07)**: The observational benchmark now compares canonical Node, legacy file-based CLI, compact stdin/profile CLI, and the development session example while verifying accepted/rejected text, comment preservation, independent batch permutations, recovery behavior, and cross-author attribution. The checked run records 82.75% fewer ordinary instruction words, one fewer shell tool turn (33.33%), and 82.22–89.96% smaller localized session requests than canonical Node envelopes without claiming unmeasured provider/model latency.
- **Explicit paragraph restoration (`type: 'restore'`)**: Restores or counterproposes another reviewer's pending whole-paragraph deletion as a separately tracked sibling paragraph. The source deletion remains untouched; the restored paragraph receives its own inserted paragraph mark, content insertion, sanitized paragraph properties, and fresh `w14:paraId`. Single paragraphs and contiguous ranges are supported, with full Accept/Reject lifecycle verification and structured refusals at unsafe table-row, move, section-break, and terminal-paragraph boundaries.
- **Compact mutation CLI contract (WP09b)**: CLI contract version 3 removes `documentXml`, OOXML/package artifacts, full inspection data, and full validation issue arrays from normal `apply`, `accept`, `reject`, and `delete-comments` stdout. Mutation responses retain actionable errors, per-operation receipts, output durability fields, code/count validation summaries, and a derived `completion` flag that cannot report success for failed, partial, or unwritten work.
- **Word-native deleted-section editing (WP09e)**: An explicit rejected-view `insert` operation can split a foreign deletion at an exact anchor-relative offset into sibling `del(A) / ins(B) / del(A)` carriers. Contiguous paragraph restorations now follow their untouched deleted source block, matching Microsoft Word's ordering. Ambiguous anchors and unsupported structural split boundaries remain fail-closed.

## 0.5.1

### Highlights & New Features

- **Cross-Author Revision Slicing (`existingRevisions: 'slice-cross-author'`)**: Adds native support for editing text inside another reviewer's pending tracked insertion without erasing their attribution or requiring prior acceptance.
  - **Word-Native Insertion Slicing**: When inserting text inside another author's pending `<w:ins>`, the engine splits the outer carrier into sibling `<w:ins>` elements at the paragraph level (`[ins(A), ins(B), ins(A)]`), ensuring strict ECMA-376 schema compliance without illegal `ins/ins` nesting.
  - **Word-Native Deletion Slicing**: When deleting text inside another author's pending `<w:ins>`, the engine nests `<w:del>` directly inside `<w:ins>` per ECMA-376 Part 1 `CT_RunTrackChange` Section 17.13.5.21 and Microsoft Word Desktop 365 native behavior.
  - **Boundary & Straddle Deletions**: Deletions straddling baseline text and pending insertions cleanly partition into separate top-level and nested `<w:del>` containers while sharing unified event/author attribution.
  - **Multi-Author Stacked Revisions**: Multiple reviewers can independently delete or insert content within the same carrier insertion without cross-author interference.
  - **Cascading & Coalescing Lifecycles**: Full round-trip lifecycle parity with Microsoft Word Desktop:
    - Rejecting Author A cleanly discards Author A's insertion and any dependent nested deletions by Author B.
    - Accepting Author A unwraps the insertion into baseline text while leaving Author B's deletions pending against the baseline.
    - Rejecting Author B restores Author B's deleted text within Author A's insertion, and coalesces adjacent split `<w:ins>` fragments back into a single continuous carrier.
  - **CLI & Facade Integration**: Fully exposed via the `docx-redline` CLI (`--existing-revisions slice-cross-author`), Node facade (`openDocx`), and batch operation runner.

### Non-breaking Changes & Improvements

- **Validation Update (`validateRedlineOoxml`)**: Refined `NESTED_REVISION` checks to permit direct `<w:ins><w:del>...</w:del></w:ins>` nesting (standard ECMA-376 and Word Desktop behavior), while continuing to strictly reject `ins/ins`, `del/del`, and `del/ins` nesting.
- **Slicing Round-Trip Guard**: Cross-author surgical edits now preserve whitespace-only insertions (including ordinary-space replacements for NBSP characters beside hyperlinks) and verify the exact accepted-view text before reporting success. A mismatch fails closed with `PATCH_ROUNDTRIP_MISMATCH` and returns the original OOXML unchanged.
- **Hyperlink-Adjacent Replacement Anchoring**: Paired replacements immediately before or after a hyperlink now retain a stable insertion point after the deletion run is split, preventing qualifiers from being relocated past the hyperlink or following formatted runs.
- **Multiple Same-Run Insertions**: Insertion-only slicing operations with multiple edit points now apply from right to left against a refreshed live span index, preventing an earlier run split from relocating later insertions.
- **Insertion Stress Hardening**: Slicing now detects edge whitespace changes exactly, uses a character-local insertion-only diff when the original text is an exact subsequence of the modified text, and coalesces new text into an existing same-author carrier when foreign revisions are also present. This prevents repeated phrases from relocating insertions and prevents invalid `w:ins/w:ins` nesting in mixed-author paragraphs.
- **Preflight Inspection**: `preflightOperations` now inspects and validates `slice-cross-author` batches, reporting pending foreign-author carrier targets as `ready` instead of `EXISTING_REVISIONS`.
- **Test Suite Expansion**: Added 7 new test suites covering 36 Word Desktop COM golden fixtures, carrier splitting invariants, the SYN-01..12d synthetic test matrix, PKG-01..06 strict package differential replay, the repeated-text/hyperlink whitespace regression, 76 deterministic insertion stress scenarios, and 12 replacement-anchor lifecycle scenarios (expanding the suite from 88 to 95 passing suites).

## 0.5.0

### ⚠️ Breaking changes

The following changes can require caller or contributor updates. No public
export or valid existing function signature was removed.

#### Runtime and result-contract changes

- **Canonical paragraph text:** `getParagraphText(...)` now returns the
  accepted/current view instead of the earlier simplified `w:t`/`w:tab`
  concatenation. It excludes deleted and `w:moveFrom` content and includes
  structural breaks, soft hyphens, and non-breaking hyphens.
  **Migration:** regenerate compared/cached text with the upgraded library. Use
  `extractCanonicalParagraphText(...)` when accepted/rejected-view semantics
  are intended; integrations that require the former raw traversal must retain
  their own legacy extractor before upgrading.
- **Paragraph fingerprints:** fingerprints now incorporate canonical paragraph
  text. Stored fingerprints can become stale for paragraphs containing
  revisions, moves, breaks, soft hyphens, or non-breaking hyphens.
  **Migration:** do not carry fingerprints across this upgrade; re-extract the
  document and regenerate target descriptors.
- **Invalid operations:** malformed or field-incompatible document-operation
  objects now return a structured `INVALID_OPERATION` error at the runner
  boundary instead of reaching later failure/no-op paths.
  **Migration:** validate operation files against the published schema and
  handle `status === 'error'` plus `error.code === 'INVALID_OPERATION'`.
- **Commented paragraph deletion:** deleting a complete paragraph that contains
  an existing comment now fails with `COMMENTED_CONTENT_DELETE` instead of
  producing OOXML with dangling comment references.
  **Migration:** preserve or deliberately remove/relocate the comment before
  deleting the paragraph, and handle the structured error. Atomic callers keep
  the original document unchanged.
- **Comment anchor failures:** missing or ambiguous comment anchors now return
  `ANCHOR_NOT_FOUND` or `AMBIGUOUS_ANCHOR` and fail/roll back an atomic batch;
  they are no longer reported as successful `no_change` operations.
  **Migration:** treat these codes as failed operations and use a unique exact
  anchor or stricter target descriptor before retrying.
- **Stricter package validation:** DOCX package validation now rejects
  unbalanced comment ranges, dangling references/usages, orphan definitions,
  and duplicate comment-definition IDs that earlier validation could allow.
  **Migration:** repair the comment anchors and `word/comments.xml` definitions
  before applying or validating further changes.

#### Development workflow change

- **Parallel tests by default:** `npm test` now runs up to four isolated child
  processes concurrently. This does not change the published runtime API, but
  contributor automation that depends on test execution order or shared files
  must be updated. **Migration:** make tests independent or set
  `DOCX_TEST_CONCURRENCY=1` for the previous serial scheduling behavior.

Strict targeting remains opt-in on existing low-level mutation APIs; it is the
default only on the newly added Node facade and CLI.

### Deprecations and future breaking changes

- Replacing an existing numbering part without a merge callback now emits a
  deprecation warning. The Node facade and CLI already merge safely by default.
  **This remains supported in this release but will become an error in the next
  major version.** Migrate low-level callers to
  `mergeNumberingXmlBySchemaOrder` before that release.
- **Permissive target resolution deprecation (`AMBIGUOUS_TARGET_HEURISTIC_USED`):**
  Lower-level document operation runners currently default to permissive targeting.
  In v1.0.0, application will default to `strictTargets: true`. In this release,
  when permissive resolution encounters multiple identical candidate paragraphs and
  heuristically chooses the first candidate, it emits an `AMBIGUOUS_TARGET_HEURISTIC_USED`
  warning with candidate count and migration guidance.
  **Migration:** Disambiguate operations using `paragraphId`, `index`, `occurrence`,
  or `fingerprint` and pass `{ strictTargets: true }`.

### ℹ️ Noteworthy Default & Usability Changes (Non-breaking)

The following improvements update default behaviors to maximize agent speed, reduce friction, and eliminate unnecessary errors during iterative editing. These changes are **not technically breaking changes** (no public APIs or exports were removed, and all existing options remain configurable), but represent important behavioral refinements that callers should note:

- **Fallback Author Default (`'AI Redliner'`):** `setDefaultAuthor`, `node/cli.js`, and `openDocx` now default to `'AI Redliner'` (configurable via the `DOCX_REDLINE_AUTHOR` environment variable or `--author`). The CLI `apply` command and core APIs no longer error with `AUTHOR_REQUIRED` when operations omit an author.
- **Output Overwrite by Default:** Destination files provided via `--output` now overwrite by default instead of failing with `OUTPUT_EXISTS`. Callers wanting protection against accidental overwrites can pass `--no-overwrite` or `--no-clobber`. The source document remains protected and is never overwritten unless `--in-place` is explicitly passed.
- **Replacement Event Pairing (`pairReplacements: true` by default):** Adjacent `<w:del>` and `<w:ins>` revisions now default to sharing linked revision metadata and identical timestamps so Microsoft Word groups them as a single replacement in the Reviewing Pane. Pass `pairReplacements: false` for independent revision timestamps.
- **Structured Content Auto-Detection (`structuredContent: true` by default):** Markdown tables, headings (`#`), and lists in replacement text automatically render as native Word elements (`w:tbl`, `w:pStyle`, `w:numPr`). Single outline-numbered legal clauses (e.g. `13.2.1.1`) continue to be treated as ordinary paragraphs without spurious list conversion. Pass `structuredContent: false` to treat replacement text strictly as plain text.
- **Progressive Batch Mode (`atomic: false` default, `atomic: true` opt-in):** Batch operations and CLI `apply` now apply edits progressively by default (`atomic: false`): valid edits commit directly, while failing edits report structured errors in `results` for faster debugging. Pass `{ atomic: true }` (or `--atomic` on the CLI) for all-or-nothing rollback when any operation fails.
- **Clean Direct Edits (`generateRedlines: false` / `--no-redlines`):** Prominently documented that tracked redlines are not always the preferred method. Callers can pass `generateRedlines: false` (or `--no-redlines` on the CLI) for clean execution drafts, document restructuring, or minor typo fixes without tracked changes markup.
- **Existing Revisions Policy (`existingRevisions: 'merge-same-author'` by default):** When editing a paragraph that already contains tracked changes from the same author, prior revisions by that author are merged against the pre-revision baseline (prior changes by that author are reverted and re-diffed to the new modified text), avoiding revision accumulation and nested markup. If a paragraph contains tracked changes from another reviewer, the operation fails closed with `EXISTING_REVISIONS` to protect third-party review marks. Commented revision content fails closed with `COMMENTED_CONTENT_MERGE` so comment anchors cannot be removed or orphaned. Callers can explicitly pass `'accept-all-first'` to normalize prior revisions or `'reject-input'` to refuse editing any revised paragraph.
- **Inline One-Liner CLI Edits:** Added `--target <text>` with `--modified <text>` or `--comment <text>` on `docx-redline apply` for fast 1–2 edit workflows without creating a JSON operations file.

### Added

- Added **Mutation Receipts** (`receipt` on `DocumentOperationResult` and `receipts`
  on `BatchOperationResult`). Every operation execution returns structured,
  commit-aware telemetry capturing durable revision IDs (with kind and target part),
  comment IDs, numbering IDs, relationship IDs, affected targets, and warnings.
  Supports `not_attempted`, `applied`, `refused`, `no_change`, and `rolled_back`
  dispositions.
- Added **Commit-Aware Output Reconciliation Oracle** (`reconcileReceiptsAgainstOutput`).
  Before completing any transaction, reported durable IDs are checked against a
  fresh parse of the generated output OOXML parts. Any reconciliation discrepancy
  immediately fails the transaction and triggers atomic rollback.
- Exported receipt primitives: `ReceiptCollector`, `createEmptyReceipt`, and
  `reconcileReceiptsAgainstOutput` along with `MutationReceipt` and
  `MutationReceiptRevisionItem` TypeScript declarations.
- Added **Paragraph Mark Revisions** (`w:pPrChange`, `markParagraphMarkInserted`,
  `markParagraphMarkDeleted`) ensuring whole-paragraph deletions and insertions
  cleanly track paragraph mark lifecycles for full Accept All / Reject All symmetry.
- Added **Paragraph Boundary Mutation Validation** (`validateParagraphBoundaryMutation`)
  guarding against invalid cross-paragraph boundary merges.

- Added `analyzeStructuredContent(...)` and `planStructuredReplacement(...)`
  for agent-authored attachments and schedules containing mixed headings,
  paragraphs, lists, and tables. The planner returns typed blocks and produces
  one atomic `structuredContent` operation only when the Markdown is valid.

- Added session-scoped paragraph metadata indexing for canonical text,
  normalized text, paragraph IDs, fingerprints, table context, and document
  indexes, plus deterministic targeting and route-profiling benchmarks.
- Added a shared list-marker grammar and parsed list-item vocabulary used by
  pipeline analysis, orchestration, normalization, and list targeting.
- Added an internal reconciliation capability matrix and opt-in route
  instrumentation. Public reconciliation result shapes and route selection
  policy remain unchanged.

- Added `scripts/extract_text.mjs` and `scripts/apply_changes.mjs` as thin
  compatibility entrypoints for legacy skill installations. They delegate to
  the supported CLI, retain strict atomic validated writes, and accept legacy
  operation files with a top-level `changes` array.
- Added typed operations, per-operation authors, strict target descriptors,
  deterministic preflight diagnostics, and auditable resolution metadata.
- Added `inspectDocumentParts(...)`, a structured document inventory using the
  same canonical text extractor as targeting and ingestion.
- Added `@ansonlai/docx-redline-js/node`. `openDocx(...)` applies batches to
  complete DOCX buffers transactionally and validates before commit.
- Added the cross-platform `docx-redline` CLI with JSON `inspect`, `extract`,
  `preflight`, `apply`, `accept`, `reject`, `delete-comments`, and `validate`
  commands, plus a published operation-file JSON Schema.

### Behavior and reliability

- Explicit structured replacements now reject missing Markdown table separator
- Generated Markdown tables repeat their header row and prevent logical data rows from splitting across pages.
  rows, missing data rows, and inconsistent column counts with
  `STRUCTURED_CONTENT_INVALID` instead of inserting literal pipe-delimited text.
- Tables created inside mixed replacements are tracked once at block scope,
  avoiding nested revisions while keeping Accept All and Reject All symmetric.

- Target preflight and live mutation reuse one document-scoped paragraph index;
  strict duplicate detection and resolution metadata retain their prior
  semantics while avoiding repeated whole-document text traversal.
- Revision-ID seeding now walks the DOM with child/sibling pointers instead of
  allocating an array containing every element.
- Single-source-paragraph list expansion uses the focused list generator
  directly. Multi-paragraph marked-list edits deliberately retain the legacy
  run-aware reconciliation pipeline for structural parity.
- Fixed paragraph-to-list expansion when a manually numbered heading carries
  `w:numId w:val="0"`. Numbering suppression is no longer reused as a generated
  list ID, so inserted bullet and numbered items receive valid positive IDs.
- Tracked heading-to-list replacements now keep the deleted source heading in
  its own paragraph and track the paragraph marks of every inserted list item.
  This prevents deleted heading text from joining the first item and preserves
  exact Accept/Reject structure without adding a full-document sentinel
  paragraph.
- Inserted list items now inherit source font family, size, language, and
  related script typography without inheriting heading-only bold, underline,
  italic, or strike formatting unless requested by the replacement markdown.
- Canonical-only table-cell targeting now consumes the shared paragraph-text
  extractor, while offset- and sentinel-producing walkers remain specialized.
  Their accepted visible projections now agree on soft hyphens as well as tabs,
  breaks, hyperlinks, and non-breaking hyphens.

- Comment preflight and application now share one anchor resolver. Exact
  matches take priority, unique ordinary-space/NBSP differences preserve raw
  offsets, and missing or repeated anchors return structured
  `ANCHOR_NOT_FOUND` or `AMBIGUOUS_ANCHOR` errors.
- Full-document operation batches now share one live DOM and revision-ID
  allocator, reducing complete-document parsing and serialization from once per
  operation to once per changed batch. Scoped reconciliation engines and their
  targeting decisions are unchanged.
- Each live-session operation uses a DOM and allocator savepoint. Errors and
  reported no-ops restore the savepoint; atomic failures and all-no-op batches
  still return the exact original XML without serializing it.
- Decomposed the standalone document-operation runner into a stable 13-line
  compatibility facade plus focused session, applier, batch-orchestrator,
  heuristic, and OOXML-mutation modules. Runtime exports and declarations are
  unchanged, and operation internals now use leaf imports instead of importing
  the root entry point.
- Package comment IDs are seeded from existing anchors and definitions, and
  the Node facade merges numbering without discarding prior definitions.
- Canonical paragraph text consistently handles revisions, moves, tabs,
  breaks, soft hyphens, and non-breaking hyphens; see the breaking-change note
  above for `getParagraphText(...)` and fingerprint compatibility.
- Structured inspection now explicitly supports the `current` view as the
  accepted/current document view and resolves comment ranges spanning multiple
  paragraphs as one exact anchor string.
- Transactional no-op results preserve the original DOCX bytes exactly,
  including when the same facade instance has already committed an edit.

### Testing and development

- Added focused and semantic visual-failure regressions for suppressed
  numbering, paragraph separation, paragraph-mark revisions, selective
  typography inheritance, and exact accepted/rejected paragraph sequences,
  plus seeded fuzz variants across list length, marker, font, and size.
  Added the independently specified
  `legal-suppressed-heading-to-bullet-list` and
  `legal-structured-attachment-mixed-blocks` Microsoft Word differential
  fixtures; the synthetic catalogue now contains 49 cases and 49 task types.

- The JavaScript test runner now uses bounded asynchronous subprocess workers,
  capped at four by default, while retaining one fresh Node process per file,
  sorted reporting, timeouts, captured diagnostics, and failure-marker checks.
  Set `DOCX_TEST_CONCURRENCY=1` for serial reproduction.
- Added `npm run benchmark:tests` and a Phase 6 runner regression. The checked
  benchmark reduced median suite time from 18.02 seconds to 7.96 seconds
  (55.80%); 20 consecutive parallel runs and a separate serial run passed all
  63 files. Parallel c8 collection remains valid at 89.90% statements/lines,
  77.40% branches, and 93.10% functions.

- Added Phase 3 targeting/traversal, Phase 4 list/text parity, and Phase 5 route
  compatibility regressions. Together with the Phase 6 runner regression, the
  suite now contains 63 passing test files.
- Added `npm run benchmark:targeting` and `npm run profile:routes`. The checked
  10,000-paragraph benchmark showed no single-operation regression and a 47.16x
  Node median improvement for 100 cached resolutions. The equivalent native
  Chrome benchmark showed no single-operation regression and a 36.45x batch
  improvement. All 47 Microsoft Word differential fixtures and all 60 real-
  document Word corpus scenarios passed after the routing changes.

- Added subprocess coverage for legacy wrapper selection, operation-file
  compatibility, successful output, atomic anchor failure, source-byte
  preservation, and protection of pre-existing output files.
- Added a live-session accuracy and instrumentation suite covering accepted and
  rejected text, structural validity, sequential equivalence, comments, lists,
  tables, highlights, one-parse/one-serialize execution, exact no-ops, and
  atomic rollback. Added `npm run benchmark:session`; the checked accuracy-first
  benchmark reduced parse/serialize counts from 10/10 to 1/1 and measured a
  1.58x median speedup with per-operation savepoints retained.
- Added a Phase 2 architecture regression covering facade export identity,
  direct leaf imports, exact session rollback, isolated context commit, and
  stable comment-first scheduling.
- Added five edge-case suites for canonical text views, cross-paragraph comment
  inspection, table/list/reference context, sequential package transactions,
  selective multi-author cleanup, ZIP rejection behavior, CLI filters, JSON
  exit contracts, and destructive-output safeguards. Together with the Phase 2
  Phase 1 and Phase 2 regressions, these established the earlier 57-file
  baseline expanded by the subsequent performance phases.

## 0.3.0

### Breaking changes

- Malformed OOXML no longer escapes public transform APIs as a raw parser
  exception. In particular, revision accept/reject and comment-deletion callers
  that previously used `try`/`catch` must now inspect `status === 'error'` and
  `error.code === 'PARSE_ERROR'` on the returned result.
- Caller content is no longer sanitized by default. Pass `sanitizeInput: true`
  to remove a standalone leading assistant-preface line. Dollar-delimited text
  and literal `\\n` / `\\r\\n` sequences are never removed implicitly.
- The exported `sanitizeAiResponse` helper no longer removes dollar-delimited
  spans or converts literal `\\n` / `\\r\\n` sequences. Direct callers that
  relied on those transformations must perform them explicitly.
- `applyOperationsToDocumentXml` is atomic by default. Failed batches return
  the original document with `rolledBack: true`; pass `atomic: false` to retain
  partial-result behavior.
- Missing multi-line targets now return `TARGET_NOT_FOUND` instead of being
  indistinguishable from a no-op. Stale batch anchors also return
  `TARGET_NOT_FOUND` instead of falling back to a visibly different paragraph.

### Behavior changes and fixes

- Revision IDs are now allocated per document. Near-limit prior IDs restart in
  a safe low range, and unrelated bookmark/relationship IDs no longer seed the
  revision counter.
- Concurrent calls with explicit authors keep their revision attribution
  isolated from process-global defaults.
- `existingRevisions: 'accept-all-first'` now preserves the original OOXML when
  the requested edit is a no-op.
- Added `existingRevisions: 'accept-all-first-keep-normalized'` for callers that
  explicitly want accepted existing revisions returned even when no redline is
  added.
- Diff output is deterministic by default. Inputs exceeding 262,144 unique
  tokens return `DIFF_TOKEN_LIMIT` with the original OOXML unchanged instead of
  risking truncated text; leading whitespace is preserved exactly.
- Reconstruction preserves structural line breaks and body section-property
  placement. A target that names only part of a reconstruction range now
  returns `PARTIAL_TARGET` instead of risking deletion of untargeted content.
- Paragraph targeting and reconstruction now preserve leading, middle, and
  trailing `w:tab` elements as literal tab characters instead of omitting or
  trimming them.
- Adjacent redline edits now preserve complex-field instructions, begin/
  separate/end markers, and unchanged cached display results in their original
  field boundary. Field instructions remain inert and are not evaluated.
- Added structured XML parse results and result-returning ingestion helpers.
- Hardened deterministic diff token handling and leading-whitespace fidelity.

### Testing and development

- Expanded the desktop Word differential from 20 to 25 cases, adding structural
  checks for bookmarks, internal hyperlinks, mixed formatting, content
  controls, and tables.
- Expanded the desktop Word differential from 25 to 28 cases with structural
  tabs and a locked PAGE field, and added tab/field shapes to focused and fuzz
  regression testing.
- Expanded the desktop Word differential from 28 to 33 cases with comments,
  footnotes, endnotes, headers/footers, and an external hyperlink.
- Extended the deterministic script-only DOCX packager with opt-in related
  parts, relationship/content-type validation, reusable fixture constructors,
  and SHA-256 verification that supplied parts remain byte-identical. Runtime
  library code and dependencies are unchanged.
- Added validated task, structure, oracle, and manual-review metadata across the
  33 synthetic and 20 SuperDoc Word cases. `npm run report:word:coverage`
  produces a deterministic coverage matrix with explicit high-priority gap
  dispositions, while `npm run review:word:prepare` creates a pending rotating
  human-review manifest without self-certifying visual results.
- Added detailed production-function coverage reporting and a checked per-file
  regression baseline. Five behavior-focused Phase 3 suites exercise numbering
  collisions/remapping, every reconciliation route, list construction and
  fallback, patch/format boundaries, table decisions, pipeline modes, and
  standalone rollback/highlight paths. Production function coverage moved from
  437/540 to 496/542 and all reachable P0 functions are covered; no runtime
  behavior or compatibility contract changed.
