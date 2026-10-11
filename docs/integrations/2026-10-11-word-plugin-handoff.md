# Hand-off for the Word add-in (AIWordPlugin): docx-redline-js changes since 0.8.4

Audience: an AI agent or developer working on the Word add-in that consumes `@ansonlai/docx-redline-js`.
Source of truth: `CHANGELOG.md` "Unreleased" and `docs/AGENT_KNOWLEDGE_BASE.md` in this repo. These changes are on
`master` and **not yet published**; they will ship as 0.9.0 (release notes: `docs/releases/0.9.0.md`). Until then, depend on
a pinned git commit or a local `file:` link.

## What changed, in one paragraph

The library now refuses any edit whose Word "Reject All" would not give back the original document, and explains why. List
work that the add-in had to do itself (typed "1." / "A." headers, changing a list's numbering style) is now available as
canonical, Word-verified operations. Several old silent successes are now warnings or refusals with recovery instructions.

## 1. Must handle (behavior changes that can break the add-in)

1. **New refusal: `REVISION_FIDELITY_VIOLATION` (default on).** Nothing is written when it fires. Handle it like other
   refusals: do not retry the same arguments.
   - `error.message` is one human-readable line (safe to log or show).
   - `error.violations[]` (max 5): `invariant`, `expected`, `actual`, `explanation`, `source.sourceParagraphIndex`
     (1-based, same index extraction reports) and `source.text`.
   - `error.recovery.action` = `narrow_or_change_operation`; `error.recovery.retryGuidance[]` has one retry instruction per
     invariant; `recovery.reportIfNarrowRetryFails` says to report, not to disable the check.
   - Invariants: `REJECT_PARAGRAPH_COUNT`, `REJECT_PARAGRAPH_TEXT`, `REJECT_PARAGRAPH_PROPS`, `REJECT_NUMBERING_REFS`,
     `REJECT_WRAPPERS`, `ACCEPT_INTENT` (plain-text `redline`: Accept All of the target must equal `modified`), `ACCEPT_NEIGHBORS` (untracked edits that changed paragraphs outside their target) and
     `UNTRACKED_CREATED_REVISION` (an untracked edit that wrote a revision id the source lacked).
   - Option `verifyRevisionFidelity: 'enforce' | 'warn' | 'off'` (CLI `--revision-fidelity`). Keep the default. `'warn'`
     applies and puts the same text in `warnings`; never ship `'off'`. The CLI agent profile rejects `off`.
2. **Typed list markers on plain paragraphs are no longer a silent `ok`.** `"Header"` -> `"A. Header"` now returns warning
   `LIST_MARKER_NOT_BOUND` (in `diagnostics[]` and `warnings`), or the error `UNSUPPORTED_LIST_CONVERSION` under the strict
   policy (`atomic: true`, CLI `--atomic` or `--profile agent`, or `strictListConversion: true`). **Atomic batches that type
   `1.`/`A.` onto plain paragraphs are now refused.** Either switch to `list-format` (section 2) or set
   `listMarkers: 'literal'` on the operation to keep the old literal-text behavior.
3. **Renamed recovery actions** (update any switch on `recovery.action`):
   - `apply_plain_text_or_convert_in_host` -> `use_list_format_convert_plain`
   - `change_list_format_in_host` -> `use_list_format_change_format`
   Both now include `recovery.suggestedOperation`, a ready-to-send `list-format` request for the same target.
4. **Numbering merge.** `ensureNumberingArtifactsInZip` always merges into an existing `word/numbering.xml`; the
   `mergeNumberingXml` callback option is ignored (remove it). A conflicting id now throws `NUMBERING_ID_CONFLICT` instead
   of silently binding paragraphs to another list. If the add-in makes several separate bare-XML calls, pass the
   document's `numberingXml` or one shared `numberingIdState` to every call so generated ids cannot collide.
5. **`list-format` scope conflicts.** A `list-format` with `scope: 'list'` counts as a write to every list item it rebinds;
   another write to one of those items in the same batch fails with `LIST_FORMAT_COMPOSITION_CONFLICT`. Send it in a
   separate apply, or use `scope: 'range'`.
6. **Small result changes:** the deleted paragraph mark of an empty-`modified` whole-paragraph delete is receipted as
   `del` (was `structural`); unknown `paragraph-format.properties` keys are now rejected with `INVALID_OPERATION`
   (`field: properties.<key>`, recovery `change_request`); the 0.8.x warning `UNKNOWN_PARAGRAPH_FORMAT_PROPERTY` no longer
   fires. Send only `alignment`, `keepNext`, `keepLines`, `pageBreakBefore`, `style`, `numbering`.
7. **Range targets.** A `targetEnd` given as a compact handle (`paragraphId` + `fingerprint`) is now honored (0.8.x
   silently ignored it for a plain `redline` and replaced only the first paragraph; remove any workaround). A range that
   crosses a table or other container boundary is refused `TARGET_RANGE_INVALID`.
8. **CLI contract version is 9.** If the add-in shells out to the CLI and only uses existing operations, requiring 8 still
   works.

## 2. Can now move off the host path (new operations)

Negotiate capabilities first (`docx-redline version` -> `capabilities`); fall back to the existing host path when missing.

| Need | Operation | Capability |
|---|---|---|
| Typed header -> real list item (`A.`, `1.`, `I.`, bullets; start values) | `{ type: 'list-format', target, targetEnd?, list: { mode: 'convert-plain' } }` (marker stripped as a tracked deletion, numbering bound as a tracked `pPrChange`) | `list-format-v1` |
| Bind a list number to text that has no marker | `list: { mode: 'convert-plain', strip: 'none', format: 'upperLetter', start: 3 }` | `list-format-v1` |
| Changed-text header (`1. Old` -> `1. New`) | one batch: `redline` with `modified: '1. New'` (keep the marker) + `list-format` `convert-plain` on the same target; the library orders them | `list-format-v1` |
| Change a list's kind/format (bullets -> 1., 1. -> A., nested level -> (i)) | `list: { mode: 'change-format', format, level?, scope: 'list' }` targeting any item | `list-format-change-v1` |
| Continue an earlier list | `list: { mode: 'join', continueFrom: <target of an earlier list item> }` | `list-format-change-v1` |
| Put a paragraph in / take it out of an existing list | `paragraph-format` with `properties.numbering: { numId, level }` or `null` | `paragraph-numbering-binding-v1` |

All of these are tracked: Reject All restores the original labels/text, Accept All shows the new list. Each was checked in
Word 16 with `ListFormat.ListString`. Refusal codes are listed in `docs/AGENT_KNOWLEDGE_BASE.md` (`LIST_MARKER_*`,
`LIST_FORMAT_*`, `PARAGRAPH_NUMBERING_NOT_FOUND`, `PARAGRAPH_FORMAT_REVISION_CONFLICT`); each carries a recovery envelope.
Not supported (refused): style-linked/heading numbering, markers inside hyperlinks/fields/revisions, `2.1.`/`(a)` markers,
text-box lists, joining plain paragraphs.

## 3. Fixed issues the add-in reported or may have worked around

- Report `2026-09-30-canonical-list-operations.md`: (a) plain header to list and (b) list format change are now
  operations (section 2); (c) parser diagnostics fixed earlier.
- `"1. Header"` on a bare `document.xml` no longer fails with `RECEIPT_RECONCILIATION_FAILED`.
- `openDocx` path no longer remaps deleted source paragraphs to a new `numId` (Reject/Accept list breakage).
- Reject All now restores: table cells' empty trailing paragraphs, text-to-table conversions, simple fields
  (`fldSimple`) in edited paragraphs, deleted manual breaks/tabs/inline objects, hyperlink membership after list
  conversion, `1. Header` list bindings.
- Text-equal letter/roman/bullet header conversions now happen in place (keeps `w14:paraId`, bookmarks, formatting).
- Historical `numPr`/`pStyle` inside `w:pPrChange` is no longer reported or used as current.
- Open reports in `AIWordPlugin/docs/library-issues/` (all reproducers exit 1 against this tree, 2026-10-11; tracked in
  this repo's `docs/plans/completed/2026-10-11-word-plugin-reported-issues.md`):
  - `2026-10-02-weak-target-soft-break-deletion.md`: fixed (`7548a9e`); a localized replacement whose accepted text does
    not match the request is now refused (`PATCH_ROUNDTRIP_MISMATCH`) instead of reported `ok`.
  - `2026-10-02-generate-redlines-false-cross-author-insertion.md`: fixed (`7ae4f00`); new invariant
    `UNTRACKED_CREATED_REVISION` refuses any `generateRedlines: false` edit that would create a revision. The add-in's
    `TRACKING_POLICY_MISMATCH` guard can stay as defense in depth.
  - `2026-10-02-separate-list-operations-restart-numbering.md`: fixed (`be539d5`); separate conversions with continuing
    markers in one batch share one list (Word `ListFormat.List.ListParagraphs.Count` confirms one list, not just labels).
    The former limit (text-changed conversion of a paragraph with direct `keepNext`/spacing refused
    `REJECT_PARAGRAPH_PROPS`) is fixed (`0d0bd2d`).
  - `2026-10-01-word-reject-table-append-paragraph.md`: on this tree Word's ribbon Reject All and
    `Document.Revisions.RejectAll()` give exactly the 7 source paragraphs and 0 revisions; the extra paragraph and
    leftover revision appear only with `Document.RejectAllRevisions()` (see section 5). The add-in's Word oracle
    `scripts/verify-wp6-word.ps1` (lines 112-113) uses `AcceptAllRevisions()`/`RejectAllRevisions()`; switch it to
    `$document.Revisions.AcceptAll()`/`$document.Revisions.RejectAll()` and re-run. A real neighbouring defect (pretty-printed bodies) was fixed in
    `0835f58`.
  - `2026-09-30-canonical-list-operations.md`: superseded by `list-format` (section 2).
- A plain `redline` keeps `w14:paraId` / `w14:textId` (`0d0bd2d`), and a text-changed header conversion with direct
  `keepNext`/spacing is no longer refused (see the `2026-10-02-separate-list...` item above).
- Any workaround in the add-in for these can be removed after upgrading; keep tests for them. Move each report to
  `completed/` once its reproducer exits 1 against the installed package.

## 3b. Known limitations (deferred, not blocking)

`list-format` `detach`/`materializeMarker`; a marker inside the author's own pending insertion; with `slice-cross-author`,
range replace or plain-to-list of another author's inserted paragraph fails late with `PATCH_ROUNDTRIP_MISMATCH` (nothing
applied); bare-XML numbering id collisions are detected, not remapped. The verifier adds roughly 30-40% per batch.

## 4. Suggested work for the add-in agent

1. Upgrade the dependency to the pinned commit (or 0.9.0 when published); remove the `mergeNumberingXml` argument; send only supported `paragraph-format` keys.
2. Add handling for `REVISION_FIDELITY_VIOLATION` (log `error.message`; for agent retries use
   `error.recovery.retryGuidance`; surface to the user if a narrower retry also fails).
3. Update `recovery.action` switches for the two renamed actions; prefer `recovery.suggestedOperation` when present.
4. Replace host-side list conversion with `list-format` behind capability checks; keep the host path as fallback.
5. Decide per call site whether typed markers should become lists (`list-format`) or stay text
   (`listMarkers: 'literal'`), especially in atomic batches.
6. Re-run the add-in's own list and Reject/Accept tests in Word; report any `REVISION_FIDELITY_VIOLATION` with the
   operation JSON and `violations` back to this repo (they indicate library gaps, not add-in bugs).

## 5. What to send back to the library

For any refusal you believe is wrong: the operation JSON, the paragraph's extracted `exactText`, the full `error` object,
and whether Word's Reject All on the add-in's own output looked right. Use `Document.Revisions.RejectAll()` /
`AcceptAll()` when checking in Word automation; `Document.RejectAllRevisions()` can leave the final paragraph mark's
revision behind and give a misleading result.
