# Known Limitations, Fail-Closed Edge Cases & Planned Architectural Evolutions

> **Note on Terminology:**  
> The items documented in this file are **not unintentional defects or unhandled exceptions**. Rather, they represent intentional, fail-closed architectural boundaries where `@ansonlai/docx-redline-js` deliberately refuses unsafe operations to protect Office Open XML (OOXML) schema validity, prevent Word package corruption, and guarantee Accept All / Reject All lifecycle symmetry.

---

## 1. Comments on Pending Deletions (`<w:del>`)

### The Scenario
Attempting to attach a comment to text that is currently marked as deleted by another reviewer (or the same reviewer), or deleting a paragraph that contains active comment anchors.

### OOXML Mechanics & Why It Is Refused
In WordprocessingML (ECMA-376 Part 1), comments are anchored to text runs via boundary markers:
```xml
<w:commentRangeStart w:id="42"/>
<w:r><w:t>commented text</w:t></w:r>
<w:commentRangeEnd w:id="42"/>
<w:r><w:commentReference w:id="42"/></w:r>
```

When text is inside a pending tracked deletion (`<w:del><w:r><w:delText>deleted text</w:delText></w:r></w:del>`), anchoring a comment to that deleted text creates severe lifecycle and schema paradoxes:

1. **Accept All Paradox (Orphaned / Dangling Definitions):**  
   If the pending deletion is accepted, all child runs and text inside `<w:del>` are removed from the document body. If comment markers were placed inside or around those deleted runs, the markers either disappear or collapse into a zero-width invalid range. Meanwhile, the corresponding comment entry in `word/comments.xml` remains defined. This produces dangling comment references or orphan definitions, which trigger Word recovery prompts and fail package validation.
2. **Reject All Ambiguity:**  
   If the deletion is rejected, the text is restored to the baseline document, but positioning and attribution of comments attached to historically rejected runs exhibit unstable or undefined behavior across different versions of Microsoft Word Desktop.

### Engine Behavior & Error Codes
- **Comment on pending deletion:** Guarded in [`services/comment-locator.js`](../../services/comment-locator.js). The locator inspects run ancestor tags; if an ancestor is `del`, it refuses with:
  ```json
  {
    "status": "error",
    "error": {
      "code": "UNSAFE_REVISION_NESTING",
      "message": "Refusing to attach comment to pending deletion."
    }
  }
  ```
- **Whole-paragraph deletion with comments:** Guarded in [`services/document-operation-mutations.js`](../../services/document-operation-mutations.js). Refuses with:
  ```json
  {
    "status": "error",
    "error": {
      "code": "COMMENTED_CONTENT_DELETE",
      "message": "Cannot delete paragraph containing active comments."
    }
  }
  ```

### Recommended Workflow for Callers
- Resolve or delete the existing comments first via `delete-comments` before deleting the containing paragraph.
- If commenting on clause removals, attach the comment to surrounding baseline text or the containing section heading rather than deleted characters.

---

## 2. Mutating or Commenting Inside Active Tracked Moves (`<w:moveFrom>` / `<w:moveTo>`)

### The Scenario
Attempting to edit, slice, format, or comment on content that is part of an unaccepted tracked move from another reviewer (e.g., text enclosed within `<w:moveFrom>` or `<w:moveTo>` blocks).

### OOXML Mechanics & Why It Is Refused
In ECMA-376, moving text is not a simple deletion followed by an insertion. Microsoft Word links moved text across two distant locations in the XML tree as a coordinated transaction:
- **Origin (Source):**
  ```xml
  <w:moveFromRangeStart w:id="10" w:name="move1"/>
  <w:moveFrom w:id="11" w:author="Reviewer" w:date="...">
    <w:r><w:t>relocated clause</w:t></w:r>
  </w:moveFrom>
  <w:moveFromRangeEnd w:id="10"/>
  ```
- **Destination (Target):**
  ```xml
  <w:moveToRangeStart w:id="12" w:name="move1"/>
  <w:moveTo w:id="13" w:author="Reviewer" w:date="...">
    <w:r><w:t>relocated clause</w:t></w:r>
  </w:moveTo>
  <w:moveToRangeEnd w:id="12"/>
  ```

The origin and destination share matching move names (`w:name`) and correlated tracking metadata.

1. **Dual-Location Parity Requirement:**  
   Microsoft Word's internal move reconciliation oracle requires the text content between `<w:moveFrom>` and `<w:moveTo>` to remain strictly identical. If a second reviewer alters text at the destination `<w:moveTo>` (e.g., editing a typo or deleting a phrase) without an identical paired mutation at the origin `<w:moveFrom>`, Word's move table corrupts: Word either unlinks the move, treats them as disconnected orphan changes, or prompts for document repair.
2. **Carrier Slicing Across Remote Trees:**  
   Applying cross-author revision slicing inside a move would require simultaneously splitting and synchronizing carriers across two completely separate subtrees in `word/document.xml`.

### Engine Behavior & Error Codes
- **Accepting / Rejecting Existing Moves (Fully Supported):**  
  The engine fully supports resolving existing moves via [`services/revision-comment-management.js`](../../services/revision-comment-management.js):
  - **Accept:** Removes `<w:moveFrom>` origin and unwraps `<w:moveTo>` destination into clean baseline text.
  - **Reject:** Unwraps `<w:moveFrom>` origin back into baseline text and removes `<w:moveTo>` destination.
- **Mutating Inside Pending Moves (Guarded & Refused):**  
  Attempting to edit or comment inside an active move fails closed in [`services/document-operation-mutations.js`](../../services/document-operation-mutations.js) and [`services/comment-locator.js`](../../services/comment-locator.js) with:
  ```json
  {
    "status": "error",
    "error": {
      "code": "UNSAFE_REVISION_NESTING",
      "message": "Refusing to mutate content with move revisions until move lifecycle is designed."
    }
  }
  ```

### Recommended Workflow for Callers
- Accept or reject prior move revisions in the document before applying new edits to the relocated sections.
- When restructuring clauses, prefer explicit insertions and deletions (or paragraph restorations via `type: 'restore'`), which provide predictable single-location lifecycles.

---

## 3. Prospective Evolution: Strict Target Resolution by Default

### Current Two-Tier Architecture (since 0.5.x, unchanged in 0.11.2)
Targeting paragraphs by plain text (e.g., `"target": "Notices"`) can be ambiguous when multiple paragraphs have identical content (repeated headers, boilerplate, empty lines, or identical table cells).

Currently, resolution behavior differs by API tier:
1. **High-Level Tier (CLI `docx-redline apply` & Node `openDocx` Facade):**  
   Defaults to **strict targeting** (`strictTargets: true`). Any duplicate target text immediately fails closed with `AMBIGUOUS_TARGET` and returns candidate diagnostics.
2. **Lower-Level Tier ([`applyOperationsToDocumentXml`](../../services/standalone-operation-runner.js)):**  
   For backwards-compatibility with earlier scripts, defaults to **permissive targeting** (`strictTargets: false`). When duplicate paragraphs match, it heuristically selects candidate #1 and emits a deprecation warning:
   ```text
   AMBIGUOUS_TARGET_HEURISTIC_USED: Target text matched <N> paragraphs; permissive resolution chose candidate 1.
   Migrate to strict targeting (e.g. strictTargets: true with paragraphId, index, occurrence, or fingerprint) before v1.0.0.
   ```

### Planned Future Direction
At an eventual major version boundary (prospective `v1.0.0`, without committing to a specific release timeline or date):
- `applyOperationsToDocumentXml` is planned to switch its default to `strictTargets: true`.
- Silent fallback to candidate #1 will be eliminated across all public APIs.
- Callers are encouraged to adopt strict target descriptors today:
  ```json
  {
    "type": "replace",
    "target": {
      "exactText": "Notices",
      "paragraphId": "1A2B3C4D",
      "occurrence": 2
    },
    "modified": "Updated Notices Clause"
  }
  ```

---

## 4. Headers and Footers

Operations edit a header or footer only when they carry `part` (see the
[knowledge base](../AGENT_KNOWLEDGE_BASE.md#headers-and-footers)). Everything below fails closed: the whole request is
refused before anything is written.

| Scenario | Why it is refused | Error code |
|---|---|---|
| A `comment`, `comment_reply` or `comment_resolve` with `part` | Word has no comments in headers or footers; an anchor there would be dropped or break the file. | `COMMENT_IN_HEADER_FOOTER` |
| An edit that would add a revision inside a field's instruction or cached result (`PAGE`, `NUMPAGES`, `DATE`, `w:fldSimple`) | Word recomputes fields. A tracked deletion inside the result leaves a stray literal after the field. | `FIELD_EDIT_REFUSED` |
| A selector that matches no part | Nothing to edit. | `PART_NOT_FOUND` |
| A selector that matches several distinct parts | Guessing would edit the wrong page. Add `section` or use the part path. | `PART_AMBIGUOUS` |

Behavior worth knowing:

- A part shared by several sections (Word's "link to previous") is one part. Editing it changes every section that uses
  it, and the result lists them in `partSections`.
- `exactText` is the whole paragraph text including field results, for example `Page 1`. Text around a field can be edited;
  the field runs are kept intact.
- Tables in headers and footers are edited like body tables.

Not supported yet:

- Creating a header or footer part that does not exist (it needs a new part, relationship, `sectPr` reference and content
  type), removing one, or changing `titlePg` / `evenAndOddHeaders`.
- Text boxes, images and drawings inside headers and footers are not inspected or edited; only their paragraph text is.
- The field guard applies only to header and footer operations. Body edits that touch a field's cached result are not
  refused the same way, so avoid targeting field text (cross-references, dates, page references) in the body.

## 5. Comment Threads

| Scenario | Behavior | Error code |
|---|---|---|
| Reply to a comment whose parent has no anchor in the body | Word does not display a reply without its own markers next to the thread's, so it is refused. | `PARENT_ANCHOR_NOT_FOUND` |
| Resolve, or delete by id, a comment that does not exist | Refused; nothing is written. | `COMMENT_NOT_FOUND` |

- Resolved is thread-level (verified in Word): resolving any comment in a thread resolves the root and every reply, and a
  reply added to a resolved thread is written resolved.
- Word threads are flat. A reply to a reply joins the root thread.
- `commentsIds.xml` and `commentsExtensible.xml` are updated only when the document already has them; they are never
  created, because Word does not require them.
- A `commentsExtensible.xml` entry we add copies the comment's `w:date` as `dateUtc`. That is exact for comments written by
  this package (real UTC), but Word writes `w:date` as local time, so an entry added for a Word-authored comment that lacked
  one would carry local time labeled as UTC.
- When the last comment is deleted, `commentsExtended.xml`, `commentsIds.xml` and `commentsExtensible.xml` stay in the
  package as empty parts, with their content types and relationships. Word opens this without complaint, but the parts are
  not removed.

## 6. Package Repair

Files saved by versions before 0.8.1 can carry the wrong `commentsExtended.xml` content type
(`application/vnd.ms-word.commentsExtended+xml`), which makes Word offer to repair them.

- Every write (`apply`, `accept`, `reject`, `delete-comments`, resolving a thread) repairs it with
  `repairKnownContentTypes`. A request that changes nothing returns the original bytes, so it does not repair; there is no
  repair-only save.
- `docx-redline validate` reports the wrong type as an error and does not repair it.
- Files that were already sent on to others are unaffected until someone re-saves them with this package or with Word.

## 7. Revision Fidelity Verification and Resolved Gap Bugs

### Post-apply verifier

Every tracked operation is checked after generation: Word's Reject All of the output must equal Word's Reject All of the
source (`services/revision-fidelity-verifier.js`). A violation fails closed with `REVISION_FIDELITY_VIOLATION` (disposition
`refused`, nothing written). Plans: `completed/2026-09-30-revision-fidelity-invariants.md`,
`completed/2026-10-11-pre-0.9.0-open-issues.md`. Current behavior and limits (0.11.2):

- The default mode is `enforce` (refuse). `verifyRevisionFidelity: 'warn'` applies and reports instead; `'off'` is for
  diagnosis only and the CLI agent profile rejects it (`--revision-fidelity`, capability `revision-fidelity-v1`).
- Untracked operations (`generateRedlines: false`) skip the Reject All invariants; instead `ACCEPT_NEIGHBORS` (single-target
  locality, `d045e7e`) and `UNTRACKED_CREATED_REVISION` (no new revision ids, `7cd74bd`) apply.
- Also enforced: `ACCEPT_INTENT` for plain-text `redline` with complete `modified` and the pre-commit localized-replacement
  intent check (`73d3d1b`, `7548a9e`); numbering-part existence for new `numId`s (`REJECT_NUMBERING_REFS`, `73d3d1b`); exact
  accept-first range scope (`73d3d1b`). Cost about +8-14% per operation (`427eb33`).
- `ACCEPT_NON_TEXT_CONTENT` (0.11.1, `851b03c`) applies to tracked and untracked operations: Accept All of the output must
  keep every footnote/endnote reference (by id), drawing (`w:drawing`, `w:pict`) and embedded object (`w:object`) of the
  source. Explicit deletions (empty `modified`, all-empty `modified` array, `delete`) are exempt. `w:sym` is not checked,
  because plain text carries it.

### Fixed gap bugs (do not re-report)

| Bug | Fixed in | Note |
|---|---|---|
| `1. Header` list conversion on a bare `document.xml` returned `RECEIPT_RECONCILIATION_FAILED` (still listed under 0.8.4 known issues in the released changelog) | `73c67e8` | Bare runner seeds numbering state; an aborted conversion records no receipt revisions. |
| Explicit-start / `1. Header` list conversion bound `numPr` without `w:pPrChange`, so Reject All left a numbered item | `f72f60a` | `orchestration/list-structural-fallback.js` |
| Text-to-table redline put tracked cell text in untracked rows, so Reject All kept the table | `c20f33c`, `4568ddc` | `engine/table-mode.js`; `ReconciliationPipeline.executeTableGeneration` (`pipeline/pipeline.js`) fixed in `4568ddc`. |
| Library Reject/Accept differed from Word for cell paragraph marks, deleted tables, and `pPrChange` paragraph marks | `cedc68a` | `tests/word_revision_semantics_tests.mjs` |
| Numbering helpers and inspection read `numPr`/`pStyle` from `w:pPrChange` snapshots as current | `1a3eeb6`..`1f89ae0` | `core/current-properties.js` plus usage lint |
| Lettered/roman whole-paragraph delete plus insert conversion lost `w14:paraId` and bookmarks | `1240842` | `list-format` `convert-plain` edits in place (tracked marker deletion plus `pPrChange`). |
| Literal underscore runs (`_____ (sign)`) were parsed as Markdown emphasis and shortened | 0.10.0 | CommonMark flanking rules in `pipeline/markdown-processor.js`; `structuredContent: false` is plain text. |
| Paragraphs added by a `structuredContent: false` redline still parsed inline Markdown and `#` headings | 0.11.0 | Also the table route and one-column `\| --- \|` separators. |
| A text edit dropped footnote/endnote references from Accept All while Reject All looked right (GitHub #6, #7, #8) | `851b03c` (0.11.1) | Semantic-diff placement in `engine/reconstruction-mapper.js`; bookmarks are zero-width; backstop `ACCEPT_NON_TEXT_CONTENT`. `tests/note_reference_preservation_tests.mjs` |
| Appending paragraphs after a target tracked the new last paragraph's mark, so Word's Reject All left an empty paragraph at the end of the body (add-in KB-9) | `e51bea6` (0.11.1) | The target's mark is the inserted break; the last new paragraph inherits the original mark. `tests/plain_text_paragraph_insertion_tests.mjs` |
| List regeneration of a container's last paragraph tracked the container's final mark (add-in KB-14) | `caeba20` (0.11.1) | `keepContainerFinalMarkUntracked`. `tests/container_final_mark_tests.mjs` |
| `convert-plain` turned `A. …<w:br/>B. …<w:br/>C. …` into one list item (add-in KB-11) | `851b03c` (0.11.1) | Now refused: `LIST_MARKER_UNSUPPORTED`, `reason: 'soft-break-markers'`. `tests/list_format_soft_break_tests.mjs` |
| A same-author delete or edit of a paragraph whose mark is that author's pending insertion (the original paragraph after "insert paragraphs after it") was refused with `REVISION_FIDELITY_VIOLATION` / `REJECT_PARAGRAPH_TEXT` (add-in KB-21) | 0.11.2 | The merge baseline keeps the paragraph's own text (`contentBaselineWithoutInsertedMark`, `engine/oxml-engine.js`). A deleted mark keeps its insertion (`w:ins` then `w:del`); the Reject/Accept projection removes it in both views. `tests/same_author_edit_after_inserted_mark_tests.mjs` |
| Deleting one of the author's own pending inserted paragraphs left an empty paragraph after Accept All | 0.11.2 | The emptied paragraph is removed; when it held the original paragraph's untracked mark, the previous inserted paragraph takes that mark (`removeEmptiedOwnInsertedParagraph`, `services/document-operation-mutations.js`). Same test file. |
| A text edit silently dropped pending same-author formatting (a bold `w:rPrChange`, an alignment `w:pPrChange`) from Accept All, alone or beside other pending revisions (add-in KB-19) | 0.11.2 | The same-author merge rejects text history only; `w:rPrChange` and `w:pPrChange` stay pending. Split runs get fresh revision ids. `tests/pending_same_author_run_formatting_tests.mjs` |
| Formatting a run whose formatting was already pending replaced the snapshot, so Reject All kept the first change; overlapping `format` operations in one batch failed with `RECEIPT_RECONCILIATION_FAILED` (add-in KB-18) | 0.11.2 | `snapshotAndAttachRPrChange` keeps a pending same-author snapshot and its id. `refreshRunPropertyChangeIds` no longer writes a second `w:id` attribute on a split run. Same test file. |
| Two text edits of one paragraph in one batch failed with `RECEIPT_RECONCILIATION_FAILED` | 0.11.2 | The second edit merges the first one's revisions; the first receipt stops listing the replaced ids and carries a warning (`supersedeReplacedRevisions`, `services/document-operation-applier.js`). |

The 0.11.2 rows were checked in Word 16 desktop on 2026-11-11 and 2026-11-14 (Accept All and Reject All text, paragraph
count, bold and italic; no repair prompt) with one-off scripts. They are not cases in the `test:word` lane.

### Still open

- **Targets that exist only after an earlier operation of the same batch resolve differently by operation type.** A
  `redline` whose `exactText` is the text an earlier operation produced resolves against the live document and applies. A
  `replace` or `format` with the same target is refused with `TARGET_NOT_FOUND`, because its target binds to the batch
  start. Both are safe; callers should target the batch-start text, or send dependent edits as separate calls.
- **A superseded receipt lists no revisions.** When a later operation of a batch re-edits a paragraph, the earlier
  operation's receipt keeps `finalDisposition: 'applied'` with the replaced ids removed from `revisionItems` and a
  warning. Consumers that map receipts to revisions one-to-one should read the later receipt.
- **Deleting an own inserted paragraph keeps an empty paragraph when paragraph properties differ.** The emptied paragraph
  is removed only when it carries the inserted mark, or when the previous inserted paragraph has identical paragraph
  properties. Otherwise the empty paragraph stays (Accept All shows it; Reject All is correct).
- **Another author's pending formatting blocks a text edit under `merge-same-author`.** A `w:rPrChange` by reviewer B in
  the target paragraph refuses author A's text edit with `EXISTING_REVISIONS`, although the two do not overlap. Use
  `existingRevisions: 'auto'` or `'slice-cross-author'`.
- **Pending table property changes are still merged to baseline.** A same-author text edit keeps `w:rPrChange` and
  `w:pPrChange`, but rejects that author's `w:tblPrChange`, `w:trPrChange` and `w:tcPrChange` in the edited scope, as
  before 0.11.2. No operation of this package writes these; they can come from Word.
- Generated numbering ids can still collide across separate calls whose runtime context is not given the document's numbering
  (`completed/2026-10-01-consumer-reported-library-issues.md`, WP-3 residual). Since `4718c36` this is detected, not silent: merging the
  parts through `mergeNumberingParts` / `ensureNumberingArtifactsInZip` (the single merge, used by `openDocx`) throws
  `NUMBERING_ID_CONFLICT` when an incoming id exists with different content. Bare callers must pass `numberingXml` (or a shared
  `numberingIdState`) to every call; no automatic remap of already-written paragraphs is attempted (deferred (aspirational, owner decision 2026-10-11)).

## 8. Consumer-Reported Markdown Reconciliation Hazards

The literal-underscore hazard reported here was fixed in 0.10.0 (see the fixed table in section 7).

### Passing a whole document as a redline scope can change untouched paragraphs

- **Where:** [`applyRedlineToOxml`](../../index.js) reconciles a paragraph or selected range. The [output shape guidance](../../README.md#output-shape-matrix) directs complete-document edits to `applyOperationToDocumentXml` or `applyOperationsToDocumentXml`.
- **Evidence (0.9.0 source):** Passing the entire `<w:document>` with two paragraphs (`**literal**` and `Old tail`) as the scope, while changing only `Old tail` to `New tail`, returns `status: 'ok'`. The first paragraph is split into an empty paragraph and a bold `literal` paragraph, so unchanged literal Markdown is reinterpreted and the paragraph structure changes.
- **Status and mitigation:** This is an unsafe whole-body call shape, not the intended targeted operation path. According to the supplied report, the add-in's integration tests now use `applyOperationsToDocumentXml` with targeted operations and production does not use this shape as of 2026-10-12.
- **Exit condition:** Keep complete-document consumers on the targeted document-operation API. Consider a guard or explicit scope declaration so a whole-body call cannot silently return `ok` after reinterpreting unrelated content.
