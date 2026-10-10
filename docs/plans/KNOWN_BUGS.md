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

### Current Two-Tier Architecture (v0.5.x)
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
`refused`, nothing written). Current limits, tracked in `2026-09-30-revision-fidelity-invariants.md`:

- The default mode is `enforce` (refuse). `verifyRevisionFidelity: 'warn'` applies and reports instead; `'off'` is for
  diagnosis only and the CLI agent profile rejects it (`--revision-fidelity`, capability `revision-fidelity-v1`).
- Untracked operations (`generateRedlines: false`) are not verified.
- Accept All intent and neighbor checks (`ACCEPT_INTENT`, `ACCEPT_NEIGHBORS`) and the numbering-part existence check are not
  implemented; `REJECT_NUMBERING_REFS` only compares restored `numId`/`ilvl` with the source.

### Fixed gap bugs (do not re-report)

| Bug | Fixed in | Note |
|---|---|---|
| `1. Header` list conversion on a bare `document.xml` returned `RECEIPT_RECONCILIATION_FAILED` (still listed under 0.8.4 known issues in the released changelog) | `73c67e8` | Bare runner seeds numbering state; an aborted conversion records no receipt revisions. |
| Explicit-start / `1. Header` list conversion bound `numPr` without `w:pPrChange`, so Reject All left a numbered item | `f72f60a` | `orchestration/list-structural-fallback.js` |
| Text-to-table redline put tracked cell text in untracked rows, so Reject All kept the table | `c20f33c` | `engine/table-mode.js`; `pipeline/pipeline.js:~340` has the same call without `trackAsBlock` and is untested. |
| Library Reject/Accept differed from Word for cell paragraph marks, deleted tables, and `pPrChange` paragraph marks | `cedc68a` | `tests/word_revision_semantics_tests.mjs` |
| Numbering helpers and inspection read `numPr`/`pStyle` from `w:pPrChange` snapshots as current | `1a3eeb6`..`1f89ae0` | `core/current-properties.js` plus usage lint |

### Still open

- Generated numbering ids can collide across separate calls with fresh runtime contexts and no numbering merge between them
  (`2026-10-01-consumer-reported-library-issues.md`, WP-3 residual; owner: the canonical-list allocator).
- Lettered and roman marker conversion still uses whole-paragraph delete plus insert (loses `w14:paraId` and bookmarks). It is
  reject-exact, so the verifier accepts it. Tracked in `2026-09-30-canonical-list-operations.md`.
