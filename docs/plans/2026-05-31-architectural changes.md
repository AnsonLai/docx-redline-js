# Redline Reliability Improvement Plan

This plan hardens the library against the ways Microsoft Word redlines are brittle.
It is written to be executed phase by phase, in order. Each phase is independently
shippable and ends with a green `npm test`.

**Scope note:** A high-level docx-in/docx-out wrapper API is explicitly OUT of scope.
Downstream tools own the packaging layer. Do not add a JSZip dependency or any
`applyRedlineToDocx`-style API to this package.

**Final audit status:** Complete. The plan has been reviewed top to bottom after
Phases 1-6. Stale status notes were removed, remaining namespace/metadata
convention gaps were closed, and the validation commands at the end of this plan
pass. The only intentionally unimplemented item is Phase 5.4 move-emission, which
is explicitly marked as a stretch/separate-PR item; move consumption is complete.

---

## Conventions (read before starting any phase)

- Tests live in `tests/*.mjs` and are auto-discovered by `scripts/run-tests.mjs`
  (run via `npm test`). Files in `tests/helpers/` and `tests/setup-xml-provider.mjs`
  are excluded from discovery. Follow the style of existing test files
  (e.g. `tests/revision_comment_management_tests.mjs`): plain `assert/strict`,
  no test framework.
- Shared assertion helpers belong in `tests/helpers/ooxml-assertions.mjs`. Add new
  helpers there rather than duplicating in test files.
- All XML element creation must go through `createWordElement` in `core/word-xml.js`
  so namespaces are correct. Never use `document.createElement` for `w:*` elements.
- Revision metadata (`w:id`, `w:author`, `w:date`) must come from
  `createRevisionMetadata(author)` in `core/types.js`. Never hand-roll these attributes.
- All public API additions must be exported through `index.js` (the only public surface,
  per ARCHITECTURE.md).
- After each phase: run `npm test` and `npm run test:isolation`. Both must pass.
- Do not change existing public function signatures. New behavior is added via new
  optional fields on options/result objects.

---

## Phase 1 — Round-trip invariant test harness (do this first)

**Status:** Complete.

- Added reusable structural assertions to `tests/helpers/ooxml-assertions.mjs`.
- Added `tests/helpers/roundtrip.mjs` with `assertRoundTrip(...)`.
- Added initial corpus in `tests/roundtrip_invariant_tests.mjs`.
- Fixed a Phase 1-discovered namespace bug in `engine/reconstruction-writer.js` by routing Word element creation through `createWordElement`.
- Added optional Word COM smoke script at `scripts/word-com-smoke.ps1`, exposed as `npm run smoke:word`.
- Verification: `npm test` passed (21/21); `npm run test:isolation` passed.

**Why:** The single most valuable check for redline correctness is:
*accepting all generated revisions must yield the modified text; rejecting all
generated revisions must yield the original text.* The library already owns both
halves of this loop (`applyRedlineToOxml` to generate, and
`acceptTrackedChangesInOoxml` / `rejectTrackedChangesInOoxml` in
`services/revision-comment-management.js` to resolve). Every later phase is
verified through this harness, so it lands first.

### 1.1 Build the harness helper

Create `tests/helpers/roundtrip.mjs` exporting:

```js
/**
 * Applies a redline, then asserts the accept/reject round-trip invariant.
 *
 * @param {string} oxml        - input OOXML (fragment, document, or package scope)
 * @param {string} original    - original plain text
 * @param {string} modified    - modified text (may contain markdown)
 * @param {object} [options]   - options forwarded to applyRedlineToOxml
 * @returns {Promise<{ redlined, accepted, rejected }>}
 */
export async function assertRoundTrip(oxml, original, modified, options = {})
```

Implementation steps inside `assertRoundTrip`:

1. Call `applyRedlineToOxml(oxml, original, modified, { generateRedlines: true, author: 'RoundTrip', ...options })`.
2. Assert the result parses as XML (use `parseXmlFragment` from
   `tests/helpers/ooxml-assertions.mjs`).
3. Run `acceptTrackedChangesInOoxml(result.oxml, { author: 'RoundTrip' })`, extract
   plain text with `ingestWordOoxmlToPlainText`, and assert it equals the
   *plain-text rendering* of `modified` (strip markdown markers the same way the
   engine does — reuse `preprocessMarkdown` from `pipeline/markdown-processor.js`
  to get `cleanText`). Compare with normalized whitespace
  (`s.replace(/\s+/g, ' ').trim()`).
4. Run `rejectTrackedChangesInOoxml(result.oxml, { author: 'RoundTrip' })`, extract
   plain text, assert it equals `original` (same whitespace normalization).
5. Structural assertions on the redlined output (add these as separate exported
   helpers so other tests can reuse them):
   - `assertNoNestedRevisions(xml)` — no `w:ins` inside `w:del` or vice versa.
   - `assertDelUsesDelText(xml)` — every `w:r` inside a `w:del` contains only
     `w:delText` (never `w:t`).
   - `assertRevisionMetadata(xml)` — every `w:ins`/`w:del` has non-empty `w:id`,
     `w:author`, and a `w:date` matching `/^\d{4}-\d{2}-\d{2}T/`.
   - `assertUniqueRevisionIds(xml)` — no duplicate `w:id` among `w:ins`/`w:del`/
     `w:rPrChange`/`w:pPrChange` elements in the output.
   - `assertSpacePreserved(xml)` — every `w:t`/`w:delText` whose text has leading
     or trailing whitespace carries `xml:space="preserve"`.

### 1.2 Build the corpus test

Create `tests/roundtrip_invariant_tests.mjs` that runs `assertRoundTrip` over a
corpus of (oxml, original, modified) cases. Reuse fixture inputs already present
in `tests/fixtures/` and `tests/sample_doc/` where possible. Minimum corpus
(each is one case; build the OOXML inline as template strings like the existing
tests do):

| # | Case |
|---|------|
| 1 | Single-run paragraph, one word replaced mid-sentence |
| 2 | Multi-run paragraph (3+ runs with different `w:rPr`), edit spanning a run boundary |
| 3 | Leading/trailing whitespace significant: replace `"foo "` with `"bar  baz "` |
| 4 | Pure insertion at start of paragraph; pure insertion at end |
| 5 | Pure deletion of an entire sentence |
| 6 | Edit inside a paragraph that contains a `w:hyperlink` (edit text *outside* the link) |
| 7 | Edit inside a table cell paragraph |
| 8 | Markdown formatting added: `**bold**` around an existing word |
| 9 | Paragraph containing `w:proofErr` markers and a simple field (`w:fldChar`/`w:instrText`) |
| 10 | Unicode: text with emoji and CJK characters replaced |
| 11 | Two consecutive edits: feed the redlined output of case 1 back through accept-all, then redline again (exercises re-entry on clean docs) |

If a case fails, do NOT weaken the assertion to make it pass — fix the engine or,
if the fix belongs to a later phase (e.g. hyperlink failures belong to Phase 4),
mark the case with a `// KNOWN-GAP: Phase N` comment and skip it with a logged
warning, so later phases un-skip it.

### 1.3 Optional Word smoke script (manual, not part of `npm test`)

Create `scripts/word-com-smoke.ps1` (Windows-only, requires desktop Word):
takes a `.docx` path, opens it via COM
(`New-Object -ComObject Word.Application`, `Documents.Open` with
`OpenAndRepair:$false`), reports whether Word opened it cleanly, counts
`document.Revisions.Count`, then closes without saving. Add an npm script
`"smoke:word": "powershell -File scripts/word-com-smoke.ps1"` and document it in
README under a new "Validating output" section. Do not wire it into CI.

**Acceptance for Phase 1:** new test file passes for all non-skipped cases;
helpers exported; `npm test` green.

---

## Phase 2 — Policy for pre-existing tracked changes in the source

**Status:** Complete.

- Added `containsTrackedChanges(xmlDoc)` in `core/word-xml.js` and exported it from `index.js`.
- Added the `existingRevisions` policy gate in `engine/oxml-engine.js`.
- Documented that `pipeline/ingestion-paragraph.js` records `w:delText` as a zero-width deletion model entry but excludes it from accepted text.
- Added `tests/existing_revisions_policy_tests.mjs`.
- Added README documentation for `existingRevisions`.
- Forwarded `existingRevisions` through `services/standalone-operation-runner.js` and updated prior-revision standalone tests to opt into `accept-all-first`.
- Final verification: `npm test` passed (21/21); `npm run test:isolation` passed.

**Why:** Running the engine over a paragraph that already contains `w:ins`/`w:del`
(from a human reviewer or a prior engine run) is the most common real-world
corruption source. Diff text extraction must treat `w:delText` as invisible and
`w:ins` content as visible, and the patcher must never nest revisions.

### 2.1 Detection

Add to `core/word-xml.js`:

```js
/**
 * Returns true if the document/fragment contains any revision markup:
 * w:ins, w:del, w:moveFrom, w:moveTo, w:rPrChange, w:pPrChange,
 * w:cellIns, w:cellDel, or a w:del/w:ins inside w:pPr/w:rPr (paragraph mark).
 */
export function containsTrackedChanges(xmlDoc)
```

Use `getElementsByTagNameNS(NS_W, localName)` per element name (see how
`services/revision-comment-management.js` does namespace-safe lookups with
`getWordElementsByLocalName`).

### 2.2 Engine policy gate

In `engine/oxml-engine.js`, immediately after the existing parse-error check in
`applyRedlineToOxml`, add:

1. Call `containsTrackedChanges(xmlDoc)`.
2. If true, behavior is controlled by a new option
   `options.existingRevisions` with values:
   - `'reject-input'` (default): return
     `{ oxml, hasChanges: false, status: 'error', error: { code: 'EXISTING_REVISIONS', message: ... } }`
     (the `status` field is introduced in Phase 6.1 — if Phase 6.1 is not yet done,
     implement the `status`/`error` fields now as part of this step; Phase 6.1
     then only extends them to other early-return paths).
   - `'accept-all-first'`: run `acceptTrackedChangesInOoxml(oxml, { allAuthors: true })`,
     re-parse, and proceed with the cleaned document. The caller's `original` text
     must then match the post-accept text (the normal targeting logic already
     verifies this and falls back to no-change if it doesn't).
3. Log which path was taken via the `log` adapter.

Do NOT attempt to diff *through* existing revisions in this phase. Normalizing
first (or refusing clearly) is the reliable behavior; transparent merge of new
revisions into already-revised text is out of scope.

### 2.3 Audit ingestion's treatment of `w:del`

`pipeline/ingestion-paragraph.js` (around lines 320–340) collects `w:delText`
content when flattening runs. Audit every caller of that code path and confirm
deleted text is **excluded** from the plain text used for diffing and from
`ingestWordOoxmlToPlainText` output (deleted text is invisible in Word's
"accepted" view and must not appear in `original` matching). If it is currently
included anywhere, fix it and add a regression test. If it is intentionally
included for some revision-management path, add a comment at the collection site
stating which caller needs it and why.

### 2.4 Tests

Create `tests/existing_revisions_policy_tests.mjs`:

- Paragraph containing a `w:del` + `w:ins` pair → default call returns
  `status: 'error'`, code `EXISTING_REVISIONS`, original oxml unchanged.
- Same input with `existingRevisions: 'accept-all-first'` and `original` set to
  the post-accept text → succeeds, and the result passes `assertRoundTrip`
  structural checks from Phase 1.
- `ingestWordOoxmlToPlainText` on a paragraph with `w:del` returns text WITHOUT
  the deleted content, and WITH `w:ins` content.
- `containsTrackedChanges` unit tests: positive for each marker type listed in
  2.1 (including paragraph-mark `w:del` inside `w:pPr/w:rPr`), negative for a
  clean paragraph and for a paragraph with only comments/bookmarks.

**Acceptance:** new tests pass; case 11 from Phase 1 still passes; README API
table gains a row for the `existingRevisions` option.

---

## Phase 3 — Paragraph-mark revisions

**Status:** Complete.

- Added `markParagraphMarkInserted(...)` and `markParagraphMarkDeleted(...)` in `engine/run-builders.js`.
- Wired reconstruction-mode paragraph-boundary insert/delete output to add `w:pPr/w:rPr/w:ins` and `w:pPr/w:rPr/w:del` paragraph-mark revisions.
- Fixed reconstruction writer paragraph state propagation across newline boundaries while wiring paragraph marks.
- Wired text-to-table transformation source paragraphs to mark deleted paragraph marks when redlines are generated.
- Extended `acceptTrackedChangesInOoxml` / `rejectTrackedChangesInOoxml` to handle paragraph-mark `w:ins`/`w:del` separately from normal run-level revisions.
- Added `tests/paragraph_mark_revision_tests.mjs` for inserted/deleted paragraph round trips and structural assertions.
- Final verification: `npm test` passed (21/21); `npm run test:isolation` passed.

**Why:** When an edit inserts or deletes whole paragraphs (or splits/merges
them), the paragraph *mark* itself must be revised. A deleted paragraph's mark
needs `w:pPr > w:rPr > w:del`; an inserted paragraph's mark needs
`w:pPr > w:rPr > w:ins`. Without this, accept/reject in Word leaves stray empty
paragraphs or fails to merge paragraphs — highly visible breakage.

### 3.1 Builders

Add to `engine/run-builders.js`:

```js
/** Marks a paragraph's mark as inserted: ensures w:pPr exists, ensures w:rPr
 *  inside it, appends <w:ins w:id w:author w:date/> (empty element). */
export function markParagraphMarkInserted(xmlDoc, paragraph, author)

/** Marks a paragraph's mark as deleted: same shape with <w:del/>. */
export function markParagraphMarkDeleted(xmlDoc, paragraph, author)
```

Rules:
- `w:pPr` must be the FIRST child of `w:p`; `w:rPr` must be the LAST child of
  `w:pPr` per schema order (check existing pPr handling in
  `engine/reconstruction-writer.js` for how the codebase orders pPr children,
  and reuse any existing ordering helper).
- Use `createRevisionMetadata` for attributes. Remove any pre-existing
  `w:ins`/`w:del` in that `w:rPr` before adding (idempotent).

### 3.2 Wire into the engine

Find every site that inserts or removes a whole `w:p` while
`generateRedlines` is true. Search hints:
`grep -n "createElement.*w:p\b\|appendChild(paragraph\|removeChild(paragraph" engine/ pipeline/`
plus read `engine/reconstruction-mode.js`, `engine/reconstruction-writer.js`,
`pipeline/list-generation.js`, and `engine/table-mode.js` (row/cell paragraph
creation). At each site:

- New paragraph created as part of a redline → call `markParagraphMarkInserted`.
- Paragraph whose entire content is wrapped in `w:del` (paragraph is going away
  on accept) → call `markParagraphMarkDeleted` instead of removing the `w:p` node.
  The paragraph node must REMAIN in the document with its content in `w:del` —
  Word removes it on accept.

### 3.3 Accept/reject support

In `services/revision-comment-management.js`, verify (and fix if missing) that:

- **Accept** of a paragraph-mark `w:del` merges the paragraph with the following
  paragraph (move this paragraph's remaining children, except `w:pPr`, into the
  next `w:p`, then remove this `w:p`; if it is the last paragraph in its parent,
  just remove the mark revision).
- **Reject** of a paragraph-mark `w:del` simply removes the `w:del` element
  from `w:pPr/w:rPr`.
- **Accept** of a paragraph-mark `w:ins` removes the `w:ins` element.
- **Reject** of a paragraph-mark `w:ins` merges the paragraph into the next one
  (inverse of accept-del).

### 3.4 Tests

Create `tests/paragraph_mark_revision_tests.mjs`:

- Modified text adds a new paragraph (`"one"` → `"one\n\ntwo"` or via markdown
  list): output's new `w:p` has `w:pPr/w:rPr/w:ins`; accept-all yields two
  paragraphs; reject-all yields one paragraph with original text.
- Modified text deletes a paragraph (`"one\n\ntwo"` → `"one"`): the second `w:p`
  still exists in the redlined output, its runs are in `w:del`, its mark has
  `w:pPr/w:rPr/w:del`; accept-all yields one paragraph; reject-all yields two.
- Round-trip via `assertRoundTrip` for both, with multi-paragraph-aware text
  comparison (join paragraphs with `\n`).

**Acceptance:** new tests pass; no existing test regresses (list generation
tests in `tests/list_tests.mjs` are the most likely to be affected — if they
assert exact XML, update expectations to include the new mark revisions).

---

## Phase 4 — Inert and structural markup safety

**Status:** Complete.

- Added reliability coverage for hyperlink-contained revisions, bookmark/comment marker preservation, `w:tab` survival, and footnote reference preservation.
- De-duplicated zero-width comment marker replay in reconstruction mode.
- Updated run builders to synthesize visible `w:tab`, `w:br`, and `w:noBreakHyphen` elements instead of writing those characters into plain `w:t` text.
- Added reconstruction preservation for missing footnote/endnote placeholder references when modified text edits adjacent content without explicitly including internal tokens.
- Added README packaging note for hyperlink/bookmark/comment/tab/break/footnote safety.
- Final verification: `npm test` passed (21/21); `npm run test:isolation` passed.

**Why:** README already documents stripping `w:fldChar`/`w:instrText`/`w:proofErr`
from the matched paragraph before diffing. The same care is needed for other
non-text and container markup, or redlines split/orphan them and Word repairs
(or rejects) the file.

Work through these sub-items one at a time, each with its own tests appended to
`tests/engine_reliability_tests.mjs` (this file already exists):

### 4.1 Hyperlinks (`w:hyperlink`)

- **Invariant:** runs created by splitting a run that lives inside a
  `w:hyperlink` must remain inside that same `w:hyperlink` element. New `w:ins`/
  `w:del` wrappers go INSIDE the hyperlink, wrapping the runs.
- Audit `engine/surgical-run-splitting.js` and `engine/surgical-diff-application.js`:
  wherever a new node is inserted with `insertBefore(node, run)` or appended to
  `run.parentNode`, the parent may be `w:hyperlink`, not `w:p`. Verify insertion
  uses `run.parentNode` (correct) and never hoists to the paragraph level
  (incorrect). Fix any site that assumes the run's parent is `w:p`.
- Test: paragraph `before [link text] after`; edit `link` → `hyperlink`. Assert
  the `w:ins`/`w:del` elements are descendants of `w:hyperlink`, `r:id`
  attribute is untouched, and round-trip holds. Also test an edit that spans
  from before the hyperlink into it (this may legitimately fall back to
  reconstruction mode — then assert the hyperlink element survives in output).

### 4.2 Bookmarks and comment range markers

- `w:bookmarkStart`/`w:bookmarkEnd`/`w:commentRangeStart`/`w:commentRangeEnd`/
  `w:commentReference` must SURVIVE the edit (never deleted, never wrapped in
  `w:del`) and must not contribute characters to diff text.
- Audit `pipeline/ingestion-paragraph.js` (text extraction) and the surgical
  splitting path. When a run range being replaced contains such markers as
  siblings, the markers must be left in place between the `w:del` and `w:ins`
  output.
- Test: paragraph with a comment range spanning a word that gets edited; assert
  `commentRangeStart/End/Reference` still present exactly once each, and
  `tests/comment_tests.mjs` still passes.

### 4.3 `w:lastRenderedPageBreak`, `w:tab`, `w:br`, `w:noBreakHyphen`

- `w:lastRenderedPageBreak` is render cache: safe to strip from the matched
  paragraph before diffing (extend the existing fldChar/proofErr stripping site —
  find it with `grep -rn "proofErr" engine/ pipeline/`).
- `w:tab` and `w:br` are VISIBLE content. Decide and document one mapping in
  ingestion: `w:tab` → `\t`, `w:br` → `\n` in extracted plain text, and ensure
  the diff/patch path can reproduce them (if the engine cannot synthesize them
  on the insert side, at minimum it must not corrupt runs containing them:
  verify a no-op edit elsewhere in the paragraph leaves them intact).
- Test: paragraph `A<w:tab/>B`, edit `B`→`C`; assert `w:tab` survives and
  round-trip holds on the textual parts.

### 4.4 Footnote/endnote references

- A run containing `w:footnoteReference`/`w:endnoteReference` must never be
  split through, deleted, or duplicated by the diff (deleting it would orphan
  the footnote part in a way this package cannot clean up).
- Implement: treat such runs like field scaffolding — exclude them from the
  editable span ranges in `engine/surgical-spans.js` (anchor text before/after
  them, same approach as existing field handling).
- Test: edit text after a footnote reference; assert exactly one
  `w:footnoteReference` with unchanged `w:id` in output.

**Acceptance:** all Phase 4 tests pass; un-skip Phase 1 corpus cases marked
`KNOWN-GAP: Phase 4`; add a sentence to README's packaging Do/Don't section
noting hyperlink/bookmark/footnote safety.

---

## Phase 5 — Move revision (`w:moveFrom` / `w:moveTo`) consumption

**Status:** Complete for move consumption.

- Ingestion treats `w:moveFrom` as deleted/invisible text and `w:moveTo` as inserted/visible text.
- `ingestWordOoxmlToPlainText` excludes moved-from text while retaining moved-to text.
- `acceptTrackedChangesInOoxml` now removes `w:moveFrom`, unwraps `w:moveTo`, and removes matching move range markers.
- `rejectTrackedChangesInOoxml` now unwraps `w:moveFrom` after converting `w:delText` back to `w:t`, removes `w:moveTo`, and removes matching move range markers.
- `containsTrackedChanges` detects move range markers in addition to `w:moveFrom`/`w:moveTo`.
- Added `tests/move_revision_tests.mjs`.
- Final verification: `npm test` passed (21/21); `npm run test:isolation` passed.

**Why:** Documents from human reviewers contain move revisions. The library
currently has zero handling, so ingestion, accept/reject, and the Phase 2
detection gate would mis-handle them. Goal of this phase is to CONSUME moves
safely. (Emitting moves from the differ is a stretch goal — see 5.4 — do not
start it unless 5.1–5.3 are done and green.)

### 5.1 Ingestion

In `pipeline/ingestion-paragraph.js` (and the table equivalent if it reads runs
independently): treat `w:moveFrom` content as deleted (excluded from plain
text, consistent with Phase 2.3) and `w:moveTo` content as inserted (included).
`w:moveFromRangeStart/End`, `w:moveToRangeStart/End` are markers — ignore for
text, preserve as nodes (Phase 4.2 rules).

### 5.2 Accept/reject

In `services/revision-comment-management.js` extend both transforms:

- **Accept:** `w:moveFrom` → remove element and contents (like `w:del`);
  `w:moveTo` → unwrap (like `w:ins`); remove all four range marker types for the
  matched author.
- **Reject:** `w:moveFrom` → unwrap, converting any `w:delText` inside back to
  `w:t`; `w:moveTo` → remove element and contents; remove range markers.
- Reuse the existing `removeNode`/`unwrapNode` helpers and the author-filter
  machinery (`resolveAuthorFilter`/`authorMatchesNode`) already in that file.
  Note: move *range markers* carry their author on the `RangeStart` element;
  the matching `RangeEnd` has only an id — match ends to starts by `w:id`.

### 5.3 Tests

Create `tests/move_revision_tests.mjs`: a two-paragraph fixture where a sentence
is wrapped in `w:moveFrom` (+ range markers) in paragraph 1 and `w:moveTo`
(+ markers) in paragraph 2.

- Accept-all → sentence appears only in paragraph 2; no move markup remains.
- Reject-all → sentence appears only in paragraph 1; no move markup remains.
- Author-filtered accept with a non-matching author → fixture unchanged.
- `ingestWordOoxmlToPlainText` shows the moved sentence exactly once (at the
  moveTo location).
- `containsTrackedChanges` (Phase 2) returns true for this fixture.

### 5.4 (Stretch, separate PR) Move emission from the differ

Only after 5.1–5.3: in `pipeline/diff-engine.js`, post-process the diff to find
delete/insert pairs with identical normalized text ≥ 15 characters; emit them as
`w:moveFrom`/`w:moveTo` pairs sharing a `w:name` attribute and linked range
markers. Verify via the Phase 1 harness plus manual Word inspection. If this
proves unstable, ship 5.1–5.3 alone — consumption is the safety-critical half.

**Acceptance:** 5.1–5.3 tests pass; README revision-management rows updated to
mention move support.

---

## Phase 6 — Hardening details

**Status:** Complete.

- Added non-breaking `status` / `error` fields to the engine result path, including `PARSE_ERROR`, `TARGET_NOT_FOUND`, and `EXISTING_REVISIONS`.
- Seeded generated revision IDs from existing document IDs via `seedRevisionIdsFromDocument(xmlDoc)`.
- Seeded revision IDs in `applyRedlineToOxml` and the standalone redline operation runner.
- Added `tests/hardening_status_tests.mjs`.
- Added `index.d.ts`, package `"types"` metadata, and `npm run check:types`.
- Added `scripts/export-validation-fixtures.mjs` and `docs/VALIDATION.md`, linked from README.
- README now documents `status`/`error`, move revision consumption, included types, and validation workflow.
- Final verification: `npm test` passed (21/21); `npm run test:isolation` passed; `npm run check:types` passed; `node scripts/export-validation-fixtures.mjs` passed.

### 6.1 Explicit error status instead of silent no-op

`applyRedlineToOxml` in `engine/oxml-engine.js` returns `{ oxml, hasChanges: false }`
on XML parse failure (the `noChanges()` early returns near the top), which is
indistinguishable from "nothing to change."

- Add optional result fields, non-breaking:
  `status: 'ok' | 'no-op' | 'error'` and
  `error?: { code: string, message: string }`.
- Error codes to introduce: `PARSE_ERROR`, `TARGET_NOT_FOUND` (where targeting
  fails and the engine currently logs + returns unchanged), `EXISTING_REVISIONS`
  (Phase 2). Successful-but-unchanged paths return `status: 'no-op'`.
- Thread the same fields through `applyRedlineToOxmlWithListFallback`,
  `reconcileMarkdownTableOoxml`, and
  `services/standalone-operation-runner.js` (`applyOperationToDocumentXml`
  result already has its own shape — add `status`/`error` alongside, do not
  rename existing fields).
- Tests: feed malformed XML → `status: 'error'`, code `PARSE_ERROR`; feed
  `original` text that doesn't exist in the document → `status` is `'error'` or
  `'no-op'` per the chosen semantics (pick one, document it in the JSDoc).
- Update README API tables to document `status`/`error`.

### 6.2 Document-unique revision IDs

`core/types.js` uses a module-global counter starting at 1000
(`revisionIdCounter`). A source document that already contains `w:id` values
≥ 1000 can collide with newly generated ids.

- Add `export function seedRevisionIdsFromDocument(xmlDoc)` in `core/types.js`:
  scan all elements for a `w:id` attribute, parse as integer, and if
  `maxFound >= revisionIdCounter`, set `revisionIdCounter = maxFound + 1`.
- Call it once in `applyRedlineToOxml` right after successful parse, and at the
  equivalent spot in `services/standalone-operation-runner.js`.
- Test: input fragment containing `w:ins w:id="5000"`; assert every generated
  revision id in the output is > 5000 and `assertUniqueRevisionIds` (Phase 1)
  passes.

### 6.3 TypeScript declarations

- Create `index.d.ts` at the repo root typing every export of `index.js`.
  Key shapes: the options bag for `applyRedlineToOxml`
  (`{ generateRedlines?, author?, targetParagraphId?, existingRevisions? }`),
  the result (`{ oxml, hasChanges, sourceType?, status?, error? }`), the
  accept/reject options (`{ author?, allAuthors? }`) and their result shape
  (read the actual return in `services/revision-comment-management.js` —
  it includes warnings), and the config functions.
- Add `"types": "index.d.ts"` to `package.json` and include it in the published
  `files` list if one exists.
- Verification: add a `scripts/check-types.mjs` step or simply run
  `npx tsc --noEmit --checkJs false index.d.ts` once locally; at minimum ensure
  the file parses (`npx tsc index.d.ts --noEmit`). Do not convert the codebase
  to TypeScript.

### 6.4 Cross-consumer fixture validation (manual, documented)

- Add `scripts/export-validation-fixtures.mjs`: runs a handful of Phase 1 corpus
  cases through `applyOperationToDocumentXml`, wraps each result in a minimal
  `.docx` using the existing package-builder/plumbing helpers
  (`services/package-builder.js`, `validateDocxPackage`), and writes them to
  `tmp/validation-docx/`. Use only existing dependencies — if zip writing isn't
  possible with what's in `package.json`, write the `word/document.xml` parts
  plus a README instructing how to assemble, instead of adding a dependency.
- If LibreOffice is installed locally, `soffice --headless --convert-to pdf`
  over the folder is a cheap "does another consumer parse it" check — document
  this (and the Word COM script from 1.3) in a new `docs/VALIDATION.md`, linked
  from README. This stays a manual/release-time step, not CI.

**Acceptance:** 6.1 and 6.2 tests pass; `index.d.ts` exists and parses;
`docs/VALIDATION.md` exists; README updated.

### Final verification

- `npm test` passed (21/21).
- `npm run test:isolation` passed.
- `npm run check:types` passed.
- `node scripts/export-validation-fixtures.mjs` passed and wrote fixtures to
  `tmp/validation-docx/`.
- Convention audit passed for `core/`, `engine/`, `pipeline/`, `services/`, and
  `index.js`: no direct `document.createElement('w:*')` or
  `createElementNS(NS_W, 'w:*')` call sites remain outside `createWordElement`,
  and tracked-change revision metadata is routed through `createRevisionMetadata`.
- Word COM validation remains a manual/release-time smoke step because it needs
  Microsoft Word and an exported `.docx` path: `npm run smoke:word -- path/to/file.docx`.

---

## Execution order and dependencies

```
Phase 1  (harness)            ← everything depends on this
Phase 2  (existing revisions) ← needs 1; introduces status/error early if 6.1 not done
Phase 3  (paragraph marks)    ← needs 1
Phase 4  (inert markup)       ← needs 1; un-skips Phase 1 KNOWN-GAP cases
Phase 5  (moves)              ← needs 1, 2
Phase 6  (hardening)          ← 6.1/6.2 anytime after 1; 6.3/6.4 last
```

Each phase should be a separate commit (or PR) with the test suite green.
Suggested commit messages: `test: add accept/reject round-trip harness`,
`feat: gate redlining on pre-existing revisions`, `feat: revise paragraph marks
on paragraph insert/delete`, `fix: preserve hyperlinks, bookmarks, and footnote
refs across redlines`, `feat: consume w:moveFrom/w:moveTo in ingestion and
accept/reject`, `feat: explicit status/error result fields and seeded revision ids`.

## Global guardrails

- Never emit `w:t` inside `w:del` (must be `w:delText`) — covered by
  `assertDelUsesDelText`; run it on every new output-producing test.
- Never nest `w:ins`/`w:del` inside each other.
- Never produce an empty `w:t`/`w:delText` element or an empty `w:ins`/`w:del`
  wrapper.
- Word merges adjacent same-author revisions in its review pane only when the
  author strings are byte-identical — always source the author through
  `createRevisionMetadata`, never trim/transform it at call sites.
- When in doubt about element ordering inside `w:rPr`, use `RPR_SCHEMA_ORDER`
  from `engine/rpr-helpers.js`; for `w:pPr`, `w:pPr` is first child of `w:p` and
  `w:rPr` is its last child.
