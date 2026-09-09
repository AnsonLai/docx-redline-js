# Cross-Author Revision Slicing: Nested Revisions via Carrier Splitting

**Status:** Completed — WP-01 through WP-07 implemented; final automated verification complete  
**Date:** 2026-09-08  
**Target releases:** v0.6.0–v1.0.0  
**Priority:** Document fidelity, multi-author contract negotiation accuracy, and strict OOXML schema compliance over single-pass simplicity.

---

## 1. Executive Summary & Findings

### The Problem
During contract review and multi-turn legal negotiations, Reviewer B often needs to edit text that was previously inserted by Reviewer A whose revision has not yet been accepted.

In previous discussions and documentation, it was assumed that deleting or inserting text inside another author's pending insertion was either impossible in WordprocessingML (OOXML) or logically paradoxical. However, empirical inspection of Microsoft Word Desktop proves otherwise:

1. **Word Desktop supports deletions inside pending insertions**: When Reviewer B deletes text inside Reviewer A's pending `<w:ins>`, Word Desktop displays Reviewer B's deletion visibly (strikethrough formatting and a deletion tooltip attributed to Reviewer B), while the surrounding insertion text remains attributed to Reviewer A.
2. **Word Desktop supports insertions inside pending insertions**: When Author B types text in the middle of Author A's `<w:ins>`, Word Desktop displays Author B's new text attributed to Author B, while preserving Author A's attribution on the preceding and succeeding words.

### The Underlying OOXML Mechanics
Empirical inspection of Microsoft Word Desktop 365 (via automated COM fixture generation in WP-01) reveals a vital asymmetry between how Word handles cross-author insertions versus deletions:

1. **Cross-Author Insertions (`insert-interior`) — Sibling Splitting**:
   ECMA-376 Part 1 `CT_RunTrackChange` does NOT allow `<w:ins>` inside `<w:ins>`. Word Desktop splits the carrier `<w:ins>` into sibling fragments at the paragraph (`<w:p>`) level, splicing Author B's new `<w:ins>` between them:
   ```xml
   <!-- Sibling 1: Author A's insertion (leading fragment) -->
    <w:ins w:id="0" w:author="Reviewer A" w:date="2026-09-08T09:29:00Z">
       <w:r><w:t xml:space="preserve">amended by this </w:t></w:r>
   </w:ins>
   <!-- Sibling 2: Author B's insertion spliced in between -->
    <w:ins w:id="1" w:author="Reviewer B" w:date="2026-09-08T09:29:00Z">
       <w:r><w:t xml:space="preserve">MASTER </w:t></w:r>
   </w:ins>
   <!-- Sibling 3: Author A's insertion (trailing fragment) -->
    <w:ins w:id="2" w:author="Reviewer A" w:date="2026-09-08T09:29:00Z">
       <w:r><w:t>Agreement.</w:t></w:r>
   </w:ins>
   ```

2. **Cross-Author Deletions (`delete-interior`) — Direct `<w:del>` Nesting inside `<w:ins>`**:
   Under ECMA-376 Part 1 Section 17.13.5.21 (`CT_RunTrackChange`), `<w:del>` is an explicitly permitted child element of `<w:ins>`. Microsoft Word Desktop does **not** split `<w:ins>` for deletions; instead, it nests `<w:del>` directly inside `<w:ins>`:
   ```xml
    <w:ins w:id="0" w:author="Reviewer A" w:date="2026-09-08T09:29:00Z">
       <w:r><w:t xml:space="preserve">The Services will process the Input to </w:t></w:r>
        <w:del w:id="1" w:author="Reviewer B" w:date="2026-09-08T09:29:00Z">
           <w:r w:rsidDel="00F70C09"><w:delText xml:space="preserve">generate </w:delText></w:r>
       </w:del>
       <w:r><w:t>outputs for Customer.</w:t></w:r>
   </w:ins>
   ```
   This nested structure ensures that:
   - If Author A's insertion is rejected, the entire container—including Author B's deletion of that unaccepted text—is discarded cleanly.
   - If Author A's insertion is accepted, Author A's text is unwrapped into the baseline while Author B's `<w:del>` remains pending against the baseline.

### Why `docx-redline-js` Currently Blocks This
`docx-redline-js` currently guards against this with `EXISTING_REVISIONS` / `UNSAFE_REVISION_NESTING` because:
1. **Container Splitting Is Unimplemented for Insertions**: In `engine/surgical-diff-application.js`, `processInsert` inserts newly created revision elements into `span.runElement.parentNode`. When that parent is `<w:ins>`, inserting `<w:ins>` directly into it produces schema-invalid nested `<w:ins><w:ins>...</w:ins></w:ins>`.
2. **Overzealous Validation Rule**: In `core/redline-validation.js`, an existing check unconditionally flagged any `<w:del>` inside `<w:ins>` as `NESTED_REVISION`, contradicting ECMA-376 and Microsoft Word Desktop's native behavior.
3. **Same-Author vs. Cross-Author Asymmetry**: The engine already has `merge-same-author`, but that strategy works by *reverting* prior revisions back to the pre-revision baseline and re-diffing against the baseline into a single revision. Applying that to another author would erase Author A's attribution.

This plan defines the architecture, mutation mechanics, lifecycle rules, and test strategy to implement **Cross-Author Revision Slicing** natively.

---

## 2. Hard Invariants

1. **Strict OOXML Schema Compliance (Word-Native Hierarchy)**:
   - `<w:del>` INSIDE `<w:ins>` IS allowed (and expected by Word Desktop and ECMA-376 `CT_RunTrackChange`).
   - `<w:ins>` must NEVER contain `<w:ins>`, `<w:moveFrom>`, or `<w:moveTo>`.
   - `<w:del>` must NEVER contain `<w:ins>`, `<w:del>`, `<w:moveFrom>`, or `<w:moveTo>`.
   - All revision wrappers must be immediate children of `<w:p>`, `<w:hyperlink>`, or (for `<w:del>`) `<w:ins>`.

2. **Durable ID Allocation & Metadata Preservation**:
   - Split fragments of Author A's `<w:ins>` MUST preserve Author A's exact author string, date, and `dateUtc`.
   - The first fragment may retain the original `w:id` if non-conflicting, but subsequent split fragments MUST be allocated globally unique `w:id` values via the active `RevisionIdAllocator`.
   - Author B's new revision element (`<w:del>` or `<w:ins>`) MUST be allocated its own fresh `w:id` under Author B's name and current timestamp.

3. **Multi-Author Revision Lifecycle Consistency**:
   - **Accept All**: Removing `<w:del>` and unwrapping all `<w:ins>` must produce the identical text string as Word Desktop.
   - **Accept Author A Only**: Author A's split `<w:ins>` elements unwrap into baseline text. Author B's `<w:del>` and `<w:ins>` remain pending against that now-baseline text.
   - **Reject Author A Only**: Author A's `<w:ins>` elements are removed. Any `<w:del>` by Author B situated inside Author A's insertion MUST be removed as well (you cannot delete text that was rejected from ever entering the document). Any `<w:ins>` by Author B remains as a pending insertion in the surrounding baseline.
   - **Accept Author B Only**: Author B's `<w:del>` is removed; Author B's `<w:ins>` unwraps into text inside Author A's pending insertion. Author A's insertion fragments remain pending.
   - **Reject Author B Only**: Author B's `<w:del>` is converted back into regular runs and coalesced back into Author A's `<w:ins>`. Author B's `<w:ins>` is removed.

4. **Zero Empty Containers**:
   - If an edit touches the exact boundary of Author A's `<w:ins>`, no empty `<w:ins>` (container with 0 text characters) may be emitted.

---

## 3. Architecture & Technical Design

```
                     Target Paragraph OOXML
                               |
                               v
            Build Surgical Text Spans (Surgical Spans)
      [Identifies text, offsets, and carrier revision element]
                               |
                               v
           Target Range Overlaps Foreign Author <w:ins>
                               |
            +------------------+------------------+
            |                                     |
            v                                     v
     [Deletion Edit]                       [Insertion Edit]
  1. Keep carrier <w:ins> intact        1. Split carrier <w:ins>
  2. Extract deleted text                  at insertion offset
     into <w:delText>                    2. Create Author B's
  3. Create Author B's                     <w:ins> with new ID
     <w:del> with new ID                3. Splice at carrier level:
  4. Nest <w:del> directly inside          [Author A Left <w:ins>]
     Author A's <w:ins>                    [Author B New <w:ins>]
                                             [Author A Right <w:ins>]
                               |
                               v
                  Output Reconciliation Oracle
        (Assert valid schema, unique IDs, only permitted nesting)
```

### 3.1 Container Splitting Primitive: `splitTrackChangeCarrier`
A new utility in `engine/surgical-run-splitting.js`:
```js
export function splitTrackChangeCarrier(xmlDoc, carrierElement, splitOffset, allocator) {
    // 1. Validate the <w:ins> carrier and preserve all carrier metadata.
    // 2. Locate run and character offset corresponding to splitOffset.
    // 3. Clone carrierElement into leftCarrier and rightCarrier.
    // 4. Distribute child runs before splitOffset to leftCarrier, after to rightCarrier.
    // 5. Allocate new unique revision ID for rightCarrier.
    // 6. Return { leftCarrier, rightCarrier }.
}
```

### 3.2 Splicing in `processDelete`
When deleting a range `[startPos, endPos]`:
1. Check if the affected spans belong to a `<w:ins>` container.
2. If `span.runElement.parentNode` is `<w:ins>`:
   - Identify whether the author of `<w:ins>` differs from the mutation `author`.
   - Under cross-author slicing mode, keep the foreign `<w:ins>` intact.
   - Create `<w:del w:author="AuthorB">` containing `<w:r><w:delText>deleted text</w:delText></w:r>`.
   - Insert the `<w:del>` inside the carrier at the deleted run position, preserving unaffected runs on either side.

### 3.3 Splicing in `processInsert`
When inserting at position `pos`:
1. Check if `pos` falls inside an existing `<w:ins>`.
2. If `targetSpan.runElement.parentNode` is `<w:ins>`:
   - Split carrier `<w:ins>` into `left_ins` and `right_ins` at `pos`.
   - Create `new_ins` for Author B with Author B's text runs.
   - Insert `[left_ins, new_ins, right_ins]` into the common `<w:p>`.
   - Remove original carrier `<w:ins>`.

### 3.4 Straddle Deletions (Boundary Spanning)
When Author B's deletion starts in baseline text and ends inside Author A's `<w:ins>` (or vice-versa):
* The deletion must be partitioned:
  1. The baseline portion emits a normal `<w:del>` child of `<w:p>`.
  2. The insertion portion remains inside Author A's `<w:ins>` and emits a nested `<w:del>` there.
  3. The baseline and insertion portions cannot be coalesced across the carrier boundary; preserve a top-level `<w:del>` plus a nested `<w:ins><w:del>...</w:del></w:ins>`, matching the Word Desktop fixture.

---

## 4. Policy Configuration & Migration

We introduce a dedicated policy setting on `existingRevisions`:

```ts
type ExistingRevisionsPolicy = 
    | 'merge-same-author'        // Default in v0.5.x: merges same author, fails cross-author
    | 'slice-cross-author'       // NEW: merges same author, slices cross-author insertions
    | 'accept-all-first'         // Normalizes all prior revisions to baseline
    | 'reject-input';            // Refuses any paragraph with revisions
```

### CLI Flag
* `--existing-revisions slice-cross-author`
* Future migration in v1.0.0: make `slice-cross-author` the default behavior of `merge-same-author` or replace it once verified across stress tests.

---

## 5. Work Packages

### WP-01: Baseline Word Desktop Fixture Generation [COMPLETED 2026-09-08]
* **Goal**: Produce reference DOCX files created directly in Microsoft Word Desktop 365 / 2021 covering:
  1. Cross-author insertion in the middle of a pending insertion (`insert-interior`).
  2. Cross-author deletion in the middle of a pending insertion (`delete-interior`).
  3. Cross-author deletion at the start of a pending insertion (`delete-boundary-start`).
  4. Cross-author deletion at the end of a pending insertion (`delete-boundary-end`).
  5. Cross-author deletion straddling baseline text and an insertion (`delete-straddle-baseline-insertion`).
  6. Multi-author stacked edits (`multi-author-stacked` - Author C editing Author B's edit inside Author A's edit).
* **Deliverables Produced**:
  1. `scripts/generate-cross-author-slicing-fixtures.ps1`: Windows PowerShell script automating Microsoft Word Desktop 365 via COM (`Word.Application`) to author the 6 scenarios, export pending/accepted/rejected `.docx` files, and extract `word/document.xml`.
  2. `tests/fixtures/cross-author-slicing/`: 36 golden fixture files (18 `.docx` and 18 UTF-8 BOM-clean `.xml` files) covering all triples:
     - `insert-interior-{pending,accepted,rejected}.{docx,xml}`
     - `delete-interior-{pending,accepted,rejected}.{docx,xml}`
     - `delete-boundary-start-{pending,accepted,rejected}.{docx,xml}`
     - `delete-boundary-end-{pending,accepted,rejected}.{docx,xml}`
     - `delete-straddle-baseline-insertion-{pending,accepted,rejected}.{docx,xml}`
     - `multi-author-stacked-{pending,accepted,rejected}.{docx,xml}`
  3. `tests/cross_author_slicing_fixtures_tests.mjs`: Comprehensive automated test suite asserting:
     - All 36 files exist and load cleanly without BOM corruption.
     - Word Desktop native OOXML structural patterns match assertions (sibling `<w:ins>` splitting vs. nested `<w:del>` inside `<w:ins>`).
     - Lifecycle oracle equivalence between `docx-redline-js` (`acceptTrackedChangesInOoxml`, `rejectTrackedChangesInOoxml`) and Word Desktop's native `AcceptAll` / `RejectAll`.
     - Selective author accept/reject lifecycle assertions on nested revisions.
* **Files Touched**:
  - `.gitignore` (MODIFIED to retain this plan in version control)
  - `scripts/generate-cross-author-slicing-fixtures.ps1` (NEW)
  - `tests/fixtures/cross-author-slicing/*` (36 files: 18 `.docx`, 18 `.xml`) (NEW)
  - `tests/cross_author_slicing_fixtures_tests.mjs` (NEW)
  - `docs/plans/2026-09-08-cross-author-revision-slicing.md` (MODIFIED)
* **Functions & Modules Touched / Created**:
  - `scripts/generate-cross-author-slicing-fixtures.ps1`:
    - `Extract-DocumentXml`: Extracts UTF-8 `word/document.xml` from DOCX ZIP archive.
    - `Save-Triple`: Saves pending DOCX, accepts all and saves accepted DOCX, reopens and rejects all and saves rejected DOCX.
    - Scenario blocks 1–6 utilizing COM Word automation (`$doc.TrackRevisions`, `$doc.Range.Text`, `$delRange.Delete()`, etc.).
  - `tests/cross_author_slicing_fixtures_tests.mjs`:
    - `loadFixtureXml`: Safe XML loader with leading BOM stripping.
    - `findChildrenByLocalName`, `findDescendantsByLocalName`: Namespace-resilient OOXML tree inspection.
    - Fixture inventory checks, Word Desktop structural checks, lifecycle parity checks (`acceptTrackedChangesInOoxml`, `rejectTrackedChangesInOoxml`), selective author accept/reject tests.
* **Key Findings & Next Phase Guidance (Crucial for WP-02 – WP-04)**:
  - **Insertion Slicing**: Word Desktop splits `<w:ins>` into siblings at `<w:p>` level: `[ins(A), ins(B), ins(A)]`.
  - **Deletion Slicing**: Word Desktop nests `<w:del>` directly inside `<w:ins>`: `<w:ins><w:r>...</w:r><w:del><w:r><w:delText>...</w:delText></w:r></w:del><w:r>...</w:r></w:ins>`. This is valid ECMA-376 `CT_RunTrackChange`.
  - **Validator Update Required in WP-04**: `core/redline-validation.js` currently flags any `<w:del>` inside `<w:ins>` as `NESTED_REVISION`. That rule must be updated to allow `<w:del>` inside `<w:ins>`, while maintaining the prohibition against `<w:ins>` inside `<w:ins>` or `<w:ins>` inside `<w:del>`.
* **WP-01 Review (2026-09-08)**:
  - Ran `node tests/cross_author_slicing_fixtures_tests.mjs`: all inventory, structural, lifecycle parity, and selective-author assertions passed.
  - Corrected the generator to derive the repository root from `$PSScriptRoot` instead of a machine-specific absolute path.
  - Added `Find-RequiredText` and changed all six scenario mutations to fail explicitly when Word cannot locate the intended text, preventing false golden fixtures.
  - Corrected stale deletion diagrams and synthetic expectations in this plan to match the native nested `<w:del>` evidence.

### WP-02: Carrier Splitting & Revision Allocator Integration [COMPLETED 2026-09-08]
* **Goal**: Implement `splitTrackChangeCarrier` in `engine/surgical-run-splitting.js`.
* **Deliverable**: Unit tests proving clean splitting of `<w:ins>` runs at character offsets, proper allocation of unique revision IDs via `RevisionIdAllocator`, and metadata preservation (`w:author`, `w:date`, `w16du:dateUtc`).
* **Implementation Summary**:
  - `splitTrackChangeCarrier` validates a `<w:ins>` carrier and accepted-view character offset, returns detached left/right fragments, and never mutates the source carrier.
  - Interior splits retain the original revision ID on the left fragment and allocate a fresh document-scoped ID for the right fragment. Exact-start and exact-end splits return `null` for the empty side and do not consume an ID.
  - Carrier attributes are cloned unchanged, including `w:author`, `w:date`, and `w16du:dateUtc`.
  - A run split preserves `w:rPr` and text-like run children (`w:t`, `w:tab`, `w:br`, and related supported nodes). If cloning duplicates a nested `w:rPrChange`, the right-hand copy receives a fresh ID.
  - Newly allocated carrier IDs are reported to an attached receipt collector when present.
* **Files Touched**:
  - `engine/surgical-run-splitting.js` (MODIFIED)
  - `tests/cross_author_carrier_splitting_tests.mjs` (NEW)
  - `scripts/generate-cross-author-slicing-fixtures.ps1` (MODIFIED during WP-01 review)
  - `docs/plans/2026-09-08-cross-author-revision-slicing.md` (MODIFIED)
* **Functions Touched / Created**:
  - `engine/surgical-run-splitting.js`:
    - `splitTrackChangeCarrier` (NEW): non-mutating `<w:ins>` split primitive.
    - `resolveAllocator` (NEW): resolves the explicit or document-scoped `RevisionIdAllocator`.
    - `setWordAttribute` (NEW): writes namespace-correct Word attributes.
    - `getLocalName` (NEW): namespace-resilient carrier validation.
    - Reuses `getRunContentPieces`, `getRunTextLength`, `sliceRunPieces`, and `createRunFromPieces`; these existing functions were not behaviorally changed.
  - `scripts/generate-cross-author-slicing-fixtures.ps1`:
    - `Find-RequiredText` (NEW): fail-fast lookup used by all six fixture scenarios.
    - Fixture directory initialization (MODIFIED): now repository-relative.
  - `tests/cross_author_carrier_splitting_tests.mjs`:
    - `parseCarrier`, `wordAttribute`, `text`, and `revisionIds` (NEW test helpers).
    - Tests cover interior/multi-run splitting, metadata and formatting preservation, unique carrier/`w:rPrChange` IDs, special run children, zero-empty-container boundaries, source immutability, and invalid inputs.
* **Verification**:
  - `node tests/cross_author_carrier_splitting_tests.mjs` — PASS.
  - `npx eslint engine/surgical-run-splitting.js tests/cross_author_carrier_splitting_tests.mjs` — PASS.
  - PowerShell parser check for `scripts/generate-cross-author-slicing-fixtures.ps1` — PASS (the COM fixtures were reviewed, not regenerated).
  - `npm test` — PASS, 90 test files passed and 0 failed.
  - `npm run lint` — PASS.
  - `npm run check:types` — PASS; all 123 runtime exports have declarations.
* **Handoff to WP-03**:
  - Import `splitTrackChangeCarrier` into `engine/surgical-diff-application.js` and use it only when a generated insertion lands inside a foreign-author `<w:ins>`.
  - The returned fragments are detached. WP-03 must splice non-null fragments plus Author B's new `<w:ins>` into the original carrier's parent, then remove the original carrier.
  - Pass the active document-scoped allocator so split-fragment IDs participate in the same operation receipt and rollback lifecycle.

### WP-03: Surgical Engine Cross-Author Insert Splicing [COMPLETED 2026-09-08]
* **Goal**: Update `processInsert` in `engine/surgical-diff-application.js` to split carrier `<w:ins>` when author differs and hoist Author B's `<w:ins>` to paragraph sibling level (`[ins(A), ins(B), ins(A)]`).
* **Deliverable**: Tests verifying cross-author insertion produces 3 sibling `<w:ins>` tags in valid schema without `NESTED_REVISION` errors.
* **Implementation Summary**:
  - `processInsert` now detects a foreign-author `<w:ins>` in ordinary, insertion-affinity, boundary, and replacement-anchor paths when `existingRevisions: 'slice-cross-author'` is active.
  - `spliceInsertionAtCarrierOffset` uses the WP-02 primitive and replaces the original carrier with `[left foreign ins, current-author ins, right foreign ins]`, omitting null boundary fragments.
  - Inserted text inherits effective run formatting but drops cloned historical `w:rPrChange` markup so revision IDs are not duplicated or misattributed.
  - `applyRedlineToOxml` retains foreign revisions and routes `slice-cross-author` text edits through surgical mode. Same-author input continues to use the established reject-to-baseline/re-diff merge behavior.
  - Pending move revisions still fail closed with `UNSAFE_REVISION_NESTING`.
* **Files Touched**:
  - `engine/surgical-diff-application.js`
  - `engine/surgical-mode.js`
  - `engine/oxml-engine.js`
  - `tests/cross_author_slicing_synthetic_tests.mjs` (NEW)
  - `tests/cross_author_slicing_fixtures_tests.mjs`
  - `docs/plans/2026-09-08-cross-author-revision-slicing.md`
* **Functions Touched / Created**:
  - `processInsert` (MODIFIED): policy-aware foreign carrier detection and replacement-anchor support.
  - `spliceInsertionAtCarrierOffset` (NEW): carrier split/splice orchestration.
  - `withoutRunPropertyChanges` (NEW): preserves effective formatting without copying historical formatting revisions.
  - `getCarrierSplitOffset`, `getCarrierGlobalStart`, `isForeignInsertion`, `isConnected` (NEW): carrier targeting helpers.
  - `applySurgicalMode` (MODIFIED): forwards the policy and permits paired replacements inside an insertion carrier under slicing mode.
  - `checkSafeAdjacencyForPairing` (MODIFIED): conditionally treats a foreign insertion carrier as a supported replacement context.
  - `applyRedlineToOxml` (MODIFIED): retains cross-author revisions, merges same-author revisions, rejects moves, and selects surgical routing.

### WP-04: Surgical Engine Cross-Author Delete Splicing & Validation Update [COMPLETED 2026-09-08]
* **Goal**: Update `processDelete` in `engine/surgical-diff-application.js` to nest Author B's `<w:del>` directly inside Author A's carrier `<w:ins>` (matching Word Desktop). Update `core/redline-validation.js` to permit `<w:del>` inside `<w:ins>`.
* **Deliverable**: Tests verifying cross-author deletion produces `<w:ins>...<w:del>...</w:del>...</w:ins>` matching Word Desktop fixtures without validation errors.
* **Implementation Summary**:
  - `processDelete` now groups adjacent affected runs sharing the same parent into one deletion wrapper, preserving each run's formatting and converting its text-like children to `w:delText`.
  - When the shared parent is a foreign `<w:ins>`, the deletion wrapper remains directly nested in that carrier; the carrier and its metadata stay intact at start, end, interior, and complete-content deletions.
  - Separate structural contexts receive separate deletion IDs but preserve shared replacement metadata/timestamps where supplied.
  - `validateRedlineOoxml` now accepts only direct `<w:ins><w:del>...</w:del></w:ins>` nesting. It continues to reject `ins/ins`, `del/del`, `ins` inside `del`, and deeper non-direct revision nesting.
* **Files Touched**:
  - `engine/surgical-diff-application.js`
  - `core/redline-validation.js`
  - `tests/redline_validation_tests.mjs`
  - `tests/cross_author_slicing_synthetic_tests.mjs` (NEW)
  - `tests/cross_author_slicing_fixtures_tests.mjs`
  - `docs/plans/2026-09-08-cross-author-revision-slicing.md`
* **Functions Touched / Created**:
  - `processDelete` (MODIFIED): plans run mutations, groups adjacent runs, emits one schema-correct deletion per parent context, and records replacement anchors.
  - `nextElementSibling` (NEW): determines whether affected runs can share a deletion wrapper without crossing structural markers.
  - `validateRedlineOoxml` (MODIFIED): implements the Word-native nested-deletion exception and retains all other nesting prohibitions.

### WP-05: Straddle & Boundary Deletion Normalization [COMPLETED 2026-09-08]
* **Goal**: Support deletions that cross between baseline text and pending insertions.
* **Deliverable**: Regression suite ensuring each deletion portion remains in its schema-correct context (top-level for baseline text, nested for pending insertion text) while sharing coherent author/event metadata.
* **Implementation Summary**:
  - Cross-boundary deletion records are partitioned by their actual DOM parent: baseline portions emit top-level `<w:del>` and insertion portions emit nested `<w:del>` without attempting schema-invalid coalescing.
  - Forward and reverse baseline/insertion straddles and deletion across two distinct foreign insertion carriers are covered.
  - `revisionInsertionAnchors` preserves the post-deletion carrier/offset/formatting location for the insertion half of a replacement after original span nodes have been removed.
  - Paired replacement metadata remains paired inside a foreign insertion: the deletion stays in the left carrier, the new insertion is hoisted, and the surviving right carrier receives its own unique ID.
* **Files Touched**:
  - `engine/surgical-diff-application.js`
  - `engine/surgical-mode.js`
  - `tests/cross_author_slicing_synthetic_tests.mjs` (NEW)
  - `tests/cross_author_slicing_fixtures_tests.mjs`
  - `docs/plans/2026-09-08-cross-author-revision-slicing.md`
* **Functions Touched / Created**:
  - `processDelete` (MODIFIED): records per-carrier replacement anchors and keeps parent contexts separate.
  - `processInsert` (MODIFIED): consumes a connected replacement anchor before consulting stale original spans.
  - `applySurgicalMode` and `checkSafeAdjacencyForPairing` (MODIFIED): allow paired events in the newly supported carrier context.
* **Verification to Date**:
  - `node tests/cross_author_slicing_synthetic_tests.mjs` — PASS.
  - `node tests/cross_author_slicing_fixtures_tests.mjs` — PASS, including engine reproduction of the five Word Desktop WP03-WP05 fixtures and AcceptAll/RejectAll parity.
  - `npm test` — PASS, 91 test files passed and 0 failed.
  - `npm run lint` — PASS.
  - `npm run check:types` — PASS; all 123 runtime exports have declarations.
  - `git diff --check` — PASS.
* **Handoff Boundary**:
  - `slice-cross-author` is implemented in the low-level `applyRedlineToOxml` engine used by these work packages.
  - CLI validation, operation preflight, document-operation/facade gating, schema enum changes, and public type exposure remain intentionally deferred to WP-07.
  - WP-06 can build on the already passing native-fixture lifecycle tests; no lifecycle service code was changed in WP03-WP05.

### WP-06: Lifecycle Oracles & Selective Accept/Reject Support [COMPLETED 2026-09-08]
* **Goal**: Update `services/revision-comment-management.js` (`acceptTrackedChangesInOoxml` and `rejectTrackedChangesInOoxml`) to handle sliced revisions correctly:
  - Selective accept/reject by author.
  - Cascading rejection (rejecting Author A cleans up internal deletions by Author B).
* **Deliverable**: Round-trip lifecycle oracle tests comparing against Word Desktop results.
* **Implementation Summary**:
  - Confirmed the existing outer-first accept/reject traversal already provides correct dependency behavior for nested deletions: accepting Author A unwraps the carrier and leaves Author B's deletion pending; rejecting Author A removes the carrier and its dependent deletion; rejecting Author B restores `w:delText` as normal text inside Author A's carrier.
  - Added post-rejection normalization for sliced insertions. When rejecting Author B removes the middle insertion, adjacent Author A `<w:ins>` fragments with identical metadata other than `w:id` are coalesced back into one carrier.
  - Coalescing is deliberately metadata-safe: different authors, dates, `dateUtc` values, or other carrier attributes are never merged.
* **Files Touched**:
  - `services/revision-comment-management.js`
  - `tests/cross_author_slicing_synthetic_tests.mjs`
  - `docs/plans/2026-09-08-cross-author-revision-slicing.md`
* **Functions Touched / Created**:
  - `rejectTrackedChangesInOoxml` (MODIFIED): runs compatible insertion coalescing after selective rejection.
  - `nextElementSibling` (NEW): finds adjacent revision carriers while tolerating formatting whitespace nodes.
  - `revisionMetadataWithoutId` (NEW): creates a stable metadata comparison key excluding only the split-specific revision ID.
  - `coalesceAdjacentCompatibleInsertions` (NEW): merges compatible same-author insertion fragments without changing lifecycle counts.
* **Lifecycle Coverage**:
  - Selective accept/reject of Author A and Author B on generated sibling insertion slices.
  - Fragment coalescing after rejecting Author B.
  - Selective rejection of nested Author B deletion and restoration inside Author A's insertion.
  - Selective acceptance of Author A with Author B's deletion remaining pending.
  - Cascading removal of Author B's nested deletion when Author A is rejected.
  - Existing Word Desktop AcceptAll/RejectAll and selective-author fixture oracles remain green.

### WP-07: CLI, Preflight, and Facade Exposure [COMPLETED 2026-09-08]
* **Goal**: Expose `slice-cross-author` in:
  - `docx-redline` CLI (`--existing-revisions slice-cross-author`)
  - `services/operation-preflight.js` (recognizes cross-author slicing as valid rather than `EXISTING_REVISIONS`)
  - `openDocx` / `applyOperationsToDocumentXml` options
* **Deliverable**: Updated schema definitions and CLI test cases.
* **Implementation Summary**:
  - Added one canonical runtime policy list and validation helper. Operation-level and batch-level invalid policies now return `INVALID_OPERATION` instead of falling through to unrelated revision errors.
  - Preflight reports foreign insertion/deletion edits as `ready` under `slice-cross-author`, preserves same-author comment safeguards, and continues to fail closed on pending moves.
  - The document mutation gate now permits slicing operations to reach `applyRedlineToOxml`; operation-level policy overrides continue to flow through `document-operation-applier`.
  - `applyOperationsToDocumentXml`, `openDocx(...).applyOperations`, and CLI `apply`/`preflight` accept the batch policy. CLI version capabilities now advertise `cross-author-revision-slicing`.
  - Public TypeScript types, JSON Schema, README, and `AGENTS.md` document the new policy.
* **Files Touched**:
  - `services/document-operation-contract.js`
  - `services/operation-preflight.js`
  - `services/document-operation-mutations.js`
  - `services/batch-operation-orchestrator.js`
  - `node/cli.js`
  - `index.d.ts`
  - `docs/schemas/document-operations.schema.json`
  - `README.md`
  - `AGENTS.md`
  - `tests/existing_revisions_modes_matrix_tests.mjs`
  - `tests/agent_operation_contract_tests.mjs`
  - `tests/agent_cli_tests.mjs`
  - `docs/plans/2026-09-08-cross-author-revision-slicing.md`
* **Functions / Types Touched or Created**:
  - `EXISTING_REVISIONS_POLICIES` and `isExistingRevisionsPolicy` (NEW): shared policy contract and validator.
  - `validateDocumentOperation` (MODIFIED): validates operation-level `existingRevisions` values.
  - `preflightOperations` (MODIFIED): validates batch policy and models same-author merge, foreign-author slicing, and unsupported moves.
  - `applyOperationsToDocumentXml` (MODIFIED): validates batch policy before creating or mutating a session.
  - Document redline mutation policy gate in `applyToParagraphByExactText` (MODIFIED): permits slicing while preserving comment and move safeguards.
  - `executeCli` (MODIFIED): validates `--existing-revisions`; apply/preflight pass-through remains shared with facade options.
  - `ExistingRevisionsPolicy` (MODIFIED): includes `'slice-cross-author'`.
* **End-to-End Coverage**:
  - CLI `apply --existing-revisions slice-cross-author` writes a package that inspects with both original and current authors.
  - CLI `preflight` accepts the same operation and reports the selected policy.
  - CLI rejects an unknown policy with `INVALID_OPERATION`.
  - Runtime operation validation accepts `slice-cross-author` and rejects misspellings.
  - JSON Schema is parsed and asserted to contain the new enum value.
  - Facade/package validation and output reconciliation run through the existing CLI application path.
* **Final Verification**:
  - `npm test` — PASS, 91 test files passed and 0 failed.
  - `npm run lint` — PASS.
  - `npm run check:types` — PASS; all 123 runtime exports have declarations.
  - `git diff --check` — PASS.

### Final Automated Test Completion [COMPLETED 2026-09-08]
* **Coverage Added**:
  - Completed the executable synthetic matrix for SYN-01 through SYN-12d. The final additions generate SYN-02 in the engine, add a second cross-author deletion to produce SYN-11, and exercise Accept All, Accept Author A, Reject Author A, Reject Author B, and selective rejection of the third author.
  - Added strict package-facade differential tests for all five core Word Desktop scenarios. Each test reconstructs the pre-Reviewer-B package, applies `slice-cross-author` through `openDocx(...).applyOperations`, requires atomic package validation, and compares engine Accept-All/Reject-All text with the checked-in Word Desktop accepted/rejected DOCX files.
  - Added a two-round package test that generates the three-author stacked-deletion fixture through Reviewer B and Reviewer C operations, asserts all three reviewers survive inspection, and compares both lifecycle endpoints with Word Desktop.
* **Files Touched**:
  - `tests/cross_author_slicing_synthetic_tests.mjs` (MODIFIED)
  - `tests/cross_author_slicing_real_tests.mjs` (NEW)
  - `docs/plans/2026-09-08-cross-author-revision-slicing.md` (MODIFIED)
* **Functions / Test Helpers Touched or Created**:
  - `tests/cross_author_slicing_synthetic_tests.mjs`: added the SYN-02/SYN-11/SYN-12 generated lifecycle block; no production functions changed during final test completion.
  - `tests/cross_author_slicing_real_tests.mjs`:
    - `fixtureBuffer` (NEW): loads a checked-in Word Desktop DOCX oracle.
    - `visibleText` (NEW): obtains stable main-document text through package inspection.
    - `resolvedText` (NEW): resolves all revisions through the package facade, requires package validation, and asserts no revision authors remain.
    - PKG-01 through PKG-05 (NEW): strict package reproduction and differential lifecycle tests.
    - PKG-06 (NEW): generated three-reviewer, two-round package negotiation test.
* **Final Verification**:
  - `node tests/cross_author_slicing_synthetic_tests.mjs` — PASS.
  - `node tests/cross_author_slicing_real_tests.mjs` — PASS.
  - `npm test` — PASS, 92 test files passed and 0 failed.
  - `npm run lint` — PASS.
  - `npm run check:types` — PASS; all 123 runtime exports have declarations.
  - `git diff --check` — PASS (line-ending conversion notices only; no whitespace errors).
* **Scope Note**:
  - The repository does not contain the private source packages referenced by REAL-01 through REAL-03. Those cases remain external acceptance scenarios rather than silently skipped automated tests.
  - REAL-04/REAL-05 require Microsoft Word Desktop COM and visual review. The checked-in fixtures were produced by Word COM, while the normal automated suite deliberately remains deterministic and non-interactive.

### Bug Follow-Up: Hyperlink Boundary Round-Trip Mismatch [FIXED 2026-09-08]
* **Report Reproduced**: A slicing edit with repeated text, hyperlink runs, and NBSP-to-space substitutions could return `status: "ok"` even though its accepted-view text differed from `modified`.
* **Root Cause**: `applySurgicalMode` discarded whitespace-only insertion diff segments by checking `textWithoutNewlines.trim().length`. The corresponding NBSP deletion still committed, changing `located at\u00a0example.com` to `located atexample.com`. Structural replacement paths could also proceed after `PAIRING_SKIPPED_STRUCTURAL_BOUNDARY` without a final exact-text oracle.
* **Fix**:
  - Whitespace-only insertions are now applied rather than silently skipped.
  - Every `slice-cross-author` surgical result reconstructs canonical accepted-view text and compares it exactly with the requested clean modified text.
  - A mismatch returns `PATCH_ROUNDTRIP_MISMATCH`, `hasChanges: false`, diagnostic excerpts and offset, and the exact original OOXML. The document runner/facade therefore treats the operation as unapplied and preserves transactional rollback.
* **Files Touched**:
  - `engine/surgical-mode.js`
  - `engine/oxml-engine.js`
  - `tests/cross_author_slicing_hyperlink_roundtrip_tests.mjs` (NEW)
  - `CHANGELOG.md`
  - `README.md`
  - `AGENTS.md`
  - `docs/TESTING.md`
  - `docs/plans/2026-09-08-cross-author-revision-slicing.md`
* **Functions Touched / Created**:
  - `applySurgicalMode` (MODIFIED): retains whitespace-only insertions and enforces the slicing accepted-view postcondition.
  - `firstMismatchOffset` (NEW): locates the first exact-text divergence.
  - `excerptAt` (NEW): provides bounded expected/actual diagnostics without returning entire contract paragraphs.
  - `applyRedlineToOxml` surgical result handling (MODIFIED): restores the exact input OOXML on `PATCH_ROUNDTRIP_MISMATCH`.
  - Test helpers `run`, `insertion`, `hyperlink`, and `acceptedParagraphText` (NEW).
* **Regression Coverage**:
  - Low-level reproduction with two hyperlink relationship containers, repeated `Widget Policy`, and three NBSP-to-space edits.
  - Full `applyOperationsToDocumentXml` atomic runner reproduction matching the CLI execution path.
  - Exact Accept-All equality with the submitted modified string.
  - Explicit fail-closed test proving mismatch status, error code, mismatch offset, and byte-exact original OOXML rollback.
* **Verification**:
  - `node tests/cross_author_slicing_replacement_anchor_tests.mjs` — PASS, 12 scenarios.
  - `node tests/cross_author_slicing_hyperlink_roundtrip_tests.mjs` — PASS.
  - `npm test` — PASS, 93 test files passed and 0 failed.
  - `npm run lint` — PASS.
  - `npm run check:types` — PASS; all 123 runtime exports have declarations.
  - `git diff --check` — PASS (line-ending conversion notices only; no whitespace errors).

### Insertion Stress Follow-Up [COMPLETED 2026-09-08]
* **Motivation**: Real usage reported failures across a wider variety of insertions after the first hyperlink/NBSP bug. A generated matrix was added to exercise location, payload, structure, lifecycle, and repeated-review dimensions rather than relying on a few fixed examples.
* **Defects Exposed and Fixed**:
  1. Leading/trailing spaces, tabs, and NBSP-only additions were classified as no-ops because slicing inherited trim-based text-change detection. `applyRedlineToOxml` now uses exact comparison for `slice-cross-author`.
  2. Word-token semantic diff cleanup could relocate a pure insertion between repeated phrases, especially inside a hyperlink. `computeInsertionOnlyDiffs` now selects a character-local, no-deletion diff whenever the original is an exact subsequence of the modified text; replacements retain the established word diff and exact round-trip guard.
  3. In a paragraph containing both current-author and foreign insertion carriers, inserting into the current-author carrier produced illegal nested `w:ins`. `processInsert` now adds a normal run to that existing carrier while continuing to split foreign carriers into siblings.
* **Coverage Added**:
  - 76 deterministic scenarios spanning carrier start/end/interior positions; single-, multi-, and formatted runs; repeated tokens; double spaces, tabs, NBSP, XML-sensitive characters, emoji, ZWJ emoji, combining characters, citations, and punctuation.
  - Hyperlink interiors and both hyperlink boundaries; bookmarks; comment anchors; nested prior deletions; adjacent foreign authors; mixed current/foreign authors; three-container edits; and consecutive second-/third-reviewer rounds.
  - Exact accepted-view, Accept-All, Reject-Current, validation, unique metadata, hyperlink preservation, and zero-empty-insertion assertions.
  - Twenty scenarios also execute through `applyOperationsToDocumentXml` with atomic and strict-target settings, matching the CLI runner path.
  - Nested hyperlink/field structures are required either to produce exact valid output or fail closed without throwing.
* **Files Touched**:
  - `pipeline/diff-engine.js`
  - `engine/oxml-engine.js`
  - `engine/surgical-mode.js`
  - `engine/surgical-diff-application.js`
  - `tests/cross_author_slicing_insertion_stress_tests.mjs` (NEW)
  - `CHANGELOG.md`
  - `README.md`
  - `docs/TESTING.md`
  - `docs/plans/2026-09-08-cross-author-revision-slicing.md`
* **Functions Touched / Created**:
  - `computeInsertionOnlyDiffs` (NEW): detects insertion-only transforms and returns a character-local diff only when it contains no deletion.
  - `applyRedlineToOxml` (MODIFIED): uses exact slicing change detection, including boundary whitespace.
  - `applySurgicalMode` (MODIFIED): selects insertion-only versus word diff without changing replacement semantics.
  - `processInsert` (MODIFIED): inserts directly into an existing same-author carrier in mixed-author paragraphs.
  - `isSameAuthorInsertion` (NEW): namespace-safe author comparison for carrier coalescing.
  - Stress helpers `escapeXml`, `run`, `insertion`, `paragraph`, `parsed`, `acceptedText`, `authorOf`, `assertValid`, and `assertInsertionRoundTrip` (NEW).
* **Verification**:
  - `node tests/cross_author_slicing_insertion_stress_tests.mjs` — PASS, 76 scenarios.
  - `npm test` — PASS, 95 test files passed and 0 failed.
  - `npm run lint` — PASS.
  - `npm run check:types` — PASS; all 123 runtime exports have declarations.
  - `git diff --check` — PASS (line-ending conversion notices only; no whitespace errors).

### Hyperlink-Adjacent Replacement Follow-Up [COMPLETED 2026-09-08]
* **Bug Report**: Replacing the space immediately after a policy hyperlink with a comma and effective-date qualifier failed with `PATCH_ROUNDTRIP_MISMATCH`. The generated intermediate OOXML moved the qualifier and URL relative to the following definition text.
* **Root Cause**: `processDelete` split and removed the run containing the replaced boundary space, but `processInsert` subsequently resolved the paired insertion through the pre-mutation span index. That span still referenced the detached source run, so insertion placement fell back to the wrong paragraph location.
* **Fix**: `processDelete` now records a stable parent/reference-node anchor at a non-carrier deletion boundary. `processInsert` consumes that anchor for the immediately paired insertion when no explicit insertion affinity was requested. Foreign `w:ins` carriers continue to use their existing carrier-splitting anchor and explicit affinity remains authoritative.
* **Additional Defect Found by the Matrix**: Two pure insertions in the same source run could detach the shared pre-mutation span after the first insertion and relocate the second insertion to the paragraph end. Multi-insertion-only slicing now applies insertions from right to left and rebuilds the live span index between mutations.
* **Files Touched**:
  - `engine/surgical-diff-application.js`
  - `engine/surgical-mode.js`
  - `tests/cross_author_slicing_hyperlink_roundtrip_tests.mjs`
  - `tests/cross_author_slicing_replacement_anchor_tests.mjs` (NEW)
  - `CHANGELOG.md`
  - `docs/plans/2026-09-08-cross-author-revision-slicing.md`
* **Functions Touched**:
  - `processDelete` (MODIFIED): records the live DOM insertion boundary while splitting a deleted run.
  - `processInsert` (MODIFIED): consumes the stable replacement anchor before consulting stale pre-mutation spans.
  - `applySurgicalMode` (MODIFIED): uses live right-to-left application for multiple insertion-only diffs.
  - `collectInsertionOperations` (NEW): records original/new offsets for insertion-only mutations.
  - Hyperlink round-trip test helpers and assertions (MODIFIED): cover low-level apply, Accept All, hyperlink relationship preservation, and the atomic strict-target document runner.
  - Replacement-anchor matrix helpers `escapeXml`, `run`, `hyperlink`, `insertion`, `paragraph`, `parse`, `acceptedText`, and `hyperlinkIds` (NEW).
* **Additional Future-Regression Coverage**:
  - 12 deterministic replacements at run starts, interiors, and ends; before and after hyperlinks; across an entire spacer run; beside bold/underlined runs, bookmarks, and comment markers; across multiple replacements; and beside/inside a foreign insertion carrier.
  - Every case asserts exact current view, Accept All, Reject Reviewer, structural validation, unique revision IDs, and hyperlink relationship preservation.
  - Four representative hyperlink and multi-replacement cases also execute through the atomic strict-target document runner used by the CLI.
* **Verification**:
  - `node tests/cross_author_slicing_replacement_anchor_tests.mjs` — PASS, 12 scenarios.
  - `node tests/cross_author_slicing_hyperlink_roundtrip_tests.mjs` — PASS.
  - `node tests/cross_author_slicing_insertion_stress_tests.mjs` — PASS, 76 scenarios.
  - `node tests/insertion_affinity_tests.mjs` — PASS.
  - `npm test` — PASS, 95 test files passed and 0 failed.
  - `npm run lint` — PASS.
  - `npm run check:types` — PASS; all 123 runtime exports have declarations.
  - `git diff --check` — PASS (line-ending conversion notices only; no whitespace errors).

### WP08a — Fail-Closed Gate for Foreign Paragraph-Mark Deletions [COMPLETED 2026-09-08]

WP08a is separable from, and a prerequisite of, WP08b. It ships on its own as a patch release: it adds no new capability, only refuses an operation that currently returns `status: 'ok'` while producing a lifecycle-unsafe document. Every prior follow-up in this document shipped the fail-closed guard before the feature; WP08 follows the same order.

#### Scope

1. At the mutation gate, refuse any operation that would add visible runs or `w:ins` content to a paragraph in the **resurrection state** defined in WP08b's trigger taxonomy (foreign paragraph-mark deletion + every pre-existing content node already deleted + new non-empty insertion). Return `FOREIGN_PARAGRAPH_MARK_DELETION` with the owning author, and the original document unchanged under atomic mode.
2. Add the same predicate to `core/redline-validation.js` as a **warning**, not an error. Validation runs over documents this engine did not author; see the Validation Predicate section below for why the broader rule is unsafe.
3. Regression fixture reproducing the reported restoration shape (foreign `w:pPr/w:rPr/w:del` plus an attempted same-paragraph `w:ins`), asserting the refusal, the error code, and byte-exact rollback.
4. Lifecycle assertions in the fixture proving *why* the shape is refused: Accept All loses Reviewer B's text, and Reject Reviewer A yields duplicate visible text.

Package validation alone is not a sufficient oracle for this case — the unsafe shape is structurally valid. Accept/Reject lifecycle checks are mandatory.

#### Implementation Record

* **Result**: Added a shared, narrowly scoped resurrection-state predicate. Both the low-level paragraph engine and the document mutation runner now refuse a non-empty edit when a different author owns the paragraph-mark deletion and every pre-existing content child is deleted. The refusal returns `FOREIGN_PARAGRAPH_MARK_DELETION`, includes `ownerAuthor`, and leaves the original input unchanged. Same-author deleted paragraphs, foreign deleted marks with surviving content, and content-only deletions without a paragraph-mark deletion remain on their existing paths.
* **Validation**: `validateRedlineOoxml` uses the same structural model to emit a warning for already-authored unsafe shapes. The warning does not make otherwise valid OOXML invalid.
* **Files Touched**:
  - `core/paragraph-revision-safety.js` (NEW)
  - `core/redline-validation.js`
  - `engine/oxml-engine.js`
  - `services/document-operation-mutations.js`
  - `index.d.ts`
  - `tests/foreign_paragraph_mark_deletion_gate_tests.mjs` (NEW)
  - `CHANGELOG.md`
  - `docs/plans/2026-09-08-cross-author-revision-slicing.md`
* **Functions and Types Touched**:
  - `inspectForeignDeletedParagraphTarget` (NEW): identifies the pre-mutation WP08 resurrection state and excludes same-author ownership.
  - `findForeignDeletedParagraphResurrections` (NEW): identifies already-authored unsafe foreign-insertion shapes for warning-only validation.
  - `paragraphMarkDeletion`, `hasVisibleInsertionContent`, `isAnchorOnlyRun`, and DOM/name/author helpers (NEW): implement the shared structural inspection without mutating the source DOM, while excluding non-visible comment/bookmark anchors from the content-state decision.
  - `applyRedlineToOxml` (MODIFIED): fails closed before existing-revision normalization or diff application for paragraph-level calls.
  - `applyToParagraphByExactText` (MODIFIED): fails closed immediately after strict target resolution and before preprocessing/mutation for document-runner calls.
  - `validateRedlineOoxml` (MODIFIED): reports `FOREIGN_PARAGRAPH_MARK_DELETION` as a warning for structurally valid but lifecycle-unsafe authored output.
  - `RedlineError` (MODIFIED): documents the new error code and optional `ownerAuthor` metadata.
  - WP08a fixture helpers and assertions (NEW): cover the low-level tracked and direct-edit paths, strict atomic runner rollback, validator severity, Accept/Reject lifecycle evidence, and every non-triggering taxonomy row.
* **Verification**:
  - `node tests/foreign_paragraph_mark_deletion_gate_tests.mjs` — PASS.
  - Focused validation, revision-policy, replacement-anchor, and paragraph-boundary suites — PASS.
  - `npm test` — PASS, 96 test files passed and 0 failed.
  - `npm run lint` — PASS.
  - `npm run check:types` — PASS; all 123 runtime exports have declarations.
  - `npm run build` — PASS.
  - `git diff --check` — PASS (line-ending conversion notices only; no whitespace errors).

---

### WP08b — Paragraph-Level Cross-Author Slicing for Deleted Paragraph Restoration [COMPLETED 2026-09-08]

#### Motivation

Restoring text from another reviewer's pending whole-paragraph deletion is the paragraph-level counterpart of run-level cross-author slicing. The accepted/current view of such a paragraph is empty, while its text exists only in the rejected view. Writing replacement text into that same paragraph can look correct before revisions are resolved, but it is lifecycle-unsafe because the foreign paragraph-mark deletion still owns the paragraph.

The observed unsafe shape is conceptually:

```xml
<w:p>
  <w:pPr><w:rPr><w:del w:author="Reviewer A"/></w:rPr></w:pPr>
  <w:del w:author="Reviewer A">...</w:del>
  <w:ins w:author="Reviewer B">restored text</w:ins>
</w:p>
```

This passes structural package validation and looks correct in the current view, but Accept All removes the entire paragraph because Reviewer A's paragraph-mark deletion remains active. Rejecting Reviewer A can also expose both the original deleted text and Reviewer B's inserted copy.

Nesting Reviewer B's `<w:ins>` inside Reviewer A's `<w:del>` is not a solution: `w:del/w:ins` nesting is invalid for this use, and accepting the outer deletion would remove the nested text.

#### Required Paragraph-Level Slicing Model

Preserve the foreign deleted paragraph and materialize the restoring reviewer's counterproposal as a new adjacent tracked paragraph:

```text
[restored/adjusted paragraph inserted by Reviewer B]
[paragraph deleted by Reviewer A]
```

#### Paragraph-Mark Semantics (Normative)

This is the part the run-level slicing model has no analogue for, and it governs every lifecycle row below.

**Accepting a paragraph-mark deletion does not remove the paragraph — it merges the paragraph into the next paragraph.** `mergeParagraphIntoNextAndRemove` in `services/revision-comment-management.js` moves the deleted paragraph's surviving children into the following `w:p` and removes the emptied paragraph; the **following** paragraph's `pPr` is the one that survives. Any design statement phrased as "Reviewer A's paragraph is removed" is imprecise and must be read as "merged forward".

Three consequences are binding on the implementation:

1. **Sibling order is a design decision, not cosmetic.** Placing Reviewer B's paragraph *after* Reviewer A's makes B the merge target when Reviewer A is accepted: A's surviving children land inside B. This is harmless only while A's content is 100% deleted, and stops being harmless the moment A retains content (a partially resolved deletion, or a third author's `w:ins` still pending inside A). Placing B *before* A leaves A's merge target exactly as it was before the restoration existed, so accepting A behaves identically with or without B.
   * **Decision: place Reviewer B's paragraph immediately BEFORE Reviewer A's**, for merge-target neutrality. The current view is unaffected (A is invisible), and All-Markup view order is a rendering preference, not a correctness property. Fixtures must assert the merge target explicitly rather than inferring it from the resulting text.
2. **Reviewer B's paragraph MUST carry its own inserted paragraph mark** (`w:pPr/w:rPr/w:ins` attributed to Reviewer B, with an allocator-issued ID). Adding a paragraph adds a paragraph mark. Without it, Reject Reviewer B removes B's content but leaves an empty stub paragraph permanently, silently violating the Reject-B lifecycle row and drifting the document's paragraph count.
3. **The existing inserted-paragraph builders do not do this today.** `wrapParagraphContentInInsertion` and `buildFallbackInsertedPlainParagraph` in `services/document-operation-mutations.js` emit no paragraph-mark revision and clone `pPr` verbatim. Cloning `pPr` verbatim from the deleted source paragraph would copy Reviewer A's `w:rPr/w:del` onto Reviewer B's paragraph, reproducing the exact unsafe shape WP08 exists to prevent. Emitting the inserted mark and sanitizing `pPr` is new work in those builders, not reuse of them.

> **WP09e supersession:** A later Microsoft Word Desktop oracle places a two-paragraph counterproposal **after** the corresponding fully deleted source block, not before it. WP09e reopens the completed WP08b placement decision and requires Word-native ordering plus native lifecycle comparison. This historical WP08b record describes the v0.5.3 implementation, not the final target behavior.

#### Paragraph Property Sanitization (Normative Allowlist)

When deriving Reviewer B's paragraph from the deleted source, copy only:

* `w:pStyle`, `w:numPr`, `w:ind`, `w:jc`, `w:spacing`, `w:tabs`, `w:keepNext`/`w:keepLines`, `w:outlineLvl`, `w:contextualSpacing`.

Strip unconditionally:

* `w:rPr/w:del` and `w:rPr/w:ins` (foreign mark revisions — replaced by Reviewer B's own inserted mark),
* `w:sectPr` (section identity must never be duplicated; see refusals),
* `w:pPrChange`, `w:rPrChange`, and every other `*Change` element (they describe a revision of the *source* paragraph and are meaningless on the clone).

#### Paragraph Identity (Normative)

Reviewer B's paragraph MUST receive a **fresh `w14:paraId`**, and MUST drop `w14:textId` and all `w:rsid*` attributes. This is unconditional, not best-effort: `extractParagraphIdFromOoxml` in `core/ooxml-identifiers.js` resolves strict targets by `w14:paraId`, so a duplicated paraId makes Reviewer A's and Reviewer B's paragraphs indistinguishable to paragraph-ID targeting — including to this work package's own strict-targeting test case.

#### Trigger Taxonomy (Normative)

"Whole-paragraph deletion" is ambiguous across the four combinations of paragraph-mark state and content state. Only one routes to WP08b:

| Paragraph mark | Pre-existing content | Route |
|:--|:--|:--|
| Deleted by foreign author | All deleted | **WP08b sibling restoration** (the resurrection state) |
| Deleted by foreign author | Intact or partially deleted | Ordinary run-level cross-author slicing — the accepted view is non-empty; a pending forward merge is legal and Word-native |
| Not deleted | All deleted | Ordinary cross-author insertion — no foreign mark owns the paragraph; MUST NOT route to WP08b |
| Deleted by current author | All deleted | `merge-same-author`; MUST NOT route to WP08b |

#### Validation Predicate (Normative)

The reconciliation rule must match the resurrection state exactly. A broader rule of the form "foreign paragraph-mark deletion plus visible insertion in the same paragraph" is **wrong** — inserting text into a paragraph whose mark is deleted by another author is legal, Word-native, and common (it is an ordinary pending merge). Flagging it would reject valid third-party documents.

The predicate is: foreign `w:pPr/w:rPr/w:del` **AND** every pre-existing content node deleted **AND** a new non-empty foreign `w:ins`. It is a **warning** in `core/redline-validation.js` and an **error** only at the mutation gate (WP08a).

> **WP09e supersession:** Word Desktop itself emits that structural predicate when a second reviewer inserts at a rejected-view offset inside the deleted text. Authored shape alone cannot distinguish unsafe generic resurrection from intentional deletion-carrier slicing. WP09e narrows validation and routing by operation intent and anchoring evidence.

#### Required Behavior

1. Detect the resurrection state per the trigger taxonomy when an operation attempts to restore non-empty text into the paragraph's empty accepted view.
2. Never append visible runs or `w:ins` content to the paragraph still owned by the foreign paragraph deletion.
3. Derive Reviewer B's paragraph as a sanitized sibling of the source paragraph per the allowlist and identity rules above, without mutating Reviewer A's original deleted paragraph.
4. Track the new paragraph's content and its paragraph mark as Reviewer B insertions with document-scoped revision IDs.
5. Preserve document order and ensure `w:sectPr`, tables, list boundaries, comments, bookmarks, and other structural anchors are neither displaced nor duplicated.
6. **Restoring a multi-paragraph range emits one inserted sibling paragraph per restored source paragraph**, as a contiguous block preserving source order, with N paragraph-mark insertions — never one merged paragraph.
7. **Idempotency**: re-running the same restoration must not emit a second Reviewer B paragraph. If an adjacent same-author inserted paragraph already carries the restoration, route the edit through ordinary run-level slicing of that paragraph.
8. Rejected-view descriptors remain read-only targeting aids until rejected-view mutation is deliberately supported. Do not silently treat `revisionView: 'rejected'` as accepted-view mutation.
9. **Contract decision (resolved, not deferred):** restoration requires **explicit caller intent** — a dedicated restore operation or an explicit restoration option. A `redline` operation with an empty accepted-view target and non-empty modified text remains fail-closed under WP08a. Automatic conversion is rejected because the same request shape is indistinguishable from an ordinary "insert text into an empty paragraph", and silently choosing restoration would move the caller's content into a different paragraph than the one they targeted.

#### Structural Anchors (Normative)

Reviewer B's paragraph MUST NOT clone `w:bookmarkStart`/`w:bookmarkEnd` or `w:commentRangeStart`/`w:commentRangeEnd`/`w:commentReference`. Bookmark names are document-unique, and duplicating a comment range attaches one comment to two disjoint locations. Anchors stay on Reviewer A's paragraph, where they remain valid until that deletion is resolved. Every anchor not carried over is reported in the receipt as a structured warning naming the bookmark or comment ID, so the caller can re-anchor deliberately.

#### Fail-Closed Refusals (Distinct Codes)

A single `UNSAFE_PARAGRAPH_BOUNDARY` code conflates unrelated conditions and reads as a near-collision with the existing `PAIRING_SKIPPED_STRUCTURAL_BOUNDARY`. Enumerate:

| Condition | Code |
|:--|:--|
| Resurrection attempted without explicit restore intent (WP08a gate) | `FOREIGN_PARAGRAPH_MARK_DELETION` |
| Source paragraph inside a row deleted via `w:trPr/w:del` — a sibling paragraph cannot survive the row | `UNSAFE_DELETED_TABLE_ROW` |
| Source paragraph is part of a move (`w:moveFrom` / `w:moveFromRangeStart`) | `UNSUPPORTED_MOVE_REVISION` |
| Source paragraph's `pPr` carries `w:sectPr` — mirrors the existing deletion refusal in `core/paragraph-targeting.js` | `SECTION_BREAK_PARAGRAPH` |
| Source paragraph is the final paragraph of the body, so no safe sibling placement exists | `UNSAFE_PARAGRAPH_PLACEMENT` |

All refusals return the original document unchanged under atomic mode.

#### Lifecycle Invariants

For the paragraph pair `[ins(B), del(A)]` in document order:

| Resolution | Expected Result |
|:--|:--|
| Current view | Reviewer B's restored/adjusted paragraph appears exactly once; Reviewer A's paragraph is invisible. |
| Accept All | Reviewer B's paragraph and mark become baseline; Reviewer A's content deletion resolves and A's mark merges A forward into its **original** successor (not into B). Net: Reviewer B's paragraph remains exactly once. |
| Reject Reviewer B | Reviewer B's content is removed and B's inserted mark is rejected, merging the now-empty B forward into A. Net: the document returns to its pre-restoration text with Reviewer A's deletion still pending. |
| Reject Reviewer A | Reviewer A's paragraph and its text return; Reviewer B's insertion remains independently pending. Both marks remain attributable and structurally valid. |
| Accept Reviewer A only | Reviewer A's content deletion resolves and A merges forward into its original successor; Reviewer B's inserted paragraph remains pending and visible. |
| Accept Reviewer B only | Reviewer B's paragraph and mark become baseline; Reviewer A's deleted paragraph remains pending and invisible in the current view. |

No lifecycle path may silently discard Reviewer B's restoration, produce invalid nested revisions, orphan comments/bookmarks, or leave duplicate visible text after all revisions are resolved.

#### Verification Oracle (Normative)

Package validation is not an oracle for this feature; neither is a paragraph-local text comparison. Accepting a paragraph-mark deletion **crosses the paragraph boundary**, so a paragraph-scoped round-trip check cannot observe the merge.

Every WP08b fixture must therefore:

1. Reconstruct **body-scoped** (or at minimum a window of source paragraph ± 2) canonical text for **both the accepted view and the rejected view**, and compare each exactly against expectation. The rejected view is where duplicate-text regressions surface; the accepted view alone would pass the reported bug.
2. Assert the merge target of each paragraph-mark resolution explicitly, not inferred from resulting text.
3. Fail closed with a structured mismatch code and byte-exact rollback on divergence, matching the established `PATCH_ROUNDTRIP_MISMATCH` pattern.

#### Planned Implementation Areas

- `services/document-operation-applier.js`: route explicit restoration intent; retain the rejected-view mutation guard for unsupported generic mutations.
- `services/document-operation-mutations.js`: add the paragraph-level restoration mutation and sibling placement; extend the inserted-paragraph builders to emit inserted paragraph marks and sanitized `pPr` (see Paragraph-Mark Semantics item 3).
- `core/paragraph-targeting.js`: resolve the deleted paragraph identity consistently across accepted and rejected metadata without allowing stale descriptors; reuse the existing `w:sectPr` refusal.
- Revision allocator and receipt collector: report every paragraph-mark and content revision ID allocated by the restoration, plus dropped-anchor warnings.
- `core/redline-validation.js`: add the narrowed resurrection-state warning (see Validation Predicate).

#### Required Test Matrix

1. Plain whole-paragraph deletion restored verbatim.
2. Restored paragraph adjusted while being restored.
3. Bold, italic, underline, and mixed-run formatting preservation.
4. Numbered and bulleted paragraph restoration without list-label drift.
5. Paragraph immediately before `w:sectPr`, **and** a paragraph whose own `pPr` carries `w:sectPr` (refusal).
6. Paragraph inside a table cell; paragraph inside a row deleted via `w:trPr/w:del` (refusal).
7. Deleted paragraph containing bookmarks or comment anchors: assert anchor counts are **unchanged**, that no name or comment ID appears twice, and that each dropped anchor is reported in the receipt.
8. Multiple adjacent deleted paragraphs restored independently and as a range, asserting one inserted sibling per source paragraph and preserved order.
9. Same-author deletion behavior remains governed by `merge-same-author` and is not routed through cross-author restoration.
10. Each non-triggering row of the trigger taxonomy routes to its stated path and not to WP08b.
11. Third-author follow-up edits to Reviewer B's restored paragraph continue to use ordinary cross-author slicing.
12. Repeat application of the same restoration is idempotent — no duplicate Reviewer B paragraph.
13. Source paragraph retains unresolved content (partial deletion, or a third author's pending `w:ins`) — proves the merge target is A's original successor and that surviving content does not land inside Reviewer B's paragraph.
14. `w:moveFrom` source paragraph (refusal); final-paragraph-of-body source (refusal).
15. Atomic runner rollback and progressive batch receipts.
16. Strict targeting by paragraph ID, index, fingerprint, and `revisionView: 'rejected'` diagnostics — including an assertion that Reviewer A's and Reviewer B's paragraphs carry distinct `w14:paraId` values.

Every successful fixture must assert exact current text, both-view body-scoped round-trip equality, structural validation, unique revision IDs, receipt reconciliation, paragraph ordering, and the six lifecycle outcomes above.

#### Implementation Record

* **Public Contract**: Added a dedicated `restore` document operation. A single restoration accepts a non-empty `modified` string; a contiguous range accepts one string per source paragraph. `generateRedlines: false` is rejected because restoration necessarily creates both a tracked content insertion and an inserted paragraph mark. Generic `redline` operations remain protected by WP08a.
* **Paragraph Model**: Each counterproposal is inserted immediately before the foreign-deleted source paragraph (or, for a range, as one contiguous inserted block before the source block). The source paragraph is not modified. The new paragraph receives two allocator-issued revisions, a fresh `w14:paraId`, no copied `w14:textId`/`w:rsid*`, and only allowlisted paragraph properties.
* **Lifecycle Oracle**: Before commit, restoration validates the authored OOXML and compares body-scoped paragraph text vectors for current view, Accept All, and Reject All against independently constructed expected documents. Any mismatch returns `PATCH_ROUNDTRIP_MISMATCH`; the operation savepoint supplies byte-exact rollback.
* **Idempotency**: An identical adjacent same-author restoration is a no-op. A changed same-author reapplication replaces the prior counterproposal rather than adding a duplicate paragraph.
* **Files Touched**:
  - `core/paragraph-revision-safety.js`
  - `core/paragraph-targeting.js`
  - `services/document-operation-contract.js`
  - `services/document-operation-applier.js`
  - `services/document-operation-mutations.js`
  - `services/operation-preflight.js`
  - `services/standalone-operation-runner.d.ts`
  - `docs/schemas/document-operations.schema.json`
  - `index.d.ts`
  - `tests/paragraph_level_cross_author_restoration_tests.mjs` (NEW)
  - `tests/types/usage.ts`
  - `README.md`
  - `AGENTS.md`
  - `CHANGELOG.md`
  - `docs/plans/2026-09-08-cross-author-revision-slicing.md`
* **Functions and Types Touched**:
  - `getParagraphRestorationRefusal` and move-range/content-state helpers (NEW): distinguish deleted-row, move-from, section-break, and unsafe-placement refusals while preserving the narrow trigger taxonomy.
  - `resolveTargetParagraph` (MODIFIED): supports strict fingerprint-only descriptors, bringing runtime targeting into alignment with the published schema.
  - `getCanonicalOperationType`, `normalizeDocumentOperation`, and `validateDocumentOperation` (MODIFIED): normalize and validate explicit single/range `restore` operations and retain full `targetEnd` descriptors.
  - `applyOperationToDocumentXml` (MODIFIED): routes restoration separately and keeps rejected-view mutation read-only for both range endpoints.
  - `restoreDeletedParagraphByExactText` (NEW): resolves the source block, enforces trigger/safety rules, reconstructs rejected-view content, inserts tracked siblings, handles idempotency, and runs the lifecycle oracle.
  - `createSanitizedRestorationPPr`, `buildRejectedRestorationTemplate`, `editRestorationTemplate`, `allocateFreshParagraphId`, `trackRestoredParagraph`, and lifecycle/anchor helpers (NEW): implement property sanitization, formatting preservation, fresh identity, dropped-anchor diagnostics, and exact round-trip checks.
  - `wrapParagraphContentInInsertion`, `buildFallbackInsertedPlainParagraph`, and `buildInsertedPlainParagraph` (MODIFIED): optionally emit inserted paragraph marks, use typed insertion receipt metadata, sanitize restoration properties, and accept fresh paragraph identity.
  - `preflightOperations` (MODIFIED): recognizes restoration state, range cardinality, and structural refusals without mutating the document.
  - `RestoreDocumentOperation` and `RedlineError` (MODIFIED/NEW): publish the operation shape and structured refusal/oracle metadata.
  - WP08b fixture helpers and assertions (NEW): exercise six lifecycle outcomes, strict descriptors, independent and range restoration, formatting/list preservation, anchor warnings, table cells/deleted rows, section/move/placement refusals, taxonomy exclusions, progressive and atomic batches, idempotency, and third-author follow-up slicing.
* **Verification**:
  - `node tests/paragraph_level_cross_author_restoration_tests.mjs` — PASS.
  - Focused list and insertion-affinity regressions — PASS.
  - `$env:DOCX_TEST_CONCURRENCY='1'; npm test` — PASS, 97 test files passed and 0 failed. The serial final run was used after concurrent attempts hit unrelated per-file timeouts under host contention; each timed-out suite also passed directly.
  - `npm run lint` — PASS.
  - `npm run check:types` — PASS; all 123 runtime exports have declarations.
  - `npm run build` — PASS.
  - `git diff --check` — PASS (line-ending conversion notices only; no whitespace errors).

---

### WP09 — Real-Document Mutation Reliability and Agent-Safe Failure Reporting [WP09a-e COMPLETE]

#### Planning and characterization update (2026-09-08)

- Updated `docs/plans/2026-09-08-cross-author-revision-slicing.md`: added the three failure classes, the Word Desktop structural oracle, WP09a-e requirements, ordered implementation handoff, sanitized operation examples, and the planned file/function map. Removed private party, person, corpus, URL, clause, and commercial wording from permanent examples.
- Added `tests/word_deleted_section_edit_oracle_tests.mjs`: introduced local fixture/test helpers `localName`, `descendants`, `elementChildren`, `revisionElements`, `parse`, `plainText`, `assertUniqueRevisionIds`, and `assertValidOracle`; added sanitized inline deletion-carrier and post-source paragraph-restoration OOXML fixtures; asserted direct-child order, authorship, paragraph-mark behavior, global revision-ID uniqueness, validation, and selective/all-author lifecycle outcomes.
- No production mutation function has been changed for WP09 yet. The new test is a characterization oracle for the structures that the later public-runner and package-facade implementation tests must generate.
- Verification: focused oracle test passed; serial `$env:DOCX_TEST_CONCURRENCY='1'; npm test` passed **98/98 test files** with the new suite included; `npm run lint` passed; `git diff --check` reported no whitespace errors.

#### WP09a-b implementation update (completed 2026-09-09)

WP09a and WP09b are implemented. WP09c, WP09d, and WP09e remain planned and are not implied by this completion record.

**Production files and functions changed:**

- `services/document-operation-mutations.js`
  - `applyToParagraphByExactText`: sends the resolved paragraph's exact accepted-view text to the engine for every text-bearing single-paragraph edit. The legacy caller-text fallback remains only for format-only field-code paragraphs whose canonical accepted view has no extractable text spans.
  - `resolveTargetParagraph`: attaches bounded `targetTextMatch` metadata to the resolved target captured in operation results and receipts.
  - Added internal `escapeInvisibleText`, `codePointLabel`, and `describeTargetTextMatch` helpers. They distinguish `exact`, one-to-one ordinary-space/NBSP `space_equivalent`, and broader `normalized` matches and report at most eight code-point differences.
- `pipeline/diff-engine.js`
  - Added `computeCharacterDiffs`, a character-local diff without semantic cleanup for refining whitespace substitutions that word tokenization grouped with unchanged content.
- `engine/surgical-mode.js`
  - `applySurgicalMode`: refines adjacent delete/insert hunks that differ only by ordinary spaces and NBSPs, preserving unchanged hyperlink runs rather than deleting and reconstructing their visible URL text.
  - Plain-text edit groups now apply from right to left. Each group rebuilds the live surgical span index, so an earlier run split/removal cannot leave a stale DOM anchor for a later hunk.
  - Added internal `refineSpaceEquivalentReplacements`, `collectTextEditOperations`, and `codePointAtOffset` helpers. `PATCH_ROUNDTRIP_MISMATCH` now includes `expectedCodePoint` and `actualCodePoint` at the first difference while still returning the exact input OOXML.
- `node/cli.js`
  - Raised `CLI_CONTRACT_VERSION` to 3 and added the `compact-mutation-results` capability.
  - `executeCli` now passes `apply`, `accept`, `reject`, and `delete-comments` results through `compactMutationResult`.
  - Added `boundedText`, `compactError`, `compactResolvedTarget`, `compactReceipt`, `compactOperationResult`, `summarizeIssues`, and `compactMutationResult`.
  - Normal mutation stdout omits `documentXml`, `oxml`, comments/numbering XML, inspection payloads, raw package buffers, and full issue arrays. It retains per-operation errors/receipts, `written`, `outputPath`, compact issue counts, and a derived `completion` flag. Resolved target text is removed while bounded match/code-point diagnostics remain.
- `services/standalone-operation-runner.d.ts`
  - Extended `ResolvedDocumentTarget` with the typed `targetTextMatch` diagnostic contract.

**Tests changed:**

- `tests/cross_author_slicing_hyperlink_roundtrip_tests.mjs`
  - Replaced the identifying policy sample with a synthetic Service Policy fixture.
  - Added an ASCII-space target against an NBSP source, two code-point assertions, hyperlink relationship/history preservation, exact Accept-All output, and exact Reject-current-author restoration of the NBSP-bearing source.
- `tests/agent_cli_tests.mjs`
  - Added a complete DOCX/CLI ASCII-space-target versus NBSP-source regression using strict paragraph identity.
  - Asserts exact written text, `space_equivalent` diagnostics, source restoration after Reject, compact validation summaries, omitted resolved clause/XML text, and truthful `completion`.
- `tests/agent_cli_edge_tests.mjs`
  - Added failed-apply assertions for `written: false`, `outputPath: null`, `completion: false`, summarized validation, no `documentXml`, bounded stdout, no unrelated body text, and the actionable per-operation error.
- `tests/plugin_wrapper_compatibility_tests.mjs`
  - Sanitized the comment-author fixture and verified compact CLI errors retain bounded comment author/text details required to resolve `COMMENTED_CONTENT_DELETE`.

The first full serial compatibility run exposed two retained-contract requirements and was not treated as final: compact errors initially omitted protected-comment details, and the strict source-truth change initially removed the established caller-text fallback for format-only field-code paragraphs with no canonical text spans. `compactError` now retains bounded comment records, and `applyToParagraphByExactText` preserves that non-text fallback. Both formerly failing suites pass directly; the final serial result is recorded below after rerun.

**Documentation changed:** `README.md`, `CHANGELOG.md`, `ARCHITECTURE.md`, `AGENTS.md`, and this plan now describe source-truth mutation alignment, invisible-character diagnostics, reverse live-span application, CLI contract version 3, compact validation summaries, and completion semantics.

**Final WP09a-b verification:**

- `node tests/cross_author_slicing_hyperlink_roundtrip_tests.mjs` — PASS.
- `node tests/agent_cli_tests.mjs` — PASS.
- `node tests/agent_cli_edge_tests.mjs` — PASS.
- `node tests/plugin_wrapper_compatibility_tests.mjs` — PASS.
- `node tests/standalone_operation_runner_tests.mjs` — PASS.
- `$env:DOCX_TEST_CONCURRENCY='1'; npm test` — PASS, **98/98 test files** and 0 failed.
- `npm run lint` — PASS.
- `npm run check:types` — PASS; all 123 runtime exports have declarations.
- `npm run build` — PASS.
- `git diff --check` — PASS (line-ending conversion notices only; no whitespace errors).

#### WP09c-e implementation update (completed 2026-09-09)

WP09c, WP09d, and WP09e are implemented. The implementation keeps generic
accepted-view mutation fail-closed and adds only the explicit rejected-view
operation described below.

**Production files and functions changed:**

- `core/validation-delta.js` (new)
  - Added internal `issueKey`, exported `subtractValidationIssueMultiset`, and
    exported `validationErrors`. Validation differences retain multiplicity and
    include source/severity/code/message in the stable signature.
- `core/revision-cloning.js`
  - Added `clonePropertiesWithoutRevisionHistory`. New paragraphs and runs may
    inherit effective properties, but cloned `ins`, `del`, move, `pPrChange`,
    `rPrChange`, table/row/cell property-change, and section-property history is
    removed rather than duplicated with stale identities.
- `core/paragraph-revision-safety.js`
  - `getParagraphRestorationRefusal` now accepts
    `requireFollowingParagraph: false` for an inline deletion-carrier edit that
    creates no paragraph. All section, move, and deleted-table-row refusals stay
    active; paragraph restoration still requires a following sibling.
- `core/redline-validation.js`
  - `validateRedlineOoxml` accepts an already-parsed document DOM for internal
    operation checks, avoiding an extra full-source parse while retaining the
    public string input. All existing validation rules are unchanged.
- `engine/surgical-spans.js`
  - `getRunChildText` and `isTextLikeRunChild` recognize `w:delText`, allowing
    the shared run-piece splitter to address rejected-view deletion text.
- `engine/surgical-run-splitting.js`
  - `splitTrackChangeCarrier` now supports `w:del` as well as `w:ins`, emits
    `w:delText` on split deletion runs, preserves the leading carrier identity,
    allocates the trailing carrier identity, refreshes duplicated
    `w:rPrChange` IDs, and records the allocation in the active receipt.
  - `cloneRunPiece` preserves tab, break, soft-hyphen, and non-breaking-hyphen
    elements while slicing deletion carriers instead of flattening those
    controls into `w:delText`.
- `services/document-operation-contract.js`
  - `getCanonicalOperationType` maps only `type: "insert"` plus a rejected-view
    target to `rejected-insert`.
  - `normalizeDocumentOperation` normalizes `anchor.exactText`, `occurrence`,
    and `offset` while retaining whether occurrence was supplied explicitly.
  - `validateDocumentOperation` requires non-empty inserted text, an in-range
    anchor-relative offset, and `existingRevisions: "slice-cross-author"`.
- `services/document-operation-mutations.js`
  - Added `insertIntoRejectedDeletedText` and its exact occurrence, deletion
    carrier, formatting-clone, and unsupported-boundary helpers. It resolves a
    strict rejected-view paragraph, requires a wholly foreign-deleted state,
    refuses repeated anchors without an explicit occurrence, splits the direct
    deletion carrier, emits the new author's sibling `w:ins`, and verifies that
    the rejected view is unchanged while the accepted view exposes the inserted
    text. Comments, bookmarks, fields, hyperlinks, moves, and other non-text
    split markup fail closed with `UNSAFE_REVISION_BOUNDARY`.
  - `buildInsertedListParagraph`, `buildEmptyParagraphTemplateFromAnchor`, and
    `wrapParagraphContentInInsertion` use
    `clonePropertiesWithoutRevisionHistory` for effective property inheritance.
  - `verifyParagraphRestorationLifecycle` validates baseline/output issue
    multisets, separately validates each inserted mutation envelope, returns
    `GENERATED_OOXML_INVALID` for generated markup defects, and preserves the
    exact current/Accept-All/Reject-All paragraph-vector checks.
  - `followingParagraphBlock`, `buildExpectedRestorationDocument`, and
    `restoreDeletedParagraphByExactText` now detect, verify, replace, and emit
    restoration blocks immediately after the complete deleted source range.
- `services/document-operation-applier.js`
  - `applyOperationToDocumentXml` dispatches `rejected-insert` and permits a
    rejected target only for explicit `rejected-insert` and `restore`
    operations. Before marking any changed operation committed, it validates
    the entire live document against its DOM savepoint by issue multiset. A
    generated error restores the savepoint and returns
    `GENERATED_OOXML_INVALID` with a refused receipt.
- `services/operation-preflight.js`
  - Same-target conflict grouping includes `rejected-insert`, so it cannot evade
    overlap diagnostics merely because it uses a distinct canonical kind.
- `node/docx-document.js`
  - `DocxDocument.applyOperations` now applies the same multiset baseline-delta
    classification to document and package validation. Unchanged legacy defects
    remain in `validation.originalIssues`; introduced errors block the write and
    are returned as generated issues.
- `docs/schemas/document-operations.schema.json`
  - Added `rejectedTextInsertionAnchor` and the optional `anchor` field on the
    compatible `insert` shape; runtime validation makes it mandatory for a
    rejected-view target.
- `services/standalone-operation-runner.d.ts`
  - Added `RejectedTextInsertionAnchor` and exposed `anchor` on insert/redline
    operation declarations.

**Tests changed or added:**

- `tests/paragraph_level_cross_author_restoration_tests.mjs`
  - Updated current-view, selective-author, range, idempotency, progressive,
    table-cell, and follow-up edit assertions for post-source restoration.
- `tests/cross_author_carrier_splitting_tests.mjs`
  - Extended the allocator receipt assertion so a refreshed cloned
    `w:rPrChange` and the trailing carrier are both recorded in allocation order.
- `tests/docx_package_facade_tests.mjs`
  - Updated the intentionally malformed-package fixture to prove an unchanged
    package defect stays in `originalIssues` while a safe document mutation may
    still be written with zero generated issues.
- `tests/merge_same_author_tests.mjs`
  - Corrected an intermittent false-positive assertion that searched all OOXML
    for the bare string `45` and therefore failed whenever a legitimate revision
    ID reached 45. It now checks the intended intermediate phrase `45 days`.
- `tests/wp09c_e_validation_identity_rejected_insert_tests.mjs` (new)
  - Proves restoration succeeds over an unrelated duplicate-ID baseline;
    paragraph expansion does not clone `pPrChange`/`rPrChange` history; rejected
    insertion produces `pPr / del(A) / ins(B) / del(A)` with globally unique
    IDs; rejected text remains exact; selective/all-author lifecycle outcomes
    match the minimized Word shape; missing policy/anchor inputs fail schema
    validation; and both dirty-baseline restoration and rejected-view insertion
    write successfully through `openDocx(...).applyOperations` with zero
    generated package issues. Oracle reinforcement covers deletion start/end,
    multi-run and formatted-run boundaries, tab/break preservation,
    repeated-anchor disambiguation, comment/hyperlink/field refusal, explicit
    rejected-view range restoration, all seven lifecycle views, and a combined
    inline-insertion plus post-source range-restoration forward-merge case.
- `tests/word_deleted_section_edit_oracle_tests.mjs`
  - Remains the sanitized Word-authored structural characterization oracle for
    inline deletion slicing and post-source range restoration.

**Documentation changed:** `README.md`, `CHANGELOG.md`, `ARCHITECTURE.md`,
`AGENTS.md`, the operation JSON Schema, declarations, and this plan describe
baseline-delta validation, effective-property clone sanitation, post-source
restoration, and the explicit rejected-view insertion contract.

**Final WP09c-e verification:**

- `node tests/paragraph_level_cross_author_restoration_tests.mjs` — PASS.
- `node tests/wp09c_e_validation_identity_rejected_insert_tests.mjs` — PASS.
- `node tests/word_deleted_section_edit_oracle_tests.mjs` — PASS.
- `node tests/cross_author_carrier_splitting_tests.mjs` — PASS.
- `node tests/docx_package_facade_tests.mjs` — PASS.
- `node tests/performance_phase1_session_tests.mjs` — PASS; the full source is
  still parsed once and the live document serialized once.
- 30 consecutive isolated `merge_same_author_tests.mjs` runs — PASS after
  correcting the bare-revision-ID false positive.
- `$env:DOCX_TEST_CONCURRENCY='1'; npm test` — PASS, **99/99 test files** and 0 failed.
- `npm run lint` — PASS.
- `npm run check:types` — PASS; all 123 runtime exports have declarations.
- `npm run build` — PASS.
- `git diff --check` — PASS (line-ending conversion notices only; no whitespace errors).

#### Motivation and First Bug Report

The first post-WP08 report is not a paragraph-restoration case and does not involve a foreign paragraph-mark deletion. It is an ordinary single-paragraph replacement in a paragraph with two external hyperlinks and non-breaking spaces (`U+00A0`) immediately before both URLs and after the Service Policy URL.

The requested edit adds an execution-date qualifier around the second URL while leaving the Processing Schedule URL alone. The operation supplied ordinary spaces (`U+0020`) in its `target.exactText` and `modified` strings. Strict descriptor resolution still selected a uniquely identified paragraph by paragraph ID/fingerprint because paragraph targeting treats ordinary spaces and NBSPs as equivalent. The mutation then failed closed:

```text
PATCH_ROUNDTRIP_MISMATCH at offset 934
expected: "Service Policy available at example.invalid/policy/ that ..."
actual:   "Service Policy available at\u00a0 example.invalid/policy/ that ..."
```

Adding another ordinary space on retry made the actual sequence `NBSP + two ordinary spaces`; it did not consume the source NBSP. A later `extract` exposed the hidden NBSPs, and the agent constructed a third operation with them, but the transcript contains no third `apply` command or successful result. It nevertheless reported the document as complete. No output path was produced by either recorded application.

The failed CLI response also serialized the entire large `documentXml` payload and repeated 21 pre-existing `MISSING_SPACE_PRESERVE` issues plus one pre-existing commentsExtended content-type issue. `generatedIssues` was empty. This noise obscured the actionable operation error and contributed to an unreliable agent recovery loop.

The exact source paragraph was recovered from the failed result. It contains no open revisions itself; its important shape is:

```xml
<w:r><w:t>... Processing Schedule available at&#xA0;</w:t></w:r>
<w:hyperlink r:id="rId7"><w:r><w:t>example.invalid/schedule</w:t></w:r></w:hyperlink>
<w:r><w:t>... Service Policy available at&#xA0;</w:t></w:r>
<w:hyperlink r:id="rId8"><w:r><w:t>example.invalid/policy/</w:t></w:r></w:hyperlink>
<w:r><w:t>&#xA0;(the “</w:t></w:r>
```

This is therefore a distinct gap in replacement alignment and failure ergonomics. WP08a/b remain unchanged.

#### Second Bug Report — Restore Blocked by Pre-Existing Validation Errors

The second report exercises the explicit WP08b `restore` operation against a paragraph wholly deleted by another reviewer. Target resolution succeeds by paragraph ID in the accepted/current view, where the paragraph text is empty; the rejected view exposes the deleted restrictions paragraph used to draft the adjusted restoration. The operation then fails before commit with:

```text
PATCH_ROUNDTRIP_MISMATCH
stage: "validation"
message: repeated MISSING_SPACE_PRESERVE errors
written: false
```

Those `MISSING_SPACE_PRESERVE` errors already appear in `validation.originalIssues`. They occur elsewhere in the source document and were not created by the restoration. The package facade knows the baseline is imperfect, but `verifyParagraphRestorationLifecycle` calls `validateRedlineOoxml(outputXml)` and rejects the complete output whenever *any* error exists. It never validates `beforeXml`, subtracts the baseline, or scopes the errors to the inserted restoration paragraphs. This makes WP08b unusable on a real document that contains an unrelated legacy validation defect, even when the proposed restoration itself is valid.

The recovery then abandons `restore` and considers inserting the paragraph after a different surviving paragraph. That is not an equivalent fallback: it can change clause order, list numbering, paragraph-mark lifecycle behavior, and the relationship between the restored counterproposal and the original deleted paragraph. The transcript says the insertion worked but contains neither its apply result nor a validated output record. WP09 must prohibit silent operation-type fallback and make successful completion independently provable.

#### Third Bug Report — Paragraph Expansion Duplicates Revision IDs

The third report asks to restore the adjacent `8.4 Changes; No Waiver` heading and body, while modifying the body so an email satisfies the signature/writing requirement. It reconfirms the WP09c dirty-baseline failure: even a trivial explicit restore of `"Synthetic restoration text"` is rejected because the same 21 pre-existing `MISSING_SPACE_PRESERVE` issues are treated as generated validation failures.

It also exposes a separate structural mutation defect. The agent tried to append a new paragraph to an existing clause with both `"\n\n"` and `"\n"`. In both cases the operation returned `results[0].status: "applied"`, but the package facade subsequently rejected the output:

```text
status: error
written: false
error: PACKAGE_OPERATION_FAILED
message: Applied operations introduced invalid revision markup (DUPLICATE_REVISION_ID)
results[0]: applied
```

The duplicate appears only after expanding one paragraph into multiple tracked paragraphs. The likely implementation surface is the plain-adjacency/structured paragraph builder: `buildEmptyParagraphTemplateFromAnchor` clones the anchor's `pPr` and first-run `rPr` verbatim, while `wrapParagraphContentInInsertion` clones those properties again. Existing `w:pPrChange` or `w:rPrChange` descendants can therefore carry their old `w:id` into a newly inserted paragraph even though the new content and paragraph mark use the live document allocator. This hypothesis must be proven by recording the duplicated ID and both owning elements before choosing whether inherited change-history elements should be stripped or deliberately cloned with fresh IDs.

The run also demonstrates why structural validity alone is insufficient. An earlier workaround replaced the text of a surviving `3.1 Restricted Use` paragraph with restored `3.2` language; the generated redline was internally coherent but semantically targeted the wrong clause. For 8.4, another workaround inserted only the heading into a preceding amendment sentence, producing same-author deletions of that sentence and an insertion of the heading. A final baseline validation reported no *structural* issues, yet the requested standalone 8.4 heading/body was delivered inline after an earlier subsection instead. WP09 must require a document-shape oracle—paragraph count, order, identity, and exact neighboring text—for structural operations and restorations.

#### Microsoft Word Desktop Oracle — Edits Inside a Deleted Section

The user supplied a complete `word/document.xml` after making two edits directly in Microsoft Word Desktop. This is a first-party structural oracle, not an engine-generated hypothesis. The full source XML must remain an external/private acceptance artifact; implementation must reduce the relevant paragraphs to minimal checked-in fixtures.

**Oracle A — inline insertion inside a deleted subsection body.** A synthetic `8.1 Standard Charges` body has a paragraph-mark deletion by `Reviewer A` and all original text is deleted. Word inserts `REVIEWER B INSERTION` at a rejected-view offset by splitting Reviewer A's deletion carrier:

```xml
<w:p>
  <w:pPr><w:rPr><w:del w:id="710" w:author="Reviewer A"/></w:rPr></w:pPr>
  <w:del w:id="711" w:author="Reviewer A">The account holder must pa</w:del>
  <w:ins w:id="712" w:author="Reviewer B">REVIEWER B INSERTION</w:ins>
  <w:del w:id="713" w:author="Reviewer A">y each undisputed invoice.</w:del>
</w:p>
```

Word preserves the leading carrier's ID/metadata, gives the trailing split carrier a fresh ID while retaining Reviewer A's author/date, and emits Reviewer B's insertion as a sibling—never nested inside `w:del`. It does not add an inserted paragraph mark because no new paragraph was created. This proves that foreign paragraph-mark deletion + all original content deleted + a foreign `w:ins` is not intrinsically invalid. The missing discriminator in WP08 is **intent and rejected-view anchoring**: a generic current-view insertion into an empty deleted paragraph remains ambiguous and fail-closed, while an explicit insertion at a uniquely resolved offset inside the foreign deletion is Word-native deletion-carrier slicing.

This operation is lifecycle-dependent by design. Rejecting Reviewer B removes only `REVIEWER B INSERTION`; rejecting Reviewer A restores the original sentence with Reviewer B's insertion at the exact `pa|y` offset. Accepting Reviewer A removes the deleted text and resolves its paragraph mark, so Reviewer B's surviving text follows Word's forward-merge semantics. WP09 must compare all selective Accept/Reject outcomes with Word Desktop, not assume the insertion remains an independent standalone paragraph.

**Oracle B — two new paragraphs between wholly deleted subsections.** Word leaves a deleted `8.2 Usage Adjustments` heading and body byte-for-byte intact, then inserts two new sibling paragraphs after that source block and before deleted `8.3 Collection Costs`:

```text
[del(A): 8.2 heading]
[del(A): 8.2 body]
[ins(B) paragraph mark + ins(B) content: 8.2 Usage Adjustments]
[ins(B) paragraph mark + ins(B) content: Usage above the stated threshold may be billed at the next tier.]
[del(A): 8.3 heading]
```

The heading paragraph uses five distinct synthetic IDs (720–724) for its paragraph-mark insertion, Word-authored property-change metadata, nested historical paragraph-mark snapshot, content insertion, and run-property change. The body uses IDs 725 and 726 for paragraph mark and content. All IDs are globally unique. The extra `w:rPrChange` nodes are Word UI artifacts and need not be reproduced byte-for-byte if the engine preserves equivalent effective bold formatting and lifecycle behavior; if emitted, however, each must receive a unique allocator ID.

Oracle B supersedes WP08b's pre-source placement decision for a fully deleted source range. The target behavior is now: insert the counterproposal block immediately **after the source range and before its original next paragraph**, matching Word Desktop. The source range remains untouched. Combined cases—such as Oracle A immediately before Oracle B—must be tested because accepting paragraph-mark deletions can cascade surviving inline content forward into the next inserted paragraph.

#### Diagnosis to Prove Before Mutation Changes

WP09 must first reduce the recovered paragraph and operation to a checked-in fixture and record the surgical diff/mutation trace. The implementation must determine which of these boundaries is wrong rather than patching the final string:

1. `computeWordDiffs` may align the source NBSP as unchanged while treating the requested ordinary space as an insertion after it when the surrounding phrase is also replaced.
2. The diff may be correct, but `processDelete` may fail to consume the NBSP at a run/hyperlink boundary.
3. `processInsert` may consume a stale or ambiguous replacement anchor and place the ordinary space beside the still-live NBSP.
4. The document runner may pass a space-equivalent caller target as the edit-coordinate baseline instead of the exact accepted-view text recovered from the resolved paragraph.

The regression must expose the diff tuples, original offsets, live DOM anchors, and final accepted text sufficiently to identify the failing layer. A fix is not accepted if it merely special-cases URLs, policy wording, or one observed offset.

#### Normative Text Contract

1. **Target selection and edit coordinates are separate concerns.** Space/NBSP equivalence may select a uniquely identified paragraph, but mutations MUST use the resolved paragraph's exact canonical accepted-view text as their source coordinate system.
2. **`modified` remains exact.** The engine must reconstruct the caller's requested `modified` string byte-for-byte at the JavaScript string level, including every `U+0020`, `U+00A0`, tab, and line break. WP09 must not make replacement text globally whitespace-normalized.
3. **Whitespace substitutions are real edits.** When exact source has NBSP and `modified` has an ordinary space, the output must track deletion/replacement of the NBSP; it must never retain the NBSP and append the ordinary space beside it.
4. **Unchanged hyperlink containers survive.** Both `rId7` and `rId8`, their `w:history` attributes, and their run formatting must be preserved. Text immediately outside a hyperlink must not be moved inside it, and URL text must not be reconstructed as a plain run.
5. **Fail closed remains mandatory.** `PATCH_ROUNDTRIP_MISMATCH` is doing the right thing by refusing incorrect OOXML. WP09 fixes the false mismatch; it must not weaken, normalize, or remove the exact accepted-view oracle.
6. **The runner must report how matching occurred.** If target resolution used space equivalence, the result/receipt should expose that fact and provide escaped source/caller excerpts or differing code points. Invisible characters must be diagnosable without a second ad hoc script.

#### WP09a — Source-Truth Replacement Alignment

1. Capture the recovered paragraph as a minimal fixture with the two hyperlinks, all three NBSP boundaries, bold/underline formatting on the defined term, and the reported replacement.
2. Make the exact resolved accepted-view text authoritative from target resolution through surgical span construction. A normalized target string may validate identity, but must never supply mutation offsets.
3. Normalize replacement hunks/anchors so an NBSP-to-space substitution adjacent to a retained or shifted hyperlink becomes one deletion plus one insertion at the same logical boundary.
4. Rebuild live span/anchor state after any deletion that detaches a run. Multiple replacement hunks in the same paragraph must not reuse stale pre-mutation nodes.
5. Preserve existing insertion-only behavior, foreign-carrier slicing, explicit insertion affinity, formatting, comments, bookmarks, fields, and hyperlink relationships.
6. Run the exact round-trip oracle after the complete paragraph mutation and return the byte-exact input OOXML on any mismatch.

#### WP09b — Compact, Actionable CLI Failures

1. Do not include full `documentXml` in normal CLI stdout for `apply`, `accept`, `reject`, or `delete-comments`. It remains available from library APIs where it is the actual programmatic result, but the CLI already communicates durability through the written DOCX and `outputPath`.
2. On a failed mutation, lead with the per-operation error and retain `written: false`, `outputPath: null`, `status`, `results`, and receipts. Include `mismatchOffset`, escaped excerpts, and code-point details for whitespace mismatches.
3. Summarize pre-existing validation issues by code/count in the normal mutation response. Keep `generatedIssues` explicit and provide full issue arrays only through `validate` or an explicit verbose diagnostics option if one is added.
4. A failed or partial result must be mechanically unmistakable. Add a compact `completion`/`success` signal only if it is derived from `written === true`, a non-error top-level status, and zero failed result entries; do not introduce a second contradictory status model.
5. CLI tests must prove that a failed application cannot emit an output path, cannot write a destination, and produces bounded stdout that does not contain `<w:document>` or contract body text.

The library cannot prevent an external agent from making a false narrative claim, but its default output must make the recorded mistake difficult: there must be no huge XML payload between the failure code and `written: false`, and no ambiguous success-looking field.

#### WP09c — Baseline-Delta Validation for Restoration and Package Mutation

1. `verifyParagraphRestorationLifecycle` must validate both `beforeXml` and `outputXml`. Pre-existing errors are baseline diagnostics, not generated-output failures.
2. Compare issues as a **multiset**, not a `Set`: code/message duplicates occur many times in real Word documents. An additional occurrence after mutation is generated even when its code/message matches a baseline issue.
3. Validate every newly inserted or modified restoration paragraph independently. A generated `<w:t>` or `<w:delText>` with missing `xml:space="preserve"` must fail even if the source already contains the same error elsewhere.
4. Full-document validation must still catch document-scoped failures such as duplicate revision IDs, invalid revision nesting, unsafe paragraph restoration state, or duplicated structural anchors. Baseline subtraction must not become a blanket bypass.
5. Apply the same baseline-delta semantics at the package facade. Pre-existing package defects must remain visible in `originalIssues`; newly introduced defects belong in `generatedIssues` and fail the write. If a package defect cannot be safely classified, fail closed with a package-validation code rather than mislabeling it as a text round-trip mismatch.
6. Reserve `PATCH_ROUNDTRIP_MISMATCH` for current/Accept-All/Reject-All text-or-lifecycle divergence. Validation failures should return a distinct structured code such as `GENERATED_OOXML_INVALID`, with `stage: "validation"`, generated issue details, and the original document unchanged.
7. A failed `restore` must never be auto-converted to `redline`, `replace`, or an insertion beside a convenient surviving paragraph. Recovery requires either correcting the reported cause and retrying the same restoration or explicit caller authorization for a semantically different operation.
8. Successful restoration through the package facade must prove `written: true`, a non-null `outputPath` at the CLI layer, one applied result with a committed receipt, zero generated validation issues, and exact lifecycle oracle results.

#### WP09d — Revision Identity and Shape Oracles for Paragraph Expansion

1. Every revision-bearing element introduced into the live document—including `w:ins`, `w:del`, paragraph-mark revisions, `w:rPrChange`, `w:pPrChange`, table/row revisions, and deliberately preserved cloned revisions—must have a document-unique allocator-issued `w:id`.
2. New paragraphs may inherit effective paragraph/run formatting, but MUST NOT blindly inherit the anchor paragraph's revision history. Define a shared clone policy:
   - strip `w:pPrChange`, `w:rPrChange`, and other historical `*Change` descendants when only effective formatting is needed; or
   - when revision history is intentionally preserved, deep-clone it and refresh every revision ID through the live document allocator.
3. `buildEmptyParagraphTemplateFromAnchor`, `buildInsertedPlainParagraph`, `buildFallbackInsertedPlainParagraph`, `wrapParagraphContentInInsertion`, and list/structured paragraph builders must use that shared policy. No builder may call `cloneNode(true)` on revision-bearing properties without explicit sanitization or ID refresh.
4. Perform an operation-level global revision-ID uniqueness check before an operation is reported as applied. A duplicate introduced by the mutation must roll back the operation savepoint, mark its receipt refused/uncommitted, and return a per-operation generated-markup error. It must not appear as `results[i].status: "applied"` followed by only a top-level package error.
5. Preserve the package-level duplicate-ID check as defense in depth. Operation-level and package-level checks must agree on the offending ID and element kinds.
6. Structural paragraph operations require shape postconditions in addition to accepted text: exact paragraph count delta, sibling order, fresh `w14:paraId`, unchanged anchor/source text where the operation is insertion-only, and correct paragraph-mark ownership.
7. Accept All and Reject Current Author must be checked at paragraph-vector scope. Rejecting a paragraph expansion must restore the original paragraph vector without empty stubs; accepting must retain the intended standalone paragraphs in order.
8. Restoring an adjacent deleted heading/body pair should use one explicit range `restore` operation with two `modified` strings, not a sequence that targets the newly created heading or a multiline replacement of an unrelated surviving clause. The body may be adjusted during restoration to add the email-sufficiency sentence.

#### WP09e — Word-Native Deleted-Section Editing

1. Add an explicit, unambiguous contract for inserting at a rejected-view offset inside foreign deleted content. Reuse `type: "insert"` only if it can require a strict deleted-text anchor/offset and `revisionView: "rejected"`; otherwise add a dedicated operation shape. Do not reinterpret an ordinary empty accepted-view `redline` as this intent.
2. Resolve the deleted carrier by paragraph identity plus an exact, unique rejected-view anchor (or explicit offset tied to a fingerprint). Stale fingerprints, repeated anchors, move revisions, comments crossing the split, and ambiguous offsets fail closed.
3. Split a foreign `w:del` carrier at the exact run-piece offset. Preserve the leading carrier identity, allocate a fresh ID for each trailing carrier, preserve original author/date on split carriers, convert/retain `w:delText` correctly, and preserve formatting, tabs, breaks, rendered page breaks, fields, bookmarks, and hyperlinks.
4. Insert Reviewer B's `w:ins` as a sibling between deletion carriers. Never nest `w:ins` inside `w:del`; never add a paragraph-mark insertion unless the operation actually creates a paragraph.
5. Narrow `inspectForeignDeletedParagraphTarget` and `findForeignDeletedParagraphResurrections`: explicit, structurally anchored deletion-carrier slicing is allowed and must not produce `FOREIGN_PARAGRAPH_MARK_DELETION`; ambiguous generic resurrection stays protected by WP08a.
6. Update range restoration placement to insert Reviewer B's sibling block immediately **after** the complete foreign-deleted source range and before the range's original next paragraph, matching Oracle B. Idempotency detection and changed-restoration replacement must recognize the new side of the source block.
7. Preserve WP08b property sanitization and fresh identity requirements. Word's incidental `rsid`, `textId`, and property-change history are not required output, but effective heading/body formatting, separate paragraph/content revisions, and global ID uniqueness are required.
8. Extend lifecycle oracles to the entire affected section and compare engine resolution with Word Desktop for Current, Accept All, Reject All, Accept A, Reject A, Accept B, and Reject B. Explicitly test forward-merge cascades across adjacent deleted paragraphs and inserted blocks.
9. The validator must accept the minimized Word-authored Oracle A shape. It may warn about lifecycle dependency, but must not label a first-party Word structure as an unsafe resurrection solely from the presence of a foreign insertion in an otherwise deleted paragraph.

#### Required Test Matrix

1. Exact synthetic two-hyperlink Service Policy fixture: ASCII-space operation against NBSP source; current view and Accept All equal `modified` exactly.
2. The same fixture through `applyOperationsToDocumentXml`, `openDocx(...).applyOperations`, and CLI `apply` with strict paragraph ID/fingerprint targeting.
3. Qualifier inserted before the second hyperlink, after it, and on both sides; each permutation with `U+0020 -> U+00A0`, `U+00A0 -> U+0020`, and unchanged NBSP.
4. Two hyperlinks in one paragraph where only the second surrounding clause changes; assert both relationship IDs and hyperlink attributes survive.
5. Repeated URL text and repeated surrounding prose so placement cannot rely on the first string occurrence.
6. NBSP as its own run, at the end of the run before a hyperlink, at the start of the run after a hyperlink, and inside a foreign `w:ins` carrier.
7. Multiple replacement hunks in one paragraph, including one earlier whitespace substitution and the later reported qualifier replacement.
8. Leading/trailing spaces, consecutive ordinary spaces, tabs, narrow NBSP (`U+202F`), word joiner (`U+2060`), and non-breaking hyphen (`U+2011`) remain distinct unless the contract explicitly declares equivalence.
9. Explicit insertion affinity at both hyperlink boundaries remains authoritative.
10. Reject-current-author restores the exact pre-operation source, including NBSPs; selective Accept/Reject of foreign authors retains valid lifecycle behavior.
11. Structural validation, unique revision IDs, receipt reconciliation, hyperlink preservation, and exact current/Accept-All/Reject-current views for every successful fixture.
12. Forced round-trip mismatch still returns `PATCH_ROUNDTRIP_MISMATCH`, byte-exact OOXML rollback, and bounded escaped diagnostics.
13. CLI failure over a large document omits `documentXml`, does not echo clause text, stays below a fixed response-size ceiling, and reports `written: false`/`outputPath: null` adjacent to the actionable error.
14. Pre-existing validation defects are summarized separately from generated defects; zero `generatedIssues` must remain obvious.
15. Progressive batch with one success and one failure reports `status: "partial"` and is never marked complete; atomic mode writes nothing and rolls back exactly.
16. Explicit restoration in a document containing pre-existing `MISSING_SPACE_PRESERVE` errors succeeds when the restoration introduces no new issues; the original issue counts remain reported.
17. Restoration that itself emits one missing `xml:space="preserve"` fails with `GENERATED_OOXML_INVALID` even when identical baseline errors already exist.
18. Duplicate baseline issues are compared by multiplicity: N baseline occurrences plus one generated occurrence yields exactly one generated issue.
19. Baseline issue removed in one location and reintroduced in a newly authored paragraph is still detected by mutation-envelope validation rather than hidden by equal aggregate counts.
20. Pre-existing document-level revision warnings remain visible but do not block a structurally valid restoration; new duplicate IDs, nested revisions, duplicated anchors, and unsafe restoration shapes still fail.
21. Package-facade fixture with a pre-existing commentsExtended content-type defect either repairs that defect through normal packaging or preserves it as a baseline issue without attributing it to the restoration; any new package defect fails.
22. A synthetic deleted restrictions paragraph restores at its original structural location with adjusted `(i)`–`(vi)` text, one new sibling paragraph, distinct paragraph identity, content/paragraph-mark revisions, exact six-way lifecycle behavior, and no numbering drift.
23. Tests assert that restore failure never invokes or reports a fallback insertion after another paragraph. A success narrative is supported only by an applied result, committed receipt, `written: true`, and a real output path.
24. Exact reported single- and double-newline paragraph expansion fixtures reproduce `DUPLICATE_REVISION_ID` from an anchor containing `w:pPrChange` and/or `w:rPrChange`, then prove all generated IDs are unique.
25. Property-clone matrix: clean `pPr`/`rPr`, `pPrChange` only, `rPrChange` only, both, nested formatting changes, and anchor content already inside a same-author or foreign `w:ins`.
26. Insert one, two, and three adjacent plain paragraphs; assert unique content and paragraph-mark revision IDs, fresh paragraph IDs, committed receipt reconciliation, and exact paragraph order.
27. Repeat the expansion through structured Markdown heading, plain adjacency, list adjacency, explicit range, low-level runner, package facade, and CLI routes that share paragraph builders.
28. A deliberately injected cloned revision ID is caught before commit: per-operation status is `error`, receipt is uncommitted, atomic output is byte-exact, and no destination is written.
29. Package validation remains a backstop and reports the same offending ID/kinds if the operation-level guard is deliberately bypassed in a test harness.
30. Shape oracle catches a syntactically valid operation that replaces anchor text instead of inserting siblings, even when accepted text contains all requested words.
31. Shape oracle catches the `8.4 Changes; No Waiver` heading/body being appended inline to an earlier subsection rather than restored as two standalone paragraphs.
32. Exact 8.4 range restoration: heading remains `8.4 Changes; No Waiver`; body retains its original substance plus an email-sufficiency adjustment; both follow their foreign-deleted sources per WP09e; paragraph identities and six lifecycle outcomes are exact.
33. Exact 3.1/3.2 neighborhood regression: restoring 3.2 cannot replace, delete, concatenate with, or otherwise mutate the surviving `3.1 Restricted Use` paragraph.
34. Minimized Word Oracle A: split deletion IDs/authors/dates, sibling insertion placement, no nested revision, no inserted paragraph mark, exact `pa|REVIEWER B INSERTION|y` rejected-author view.
35. Oracle A at deletion start, end, run boundary, formatted-run boundary, and across multiple runs; repeated anchor text must require an occurrence/offset discriminator.
36. Oracle A preserves `w:lastRenderedPageBreak`, bold/underline runs, NBSPs, tabs, hyperlinks, and field boundaries or fails closed with a specific unsupported-boundary code.
37. Oracle A six-way lifecycle matrix is compared with a Word Desktop accepted/rejected oracle; Reject B reconstructs the original split deletion semantically and Reject A restores original text with B at the exact offset.
38. Generic non-explicit insertion into an empty foreign-deleted paragraph remains refused, proving WP09e does not weaken WP08a's ambiguity guard.
39. `validateRedlineOoxml` accepts the minimized Word-authored Oracle A structure without `FOREIGN_PARAGRAPH_MARK_DELETION` as an error; any informational warning must identify it as dependent inline content, not corruption.
40. Minimized Word Oracle B: two inserted paragraphs appear after the untouched deleted 8.2 heading/body and before deleted 8.3, with fresh paragraph IDs and distinct paragraph/content revision IDs.
41. Oracle B heading preserves effective bold formatting without requiring Word's incidental `rPrChange` history; if property changes are emitted, synthetic IDs 720–726 are modeled as seven distinct revision identities.
42. Combined Oracle A + B fixture checks forward-merge destination and exact section paragraph vectors under all selective lifecycle resolutions, including whether accepted inline text joins the next surviving paragraph exactly as Word does.
43. Restoration idempotency and changed reapplication operate on the post-source block; they neither prepend a second block nor mistake the original deleted paragraphs for the restoration.
44. Existing pre-source v0.5.3 restoration output is detected deliberately: migrate/reapply only under explicit policy, otherwise return a structured placement-version diagnostic rather than duplicating it.

#### Planned Implementation Areas

* `core/paragraph-targeting.js`
  - Extend `resolveTargetParagraph` results with exact canonical source text and an explicit resolution/match mode without changing strict identity checks.
  - Keep `normalizeWhitespaceForTargeting` confined to candidate comparison; do not reuse its output as mutation text.
  - Add strict rejected-view deleted-text anchor/offset resolution for Word-native deletion-carrier insertion.
* `core/paragraph-revision-safety.js`
  - Separate ambiguous same-paragraph resurrection from explicitly anchored inline insertion inside a foreign deletion carrier.
* `services/document-operation-mutations.js`
  - Update `applyToParagraphByExactText` to pass exact resolved source text and space-equivalence diagnostics into the engine and receipt path.
  - Update `verifyParagraphRestorationLifecycle` and `restoreDeletedParagraphByExactText` to use baseline-delta plus mutation-envelope validation while preserving the current/Accept-All/Reject-All oracle and rollback.
* `core/validation-delta.js` (NEW, or an equivalent shared validation helper)
  - Centralize stable issue signatures, multiset subtraction, issue summaries, and generated-versus-baseline classification so restoration, package, and CLI paths cannot drift.
* `core/revision-cloning.js`
  - Generalize revision-ID refresh/sanitization beyond `w:rPrChange`, or provide separate effective-property clone helpers that deliberately remove historical change elements.
* `core/redline-validation.js`
  - Expose duplicate-ID details sufficient to identify the repeated ID and owning element kinds for operation-level diagnostics.
* `pipeline/diff-engine.js`
  - Add or adjust deterministic replacement-hunk normalization for exact whitespace substitutions and repeated-token alignment.
* `engine/surgical-mode.js`
  - Add a pre-mutation diff replay assertion against exact source/modified text and rebuild live spans between mutation-dependent replacement hunks where required.
* `engine/surgical-diff-application.js`
  - Correct deletion/insertion anchors at run and hyperlink boundaries; make anchor consumption single-use and connected-node checked.
* `engine/surgical-run-splitting.js`
  - Extend carrier splitting to `w:del` with fresh trailing IDs and preserved foreign metadata/run pieces.
* `engine/oxml-engine.js`
  - Preserve the exact final accepted-view oracle and enrich whitespace mismatch metadata without leaking the whole payload.
* `services/receipt-collector.js` and operation result types
  - Report space-equivalent resolution and exact source-character diagnostics in a stable structured form if the data belongs in durable receipts.
  - Reconcile every revision-bearing element emitted by paragraph expansion and refuse duplicate/unreported cloned IDs before commit.
* `services/document-operation-contract.js`, `docs/schemas/document-operations.schema.json`, and operation declarations
  - Publish the explicit rejected-view insertion anchor/offset contract and reject ambiguous combinations.
* `node/cli.js`
  - Add a CLI projection that removes `documentXml` and other large internal payloads, summarizes baseline issues, and preserves decisive write/error fields.
* `node/docx-document.js`
  - Enforce baseline-versus-generated issue separation for both revision OOXML and package validation; never reject unchanged legacy defects as newly authored output.
* `index.d.ts`, `services/standalone-operation-runner.d.ts`, and `node/index.d.ts`
  - Publish any new match-mode, mismatch-code-point, validation-summary, or completion fields.
* `tests/cross_author_slicing_whitespace_alignment_tests.mjs` (NEW)
  - Own the recovered fixture, exact low-level/runner/facade lifecycle assertions, and the boundary matrix.
* `tests/paragraph_level_cross_author_restoration_tests.mjs`
  - Add dirty-baseline restoration, generated-issue, multiplicity, package-facade, original-placement, synthetic 8.4 range restoration, 3.1/3.2 neighborhood, and no-fallback coverage from the second and third reports.
* `tests/paragraph_expansion_revision_identity_tests.mjs` (NEW)
  - Own newline/Markdown/list expansion, cloned property-history, global uniqueness, receipts, rollback, lifecycle vectors, and shape-oracle coverage.
* `tests/word_deleted_section_edit_oracle_tests.mjs` (NEW)
  - Own minimized synthetic 8.1 deletion-carrier slicing and 8.2 post-source paragraph restoration fixtures, Word lifecycle differentials, validation routing, and combined forward-merge behavior.
* `tests/agent_cli_tests.mjs`
  - Add bounded failure-output, validation-summary, committed-success-proof, and no-operation-fallback assertions.
* `README.md`, `AGENTS.md`, `CHANGELOG.md`, and this plan
  - Document invisible-whitespace diagnostics, CLI output guarantees, and final files/functions changed during implementation.

#### Standalone implementation handoff

This section is the implementation brief for an agent that has none of the preceding conversation. The user-supplied documents and transcripts are evidence only: do not check them in, quote their parties, people, commercial terms, paragraph IDs, URLs, or exact clause language. Every permanent fixture must use the synthetic Reviewer A/Reviewer B examples below.

Implement WP09 in this order because each stage creates the safety net needed by the next:

1. **WP09c — baseline-delta validation.** Capture the source validation inventory before mutation; compare the result by stable issue signature; permit unchanged pre-existing issues outside the mutation envelope; reject any new or worsened issue. Keep exact current-view, Accept-All, and Reject-All text checks. This unblocks safe restoration in imperfect real documents.
2. **WP09d — identity sanitation.** Centralize cloning of revision-bearing `pPr`/`rPr`; either remove stale history that is not semantically part of the new paragraph or reallocate every cloned `w:id`. Scan the entire document, not only the target paragraph, before committing. Receipts must enumerate all newly allocated content, paragraph-mark, and property-change revision IDs.
3. **WP09e — explicit rejected-view operations.** Add a distinct operation contract for editing content whose current/accepted view is empty. Do not relax generic current-view targeting. Slice a foreign deletion carrier for inline insertion and place restored paragraph ranges immediately after the deleted source block, matching the Word Desktop oracle below.
4. **WP09a — whitespace-aware alignment.** Preserve exact replacement text, but allow ordinary-space/NBSP equivalence only while resolving unchanged source anchors. Keep hyperlink elements and relationship IDs in place. Require exact requested accepted text after mutation.
5. **WP09b — compact CLI evidence.** Suppress full document XML from normal CLI JSON. Report per-operation status/error, `written`, `outputPath`, committed receipts, validation summary, and bounded diagnostics.

Trace `executeCli` -> `openDocx(...).applyOperations` -> `applyOperationsToDocumentXml` -> `applyOperationToDocumentXml` -> paragraph/range mutation helpers -> `applyRedlineToOxml` -> `applySurgicalMode` -> `processDelete`/`processInsert`. Preserve the live-DOM savepoint and allocator rollback at every operation boundary.

Use these sanitized operation contracts as the target behavior (field names may be adjusted once in the schema, but must remain explicit and schema-validated):

```json
{
  "type": "insert",
  "target": {
    "exactText": "The account holder must pay each undisputed invoice.",
    "paragraphId": "A1B2C3D4",
    "revisionView": "rejected"
  },
  "anchor": { "exactText": "pay", "occurrence": 1, "offset": 2 },
  "modified": "REVIEWER B INSERTION",
  "author": "Reviewer B",
  "existingRevisions": "slice-cross-author"
}
```

This must convert Reviewer A's single deletion carrier into sibling `del(A prefix)`, `ins(B text)`, `del(A suffix)` nodes. The trailing Reviewer A carrier gets a fresh revision ID; the insertion must not inherit a deleted paragraph mark and must not be nested inside `w:del`. Rejecting Reviewer B restores Reviewer A's original deletion-carrier text; rejecting Reviewer A exposes the original sentence plus Reviewer B's pending insertion; Accept-All retains Reviewer B's inserted text at the deletion boundary with Word-compatible paragraph merging.

```json
{
  "type": "restore",
  "target": {
    "exactText": "8.2 Usage Adjustments",
    "paragraphId": "B1C2D3E4",
    "revisionView": "rejected"
  },
  "targetEnd": {
    "exactText": "Usage above the stated threshold may be billed at the next tier.",
    "paragraphId": "B1C2D3E5",
    "revisionView": "rejected"
  },
  "modified": "8.2 Usage Adjustments\n\nUsage above the stated threshold may be billed at the next tier.",
  "author": "Reviewer B"
}
```

The restored heading and body must be newly inserted sibling paragraphs immediately **after** the two-paragraph Reviewer A deletion block and before the following source paragraph. Each inserted paragraph needs a fresh paragraph identity, a Reviewer B paragraph-mark insertion, a Reviewer B content insertion, sanitized properties, and globally unique IDs. The original deleted block remains unchanged outside allocator-neutral serialization. A repeated request is an idempotent no-op or a specific duplicate-restoration error; it must never add a second copy.

Required red/green fixtures are: (a) two hyperlinks separated by NBSP/ordinary spaces with a small insertion; (b) a source document containing an unrelated missing-`xml:space` warning plus a safe restoration that adds no issue; (c) newline expansion from a paragraph whose `pPr` and first run contain prior property-change IDs; (d) the inline deletion-carrier oracle above; and (e) the two-paragraph post-source restoration oracle above. For each, assert structural order, authorship, global ID uniqueness, exact accepted and rejected paragraph vectors, selective Accept/Reject outcomes, receipt reconciliation, atomic rollback, and package validation. Expected failures must use a specific code such as `AMBIGUOUS_TARGET`, `UNSAFE_REVISION_NESTING`, `UNSAFE_PARAGRAPH_BOUNDARY`, or a new rejected-view anchor error—not a generic exception.

The checked-in characterization suite `tests/word_deleted_section_edit_oracle_tests.mjs` records the Word-authored shapes without private source text. It is an oracle scaffold, not proof that WP09e mutation support already exists; implementation tests must additionally create those shapes through the public runner and package facade.

#### Exit Criteria

WP09 is complete only when all reported workflows and both sanitized Word Desktop oracles succeed through the real package facade: the Service Policy edit must require no manual NBSP discovery, produce exact requested accepted text, preserve both hyperlinks, and reject back to the exact source; explicit paragraph restorations must tolerate unrelated baseline defects, introduce zero new validation issues, follow Word's post-source placement, and satisfy all selective lifecycle outcomes; rejected-view insertion inside a foreign deletion must split the carrier exactly like Word; and multiline paragraph expansion must allocate globally unique revision IDs and satisfy exact paragraph-shape oracles. The synthetic 8.4 pair must exist as standalone paragraphs in its original location, while the neighboring 3.1 paragraph remains unchanged. The CLI must provide a compact success or failure record that an agent can verify from `status`, every `results[i].status`, committed receipts, `written`, and `outputPath` without receiving full document XML.

Final implementation notes must replace the planned file/function list above with the actual touched files and functions, record any deliberately deferred cases, and include focused tests, full serial `npm test`, lint, type declarations, build, schema parsing if changed, and `git diff --check`.

---

## 6. Comprehensive Verification Plan (Synthetic & Real Test Series)

To prove correctness across all layers of the stack, this plan defines two comprehensive test suites:
1. **Synthetic Unit & Boundary Suites** (`tests/cross_author_slicing_synthetic_tests.mjs`): Isolated OOXML fixtures testing edge cases, boundary alignments, and lifecycle mechanics.
2. **Checked-In Word Package Differential Suite** (`tests/cross_author_slicing_real_tests.mjs`): End-to-end strict-facade replay against actual DOCX packages created by Microsoft Word Desktop, including package validation and lifecycle comparison with Word-generated oracles.

The synthetic matrix below is fully automated. The checked-in package suite is also fully automated as PKG-01 through PKG-06. REAL-01 through REAL-05 remain an environment/input-dependent acceptance matrix using sanitized documents and live Word COM/visual checks.

---

### 6.1 Synthetic Test Series (`tests/cross_author_slicing_synthetic_tests.mjs`)

| ID | Test Case Name | Input Structure | Operation (Author B) | Expected OOXML Structure | Invariant Assertions |
|:---|:---|:---|:---|:---|:---|
| **SYN-01** | Pure Interior Insertion | `<w:ins author="Reviewer A">amended by this agreement</w:ins>` | Reviewer B inserts `"MASTER "` before `"agreement"` | `[ins(A): "amended by this "][ins(B): "MASTER "][ins(A): "agreement"]` | 3 sibling `<w:ins>` nodes; no `NESTED_REVISION`; unique IDs for Reviewer B and the trailing Reviewer A carrier. |
| **SYN-02** | Pure Interior Deletion | `<w:ins author="Reviewer A">The service processes input to produce output</w:ins>` | Reviewer B deletes `"produce "` | `<w:ins author="Reviewer A">...<w:del author="Reviewer B">produce </w:del>...</w:ins>` | One Reviewer A carrier remains; nested `w:del` uses `<w:delText>` and is authored by Reviewer B. |
| **SYN-03** | Boundary Deletion at Insertion Start | `<w:ins author="Reviewer A">Subject to the exception, the policy remains</w:ins>` | Delete `"Subject to the exception, "` | `<w:ins author="Reviewer A"><w:del author="Reviewer B">Subject...</w:del>the policy remains</w:ins>` | Nested deletion is the carrier's first content node; Reviewer A metadata remains intact. |
| **SYN-04** | Boundary Deletion at Insertion End | `<w:ins author="Reviewer A">subject to Section 2 and applicable rules</w:ins>` | Delete `" and applicable rules"` | `<w:ins author="Reviewer A">subject...<w:del author="Reviewer B"> and applicable rules</w:del></w:ins>` | Nested deletion is the carrier's final content node. |
| **SYN-05** | Complete Deletion of Pending Insertion Text | `<w:ins author="Reviewer A">Obsolete inserted clause.</w:ins>` | Delete the entire inserted string | `<w:ins author="Reviewer A"><w:del author="Reviewer B">Obsolete inserted clause.</w:del></w:ins>` | Reviewer A's carrier remains so rejecting A still cascades away B's dependent deletion. |
| **SYN-06** | Straddle Deletion (Baseline to Insertion) | `<w:r><w:t>Baseline start </w:t></w:r><w:ins author="Reviewer A">inserted finish</w:ins>` | Delete `"start inserted"` | `[r: "Baseline "][del(B): "start "][ins(A): [del(B): "inserted"] " finish"]` | Top-level and nested deletion portions remain structurally separate. |
| **SYN-07** | Straddle Deletion (Insertion to Baseline) | `<w:ins author="Reviewer A">Inserted start</w:ins><w:r><w:t> baseline finish</w:t></w:r>` | Delete `"start baseline"` | `[ins(A): "Inserted " [del(B): "start"]][del(B): " baseline"][r: " finish"]` | Nested and top-level deletion portions preserve their respective carrier contexts. |
| **SYN-08** | Multi-Insertion Straddle | `<w:ins author="Reviewer A">Alpha text </w:ins><w:ins author="Reviewer C">Gamma text</w:ins>` | Reviewer B deletes `"text Gamma"` | `[ins(A): "Alpha " [del(B): "text "]][ins(C): [del(B): "Gamma"] " text"]` | Each foreign carrier owns its nested deletion portion. |
| **SYN-09** | Multi-Run Formatting Preservation | `<w:ins author="Reviewer A"><w:r><w:rPr><w:b/></w:rPr><w:t>Bold text </w:t></w:r><w:r><w:t>plain text</w:t></w:r></w:ins>` | Delete `"text plain"` | Reviewer A's `<w:ins>` remains intact around Reviewer B's nested deletion with bold and plain runs. | Exact run-level formatting is preserved inside `<w:delText>` and unaffected insertion runs. |
| **SYN-10** | Paired Replacement Event inside Insertion | `<w:ins author="Reviewer A">process input to produce output</w:ins>` | Replace `"produce"` with `"create"` | `[ins(A): prefix + nested del(B)][ins(B): "create"][ins(A): suffix]` | Deletion and insertion share a timestamp; only the insertion requires carrier splitting/hoisting. |
| **SYN-11** | 3-Author Stacked Deletions | Output of **SYN-02** | Reviewer C deletes `"processes"` in Reviewer A's carrier | `<w:ins author="Reviewer A">...<w:del author="Reviewer C">processes</w:del>...<w:del author="Reviewer B">produce</w:del>...</w:ins>` | Multiple reviewer deletions coexist safely inside the same foreign insertion. |
| **SYN-12a** | Lifecycle Oracle: Accept All | Output of **SYN-02** | `acceptTrackedChanges({ allAuthors: true })` | Clean baseline string: `"The service processes input to output"` | All `<w:del>` removed, all `<w:ins>` unwrapped; zero revision tags remaining. |
| **SYN-12b** | Lifecycle Oracle: Accept Reviewer A Only | Output of **SYN-02** | `acceptTrackedChanges({ author: 'Reviewer A' })` | Reviewer A's text becomes baseline; Reviewer B's `<w:del>` remains pending. | Reviewer B's deletion remains intact and reviewable. |
| **SYN-12c** | Lifecycle Oracle: Reject Reviewer A Only | Output of **SYN-02** | `rejectTrackedChanges({ author: 'Reviewer A' })` | Reviewer A's insertion is removed with Reviewer B's dependent deletion. | Prevents an orphaned deletion of text that never became baseline. |
| **SYN-12d** | Lifecycle Oracle: Reject Reviewer B Only | Output of **SYN-02** | `rejectTrackedChanges({ author: 'Reviewer B' })` | Reviewer B's nested deletion is unwrapped inside Reviewer A's carrier. | Full restoration of Reviewer A's original insertion. |

---

### 6.2 External Real-World Acceptance Series

| ID | Scenario & Source Document | Workflow & Operations | Expected Real-World Behavior | Verification Oracle |
|:---|:---|:---|:---|:---|
| **REAL-01** | **Synthetic Service Agreement: Pending Feature Clause**<br>Source: sanitized Word-authored package | 1. Reviewer A has a pending insertion.<br>2. Reviewer B deletes one interior verb from it.<br>3. Apply with `--existing-revisions slice-cross-author`. | Batch commits with `status: "ok"`, `written: true`.<br>Reviewer A's insertion remains intact and contains Reviewer B's visible, attributed nested deletion. | Output validates via `validateDocxPackage`; `inspect` reports both generic reviewers. |
| **REAL-02** | **Synthetic Master Agreement: Policy Carve-Out**<br>Source: sanitized Word-authored package | 1. Reviewer A has a pending amendment sentence.<br>2. Reviewer B inserts a policy carve-out in its middle.<br>3. Apply with `--existing-revisions slice-cross-author`. | Reviewer A's insertion remains attributed to Reviewer A rather than becoming baseline; Reviewer B's carve-out is an adjacent/spliced insertion. | `extract` and `inspect` report Reviewer A and Reviewer B without private source metadata. |
| **REAL-03** | **Synthetic Three-Reviewer Negotiation**<br>Source: sanitized multi-round corpus package | Round 1: Reviewer A inserts text in two clauses.<br>Round 2: Reviewer B edits inside those insertions.<br>Round 3: Reviewer C applies further modifications. | Three authors' overlapping and sliced edits commit without merge corruption or lost attribution. | Document hash checks; zero new schema errors across all rounds. |
| **REAL-04** | **Desktop Word COM Automation Oracle**<br>Execution on a Windows runner with Word | Open the sanitized outputs from **REAL-01**, **REAL-02**, and **REAL-03** through the native Word automation client. | 1. Word opens each file with no repair prompt.<br>2. `Document.Revisions.Count` matches committed receipt counts.<br>3. Word Accept-All and Reject-All text matches the engine lifecycle results. | Automation asserts identical paragraph vectors after native and engine lifecycle operations. |
| **REAL-05** | **Word Visual Rendering and PDF Proof**<br>Sanitized evidence pipeline | Export **REAL-01** and **REAL-02** to PDF through Word, then render the affected pages. | Reviewer B's deletions and Reviewer A's insertions have distinct reviewer colors; the Reviewing Pane attributes both correctly without overlapping or misplaced balloons. | Sanitized visual artifacts are generated outside the repository unless explicitly approved for check-in. |

---

### 6.3 Test Execution Matrix

```bash
# 1. Run the sanitized Word deleted-section characterization oracle
node tests/word_deleted_section_edit_oracle_tests.mjs

# 2. Run synthetic unit & boundary suite
node tests/cross_author_slicing_synthetic_tests.mjs

# 3. Run checked-in Word DOCX package differential suite (PKG-01..06)
node tests/cross_author_slicing_real_tests.mjs

# 4. Optional external acceptance: Word Desktop COM differential oracle (Windows desktop)
npm run test:word

# 5. Verify existing mode matrix remains 100% backward compatible
node tests/existing_revisions_modes_matrix_tests.mjs

# 6. Full regression check
npm test
```
