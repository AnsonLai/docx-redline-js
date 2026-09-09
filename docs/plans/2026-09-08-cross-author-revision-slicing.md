# Cross-Author Revision Slicing: Nested Revisions via Carrier Splitting

**Status:** Completed — WP-01 through WP-07 implemented; final automated verification complete  
**Date:** 2026-09-08  
**Target releases:** v0.6.0–v1.0.0  
**Priority:** Document fidelity, multi-author contract negotiation accuracy, and strict OOXML schema compliance over single-pass simplicity.

---

## 1. Executive Summary & Findings

### The Problem
During contract review and multi-turn legal negotiations, a reviewer (Author B, e.g., "Lai, Anson") often needs to edit text that was previously inserted by another reviewer (Author A, e.g., "Lai, Barry") whose revision has not yet been accepted.

In previous discussions and documentation, it was assumed that deleting or inserting text inside another author's pending insertion was either impossible in WordprocessingML (OOXML) or logically paradoxical. However, empirical inspection of Microsoft Word Desktop proves otherwise:

1. **Word Desktop supports deletions inside pending insertions**: When Author B deletes text inside Author A's pending `<w:ins>`, Word Desktop displays Author B's deletion visibly (strikethrough formatting, deletion tooltip attributed to Author B with timestamp: `Lai, Anson deleted: <text>`), while the surrounding insertion text remains attributed to Author A.
2. **Word Desktop supports insertions inside pending insertions**: When Author B types text in the middle of Author A's `<w:ins>`, Word Desktop displays Author B's new text attributed to Author B, while preserving Author A's attribution on the preceding and succeeding words.

### The Underlying OOXML Mechanics
Empirical inspection of Microsoft Word Desktop 365 (via automated COM fixture generation in WP-01) reveals a vital asymmetry between how Word handles cross-author insertions versus deletions:

1. **Cross-Author Insertions (`insert-interior`) — Sibling Splitting**:
   ECMA-376 Part 1 `CT_RunTrackChange` does NOT allow `<w:ins>` inside `<w:ins>`. Word Desktop splits the carrier `<w:ins>` into sibling fragments at the paragraph (`<w:p>`) level, splicing Author B's new `<w:ins>` between them:
   ```xml
   <!-- Sibling 1: Author A's insertion (leading fragment) -->
   <w:ins w:id="0" w:author="Barry Lai" w:date="2026-09-08T09:29:00Z">
       <w:r><w:t xml:space="preserve">amended by this </w:t></w:r>
   </w:ins>
   <!-- Sibling 2: Author B's insertion spliced in between -->
   <w:ins w:id="1" w:author="Anson Lai" w:date="2026-09-08T09:29:00Z">
       <w:r><w:t xml:space="preserve">MASTER </w:t></w:r>
   </w:ins>
   <!-- Sibling 3: Author A's insertion (trailing fragment) -->
   <w:ins w:id="2" w:author="Barry Lai" w:date="2026-09-08T09:29:00Z">
       <w:r><w:t>Agreement.</w:t></w:r>
   </w:ins>
   ```

2. **Cross-Author Deletions (`delete-interior`) — Direct `<w:del>` Nesting inside `<w:ins>`**:
   Under ECMA-376 Part 1 Section 17.13.5.21 (`CT_RunTrackChange`), `<w:del>` is an explicitly permitted child element of `<w:ins>`. Microsoft Word Desktop does **not** split `<w:ins>` for deletions; instead, it nests `<w:del>` directly inside `<w:ins>`:
   ```xml
   <w:ins w:id="0" w:author="Barry Lai" w:date="2026-09-08T09:29:00Z">
       <w:r><w:t xml:space="preserve">The Services will process the Input to </w:t></w:r>
       <w:del w:id="1" w:author="Anson Lai" w:date="2026-09-08T09:29:00Z">
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
  - Added strict package-facade differential tests for all five core Word Desktop scenarios. Each test reconstructs the pre-Anson package, applies `slice-cross-author` through `openDocx(...).applyOperations`, requires atomic package validation, and compares engine Accept-All/Reject-All text with the checked-in Word Desktop accepted/rejected DOCX files.
  - Added a two-round package test that generates the three-author stacked-deletion fixture through Anson and Chris operations, asserts all three reviewers survive inspection, and compares both lifecycle endpoints with Word Desktop.
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
  - The repository does not contain the private `agreement.docx` referenced by REAL-01/REAL-02 or the exact `c5bb43ede5...` corpus package referenced by REAL-03. Those named cases remain external acceptance scenarios rather than silently skipped automated tests.
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
  - `npm test` — PASS, 94 test files passed and 0 failed.
  - `npm run lint` — PASS.
  - `npm run check:types` — PASS; all 123 runtime exports have declarations.
  - `git diff --check` — PASS (line-ending conversion notices only; no whitespace errors).

---

## 6. Comprehensive Verification Plan (Synthetic & Real Test Series)

To prove correctness across all layers of the stack, this plan defines two comprehensive test suites:
1. **Synthetic Unit & Boundary Suites** (`tests/cross_author_slicing_synthetic_tests.mjs`): Isolated OOXML fixtures testing edge cases, boundary alignments, and lifecycle mechanics.
2. **Checked-In Word Package Differential Suite** (`tests/cross_author_slicing_real_tests.mjs`): End-to-end strict-facade replay against actual DOCX packages created by Microsoft Word Desktop, including package validation and lifecycle comparison with Word-generated oracles.

The synthetic matrix below is fully automated. The checked-in package suite is also fully automated as PKG-01 through PKG-06. REAL-01 through REAL-05 remain an environment/input-dependent acceptance matrix for the named private/corpus documents and live Word COM/visual checks.

---

### 6.1 Synthetic Test Series (`tests/cross_author_slicing_synthetic_tests.mjs`)

| ID | Test Case Name | Input Structure | Operation (Author B) | Expected OOXML Structure | Invariant Assertions |
|:---|:---|:---|:---|:---|:---|
| **SYN-01** | Pure Interior Insertion | `<w:ins author="Barry">amended by this Agreement</w:ins>` | Insert `"MASTER "` before `"Agreement"` | `[ins(Barry): "amended by this "][ins(Anson): "MASTER "][ins(Barry): "Agreement"]` | 3 sibling `<w:ins>` nodes; no `NESTED_REVISION`; unique IDs allocated for `ins(Anson)` and trailing `ins(Barry)`. |
| **SYN-02** | Pure Interior Deletion | `<w:ins author="Barry">The Services will process the Input to generate outputs</w:ins>` | Delete `"generate "` | `<w:ins author="Barry">...<w:del author="Anson">generate </w:del>...</w:ins>` | One Barry carrier remains; nested `w:del` uses `<w:delText>` and is authored by Anson. |
| **SYN-03** | Boundary Deletion at Insertion Start | `<w:ins author="Barry">Notwithstanding the foregoing, the NDA remains</w:ins>` | Delete `"Notwithstanding the foregoing, "` | `<w:ins author="Barry"><w:del author="Anson">Notwithstanding...</w:del>the NDA remains</w:ins>` | Nested deletion is the carrier's first content node; Barry metadata remains intact. |
| **SYN-04** | Boundary Deletion at Insertion End | `<w:ins author="Barry">subject to Section 2.8 and applicable law</w:ins>` | Delete `" and applicable law"` | `<w:ins author="Barry">subject...<w:del author="Anson"> and applicable law</w:del></w:ins>` | Nested deletion is the carrier's final content node. |
| **SYN-05** | Complete Deletion of Pending Insertion Text | `<w:ins author="Barry">Obsolete clause insertion.</w:ins>` | Delete entire string `"Obsolete clause insertion."` | `<w:ins author="Barry"><w:del author="Anson">Obsolete clause insertion.</w:del></w:ins>` | Barry's carrier remains so rejecting Barry still cascades away Anson's dependent deletion. |
| **SYN-06** | Straddle Deletion (Baseline to Insertion) | `<w:r><w:t>Baseline start </w:t></w:r><w:ins author="Barry">inserted finish</w:ins>` | Delete `"start inserted"` | `[r: "Baseline "][del(Anson): "start "][ins(Barry): [del(Anson): "inserted"] " finish"]` | Top-level and nested deletion portions remain structurally separate, matching Word Desktop. |
| **SYN-07** | Straddle Deletion (Insertion to Baseline) | `<w:ins author="Barry">Inserted start</w:ins><w:r><w:t> baseline finish</w:t></w:r>` | Delete `"start baseline"` | `[ins(Barry): "Inserted " [del(Anson): "start"]][del(Anson): " baseline"][r: " finish"]` | Nested and top-level deletion portions preserve their respective carrier contexts. |
| **SYN-08** | Multi-Insertion Straddle (Author A to Author C) | `<w:ins author="Barry">Barry text </w:ins><w:ins author="Carl">Carl text</w:ins>` | Author B deletes `"text Carl"` | `[ins(Barry): "Barry " [del(Anson): "text "]][ins(Carl): [del(Anson): "Carl"] " text"]` | Each foreign carrier owns its nested deletion portion; neither carrier is sliced for deletion. |
| **SYN-09** | Multi-Run Formatting Preservation | `<w:ins author="Barry"><w:r><w:rPr><w:b/></w:rPr><w:t>Bold text </w:t></w:r><w:r><w:t>plain text</w:t></w:r></w:ins>` | Delete `"text plain"` | Barry's `<w:ins>` remains intact around a nested `<w:del>` containing a bold run for `"text "` and a plain run for `"plain"`. | Exact run-level formatting is preserved inside `<w:delText>` and unaffected insertion runs. |
| **SYN-10** | Paired Replacement Event inside Insertion | `<w:ins author="Barry">process the Input to generate outputs</w:ins>` | Replace `"generate"` with `"synthesize"` (`pairReplacements: true`) | `[ins(Barry): prefix + nested del(Anson)][ins(Anson): "synthesize"][ins(Barry): suffix]` | Deletion and insertion share timestamp; only the insertion requires carrier splitting/hoisting. |
| **SYN-11** | 3-Author Stacked Deletions | Output of **SYN-02** | Author C ("Davis, Chris") deletes `"process"` in Barry's carrier | `<w:ins author="Barry">...<w:del author="Davis">process</w:del>...<w:del author="Anson">generate</w:del>...</w:ins>` | Multiple distinct reviewer deletions coexist safely inside the same foreign insertion. |
| **SYN-12a** | Lifecycle Oracle: Accept All | Output of **SYN-02** | `acceptTrackedChanges({ allAuthors: true })` | Clean baseline string: `"The Services will process the Input to outputs"` | All `<w:del>` removed, all `<w:ins>` unwrapped; zero revision tags remaining. |
| **SYN-12b** | Lifecycle Oracle: Accept Author A Only | Output of **SYN-02** | `acceptTrackedChanges({ author: 'Barry' })` | Barry's text becomes baseline; Anson's `<w:del>` remains pending against the baseline. | Anson's `<w:del>` remains intact and reviewable. |
| **SYN-12c** | Lifecycle Oracle: Reject Author A Only | Output of **SYN-02** | `rejectTrackedChanges({ author: 'Barry' })` | Barry's insertion is deleted from the document. Anson's internal `<w:del>` is cascaded and pruned. | Prevents orphaned deletion of text that was rejected from ever existing. |
| **SYN-12d** | Lifecycle Oracle: Reject Author B Only | Output of **SYN-02** | `rejectTrackedChanges({ author: 'Anson' })` | Anson's nested `<w:del>` is unwrapped back into regular runs within Barry's `<w:ins>`. | Full restoration of Barry's original insertion. |

---

### 6.2 External Real-World Acceptance Series

| ID | Scenario & Source Document | Workflow & Operations | Expected Real-World Behavior | Verification Oracle |
|:---|:---|:---|:---|:---|
| **REAL-01** | **Salary.com Agreement: The Motivating AI Terms Deletion**<br>Source: `agreement.docx` (Section 2.1 Customer Data, `P40`) | 1. Document contains Barry Lai's pending insertion (`P40`, ID `45`).<br>2. Anson Lai runs operation to delete `"generate"` from `"to generate outputs"`.<br>3. `--existing-revisions slice-cross-author`. | Batch commits with `status: "ok"`, `written: true`.<br>Barry's insertion remains intact and contains Anson's visible, attributed nested deletion.<br>No `COMMENTED_CONTENT_DELETE` (since comment 144 is on Section 2.8, not 2.1). | Output file validated via `validateDocxPackage`. Revisions inspectable via `docx-redline inspect`. |
| **REAL-02** | **Salary.com Agreement: §14.1 NDA Carve-Out**<br>Source: `agreement.docx` (Section 14.1 Entire Agreement, `P131`) | 1. Barry has pending insertion of the amendment sentence.<br>2. Anson inserts August 25, 2026 NDA carve-out in the middle.<br>3. `--existing-revisions slice-cross-author`. | Barry's insertion remains visibly attributed to Barry (not baked into baseline).<br>Anson's carve-out sits as an adjacent/spliced insertion attributed to Anson. | `extract` and `inspect` show both `Lai, Barry` and `Lai, Anson` in `revisionAuthors`. |
| **REAL-03** | **SuperDoc Corpus: Interagency Multi-Counsel Negotiation**<br>Source: Corpus ID `c5bb43ede5...` (Joint Communications Protocol) | Round 1: BCHD Lead Agency Counsel applies insertions to Sections 3.2 and 4.1.<br>Round 2: MOHS Counterparty Counsel edits directly inside BCHD's insertions.<br>Round 3: Third Reviewer applies further modifications. | 3 distinct institutional authors with overlapping and sliced edits commit across rounds without merge corruption or loss of attribution. | Document hash checks; zero schema errors across all 3 rounds. |
| **REAL-04** | **Desktop Word 365 COM Automation Oracle**<br>Execution on Windows runner via native Word | Open the output DOCX files from **REAL-01**, **REAL-02**, and **REAL-03** via Windows COM automation (`word-client.mjs`). | 1. Word opens each file with **0 repair prompts** / corruption dialogs.<br>2. `Document.Revisions.Count` matches exact receipt counts.<br>3. `Document.Revisions.AcceptAll()` in Word matches engine `acceptAll()` bit-for-bit.<br>4. `Document.Revisions.RejectAll()` in Word matches engine `rejectAll()` bit-for-bit. | COM automation script asserts identical string contents after Word native Accept/Reject. |
| **REAL-05** | **Word Visual Rendering & PDF Export Proof**<br>Visual Evidence Pipeline | Export pages of **REAL-01** and **REAL-02** to PDF via Word COM `ExportAsFixedFormat`. Convert PDF pages to PNG. | 1. Deletions show strikethrough in Anson's reviewer color.<br>2. Insertions show underline in Barry's reviewer color.<br>3. Word Reviewing Pane displays balloons for both authors correctly without overlap or misaligned leader lines. | Visual evidence artifact generated in `tests/visual-evidence/` for human review sign-off. |

---

### 6.3 Test Execution Matrix

```bash
# 1. Run synthetic unit & boundary suite
node tests/cross_author_slicing_synthetic_tests.mjs

# 2. Run checked-in Word DOCX package differential suite (PKG-01..06)
node tests/cross_author_slicing_real_tests.mjs

# 3. Optional external acceptance: Word Desktop COM differential oracle (Windows desktop)
npm run test:word

# 4. Verify existing mode matrix remains 100% backward compatible
node tests/existing_revisions_modes_matrix_tests.mjs

# 5. Full regression check
npm test
```
