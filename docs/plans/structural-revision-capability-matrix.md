# Structural Revision Capability Matrix

**Date:** 2026-09-05  
**Baseline Version:** v0.4.0  
**Companion Plan:** [2026-09-05-structural-revisions-and-fidelity-oracles.md](./2026-09-05-structural-revisions-and-fidelity-oracles.md)

---

## 1. Executive Summary

This document establishes the required capability and gap baseline for `@ansonlai/docx-redline-js` prior to implementing new mutation semantics in the v0.5.0–v1.0.0 series.

Each row corresponds to an architectural behavior cataloged in Section 2 of the companion plan. Per Section 18 (WP-00), a capability is marked:
- `implemented`: generation, validation, Accept All, and Reject All are all verified by passing regression tests.
- `partial`: implemented only for specific sub-cases (e.g. lists, whole paragraphs) or missing specific lifecycle guarantees/metadata.
- `missing`: no production implementation exists.
- `intentionally unsupported`: explicitly refused by design with structured diagnostics.

---

## 2. Capability & Gap Matrix

| Capability Category | Behavior / Sub-Case | Status | Production Symbol(s) | Existing Regression Test(s) | Missing Case / Gap | Work Package |
|---|---|---|---|---|---|---|
| **Text Mutations** | Text insertion (`w:ins`) | `implemented` | `processInsert` in `engine/surgical-diff-application.js`, `createTrackChange` in `engine/run-builders.js` | `tests/diff_engine_reliability_tests.mjs`, `tests/standalone_smoke.mjs` | None for basic insertion | Complete |
| **Text Mutations** | Text deletion (`w:del`) | `implemented` | `processDelete` in `engine/surgical-diff-application.js`, `createTrackChange` in `engine/run-builders.js` | `tests/diff_engine_reliability_tests.mjs`, `tests/standalone_smoke.mjs` | None for basic deletion | Complete |
| **Text Mutations** | Paired text replacement (`w:del` + `w:ins`) | `partial` | `processDelete` + `processInsert` in `engine/surgical-diff-application.js` | `tests/standalone_smoke.mjs`, `tests/redline_validation_tests.mjs` | Replacement events currently allocate revision IDs and timestamps independently without paired metadata synchronization required for Word unified review balloons; structural boundary preservation rules not formal. | WP-07 |
| **Text Mutations** | Insertion container boundary affinity | `partial` | `processInsert` in `engine/surgical-diff-application.js`, `surgical-spans.js` | `tests/structural_tab_field_tests.mjs` | No explicit `InsertionAffinity` contract; run splitting at hyperlink/bookmark/comment boundaries relies on implicit order without `inside` vs `outside` control. | WP-08 |
| **Paragraph Boundaries** | Paragraph-mark insertion / split | `partial` | `markParagraphMarkInserted` in `engine/run-builders.js`, `applyReconstructionMode` in `engine/reconstruction-mode.js` | `tests/paragraph_mark_revision_tests.mjs` | Text splits spanning newlines fall back to destructive reconstruction mode, dropping paragraph properties (`w:spacing`, `w:ind`, `w:pStyle`) and bookmarks. | WP-09, WP-10 |
| **Paragraph Boundaries** | Paragraph boundary removal (join) | `partial` | `markParagraphMarkDeleted` in `engine/run-builders.js`, `mergeParagraphIntoNextAndRemove` in `services/revision-comment-management.js` | `tests/paragraph_mark_revision_tests.mjs` | Joins retaining both texts fall back to block reconstruction; native pilcrow deletion join algorithm not wired to surgical diff application. | WP-09, WP-10 |
| **Paragraph Boundaries** | Whole paragraph deletion | `implemented` | `markParagraphMarkDeleted`, `processDelete` in `engine/surgical-diff-application.js`, `services/document-operation-mutations.js` | `tests/paragraph_mark_revision_tests.mjs`, `tests/interagency_agreement_multi_author_tests.mjs` | None for whole paragraph deletion | Complete |
| **Property Revisions** | Run property revisions (`w:rPrChange`) | `partial` | `injectFormattingToRPr`, `snapshotAndAttachRPrChange` in `engine/run-builders.js`, `rejectPropertyChangeNode` in `services/revision-comment-management.js` | `tests/formatting_tests.mjs`, `tests/revision_comment_management_tests.mjs` | Gated behind Markdown formatting hints rather than explicit `CharacterFormatOperation`; author-aware coalescing (editing own pending insertion) missing. | WP-11 |
| **Property Revisions** | Paragraph property revisions (`w:pPrChange`) | `partial` | `rejectPropertyChangeNode`, `acceptTrackedChangesInOoxml` in `services/revision-comment-management.js` | `tests/revision_comment_management_tests.mjs` | Supported during Accept/Reject lifecycle, but no public operation exists to author a `w:pPrChange` in surgical mode. | WP-11 |
| **Structural Elements** | Hyperlinks (`w:hyperlink`) | `implemented` | `readCanonicalRunText` in `core/paragraph-text.js`, `surgical-spans.js` | `tests/structural_tab_field_tests.mjs` | Boundary affinity (`inside` vs `outside` link) when prepending/appending text. | WP-08 |
| **Structural Elements** | Bookmarks (`w:bookmarkStart`, `w:bookmarkEnd`) | `implemented` | `surgical-spans.js`, `core/word-xml.js` | `tests/structural_tab_field_tests.mjs`, `tests/docx_fixture_tests.mjs` | Range markers preserved during edits; boundary positioning needs explicit tests. | WP-08 |
| **Structural Elements** | Comments (`word/comments.xml` & markers) | `implemented` | `injectCommentsIntoOoxml` in `services/comment-engine.js`, `deleteCommentsByAuthorInOoxml` in `services/revision-comment-management.js` | `tests/comment_tests.mjs`, `tests/comment_anchor_locator_tests.mjs` | Targeting deleted text in `rejected` view for comment placement is currently unsupported. | WP-05, WP-06 |
| **Structural Elements** | Fields (`w:fldSimple`, `w:fldChar`) | `implemented` | `appendVisibleTextPieces` in `engine/run-builders.js`, `surgical-spans.js` | `tests/structural_tab_field_tests.mjs` | Splitting inside complex field instruction blocks is intentionally unsupported. | Intentionally Unsupported |
| **Structural Elements** | Structured Document Tags (`w:sdt`) | `implemented` | `core/word-xml.js`, `core/paragraph-targeting.js` | `tests/docx_fixture_tests.mjs` | SDTs survive as containers; container boundary affinity needs explicit tests. | WP-08 |
| **Structural Elements** | Tabs, breaks, hyphens (`w:tab`, `w:br`, `w:cr`, `w:noBreakHyphen`) | `implemented` | `appendVisibleTextPieces` in `engine/run-builders.js`, `cloneRunPiece` in `engine/surgical-run-splitting.js`, `readCanonicalRunText` in `core/paragraph-text.js` | `tests/structural_tab_field_tests.mjs`, `tests/canonical_paragraph_text_tests.mjs` | None | Complete |
| **Structural Elements** | Footnotes & Endnotes (`w:footnoteReference`) | `implemented` | `core/word-xml.js`, `surgical-spans.js` | `tests/docx_fixture_tests.mjs` | Footnote reference markers preserved during adjacent text edits. | Complete |
| **Lists & Tables** | List generation from markdown | `implemented` | `executeListGeneration` in `pipeline/list-generation.js`, `numbering-service.js` | `tests/list_tests.mjs`, `tests/list_replacement_structure_tests.mjs` | None | Complete |
| **Lists & Tables** | Markdown table reconciliation | `implemented` | `reconcileMarkdownTableOoxml`, `generateTableOoxml` in `services/table-reconciliation.js` | `tests/table_tests.mjs`, `tests/table_targeting_and_format_flags.mjs` | None | Complete |
| **Lists & Tables** | Structured content planning | `implemented` | `analyzeStructuredContent`, `planStructuredReplacement` in `pipeline/structured-content.js` | `tests/structured_content_planner_tests.mjs` | None | Complete |
| **Text Extraction** | Accepted text extraction | `implemented` | `extractCanonicalParagraphText`, `readCanonicalRunText` in `core/paragraph-text.js` | `tests/canonical_paragraph_text_tests.mjs` | None | Complete |
| **Text Extraction** | Rejected text extraction | `implemented` | `extractCanonicalParagraphText(p, { revisionView: 'rejected' })` in `core/paragraph-text.js` | `tests/canonical_paragraph_text_tests.mjs` | Detailed segment breakdown (`RevisionTextSegment[]`) with offsets and kinds missing. | WP-04 |
| **Target Resolution** | Accepted-view targeting | `implemented` | `resolveTargetParagraph`, `buildParagraphMetadataIndex` in `core/paragraph-targeting.js` | `tests/standalone_operation_runner_tests.mjs`, `tests/error_contract_tests.mjs` | None | Complete |
| **Target Resolution** | Rejected-view targeting | `missing` | None | None | `revisionView: 'rejected'` not supported in `core/paragraph-targeting.js` or `services/operation-preflight.js`. | WP-05 |
| **Target Resolution** | Fail-closed strict targeting | `implemented` (preflight & node facade) / `partial` (runner default) | `preflightOperations` in `services/operation-preflight.js`, `DocxDocument` in `node/docx-document.js` | `tests/agent_operation_contract_tests.mjs`, `tests/agent_cli_tests.mjs` | Lower-level runner currently defaults to permissive matching; strict mode must become default in v1.0.0. | WP-18 |
| **Session & Batch** | Live DOM session & savepoints | `implemented` | `DocumentOperationSession` in `services/document-operation-session.js` | `tests/performance_phase1_session_tests.mjs` | Capture table state not yet part of savepoint tracking. | WP-15 |
| **Session & Batch** | Stepwise capture chaining | `missing` | None | None | Batches cannot target entities produced by earlier steps within the same batch. | WP-14, WP-15 |
| **Session & Batch** | Optimistic concurrency tokens | `missing` | None | None | No `expectedRevision` precondition assertion on document/package state. | WP-12, WP-13 |
| **Session & Batch** | Structured mutation receipts | `missing` | None | None | Results return high-level boolean `hasChanges` but do not report allocated `w:id` revision or comment telemetry. | WP-16, WP-17 |
| **Package Plumbing** | Package artifact merging (comments, numbering) | `implemented` | `ensureCommentsArtifactsInZip`, `ensureNumberingArtifactsInZip` in `services/standalone-docx-plumbing.js` | `tests/docx_package_transaction_edge_tests.mjs`, `tests/docx_package_facade_tests.mjs` | None | Complete |
| **Package Plumbing** | Package entry byte-fidelity verification | `missing` | None | None | No automated oracle verifying bitwise invariance of non-targeted ZIP entries. | WP-01, WP-03 |
| **Package Plumbing** | Canonical subtree XML fidelity verification | `missing` | None | None | No automated oracle verifying canonical invariance of non-targeted `<w:p>` subtrees. | WP-02, WP-03 |

---

## 3. Canonical Serialized Output for Partial Areas

### 3.1 Paired Text Replacement (Current Output)
Currently, replacement operations emit independent `<w:del>` and `<w:ins>` elements with separate revision IDs and independent timestamps:

```xml
<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:r><w:t>The term shall be </w:t></w:r>
  <w:del w:id="1" w:author="Reviewer" w:date="2026-09-05T12:00:00Z">
    <w:r><w:delText xml:space="preserve">thirty (30) days</w:delText></w:r>
  </w:del>
  <w:ins w:id="2" w:author="Reviewer" w:date="2026-09-05T12:00:00Z">
    <w:r><w:t xml:space="preserve">sixty (60) days</w:t></w:r>
  </w:ins>
  <w:r><w:t> from receipt.</w:t></w:r>
</w:p>
```
*Gap:* WP-07 will guarantee unified event-level metadata synchronization, zero-length run removal, and structural container boundaries.

### 3.2 Paragraph Boundary Split / Join (Current Output)
Currently, paragraph splits and joins trigger block reconstruction in `engine/reconstruction-mode.js`, which discards original paragraph properties and IDs:

```xml
<!-- Reconstruction generates fresh paragraphs, losing original w14:paraId and custom styles -->
<w:p w14:paraId="NEW_01">
  <w:pPr><w:rPr><w:ins w:id="3" w:author="Reviewer" w:date="..."/></w:rPr></w:pPr>
  <w:r><w:t>Split second paragraph content.</w:t></w:r>
</w:p>
```
*Gap:* WP-09/WP-10 will preserve original paragraph nodes, transferring trailing runs and marking paragraph-mark revisions natively.

### 3.3 Run Property Changes (Current Output)
Currently, `injectFormattingToRPr` creates `<w:rPrChange>`:

```xml
<w:r>
  <w:rPr>
    <w:b w:val="1"/>
    <w:rPrChange w:id="4" w:author="Reviewer" w:date="2026-09-05T12:00:00Z">
      <w:rPr>
        <w:b w:val="0"/>
      </w:rPr>
    </w:rPrChange>
  </w:rPr>
  <w:t>Indemnified Parties</w:t>
</w:r>
```
*Gap:* WP-11 will expose this via explicit `CharacterFormatOperation` and `ParagraphFormatOperation`, with author-aware coalescing.

---

## 4. Audit Sign-Off

- **Baseline Tests Verified:** All 67 regression test suites pass (`npm test` $\rightarrow$ 67 passed, 0 failed).
- **Isolation Checks Verified:** `npm run test:isolation` passes.
- **Type Declarations Verified:** `npm run check:types` passes (108 runtime exports).
- **Audit Gate:** Every gap identified above maps directly to WP-01 through WP-18 in the companion plan.
