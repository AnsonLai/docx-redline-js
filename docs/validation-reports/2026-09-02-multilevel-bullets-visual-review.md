# Multi-Level Bullets & Numbered Lists Visual & Non-Visual Review — 2026-09-02

- **Evaluation Type:** Multi-Level List Verification (Visual + Non-Visual Semantic Suite)
- **Target Structure:** Bullets and Numbered Lists with Multiple Changes Across Multiple Levels (Levels 0, 1, 2)
- **Renderer:** Microsoft Word 16.0 (Build 16.0.20326) via COM `ExportAsFixedFormat`
- **Non-Visual Test Suite:** [`tests/multilevel_bullet_tests.mjs`](../../tests/multilevel_bullet_tests.mjs) (5/5 passing)
- **Visual Evidence Directory:** `tmp/multilevel-bullet-visual/`

---

## Executive Summary

Nested lists with changes at multiple levels represent one of the highest-complexity challenges in OOXML tracked changes. Potential failure modes include:
1. **Level Flattening:** Edits inside nested items (Level 1 or 2) dropping `w:ilvl`, causing Word to render the item at the margin (Level 0).
2. **Double Promotion:** Child bullets indented with leading spaces receiving an additional level bump, placing them at Level 3 instead of Level 2.
3. **Ghost Bullets:** Deleting or inserting items leaving orphaned `<w:pPr>` paragraph markers that render as empty bullets in Word.
4. **Numbering Stream Disruption:** Inserting a sub-item disrupting the sequential numbering of sibling items or outer parent items.

To thoroughly address this, we implemented a dual validation strategy:
- **Non-Visual Suite:** 5 automated tests in [`tests/multilevel_bullet_tests.mjs`](../../tests/multilevel_bullet_tests.mjs) asserting on exact OOXML `w:numPr` (`w:numId`, `w:ilvl`), paragraph mark tracking, format boundary isolation, and accept/reject symmetry.
- **Visual Suite:** Direct desktop Microsoft Word COM rendering across 3 layout-sensitive multi-level test cases (both synthetic and authentic legal/administrative documents), inspected with multimodal vision across `allMarkup`, `acceptAll`, and `rejectAll` views.

---

## Non-Visual Semantic Test Suite

All 5 tests in [`tests/multilevel_bullet_tests.mjs`](../../tests/multilevel_bullet_tests.mjs) execute as part of `npm test`:

| Test Name | Scenario & Coverage | Assertions & Invariants | Result |
|---|---|---|---|
| `testNestedChildInsertionAtMultipleLevels` | Inserting a Level 2 bullet under a Level 1 parent item. | Verifies `ilvl="2"` assigned to child; `acceptAll` yields 7 paragraphs with levels `[0, 1, 2, 2, 2, 1, 0]`; `rejectAll` restores original 6 paragraphs. | **PASS** |
| `testConcurrentEditsAcrossAllThreeLevels` | Batch of concurrent edits modifying Level 0, Level 1, and two Level 2 items simultaneously. | Verifies all 3 levels retain exact `ilvl` without level drift (`[0, 1, 2, 2, 1, 0]`); text replacement exact; zero formatting leakage. | **PASS** |
| `testLevelFlatteningGuard` | Directly guards against the visual defect of level flattening when editing a nested item. | Asserts that modifying a Level 2 item strictly maintains `w:ilvl w:val="2"` and never omits `w:ilvl`. | **PASS** |
| `testFormattingPreservationAcrossLevels` | Applying bold markdown (`**text**`) to a Level 2 item. | Verifies Level 2 item retains `ilvl="2"`, while parent Level 1 and sibling Level 2 items contain no bold run properties. | **PASS** |
| `testMultiLevelRangeDeletionAndAcceptance` | Consolidating two adjacent Level 2 items into a single modified item. | Verifies paragraph count reduces from 6 to 5; remaining consolidated item retains `ilvl="2"`; rejection cleanly restores both items. | **PASS** |

### Core Engine Improvement: Child Promotion Guard
During development, we identified and fixed a subtle double-promotion bug in [`core/list-targeting.js`](../../core/list-targeting.js#L123): when an author provided markdown input that was *already* indented with leading spaces (e.g. `  - Sub-item`), `resolveInsertionLevel` placed it at `anchorLevel + 1`. The subsequent `shouldPromoteBulletInsertionsToChildDepth` check added a second level increment, pushing it to `anchorLevel + 2` (`ilvl=3`). Adding an `alreadyIndented` check prevents this double bump, ensuring child bullets consistently settle at `ilvl=2`.

---

## Visual Word COM Evidence & Multimodal Inspection

We rendered three layout-sensitive multi-level list documents in desktop Microsoft Word 16.0:

### Case 1: Synthetic Multi-Level List (`administrative-list-change-nested-child`)
- **Structure:** 2-level numbered outline (Level 0 items `1.`, `2.`, and Level 1 sub-items `1.`, `2.`, `3.`).
- **Operation:** Inserting a new child item (`Escalate unresolved notifications.`) directly below `Disclosure obligations.`
- **Visual Observations:**
  - **All Markup:** Red underlined insertion appears at the exact child indent; sibling items below (`Notify affected parties.`, `Preserve supporting records.`) shift down smoothly with renumbering strikes (`1.2.`, `2.3.`). Margin revision bar is present.
  - **Accept All:** Sub-items cleanly renumber `1.`, `2.`, `3.` under Item 1; Item 2 (`Remediation obligations.`) remains at the outer margin.
  - **Reject All:** Exactly restores original 4 paragraphs without blank space or ghost markers.
  - **Verdict:** **PASS**

### Case 2: Municipal Board Agenda (`superdoc:administrative-list-change-board-agenda-multiple-children`)
- **Structure:** Mixed hierarchy with lettered sections (`A.`, `B.`, `C.`, `D.`), numbered reports (`1.`, `2.`, `3.`, `4.`, `5.`), and dash/bullet sub-items.
- **Operation:** Inserting two new bullet children (`Review meter replacement progress.`, `Confirm hydrant inspection dates.`) under `2. Water District Report`.
- **Visual Observations:**
  - **All Markup:** The two newly inserted bullets appear with red underline and bullet glyphs (`•`) at hanging indent 0.5 in. Margin bar indicates change.
  - **Accept All:** Bullets line up with sub-item indent; numbered items `3. Code Enforcement`, `4. WWTP Report`, `5. Tax Collector Report` maintain their outer alignment and numbering.
  - **Reject All:** Returns to exact 5 numbered report items with zero extra bullets.
  - **Verdict:** **PASS**

### Case 3: Corporate Bylaws Multi-Level Batch (`superdoc:legal-bylaws-nested-list-batch`)
- **Structure:** 7-page legal bylaws document containing Roman numeral Articles (`I.`, `II.`, `III.`), lettered sections (`A. Officers`, `B. Time of Election`), and bullet items.
- **Operation:** Concurrent multi-level edits on Page 2:
  - Level 0: Section II title text replacement (`comprised` -> `composed`).
  - Level 1: Four bullet item edits under Section II removing parenthesized numbers and updating city representation text.
  - Level 1: Subsection A title text update (`Vice-Chair` hyphenation).
- **Visual Observations:**
  - **All Markup:** Red strikethroughs and additions render inline within the bullets; hanging indents for Roman numerals (`II.`, `III.`) and bullet points remain completely aligned.
  - **Accept All:** Text flows seamlessly; bullet points retain uniform left margin; subsection letters `A.` and `B.` align with Section III text.
  - **Pagination:** Exactly 7 pages across all 3 views (`allMarkup`, `acceptAll`, `rejectAll`).
  - **Verdict:** **PASS**

---

## Conclusion

The dual test suite confirms that `@ansonlai/docx-redline-js` robustly handles multi-level lists with concurrent edits across multiple levels:
1. `w:ilvl` values are strictly preserved across single-item edits, subtree additions, and range consolidations.
2. Microsoft Word renders the resulting OOXML with exact typography, proper hanging indents, and zero ghost markers or level collapsing.
