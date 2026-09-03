# Multimodal LLM Visual Inspection Samples Report — 2026-09-02

- **Evaluation Type:** Multimodal LLM Visual Inspection (Real-Document Spot Check)
- **Source Corpus:** SuperDoc Legal & Administrative Corpus
- **Evaluator:** Multimodal Vision AI Model
- **Target Samples:** 3 Real Documents (Municipal Board Agenda, Parish Council Financial Minutes, City Zoning Ordinance)
- **Renderer:** Microsoft Word 16.0 (Build 16.0.20326) via `ExportAsFixedFormat`
- **Sample Generation Command:** `npm run test:visual:sample` / `node scripts/sample-multimodal-visual-check.mjs`

---

## Executive Summary

As defined in [`docs/TESTING.md`](../TESTING.md#multimodal-llm-visual-inspection-real-document-spot-checks), multimodal LLM visual inspection is an on-demand, sampled evaluation conducted on real-world documents to detect subtle layout, table, list, and typographical regressions that programmatic string/XML diffs cannot observe.

This report documents a complete end-to-end execution of the multimodal visual spot-check workflow across three authentic documents representing different layout complexities:

| Sample | Scenario Identity | Category & Shape | Document Size | Visual Evaluation |
|---|---|---|---|---|
| **Sample 1** | `superdoc:administrative-multi-bullet-board-agenda` | Administrative / List (Logo & Mixed Hierarchy) | 1 Page | **PASS** — Flawless list levels, zero ghost markers, image header intact |
| **Sample 2** | `superdoc:administrative-multi-table-council-minutes` | Administrative / Table (Financial Grid) | 2 Pages | **PASS** — Shaded table header, grid lines, and column widths preserved |
| **Sample 3** | `superdoc:legal-complex-zoning-list-table-batch` | Legal / Multi-Page Complex Ordinance | 25 Pages | **PASS** — Page 4 sub-clauses aligned, font size stable, 25-page count preserved |

---

## Detailed Sample Evaluations

### Sample 1: Municipal Board Agenda (`administrative-multi-bullet-board-agenda`)

- **Document Type:** Town of Prattsville Board Agenda (Public Meeting)
- **Visual Features:** Graphic logo header (`PRATTSVILLE Est. 1824`), lettered agenda sections (A., B., C., D.), dash sub-bullets, and numbered report sections (1–5, 1–8).
- **Operations Evaluated:**
  - Item A: Insertion of `THE` (`CALL THE MEETING TO ORDER`).
  - Item C: Replacement of `FINANCIALS` with `FINANCIAL` (`APPROVAL OF THE MONTHLY FINANCIAL REPORT`).
  - Item D: Replacement of `/ CORRESPONDENCE FROM THE COMMUNITY` with `AND COMMUNITY CORRESPONDENCE`.

#### Multimodal Visual Checklist:

1. **Markup Visibility (`allMarkup`):**
   - Insertions (`THE`, `FINANCIAL`, `AND COMMUNITY CORRESPONDENCE`) appear cleanly in red underline.
   - Deletions (`FINANCIALS`, `/ CORRESPONDENCE FROM THE COMMUNITY`) show distinct strikethrough without clipping surrounding text.
   - Revision change bars are correctly rendered in the left margin.
2. **Accepted State (`acceptAll`):**
   - Text flows naturally with correct single spacing between words.
   - Line wrapping on Item D (`D. COMMENTS FROM THE FLOOR AND COMMUNITY\nCORRESPONDENCE`) conforms cleanly to document margins.
   - Zero stray revision artifacts or misplaced spaces.
3. **Rejected State (`rejectAll`):**
   - Accurately restores the exact original wording and layout.
   - Dash bullets (`- Pledge of Allegiance`, `- Roll Call`) and numbered items remain in place without marker displacement.
4. **Structural Integrity:**
   - The graphic logo and centered meeting header remain in identical coordinate positions.
   - **Verdict:** **PASS**

---

### Sample 2: Parish Council Financial Minutes (`administrative-multi-table-council-minutes`)

- **Document Type:** Cholmondeley & Chorley Parish Council Minutes
- **Visual Features:** Multi-table document with light-blue shaded header rows, grid borders, numeric currency columns (`Amount`), multi-line descriptions, and two-column attendee headers.
- **Operations Evaluated:**
  - Table Header: Replacement of `date` with `Date` (`Invoice Date`).
  - Row 1: Replacement of `Payroll` with `Payroll,` (`Payroll, April 2023-March 2024`).
  - Row 3: Replacement of `Jan-March2024tax,` with `, January-March 2024` (`HMRC tax, January-March 2024`).

#### Multimodal Visual Checklist:

1. **Markup Visibility (`allMarkup`):**
   - Strikethroughs and red underlines inside table cells remain strictly contained within their respective cells.
   - No text overflow into adjacent columns (`Account name`, `Amount`).
2. **Table Grid & Cell Properties:**
   - Header shading (`fill="#EBF1F5"`) remains unbroken across all four columns.
   - Column widths remain stable; the narrow `Amount` column does not expand or collapse.
   - Outer border and interior cell gridlines render continuously without hairline gaps.
3. **Accepted State (`acceptAll`):**
   - "Invoice Date" renders with proper capitalization and bold formatting.
   - Punctuation and spacing around currency amounts and date ranges are clean and readable.
4. **Page Layout:**
   - Two-column layout at top (`PRESENT:` vs `Vice Chair`, `APPOLOGIES:` vs `Chairman`) is undisturbed.
   - Document length is exactly 2 pages with centered footer (`Page: 1`, `Page: 2`).
   - **Verdict:** **PASS**

---

### Sample 3: Municipal Zoning Ordinance Resolution (`legal-complex-zoning-list-table-batch`)

- **Document Type:** City of Cupertino Planning Commission Draft Resolution (25 Pages)
- **Visual Features:** Formal legal resolution apparatus with double-line borders, legal definitions, nested statutory citations, and zoning schedule tables.
- **Operations Evaluated (Page 4):**
  - Definition "Accessory dwelling unit", sub-item 1: Insertion of `the California` (`...Section 17958.1 of the California Health and Safety Code.`).
  - Definition "Accessory dwelling unit", sub-item 2: Insertion of `California` (`...Section 18007 of the California Health and Safety Code.`).

#### Multimodal Visual Checklist:

1. **Markup Visibility (`allMarkup`):**
   - Red underlined additions appear precisely adjacent to statutory section numbers on page 4.
   - Vertical revision bars appear cleanly in the left margin.
2. **Accepted State (`acceptAll`):**
   - Statutory phrases flow seamlessly into the legal definitions.
   - List numbering (`1.`, `2.`) maintains uniform hanging indent alignment.
   - Font family and point size remain completely uniform with the surrounding body text (no unstyled font fallback).
3. **Document-Wide Pagination Stability:**
   - Across the entire 25-page document, page counts are identical (25 pages in `allMarkup`, 25 pages in `acceptAll`, 25 pages in `rejectAll`).
   - Centered footer (`- 4 -`) and top section headings (`SECTION 1. Section 19.08.030...`) remain aligned.
   - **Verdict:** **PASS**

---

## Conclusion & Recommendations

The multimodal visual spot checks confirmed that desktop Microsoft Word renders the engine's OOXML output with complete layout, typographical, and structural fidelity:
- Table grids, borders, and cell background shading are preserved.
- Mixed list hierarchies (dashes, numbers, letters) maintain exact hanging indents without ghost markers.
- Long-document pagination (up to 25 pages) does not drift across revision states.
- Running `npm run test:visual:sample` provides a fast, repeatable mechanism to package and inspect real-document spot checks on demand.
