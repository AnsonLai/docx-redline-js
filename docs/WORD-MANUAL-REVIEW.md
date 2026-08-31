# Microsoft Word Manual Review

This checklist is the human visual companion to `npm run test:word` and
`npm run test:corpus:word`. The automated differential proves Word revision
semantics by comparing text after Accept All and Reject All. This review checks
the layout and interaction details that `Document.Content.Text` cannot see.

An AI agent may use this same checklist as a visual preflight by controlling the
local Word UI and inspecting screenshots. Mark that report **AI visual
preflight**; it helps select and triage cases but does not replace the human
release sign-off described below.

## Prepare the review set

Run the automated lane first:

```powershell
npm run test:word
npm run test:corpus:word
npm run report:word:coverage
npm run review:word:prepare -- --cycle=0
```

The last command writes a pending review manifest under ignored
`tmp/word-manual-review/` containing changed catalogue families, the rotating
20% synthetic release sample, and legal/administrative corpus representatives.
It only prepares the selection: it never records a visual pass or human
sign-off. Use a new cycle number for each release rotation.

Synthetic documents are generated under `tmp/word-validation/`; reviewed
SuperDoc results are under `tmp/superdoc-word-fixtures/`. Do not commit generated
or downloaded `.docx` files.

Select:

- every new or changed case;
- every case required by the triggers in `docs/TESTING.md`;
- at least 20% of unchanged synthetic cases, rotating from the prior release;
- at least one legal and one administrative SuperDoc result; and
- representative list, table, formatted-run, and anchor/field/content-control
  structures.

Record the installed Word version and build from **File → Account → About
Word**. Differences in rendering can be version-specific.

## Configure Word

For the tracked-change inspection:

1. Open the generated fixture directly in desktop Word.
2. On **Review**, select **All Markup**.
3. Under **Show Markup**, enable insertions/deletions, formatting, comments, and
   all reviewers relevant to the case.
4. Use the expected balloons/inline display for the document and enable
   paragraph marks when checking whitespace, tabs, breaks, lists, and empty
   paragraphs.
5. Do not overwrite the generated fixture. Work on disposable copies if a
   saved accepted or rejected view is useful.

## Inspect each case

### Tracked-change view

- The document opens without a repair, conversion, or unreadable-content prompt.
- Word shows the expected revision count and author attribution.
- Insertions and deletions are anchored at the intended words or paragraph
  marks; a small edit has not become an unexplained whole-paragraph rewrite.
- Untargeted text and surrounding revisions remain unchanged.
- Existing bold, italic, underline, highlighting, fonts, styles, and language
  settings remain visually consistent.
- Spaces, tabs, manual breaks, paragraph spacing, indentation, and alignment
  look intentional with formatting marks visible.
- Lists retain numbering, levels, continuation, indentation, and marker style.
- Tables retain widths, borders, merged cells, row heights, and alignment.
- Bookmarks, hyperlinks, fields, content controls, comments, and note references
  remain in the correct visible location and still behave when activated.
- Headers, footers, section boundaries, page breaks, and pagination remain
  stable around the edit.
- Revision balloons and comment balloons point to the correct content and do
  not obscure or displace unrelated layout unexpectedly.

### Accept All view

On a disposable copy, choose **Accept All Changes** and verify:

- the resulting visible text expresses the intended edit;
- no deletion residue, empty revision wrapper, unexpected blank line, or stale
  formatting remains;
- lists, tables, fields, links, comments, notes, and page layout still work; and
- untargeted content is visually unchanged.

Close the copy without replacing the generated fixture.

### Reject All view

Reopen a fresh copy, choose **Reject All Changes**, and verify:

- the original visible text and formatting are restored;
- original list numbering, table layout, fields, anchors, and pagination return;
  and
- no content introduced by the edit remains.

## Record the result

Keep the review report with the release-validation artifacts. Screenshots may be
stored under ignored `tmp/word-manual-review/<date>/` when useful, but do not
commit corpus document images or document contents without checking their
rights and sensitivity.

Suggested report:

```markdown
# Word visual review — <release/date>

- Reviewer:
- Review date:
- Word version/build:
- Automated synthetic result:
- Automated corpus result:
- Review type: Human sign-off | AI visual preflight

| Case | Why selected | All Markup | Accept All | Reject All | Result | Notes |
|---|---|---|---|---|---|---|
| example-case | New table structure | Pass | Pass | Pass | Pass | No layout shift |

## Failures or follow-ups

- None.
```

A **Pass** requires all three views to pass. Record **Fail** if the rendering or
interaction is wrong even when automated text comparison passes. Turn a failure
into a fixed regression case when possible; otherwise record the exact harness
or Word-version limitation in the active reliability plan.

For an AI preflight, also record screenshot paths and confidence/uncertainty in
the notes. A human reviewer should revisit every AI failure or uncertain result
and must still complete the release sample independently.
