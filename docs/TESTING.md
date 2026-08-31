# Testing Guide

This project uses several test lanes because no single oracle can prove that a
tracked-change document is correct. JavaScript assertions can verify exact XML
and return contracts; Microsoft Word can verify what the primary consumer
actually accepts and rejects; real documents expose structures that small
fixtures rarely contain.

## Test lanes at a glance

| Lane | Command | What it proves | What it does not prove |
|---|---|---|---|
| Automated regression suite | `npm test` | API behavior, exact text invariants, OOXML structure, deterministic fuzz cases, and catalogue integrity | That desktop Word accepts the generated package |
| Isolation, types, and lint | `npm run test:isolation`, `npm run check:types`, `npm run lint` | Runtime boundaries, declaration alignment, and static repository rules | Document correctness |
| JavaScript coverage | `npm run test:coverage` | Which source lines, functions, and branches the automated suite executes | That an executed path is correct |
| Synthetic Word differential | `npm run test:word` | Word opens generated packages, sees revisions, and produces the intended text after Accept All and Reject All | The diversity of real-world packages |
| SuperDoc Word corpus | `npm run test:corpus:word` | The same Word differential on 20 reviewed, pinned real English legal/administrative documents while untouched package parts remain byte-identical | Every possible DOCX producer or document type |
| XSD and LibreOffice | See `docs/VALIDATION.md` | Schema conformance and acceptance by a second consumer | Word-specific revision semantics |

## Automated JavaScript tests

Files matching `tests/*.mjs` are discovered by `scripts/run-tests.mjs`. Tests
use `assert/strict` and run as separate Node processes. Shared OOXML assertions
belong in `tests/helpers/ooxml-assertions.mjs`; XML-provider setup belongs in
`tests/setup-xml-provider.mjs`.

For an engine regression:

1. Add the smallest fixed case that reproduces the bug and asserts the exact
   result or structured error.
2. Assert both sides of tracked changes: accepting must produce the requested
   text and rejecting must restore the original text exactly.
3. Assert structural invariants with `validateRedlineOoxml` and the shared OOXML
   helpers. Do not use a whitespace-collapsing reader as the oracle.
4. If engine behavior changed, add the same shape to the deterministic fuzz
   corpus so nearby inputs are exercised too.
5. Run `npm test`, isolation, types, and lint.

## Synthetic Microsoft Word tests

The runtime library contains no COM or Word dependency. Word automation exists
only in Windows development scripts and is intentionally excluded from package
runtime paths.

The synthetic catalogue is `tests/fixtures/word-task-cases.mjs`. During
`npm run test:word`:

1. `scripts/export-validation-fixtures.mjs` applies each operation, runs the
   library's structural validator, checks any case-specific `requiredElements`,
   and builds a minimal `.docx` under ignored `tmp/word-validation/` storage.
2. Expected Accept All and Reject All text is derived from the test's edit
   intent, not from the library's own revision-management functions.
3. `scripts/word-com-differential.ps1` opens each fixture through desktop Word
   without a repair dialog and confirms Word sees at least one revision.
4. Word accepts every revision in one fresh document and rejects every revision
   in another. Both resulting texts are compared with the intent-derived
   expectations.

To add a synthetic Word case:

1. Add a unique lowercase-hyphenated entry to `WORD_TASK_CASES` with
   `category: 'legal'` or `'administrative'`, a distinct `task`, and `original`
   and `modified` text.
2. Use `sourceDocumentXml` when the case needs bookmarks, hyperlinks, tables,
   content controls, prior revisions, or another deliberate structure.
3. Add `requiredElements`, such as `{ bookmarkStart: 1, bookmarkEnd: 1 }`, when
   retaining a structure is part of the claim. These are minimum namespace-aware
   element counts checked before Word runs.
4. Supply `expectedAcceptedText` and `expectedRejectedText` when the source is a
   multi-paragraph document or Word exposes structural separators such as table
   row boundaries. Exact comparison is the default.
5. Run `node tests/word_task_catalog_tests.mjs`, then `npm run test:word` on a
   Windows machine with desktop Word installed.

### Related-part fixture schema

The script-only packager supports optional numbering plus comments, footnotes,
endnotes, headers, footers, and external hyperlinks. Add them to a catalogue
case through `packageParts`:

```js
{
  sourceDocumentXml,
  packageParts: {
    commentsXml,
    footnotesXml,
    endnotesXml,
    headers: [{
      partName: 'header1.xml',
      relationshipId: 'rIdHeader1',
      xml: headerXml
    }],
    footers: [{
      partName: 'footer1.xml',
      relationshipId: 'rIdFooter1',
      xml: footerXml
    }],
    externalHyperlinks: [{
      relationshipId: 'rIdPolicy',
      target: 'https://example.com/policy'
    }]
  }
}
```

`createCommentsPart`, `createNotesPart`, and `createHeaderFooterPart` in
`tests/fixtures/word-package-parts.mjs` provide small escaped constructors for
the common XML parts. Header/footer `partName` and `relationshipId` values have
deterministic defaults, but explicit values make the corresponding
`w:headerReference`, `w:footerReference`, or `w:hyperlink` easier to audit.

Before ZIP emission, the packager rejects malformed XML, duplicate relationship
IDs or part names, missing relationship targets, undefined comment/note IDs,
and note parts without separator IDs `-1` and `0`. It generates the required
content-type overrides and document relationships. Every supplied related part
is compared byte-for-byte with the packaged entry, recorded as SHA-256 in the
case sidecar, and rechecked by the Word differential before opening the DOCX.
These helpers remain under `scripts/` and `tests/`; none are shipped through the
runtime entry point.

### What the automated Word differential proves

The Word process is real desktop Microsoft Word, but it is driven invisibly
through COM. For each fixture, the script:

1. starts `Word.Application` with alerts and the window disabled;
2. opens the generated package with `OpenNoRepairDialog`;
3. fails if Word cannot open it or sees zero tracked revisions;
4. accepts all revisions and reads `Document.Content.Text`;
5. closes without saving, reopens the untouched fixture, rejects all revisions,
   and reads the text again; and
6. compares both results with expectations that were not calculated by this
   library's own accept/reject implementation.

Exact comparison is the default. The harness removes Word's terminal paragraph
mark, normalizes CR/LF representation, and removes the characters Word exposes
as table-cell and footnote/endnote reference boundaries; case-specific normalized comparison requires an
explicit reason. Documents are never saved by the automated differential.

For synthetic fixtures, expectations come directly from `original`, `modified`,
and any explicit full-document expectations in the case. For SuperDoc fixtures,
Word first reads the original pinned source document; the target must occur
exactly once, and the accepted expectation is formed by replacing that unique
target in Word's own source text. The corpus packager also hashes every package
part other than the intentionally replaced `word/document.xml`.

This proves that Word can consume the package, recognizes the revision markup,
and resolves Accept All and Reject All to the intended text. It does **not**
prove that the document looks right on the page. `Content.Text` cannot detect
bad pagination, awkward revision balloons, shifted table widths, broken tab
alignment, font substitution, changed list indentation, clipped headers,
visually stale fields, or a comment/footnote marker that is technically present
but poorly placed.

### Human Word visual review

Human visual review complements—not replaces—the automated differential. Use
the checklist and report template in
[`WORD-MANUAL-REVIEW.md`](./WORD-MANUAL-REVIEW.md).

Review is required for:

- every new or materially changed synthetic Word case;
- the first case for a new structure, operation type, or related package part;
- any engine change affecting reconstruction, formatting, lists, tables,
  fields, tabs/breaks, comments, notes, headers/footers, or revision metadata;
- any case that needs normalized rather than exact text comparison; and
- any automated Word failure whose cause is not immediately textual.

Before a release, review all new/changed cases plus a rotating sample of at
least 20% of the unchanged synthetic catalogue. The sample must include legal
and administrative content and at least one list, table, formatted-run, and
structural-anchor case. Also review at least one legal and one administrative
SuperDoc result. Rotate the sample so every retained synthetic case receives a
human review over five release cycles. A major release or a change to package
assembly requires a full visual sweep of affected structure families.

The reviewer inspects three states in Word: tracked changes with **All Markup**,
the result after **Accept All**, and a fresh copy after **Reject All**. Record
the reviewer, date, Word version/build, cases selected, pass/fail result, and
notes. A visual failure becomes a regression case or a documented harness gap;
do not waive it merely because the COM text differential passed.

### AI-assisted Word visual preflight

An AI agent with Windows computer control may also open the generated fixtures
in the installed desktop Word application, switch among All Markup, Accept All,
and Reject All views, capture screenshots, and inspect them for visible
regressions. This is a useful intermediate oracle because it exercises the same
real UI a person sees rather than only `Document.Content.Text`.

Run an AI visual preflight:

- after a phase that changes reconstruction, formatting, list/table behavior,
  package assembly, or revision display;
- after every five to ten new Word fixtures;
- for all newly introduced structure families;
- when automated Word passes but the XML change is unusually broad; and
- before asking a human to perform the release sample, so obvious failures are
  found first.

The AI should use the same selection rules and checklist as a human, inspect at
least the changed cases plus representative legal and administrative samples,
and save screenshots under ignored `tmp/word-manual-review/<date>/` storage.
Its report must identify itself as **AI visual preflight**, record the Word
version/build when available, list the exact cases and views inspected, and
separate observed facts from uncertain visual judgments.

AI visual review is advisory. It can catch missing or misplaced revisions,
unexpected whole-paragraph markup, obvious pagination shifts, broken tables,
lost indentation, clipped headers, and visibly misplaced anchors. It may miss
subtle font metrics, accessibility issues, field behavior that requires domain
knowledge, or a legally meaningful formatting distinction. An AI pass does not
satisfy the human release sign-off, and sensitive/private documents must not be
opened for AI review without explicit authorization.

## SuperDoc real-document Word tests

The corpus lane references selected documents from SuperDoc's
[docx-corpus](https://docxcorp.us/) under ODC-By 1.0. Source documents are never
committed. `tests/corpus/superdoc-english-legal-administrative.json` pins each
reviewed source and observed SHA-256; `tests/corpus/superdoc-word-scenarios.json`
defines one deterministic edit and records the structural coverage it adds.

To add a corpus case:

1. Select an English legal or administrative document that adds a structure or
   task not already represented. Review its content and record why it is
   suitable.
2. Add an explicit pinned reference and observed hash to the manifest. Do not
   use a floating or bulk corpus download.
3. Fetch only that reference with
   `npm run corpus:fetch:superdoc -- --id <pinned-id>`.
4. Inspect the document in Word and choose a target that occurs exactly once in
   Word's source text.
5. Add the deterministic operation, shape, coverage labels, and review note to
   `superdoc-word-scenarios.json`.
6. Run `npm test` for manifest/catalogue checks and `npm run test:corpus:word`
   for package hashing plus the Word Accept All/Reject All differential.

The corpus packager starts from the original `.docx`, replaces only
`word/document.xml`, and verifies every untouched ZIP part byte-for-byte before
Word opens the result.

## Choosing where a new test belongs

- A fixed bug always gets an automated regression.
- A broad engine rule also gets deterministic fuzz coverage.
- A claim about what Word renders or accepts gets a synthetic Word case.
- A claim involving package complexity or realistic authoring structures gets
  a reviewed corpus case.
- A structural XML rule should be checked by runtime validation and, where
  practical, the ECMA-376 XSD lane.

Keep expectations independent of the code under test. In particular, never use
this library's Accept/Reject helpers to calculate the expected result for the
Word differential.
