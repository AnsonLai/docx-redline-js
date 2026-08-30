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

The minimal synthetic packager currently supports `word/document.xml` and
optional numbering. Tests involving comments, headers, footers, footnotes,
endnotes, or external hyperlinks first require the packager to emit the related
parts, content-type entries, and relationships.

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
