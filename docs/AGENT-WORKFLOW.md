# Agent Document Workflow

Use the `docx-redline` CLI for complete `.docx` files. It emits JSON on stdout,
keeps exact text intact, and never overwrites the source unless `--in-place` is
explicitly supplied.

## Recommended sequence

```powershell
docx-redline inspect contract.docx --non-empty
docx-redline extract contract.docx --range 10:30 > paragraphs.json
docx-redline preflight contract.docx --operations operations.json --author "Editor"
docx-redline apply contract.docx --operations operations.json --author "Editor" --output reviewed.docx
docx-redline validate reviewed.docx
```

Copy `exactText`, `paragraphId`, and `fingerprint` from `extract` into operation
targets. Never normalize or reconstruct `exactText`. Operation files follow
[`document-operations.schema.json`](schemas/document-operations.schema.json).

## Commands

- `inspect` returns the structured inventory, comments, authors, and counts.
- `extract` returns a compact target inventory with exact text.
- `preflight` checks targets, anchors, revisions, conflicts, authors, and needed artifacts without mutation.
- `apply` applies an operation file transactionally.
- `accept` and `reject` resolve revisions selected by `--author` or `--all-authors`.
- `delete-comments` removes matching definitions and document anchors together.
- A whole-paragraph delete stops with `COMMENTED_CONTENT_DELETE` when the
  paragraph has an existing comment. Surface the returned reviewer and comment
  text for human follow-up; do not silently convert this into comment removal.
- `validate` checks revision markup and DOCX package wiring.

Paragraph indexes are 1-based. Inspection filters are `--index 12`,
`--range 10:30`, `--indexes 2,5,8`, `--search text`, `--revised`, `--table`,
`--body`, `--non-empty`, and `--view accepted|rejected|current`. A malformed
filter or unknown option is an error rather than an unfiltered fallback.

Mutating commands require `--author`, authors on every operation, or
`--all-authors` where applicable. Without `--output`, a sibling such as
`contract.redlined.docx` is chosen. Existing outputs are refused unless
`--force` is present. `--in-place` is the only way to overwrite the input.

Treat a nonzero exit code or JSON `status: "error"` as failure. A failed atomic
operation reports `written: false` and does not write an output file.
Missing or repeated comment anchors are errors rather than no-ops. Explicit
anchors match exact text first and then a unique ordinary-space/NBSP equivalent;
omit `textToComment` to comment the entire resolved paragraph.

To reply inside an existing Word comment thread, use the comment ID returned by
`inspect` and do not supply a paragraph target:

```json
{ "type": "comment_reply", "parentCommentId": "8", "commentContent": "Agreed; updated.", "author": "Editor" }
```

Replies are represented in `word/commentsExtended.xml` and deliberately add no
new `commentRangeStart`, `commentRangeEnd`, or `commentReference` to the body.

## Legacy skill wrapper migration

Older skills that invoke `scripts/extract_text.mjs` and
`scripts/apply_changes.mjs` should use the compatibility entrypoints published
with this package rather than carrying copied targeting or ZIP logic. The
legacy positional apply form remains supported:

```bash
node scripts/apply_changes.mjs input.docx changes.json output.docx --author "Editor"
```

Operation files may contain an array, an `operations` array, or a legacy
`changes` array. The wrapper delegates to the same strict, atomic, validated
CLI described above. If `--author` and operation authors are absent, its
compatibility fallback is `DOCX_REDLINE_AUTHOR` and then `Agent`. Consumers
must use the JSON status and process exit code; failed atomic work has
`written: false`, `outputPath: null`, and does not modify the output path.
