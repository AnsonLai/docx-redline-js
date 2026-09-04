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
- `validate` checks revision markup and DOCX package wiring.

Inspection filters are `--range 10:30`, `--indexes 2,5,8`, `--search text`,
`--revised`, `--table`, `--body`, `--non-empty`, and
`--view accepted|rejected`.

Mutating commands require `--author`, authors on every operation, or
`--all-authors` where applicable. Without `--output`, a sibling such as
`contract.redlined.docx` is chosen. Existing outputs are refused unless
`--force` is present. `--in-place` is the only way to overwrite the input.

Treat a nonzero exit code or JSON `status: "error"` as failure. A failed atomic
operation reports `written: false` and does not write an output file.
