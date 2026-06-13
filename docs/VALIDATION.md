# Validation

This package works on OOXML strings and intentionally leaves `.docx` zip
packaging to consumers. Release validation should therefore check both XML
invariants and at least one real OOXML consumer.

## Automated Checks

Run:

```bash
npm test
npm run test:isolation
npm run check:types
```

## Export Fixture Parts

Run:

```bash
node scripts/export-validation-fixtures.mjs
```

The script writes generated `word/document.xml` payloads to
`tmp/validation-docx/`. If a case emits numbering XML, it writes that alongside
the document XML. The script does not create full `.docx` files because the
library intentionally avoids a zip dependency.

## Manual Consumer Checks

To inspect in Microsoft Word or LibreOffice:

1. Start from a minimal valid `.docx` package.
2. Replace `word/document.xml` with one generated `*.document.xml` fixture.
3. Add matching `word/numbering.xml` when present.
4. Open the document and confirm the file opens without repair prompts and the
   expected tracked changes are visible.

On Windows with desktop Word installed, the optional COM smoke script described
in the improvement plan can be used to open a completed `.docx` and count
revisions.

With LibreOffice installed, a quick parser check is:

```bash
soffice --headless --convert-to pdf path/to/fixtures/*.docx
```
