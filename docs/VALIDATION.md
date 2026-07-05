# Validation

This package works on OOXML strings and intentionally leaves `.docx` zip
packaging to consumers (release *tooling* assembles minimal `.docx` fixtures
with a script-local zip writer; the published library still has no zip
dependency).

The test suite verifies the accept/reject round-trip invariant using the
library's own transforms. Because a shared misconception between the
generator and the resolver would pass those tests silently, release
validation adds **independent oracles**: Microsoft Word, LibreOffice, and
the ECMA-376 schemas.

## Automated Checks (every `npm test`)

```bash
npm test              # includes tests/roundtrip_fuzz_tests.mjs (seeded, deterministic)
npm run test:isolation
npm run check:types
```

The fuzz harness generates random paragraph structures and edits, then
asserts the round-trip invariant plus `validateRedlineOoxml` on each case.
Tune or reproduce with:

```bash
FUZZ_SEED=<seed> FUZZ_ITERATIONS=<n> node tests/roundtrip_fuzz_tests.mjs
```

A failing case prints its exact reproduction command.

## Runtime Guardrail

`validateRedlineOoxml(oxml)` (exported from `index.js`) runs the structural
invariants at runtime and returns `{ valid, issues }`. Downstream packagers
should call it before writing engine output into `word/document.xml`.

## Export Fixtures

```bash
node scripts/export-validation-fixtures.mjs
```

Writes to `tmp/validation-docx/`, per case:

- `<name>.document.xml` — generated `word/document.xml` payload
- `<name>.docx` — minimal assembled package
- `<name>.expected.json` — expected accept-all / reject-all plain text,
  derived from edit *intent* (not from this library's transforms), so
  external consumers act as independent oracles

## Word Differential Check (Windows, desktop Word)

```bash
node scripts/export-validation-fixtures.mjs
npm run smoke:word:diff
```

For each fixture, desktop Word opens the `.docx`, confirms revisions are
visible, runs **AcceptAllRevisions**, and compares the document text to the
expected modified text; then reopens and runs **RejectAllRevisions** and
compares to the original text. This is the strongest check available: Word
itself resolves the revisions this library generated.

The older `npm run smoke:word -- path/to/file.docx` open-only smoke check
remains available for ad-hoc files.

## Schema Validation (ECMA-376 transitional XSD)

```bash
node scripts/export-validation-fixtures.mjs
bash scripts/validate-fixtures-xsd.sh
```

Downloads (and caches in `.cache/ooxml-schemas/`) the transitional
wordprocessingml schemas from ECMA-376 Part 4, patches the `xml:` namespace
import to resolve offline, and validates every `*.document.xml` fixture with
`xmllint`. Requires `curl`, `unzip`, and `xmllint` (`libxml2-utils` on
Debian/Ubuntu; available on Windows via conda/msys).

## LibreOffice Consumer Check

```bash
cd tmp/validation-docx
soffice --headless --convert-to pdf --outdir converted *.docx
```

A second independent OOXML consumer parsing the fixtures without error.

## Continuous Validation

`.github/workflows/validation.yml` runs nightly (and on demand via
`workflow_dispatch`):

1. **xsd-schema** — exports fixtures and validates them against the
   ECMA-376 transitional `wml.xsd`.
2. **libreoffice** — exports fixtures and converts them with headless
   LibreOffice.
3. **fuzz-extended** — 20,000 fuzz round-trip cases with a date-derived
   seed, so every night explores new inputs. A failure log includes the
   exact `FUZZ_SEED` reproduction command.

The Word differential check stays manual because it requires desktop Word;
run it before tagging a release.
