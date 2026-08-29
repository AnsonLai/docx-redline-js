# Redline Reliability Improvement Plan — Round 2

**Status:** In Progress (Phases 1 and 8 Complete; Phases 2–7 Open)

Follow-on to `completed/2026-05-31-architectural changes.md`, which is complete. That plan
hardened the *OOXML shape* of generated redlines (paragraph marks, moves, inert
markup, structural invariants). This plan targets a different class of problem:
**places where the library silently produces wrong output, or where the test
suite cannot see that it did.**

Every finding below was reproduced against the current `master` (`73ef9c5`)
before being written down. Reproduction snippets are included so each phase can
be re-confirmed before work starts.

**Scope note (unchanged):** a docx-in/docx-out wrapper API is still OUT of scope.
Do not add a JSZip dependency or an `applyRedlineToDocx`-style API.

## Production API compatibility map

This plan distinguishes production API compatibility from test quality. Changes
to test helpers, fuzz generators, assertions, linting, coverage, and other
development-only tooling are **not** marked as production API breaks here.

In this document:

- **BREAKING — production API** means an existing caller may need code changes
  because an exception/return contract, default option, or documented result
  behavior changes.
- **CONDITIONAL — production behavior** means the public API shape remains
  compatible, but callers exercising the affected edge case may observe a
  different result or generated OOXML.
- **NOT BREAKING — production API** means the change is internal, additive, or
  limited to test/development tooling.

| Phase | Production API impact |
|---|---|
| 1 | **NOT BREAKING — production API.** Test-only verification and fuzz-corpus changes. |
| 2 | **CONDITIONAL — production behavior.** `DIFF_TOKEN_LIMIT` introduces an error path for oversized inputs; whitespace/schema fixes and deterministic diffs may change generated OOXML for affected inputs. |
| 3 | **BREAKING — production API.** Several malformed-input paths change from throwing to returning result objects; multi-line target misses can change from no-op to error. |
| 4 | **BREAKING — production API if `sanitizeInput` defaults to `false`.** Existing callers relying on implicit sanitization receive different edits. The existing-revision fix is **CONDITIONAL — production behavior**. |
| 5 | **NOT BREAKING — production API.** Public signatures remain compatible; revision IDs may change, which matters only to callers treating generated IDs as stable cross-document identifiers. |
| 6 | **BREAKING — production API.** `atomic: true` as the default changes batch results from partial application to rollback on error. |
| 7 | **NOT BREAKING — production API.** Development checks and coverage tooling only; declaration tightening may require fixes in the repository's type fixtures. |
| 8 | **CONDITIONAL — production behavior.** Structural fixes alter malformed output, and `PARTIAL_TARGET` changes the result for previously mis-scoped document edits. |

If all proposed defaults and contracts land as written, publish the change as a
major version. A minor release would require retaining the current defaults and
error behavior, or introducing the new behavior as opt-in/deprecated paths.

---

## Conventions (read before starting any phase)

Same as the previous plan, repeated here so this file stands alone:

- Tests live in `tests/*.mjs`, auto-discovered by `scripts/run-tests.mjs` (`npm test`).
  `tests/helpers/` and `tests/setup-xml-provider.mjs` are excluded from discovery.
  Plain `assert/strict`, no test framework.
- Shared assertion helpers go in `tests/helpers/ooxml-assertions.mjs`.
- All `w:*` element creation goes through `createWordElement` (`core/word-xml.js`).
- Revision metadata comes from `createRevisionMetadata(author)` (`core/types.js`).
- New public API is re-exported through `index.js`.
- After each phase: `npm test`, `npm run test:isolation`, `npm run check:types` all pass.
- Do not change existing public signatures. New behavior arrives as optional
  fields on options/result objects.

Additional convention for this plan:

- **Every phase that changes engine behavior must add its regression case to the
  fuzz corpus (Phase 1.3), not only to a hand-written test.** Fixed cases prove a
  bug is gone; the corpus proves it stays gone.

---

## Findings summary

| # | Finding | Severity | Phase | Status |
|---|---------|----------|-------|--------|
| F1 | Round-trip oracle is lossy — whitespace/tab/break corruption is invisible to the entire suite | Critical (test blindness) | 1 | **Fixed** |
| F2 | Diff engine silently drops text in documents with >65,536 unique tokens | Critical (data loss) | 2 | Open |
| F3 | Diff output depends on wall-clock time (`Diff_Timeout = 1s`) | High (non-reproducible output) | 2 | Open |
| F4 | Public API error contract is inconsistent: some functions throw raw `ParseError`, some return `''`, some return `status:'error'` | High | 3 | Open |
| F5 | `existingRevisions: 'accept-all-first'` can silently discard another reviewer's revisions while reporting `hasChanges: false` | High (data loss) | 4 | Open |
| F6 | `sanitizeAiResponse` unconditionally mutates legitimate document text | High (data corruption) | 4 | Open |
| F7 | Revision-ID counter is process-global and permanently poisonable | Medium | 5 | Open |
| F8 | `TARGET_NOT_FOUND` detection is disabled whenever `original` contains a newline | Medium | 3 | Open |
| F9 | `npm run check:types` does not type-check anything | Medium | 7 | Open |
| F10 | `applyOperationsToDocumentXml` is non-atomic; a mid-batch failure returns a half-applied document | Medium | 6 | Open |
| F11 | Fuzz corpus is single-paragraph only — no tables, lists, multi-paragraph, or pre-existing revisions | Medium | 1 | **Fixed** |

### Found by Phase 1 once the oracle could see (added 2026-08-02)

| # | Finding | Severity | Phase | Status |
|---|---------|----------|-------|--------|
| F12 | Diff tokenizer `/(\S+)(\s*)/g` cannot match whitespace before the first word, so **leading whitespace is dropped from both sides of every diff** and all diff offsets shift by its length | High (data loss) | 2 | Open |
| F13 | An edit next to a `w:br` emits a **`w:p` nested inside a `w:p`** (schema-invalid, Word reports corruption) and destroys the `w:br` | Critical (corrupt output) | 8 | Open |
| F14 | Reconstruction moves `w:sectPr` to the **front** of `w:body`, violating `CT_Body` and the package's own plumbing validator | High (corrupt output) | 8 | Open |
| F15 | `applyRedlineToOxml` on a document whose `original` covers only some paragraphs **silently deletes the untargeted paragraphs** and returns `status: 'ok'` | High (data loss) | 8 | Open |

---

## Phase 1 — Make the verification oracle honest (do this first)

**Status: Complete (2026-08-02).** Suite is green at 24/24 (was 21/21).

What landed:

- `extractExactVisibleText(xml)` and `normalizeParagraphBreaks(text)` in
  `tests/helpers/ooxml-assertions.mjs` — a lossless extractor written against the
  DOM independently of `pipeline/ingestion-export.js`, modelling Word's accepted
  view (`w:del`/`w:moveFrom` hidden, `w:ins`/`w:moveTo` visible, deleted
  paragraph mark merges into the next paragraph).
- `assertRoundTrip(..., { fidelity })` in `tests/helpers/roundtrip.mjs`,
  defaulting to `'exact'`. `'normalized'` remains for markdown cases that
  legitimately rewrite whitespace; no call site needs it yet.
- Two new structural invariants, both of which the previous assertion set was
  blind to: `assertNoNestedParagraphs` and `assertSectPrLast`. Both are wired
  into `assertRoundTripStructure`.
- `tests/roundtrip_oracle_tests.mjs` — **self-tests for the oracle itself.** A
  verification helper that cannot fail is worthless, so these pin down that the
  exact extractor really does distinguish a tab from a space, a double space from
  a single, and a lost trailing space, and that the old normalized comparison
  genuinely could not.
- Seven whitespace-hostile cases added to `tests/roundtrip_invariant_tests.mjs`
  (double spaces, `w:tab`, trailing/leading space, `w:br`, multi-paragraph),
  plus a `knownGap` skip mechanism that logs each skip every run.
- `tests/roundtrip_fuzz_tests.mjs` widened from one shape to five —
  `paragraph`, `multiParagraph`, `tableCell`, `whitespace`, `existingRevisions`
  (the last driven with `existingRevisions: 'accept-all-first'`) — with a
  per-shape count printed on every run, plus a narrow known-gap classifier so
  registered defects are *counted and reported* rather than either failing the
  build or vanishing.

Verification run: `npm test` 24/24, `npm run test:isolation`, `npm run check:types`,
`node scripts/export-validation-fixtures.mjs`, and a 12,000-case sweep
(`FUZZ_ITERATIONS=12000`) all pass.

**The oracle immediately earned its keep.** Under the old normalized comparison
every one of these passed; under `'exact'` plus the new structural invariants
they are F12, F13, F14 and F15 above. Three are registered as `KNOWN-GAP`
skips/classifiers rather than fixed here, because they are engine defects owned
by Phase 2 and the new Phase 8:

| Defect | Where it shows up | Rate |
|---|---|---|
| F12 leading whitespace dropped | corpus case + fuzz classifier `leading-whitespace-dropped` | 12 / 12,000 fuzz cases |
| F13 `w:br` → nested `w:p` | corpus case `w:br survives an edit in an adjacent run` | — |
| F14 `w:sectPr` moved to front | corpus case + fuzz classifier `sectPr-not-last` | 10,000 / 12,000 fuzz cases (i.e. **every** case that has a `sectPr` at all) |

Grep `KNOWN-GAP` to find every suppression; each names the phase that owns it.
The fuzz harness prints `KNOWN-GAP <id>: 0 cases -- possibly fixed` when a
registered gap stops reproducing, so a stale entry announces itself.

Two deliberate deviations from the plan as written:

1. `assertSectPrLast` needed an opt-out. F14 fires on essentially all
   document-scoped reconstruction output, including the pre-existing
   `tests/paragraph_mark_revision_tests.mjs` fixtures. Rather than strip
   `w:sectPr` from those fixtures (which would have made them less realistic and
   hidden the defect), `assertRoundTripStructure` takes
   `{ knownGaps: ['sectPr-not-last'] }` and each use carries a
   `KNOWN-GAP: Phase 8` comment.
2. The planned list-paragraph fuzz shape is **not** included. List generation
   routes through markdown preprocessing and numbering allocation, which
   legitimately rewrite whitespace and would have needed `'normalized'` fidelity —
   a noisy shape that cannot see the very class of bug this phase exists to
   catch. Tracked as Phase 1 follow-up below rather than shipped weak.

### Phase 1 follow-up (not blocking)

- Add the list-paragraph fuzz shape, with `w:numPr` paragraphs and markdown-list
  target text, once there is an exact-fidelity story for list output. Until then
  `tests/list_tests.mjs` (845 lines) remains the only list coverage.
- `validateRedlineOoxml` (the *runtime* guardrail exported from `index.js`) still
  does not check nested paragraphs or `sectPr` placement — the two invariants
  added here live only in the test helpers. Port both into
  `core/redline-validation.js` when Phase 8 fixes the underlying defects, so
  downstream packagers get the same protection.

**Why:** This is the highest-leverage item in the plan, because it is the reason
the other findings survived a suite that already has a 20,000-case nightly fuzz
sweep.

`assertRoundTrip` (`tests/helpers/roundtrip.mjs`) extracts text with
`ingestWordOoxmlToPlainText`, then compares with
`normalizeVisibleText` (`.replace(/\s+/g, ' ').trim()`).
But `ingestWordOoxmlToPlainText` is a *display-oriented, deliberately lossy*
reader: `normalizeInlineWhitespace` in `pipeline/ingestion-export.js:126-132`
already collapses `[ \t]+` → `' '` and trims every line. So the invariant the
suite actually enforces is "accept-all yields the modified text, up to
whitespace collapsing applied twice."

Reproduced — the engine is **correct** here; the oracle is blind:

```js
// input:  <w:t xml:space="preserve">Section  1 applies.</w:t>   (two spaces)
// redlined XML retains "Section  1 " exactly  ✓
// accepted XML retains "Section  1 " exactly  ✓
// ingestWordOoxmlToPlainText(accepted) === "Section 1 governs."  ✗ collapsed
```

Same for `w:tab`: `readRunText` maps it to `\t` correctly, then
`normalizeInlineWhitespace` turns it into a space. A future change that drops a
`w:tab`, loses an `xml:space="preserve"`, or doubles a space at a splice point
would pass `npm test` and all 20,000 nightly fuzz cases.

Do not "fix" `ingestWordOoxmlToPlainText` — lossy normalization is the right
behavior for a display/markdown reader, and downstream consumers depend on it.
Build a second, lossless extractor for verification.

### 1.1 Lossless verification extractor — done

Shipped as described, with one addition: `collectExactText` refuses to recurse
into a nested `w:p`, so extraction stays well-defined even on the malformed
output F13 produces (otherwise the nested paragraph's text is counted twice and
the failure message misleads).

Add to `tests/helpers/ooxml-assertions.mjs`:

```js
/**
 * Extracts the exact visible text of an OOXML fragment with NO normalization.
 * Unlike ingestWordOoxmlToPlainText (which collapses runs of whitespace for
 * display), this preserves every space, tab, and break so that whitespace
 * regressions are detectable.
 *
 * Mapping: w:t -> textContent, w:tab -> '\t', w:br|w:cr -> '\n',
 * w:noBreakHyphen -> '‑'. Skips runs inside w:del / w:moveFrom.
 * Paragraph boundaries emit '\n'.
 */
export function extractExactVisibleText(xml)
```

Implement it directly against the DOM (mirror the traversal in
`collectParagraphSegments` / `readRunText` in `pipeline/ingestion-export.js`,
minus the `normalizeInlineWhitespace` call). It must be independent of the
production reader so a bug in ingestion cannot mask itself.

### 1.2 Two-tier round-trip assertion — done

Note on what `'exact'` actually means in the shipped version: it normalizes
**paragraph separators only** (`normalizeParagraphBreaks` folds `\n{2,}` to `\n`
and `\r\n` to `\n`), because markdown treats a blank line as a paragraph
separator while OOXML represents one paragraph break as one boundary. Spaces,
tabs, and every other whitespace difference are compared byte-exact. The
existing eleven-case corpus migrated to `'exact'` with **no** failures — the
engine was already correct on all of them; only the oracle had been weak. The
failures came from the newly added whitespace-hostile cases.

Rework `assertRoundTrip` in `tests/helpers/roundtrip.mjs` to take a
`fidelity` option:

- `fidelity: 'exact'` (**new default for all new cases**) — compares
  `extractExactVisibleText` output against the expected string with no
  normalization at all.
- `fidelity: 'normalized'` — current behavior, retained only for cases where
  markdown preprocessing legitimately changes whitespace (list generation, table
  reconciliation). Every call site that uses `'normalized'` must carry a comment
  saying why exact comparison does not apply.

Migrate the existing corpus in `tests/roundtrip_invariant_tests.mjs` to
`'exact'` one case at a time. **Expect failures.** For each one, determine
whether it is engine whitespace corruption (fix the engine) or an artifact of
markdown preprocessing (document it and use `'normalized'` for that case only).
Do not bulk-migrate and then bulk-downgrade the failures.

### 1.3 Widen the fuzz corpus (F11) — done except the list shape

Four of the five families shipped (`multiParagraph`, `tableCell`, `whitespace`,
`existingRevisions`, alongside the original `paragraph`). The list shape is
deferred — see the Phase 1 follow-up above for why. Runtime for the default
100-case sweep is unchanged; a 12,000-case sweep completes well inside the
nightly budget.

`generateParagraph` in `tests/roundtrip_fuzz_tests.mjs` only ever emits a single
`<w:p>` with runs, an optional hyperlink, and optional bookmarks. Structures
that the engine routes very differently are never generated. Add generators for:

1. **Multi-paragraph bodies** (2-5 `w:p`), including edits that delete a whole
   paragraph, insert one, and merge two — the Phase 3 paragraph-mark paths from
   the previous plan currently have only hand-written coverage.
2. **Tables** — a `w:tbl` with 2-4 rows, edits aimed at a single cell paragraph
   (exercises `detectTableCellContext` and the isolate-then-recurse path in
   `engine/oxml-engine.js:119-127`).
3. **List paragraphs** — `w:numPr` bearing paragraphs, plus markdown-list target
   text (exercises `orchestration/list-structural-fallback.js`, 530 lines with
   comparatively little targeted coverage).
4. **Pre-existing revisions** — paragraphs already containing `w:ins`/`w:del`
   from a different author, driven with `existingRevisions: 'accept-all-first'`.
5. **Whitespace-hostile text** — deliberate double spaces, leading/trailing
   spaces, `w:tab` and `w:br` elements mid-paragraph. These only become
   meaningful once 1.1/1.2 land.

Keep the harness seeded and deterministic. Raise the default `FUZZ_ITERATIONS`
only if wall-clock stays under ~15s for `npm test`.

**Acceptance for Phase 1:** met. `extractExactVisibleText` exported;
`assertRoundTrip` defaults to `'exact'` and no call site needs `'normalized'`;
fuzz corpus emits four of five structure families (list deferred with a written
reason); `npm test` green at 24/24.

---

## Phase 2 — Diff engine correctness and determinism

> **Production API: CONDITIONAL — production behavior.** The public signatures
> remain compatible, but oversized token streams may now return the new
> `DIFF_TOKEN_LIMIT` error instead of producing corrupted output. The whitespace
> and determinism fixes can also change generated OOXML for affected inputs.

### 2.1 Token-space overflow silently destroys text (F2)

`wordsToChars` in `pipeline/diff-engine.js:35-48` assigns each unique token a
code unit via `String.fromCharCode(wordArray.length)`. `String.fromCharCode`
takes its argument **modulo 0x10000**, so token 65,536 collides with token 0.
`charsToWords` (`:67-82`) then *silently drops* any code whose value exceeds
`wordArray.length`:

```js
if (charCode < wordArray.length) parts.push(wordArray[charCode]);
// else: token vanishes, no error, no warning
```

Reproduced with 70,001 unique tokens:

```
String.fromCharCode(70000).charCodeAt(0) === 4464     // wrapped, not 70000
wordArray[5] = 't4'   wordArray[65541] = 't65540'     // collide to the same code unit
diff reconstructs original exactly: false
original length 478889 -> reconstructed 473310        // 5,579 characters silently lost
```

A large agreement, a full `word/document.xml` fed through the standalone runner,
or any document with heavy unique-token content (IDs, part numbers, citations,
multilingual text) can cross 65,536 unique tokens. The failure mode is silent
text loss inside generated redlines — the worst possible outcome for this
library.

Additionally, token codes in `0xD800-0xDFFF` are lone surrogates. They survive
`charCodeAt` round-tripping, but they make the intermediate strings ill-formed
UTF-16 and are fragile under any future change to diff-match-patch internals.

**Fix, in order:**

1. **Guard first, hard.** Before diffing, if `wordArray.length` would exceed the
   safe token ceiling, do not produce a wrong answer. Return a structured error
   (`DIFF_TOKEN_LIMIT`, threaded through the Phase 3 error contract) so the
   caller can split the work. Silent corruption must become a loud refusal
   *before* any capacity work lands — ship this step on its own if needed.
2. **Then raise the ceiling.** Switch the encoding to `String.fromCodePoint`
   over a non-surrogate plane (e.g. base `0x10000`, giving ~1M tokens) and
   decode with `codePointAt` + correct index advancement, or keep BMP encoding
   but skip the surrogate range. Whichever is chosen, `charsToWords` must
   **throw on an out-of-range code** rather than dropping it — an unmappable
   code is a bug, never a value to discard.
3. Add a unit test asserting exact reconstruction (`equal + delete` recovers the
   original, `equal + insert` recovers the modified) for token counts spanning
   the old boundary: 1,000 / 65,535 / 65,537 / 200,000.

### 2.1b Leading whitespace is dropped from every diff (F12)

Found by Phase 1. Same function as 2.1, separate defect. `tokenize` in
`pipeline/diff-engine.js:24-33` scans with `/(\S+)(\s*)/g`, which can only start
matching at a non-space character. Whitespace *before the first word* is never
captured by any token:

```js
computeWordDiffs('  indented text', '  indented copy')
// -> [[0,"indented "],[-1,"text"],[1,"copy"]]
//     the two leading spaces are absent from both sides
```

Two consequences, the second worse than the first:

1. The leading whitespace is deleted from the output — and because it lands on
   *unchanged* text outside any revision, rejecting the redline does not bring it
   back.
2. Every diff offset is short by the length of that whitespace, so
   `computeWordLevelDiffOps` hands surgical run splitting positions that are
   misaligned with the real text for any paragraph starting with whitespace.

Fix: capture a leading-whitespace prefix as its own token before the main scan.

```js
function tokenize(text) {
    const tokens = [];
    const leading = text.match(/^\s+/);
    if (leading) tokens.push(leading[0]);
    // ...existing /(\S+)(\s*)/g scan
}
```

Then un-skip the `leading whitespace is preserved` case in
`tests/roundtrip_invariant_tests.mjs` and delete the `leading-whitespace-dropped`
entry from `KNOWN_GAPS` in `tests/roundtrip_fuzz_tests.mjs`. Expect existing
expectations to shift for any fixture whose paragraph starts with whitespace —
that shift is the bug being fixed, not a regression.

### 2.2 Diff output depends on wall-clock time (F3)

`pipeline/diff-engine.js:10` creates one module-level instance:

```js
const DMP = new diff_match_patch();   // Diff_Timeout = 1 (seconds), confirmed at runtime
```

When `diff_main` exceeds one second of wall clock it abandons the optimal
bisection and returns a valid-but-cruder diff. The same input therefore produces
**different redlines on a slower or more loaded machine** — different `w:ins`/
`w:del` boundaries, different revision counts. That breaks reproducibility,
makes user-reported bugs hard to reproduce, and can make the nightly 20k fuzz
sweep flake for reasons unrelated to any code change.

**Fix:**

- Set `DMP.Diff_Timeout = 0` (no timeout → deterministic output) and rely on the
  Phase 2.1 size guard to bound worst-case work instead of a timer.
- If unbounded time is unacceptable for some consumer, expose it as an explicit
  option (`diffTimeoutSeconds`) that defaults to `0`, and document that any
  non-zero value makes output non-deterministic.
- The shared mutable instance is also a latent hazard if timeouts ever become
  per-call configurable — construct per call, or snapshot/restore the setting.
- Test: diff a large paragraph pair twice with an artificially tiny timeout and
  assert the default (`0`) path produces byte-identical output across runs.

**Acceptance for Phase 2:** overflow guard returns a structured error; encoding
handles ≥200k unique tokens with exact reconstruction; `charsToWords` throws
rather than dropping; diff output is deterministic; fuzz corpus gains a
high-unique-token case.

---

## Phase 3 — One error contract across the public API

> **Production API: BREAKING.** Existing callers that catch parse exceptions, or
> assume the current return shape from the affected functions, must adapt when
> malformed-input handling changes to structured results. The new sibling
> ingestion result helpers are additive, but do not remove this break for the
> existing throwing functions.

### 3.1 Unify parse-failure behavior (F4)

The library has four different behaviors for the same malformed input. Confirmed
by feeding `<w:p ...><w:r><w:t>hello</w:t></w:p>` (unclosed `w:r`) to each
public entry point:

| API | Behavior on malformed XML |
|---|---|
| `applyRedlineToOxml` | returns `{ status: 'error', error: { code: 'PARSE_ERROR' } }` ✓ |
| `validateRedlineOoxml` | returns `{ valid: false, issues: [PARSE_ERROR] }` ✓ |
| `injectCommentsIntoOoxml` | returns original oxml + a warning |
| `ingestWordOoxmlToPlainText` | returns `''` — indistinguishable from an empty document |
| `acceptTrackedChangesInOoxml` | **throws raw `ParseError`** |
| `rejectTrackedChangesInOoxml` | **throws raw `ParseError`** |
| `deleteCommentsByAuthorInOoxml` | **throws raw `ParseError`** |

The three throwing functions are exactly the ones a downstream tool calls in a
cleanup pass, often in a loop over many documents — an uncaught `ParseError`
from a dependency's internals takes down the batch.

Of the 31 `parseFromString`/`parseXml` call sites across `core/`, `engine/`,
`pipeline/`, `services/`, and `orchestration/`, only about five are inside a
`try`. Under `@xmldom/xmldom` (the Node peer dependency) a fatal parse **throws**
`ParseError`; it does not return a document.

**Fix:**

1. Add `parseOoxmlSafe(xmlString)` to `adapters/xml-adapter.js` returning
   `{ doc, error }` — never throwing. Route every parse site through it.
2. Every public function that currently throws returns the established shape
   instead: `{ oxml, hasChanges: false, status: 'error', error: { code: 'PARSE_ERROR', message } }`
   for transforms, `{ valid: false, issues: [...] }` for validators.
3. `ingestWordOoxmlToPlainText` / `ingestWordOoxmlToMarkdown` keep returning a
   string (signature change would be breaking), but add sibling
   `ingestWordOoxmlToPlainTextResult(oxml)` returning `{ text, status, error }`
   so callers can distinguish "empty document" from "unparseable input".
   Re-export both from `index.js`, and note the distinction in `AGENTS.md`.
4. Add `tests/error_contract_tests.mjs` asserting **every** exported function
   that accepts an OOXML string returns (never throws) for: malformed XML, empty
   string, `null`, `undefined`, and a non-OOXML but well-formed document
   (`<html><body/></html>`). Drive it off `Object.keys(await import('../index.js'))`
   so newly added exports are covered automatically.

### 3.2 Retire or fix the browser-only parse-error probe

`getXmlParseError` (`core/xml-query.js:97`) looks for a `<parsererror>` element.
That is a browser-DOMParser convention; `@xmldom/xmldom` throws instead and never
produces such an element. On the Node path the check is dead code, which makes
the `PARSE_ERROR` branches in `engine/oxml-engine.js:72-81` and `:91-103` look
better covered than they are. Keep the function for browser hosts, but document
that it is browser-only and make `parseOoxmlSafe` (3.1) the single source of
truth for both runtimes.

### 3.3 Route xmldom diagnostics through the logger

`@xmldom/xmldom` writes `[xmldom error]` / `[xmldom warning]` / `[xmldom fatalError]`
straight to the console, bypassing `adapters/logger.js`. A package that offers
injectable logging should not print to a host's stdout behind its back. Note
also that non-fatal errors do **not** throw and **do** silently alter content —
`<a>&nosuch;</a>` parses to `<a>&amp;nosuch;</a>` with only a console line.

- Pass `@xmldom/xmldom`'s `onError`/error-handler option (see its `DOMParser`
  options) from `parseOoxmlSafe`, forwarding messages to `adapters/logger.js`.
- Surface non-fatal parse diagnostics as `warnings[]` on the result so callers
  can detect content-altering recoveries such as an undefined entity.

### 3.4 Fix `TARGET_NOT_FOUND` for multi-line originals (F8)

`engine/oxml-engine.js:141` disables the whole target-existence check when the
original text contains a newline:

```js
&& !originalText.includes('\n')
```

So a multi-paragraph edit whose `original` does not appear in the document falls
through to a mode handler and silently returns `hasChanges: false` — the caller
cannot tell "nothing to do" from "I aimed at text that isn't there," which is
precisely the distinction Phase 6.1 of the previous plan set out to create.

Extend the check to multi-line originals: normalize both sides per paragraph
(split on `\n`, apply `normalizeTargetText` to each) and require every non-empty
line to be present in the document's visible text. Add tests for a multi-line
original that does match, and one that does not.

**Acceptance for Phase 3:** no exported function throws on any malformed input in
`tests/error_contract_tests.mjs`; xmldom diagnostics reach the injected logger;
multi-line `TARGET_NOT_FOUND` covered; README/AGENTS error-code tables updated
with `DIFF_TOKEN_LIMIT` (Phase 2) and the new result helpers.

---

## Phase 4 — Stop silent mutation of caller content

> **Production API: BREAKING if `sanitizeInput` defaults to `false`.** That
> default changes the behavior of existing callers that rely on implicit input
> sanitization. The `accept-all-first` correction below is a conditional output
> change, not a signature change.

### 4.1 `accept-all-first` can destroy another reviewer's work (F5)

In `engine/oxml-engine.js`, `noChanges` closes over the `oxml` parameter:

```js
const noChanges = () => finalize({ oxml, hasChanges: false });   // :57
...
oxml = accepted.oxml;                                            // :89 — parameter reassigned
```

So when `existingRevisions: 'accept-all-first'` normalizes the input and the edit
then turns out to be a no-op, the caller receives OOXML with **another author's
tracked changes already accepted and stripped**, labelled `hasChanges: false,
status: 'no-op'`. Reproduced:

```
input  had w:ins (author "Prior"): true
result: hasChanges=false, status='no-op'
output has w:ins: false        // the prior reviewer's revision is gone
```

Any caller that follows the natural contract — "`hasChanges: false`, so writing
the payload back is harmless" — silently discards a human reviewer's revisions.

**Fix:**

- Introduce `finalizeUnchanged()` that returns the **original, pre-normalization**
  `oxml` string. Capture it as `const inputOoxml = oxml;` at function entry and
  never reassign that binding.
- When normalization did occur and the edit was a no-op, that is a real change to
  the payload: return the normalized oxml with `hasChanges: true` and a
  `warnings: ['existing revisions were accepted before redlining']` entry, OR
  return the untouched input with `hasChanges: false`. **Pick the second** — it
  is the non-destructive default — and add an explicit
  `existingRevisions: 'accept-all-first-keep-normalized'` value for callers who
  actually want the normalized document back.
- Test: prior-author `w:ins` + a no-op edit under `'accept-all-first'` →
  output still contains the prior `w:ins`. Add a fuzz-corpus case (Phase 1.3 item 4).

### 4.2 `sanitizeAiResponse` corrupts legitimate document text (F6)

`engine/oxml-engine.js:353-360` runs unconditionally on `modifiedText` for every
`applyRedlineToOxml` call, with no opt-out. Reproduced on realistic contract text:

| Input | Output |
|---|---|
| `The rate is $X$ per unit as defined in Schedule A.` | `The rate is X per unit...` — `$` delimiters eaten |
| `Costs range from $ten thousand$ upward.` | `Costs range from ten thousand upward.` |
| `Escape sequences such as \n and \r\n must be preserved literally.` | literal `\n` converted to real newlines — **changes paragraph structure** |
| `Here is the text: this clause is part of the actual contract body.` | `this clause is part of the actual contract body.` — sentence truncated |

These are LaTeX/chat-response heuristics applied to what is, by the library's own
contract, *document content*. The `\n` case is the most damaging: it silently
splits one paragraph into several.

**Fix:**

- Add `options.sanitizeInput`, defaulting to **`false`** (do not touch caller
  content). This is a behavior change; call it out in the changelog and bump the
  minor version. Hosts that genuinely feed raw LLM output opt in with `true`.
- If a fully backward-compatible landing is required, default to `true` for one
  minor release while emitting `warnings: ['input was sanitized; pass sanitizeInput:false to disable']`
  whenever sanitization actually altered the text, then flip the default.
  Either way the caller must be able to find out that their text was rewritten.
- Narrow the transforms themselves regardless of the default: drop the `$...$`
  and `\n`-unescaping rules entirely (they are unsafe on document text), and
  anchor prefix stripping to a full leading line rather than a sentence prefix.
- Test each row of the table above as a regression case.

**Acceptance for Phase 4:** no-op under `'accept-all-first'` preserves input;
sanitization is opt-in (or loudly warned); all six corruption samples covered by
tests; README documents `sanitizeInput`.

---

## Phase 5 — Remove process-global mutable state

**Why (F7):** `revisionIdCounter` in `core/types.js:161` is a module-level
counter that `seedRevisionIdsFromDocument` only ever raises, never resets or
scopes. Reproduced:

```
start id: 1000
after seeding from a doc containing w:id="2147483000"  -> next id 2147483001
next id for a COMPLETELY UNRELATED clean document      -> 2147483002
ids remaining before int32 overflow: 644
```

`w:id` is `ST_DecimalNumber`, so a large value is schema-legal; a single hostile
or merely unusual document permanently poisons every later document processed by
that Node process. Past int32 the ids Word receives are out of its practical
range. In a long-running server (the primary deployment shape for this package)
this is a slow-burning corruption source with no signal.

`seedRevisionIdsFromDocument` also scans **every** attribute with local name
`id` — including `w:bookmarkStart/@w:id`, `w:comment/@w:id`, and any `r:id` that
happens to parse as an integer — so it inherits the maximum of an unrelated id
space.

`adapters/config.js` (`_defaultAuthor`, `_platform`) and `adapters/xml-adapter.js`
(`_DOMParser`, `_XMLSerializer`) are process-globals too. That is acceptable for
providers, but a multi-tenant server cannot safely vary the default author per
request.

### 5.1 Scope revision ids to a document

- Add an internal `RevisionIdAllocator` (`core/types.js`) — a small object with
  `next()` seeded from one document — and create one per `applyRedlineToOxml` /
  per `applyOperationToDocumentXml` invocation.
- Thread it through the engine and builders. Keep `createRevisionMetadata(author)`
  working against a module-level default allocator so no public signature breaks;
  add `createRevisionMetadata(author, allocator)` as the internal path.
- Narrow seeding to the id spaces that actually matter: revision-bearing elements
  (`w:ins`, `w:del`, `w:moveFrom`, `w:moveTo`, `w:rPrChange`, `w:pPrChange`,
  `w:cellIns`, `w:cellDel`) plus `w:comment`. Do not seed from `r:id` or
  bookmark ids.
- Clamp: if a seeded value would push allocation within a safety margin of
  `2^31 - 1`, restart low and rely on per-document uniqueness (already asserted
  by `assertUniqueRevisionIds`) instead of global monotonicity.
- Test: process document A with `w:id="2147483000"`, then document B (clean);
  assert B's generated ids are small and `assertUniqueRevisionIds` passes on both.

### 5.2 Per-call author override

`options.author` already exists and takes precedence; verify no code path reads
`getDefaultAuthor()` after a caller supplied `options.author`. Add a test that
two interleaved `await`ed `applyRedlineToOxml` calls with different `author`
values produce correctly-attributed revisions — the engine is `async`, so
interleaving is real.

**Acceptance for Phase 5:** id allocation is per-document; a hostile id in one
document does not affect the next; interleaved concurrent calls attribute
authors correctly.

---

## Phase 6 — Batch operation atomicity (F10)

> **Production API: BREAKING.** Making `atomic` default to `true` changes the
> established batch result on failure from a partially applied document to the
> untouched original. Existing callers that intentionally consume partial
> results must pass `atomic: false` or migrate their handling.

`applyOperationsToDocumentXml` (`services/standalone-operation-runner.js:1289-1348`)
applies operations sequentially onto a running `currentDocumentXml`. On failure
it records an error entry and — unless `options.continueOnError === false` —
keeps going. The returned `documentXml` is therefore a **partially applied
batch**, returned with `hasChanges: true`, and the caller has to reconstruct what
landed by walking `results`.

For the package's primary full-document entry point, "half the edits applied" is
usually worse than "none applied": the document no longer matches either the
original or the intended outcome, and the operation list is not replayable
because earlier edits have moved the anchors.

- Add `options.atomic` (default **`true`**): on any operation error, return the
  **original** `documentXml`, `hasChanges: false`, and the full `results` array
  describing what would have applied. Callers wanting today's behavior pass
  `atomic: false`.
- Make the current `continueOnError` semantics explicit in the JSDoc — right now
  the default (`undefined`) means "continue," which reads backwards.
- Also audit `context.targetRefSnapshot` (`:1296-1299`): it is built once from
  the pre-mutation document and reused for every operation. That is deliberate
  (comment anchors are resolved before replacements), but it means a replacement
  operation late in the batch resolves anchors against a document state that no
  longer exists. Add a test with two overlapping replacements targeting adjacent
  text and assert the second either applies correctly or reports
  `TARGET_NOT_FOUND` — never silently edits the wrong span.
- Tests in `tests/standalone_operation_runner_tests.mjs`: 5-operation batch with
  op 3 failing → `atomic: true` returns the untouched original; `atomic: false`
  reproduces today's partial result.

**Acceptance for Phase 6:** atomic batches by default; overlapping-anchor case
covered; `AGENTS.md` batch section documents the new default.

---

## Phase 7 — Make the tooling checks real

> **Production API: NOT BREAKING.** This phase changes repository validation,
> type fixtures, linting, and coverage visibility. It does not change runtime
> behavior or the public API contract.

### 7.1 `check:types` does not type-check (F9)

`scripts/check-types.mjs` greps `index.d.ts` for seven fixed substrings and
counts curly braces. It never invokes `tsc`. Consequence, measured:

```
runtime exports:            97
typed in index.d.ts:        51
exported but NOT typed:     63   (ContainerKind, DiffOp, NS_W, escapeXml,
                                  buildReconciliationPlan, getParagraphText, ...)
npm run check:types:        PASS
```

TypeScript consumers get `any`/implicit errors on two thirds of the surface, and
declaration drift is structurally undetectable.

- Add `typescript` as a devDependency and run a real `tsc --noEmit` over
  `index.d.ts` plus a small `tests/types/usage.ts` fixture that exercises the
  documented shapes (options bag, result with `status`/`error`, accept/reject
  results, config functions).
- Add a **completeness check**: import `index.js`, diff `Object.keys` against the
  declarations, and fail on any untyped export. Land it with an explicit
  allowlist of the 63 current gaps so CI goes green immediately, then burn the
  allowlist down — new exports are typed from day one.
- Wire both into the existing `check:types` script so `.github/workflows/ci.yml`
  picks them up with no workflow change.

### 7.2 Add a linter

There is no ESLint/Prettier/tsconfig in the repo. For a 14,355-line library
whose failure mode is silent wrong output, the highest-value rules are the ones
that catch the bug classes this plan documents:

- `no-unused-vars`, `no-undef`, `require-atomic-updates` (would flag the F5
  reassigned-parameter-captured-in-closure pattern).
- A `no-restricted-syntax` rule banning `document.createElement('w:...')` and
  `createElementNS(NS_W, ...)` outside `core/word-xml.js`, and bare
  `parseFromString` outside `adapters/xml-adapter.js` — this converts two
  hand-audited conventions from the previous plan (its "Final verification"
  section audited these manually) into enforced ones.
- `no-empty` with `allowEmptyCatch: false` — the silent `catch {}` in
  `services/numbering-helpers.js:75` swallows malformed numbering XML.

Add `npm run lint` and a CI step. Do not add Prettier or reformat the codebase in
the same change.

### 7.3 Coverage visibility

Source is 14,355 lines against 4,976 lines of test, and the largest module
(`services/standalone-operation-runner.js`, 1,348 lines) and
`orchestration/list-structural-fallback.js` (530 lines) are among the least
directly covered. Add `node --experimental-test-coverage` (or `c8`) as
`npm run test:coverage`, report per-file, and record a baseline in
`docs/VALIDATION.md`. Do not gate CI on a threshold yet — get the number visible
first, and use it to aim Phase 1.3's corpus work.

**Acceptance for Phase 7:** `npm run check:types` runs `tsc` and fails on drift;
`npm run lint` passes and is in CI; coverage baseline recorded.

---

## Phase 8 — Reconstruction structural correctness (found by Phase 1)

**Status: Complete (2026-08-29).** Reconstruction now restores content at its
original container position, preserves `w:br`/`w:cr` as structural sentinels,
and scopes whole-paragraph reconstruction to the contiguous range named by the
caller. Partial paragraph targets return `PARTIAL_TARGET`. Runtime validation
now rejects nested paragraphs and misplaced or duplicate body-level `w:sectPr`.
The 12,000-case sweep passes with no Phase 8 suppression; its remaining 45
classified cases are the Phase 2 leading-whitespace gap.

> **Production API: CONDITIONAL — production behavior.** These fixes preserve
> the API shape but change results for malformed or mis-scoped inputs: invalid
> nested paragraphs/section placement are corrected, and partial targeting may
> return `PARTIAL_TARGET` instead of silently reporting success.

**Priority: high.** F13 and F14 produce output that is not valid
WordprocessingML. Word reports such files as corrupt, and F14 is rejected by
this package's own plumbing validator, so it is not a theoretical concern.
Despite the number, schedule this alongside Phase 2 rather than last.

All three defects live in the same code path: reconstruction mode rebuilds
paragraphs into a fragment, removes the originals, and re-inserts the fragment
(`engine/reconstruction-writer.js:134-156`).

### 8.1 `w:sectPr` is moved to the front of `w:body` (F14)

```js
paragraphs.forEach(paragraph => {
    if (paragraph.parentNode) paragraph.parentNode.removeChild(paragraph);
});
// ...
target.appendChild(fragment);   // lands AFTER whatever is left in the body
```

Removing the target paragraphs and then *appending* means the rebuilt content is
placed after every remaining sibling — including `w:sectPr`, which is normally
the body's last child. Result: `<w:body><w:sectPr/><w:p>…</w:p></w:body>`.
`CT_Body` requires section properties last, and
`services/standalone-docx-plumbing.js:361-362` throws
`Validation failed: w:sectPr not last` on exactly this shape.

Measured rate: **10,000 of 12,000 fuzz cases** — every generated case that
contains a `w:sectPr`.

Fix: capture the position of the first removed paragraph (its `nextSibling`, or
its index among the container's children) *before* removal, and `insertBefore`
the rebuilt fragment at that position instead of appending. That also fixes
ordering against any other trailing sibling, not just `w:sectPr`. Reuse
`insertBodyElementBeforeSectPr` from `services/standalone-docx-plumbing.js` if it
fits, rather than adding a second placement rule.

Then delete the `sectPr-not-last` entry from `KNOWN_GAPS` in
`tests/roundtrip_fuzz_tests.mjs`, un-skip the corpus case, and remove the two
`{ knownGaps: ['sectPr-not-last'] }` opt-outs in
`tests/paragraph_mark_revision_tests.mjs` (grep `KNOWN-GAP` to confirm none
remain).

### 8.2 `w:br` round-trips into a nested `w:p` (F13)

Input `first line<w:br/>second line`, editing only the second line, produces:

```xml
<w:p><w:r><w:t>first line</w:t></w:r>
  <w:p>…second row…</w:p>          <!-- a paragraph INSIDE a paragraph -->
</w:p>
```

The `w:br` is gone (0 in output) and `CT_P` has no paragraph child, so the file
is corrupt. Root cause is an ambiguity, not a typo: ingestion maps `w:br` to
`'\n'` (`pipeline/ingestion-export.js:88`, `pipeline/ingestion-paragraph.js:326`),
and reconstruction reads `'\n'` back as a *paragraph* boundary. The two meanings
of `'\n'` are not distinguished anywhere.

Fix requires picking a representation and applying it consistently:

- Preferred: keep `w:br` out of the text stream entirely — treat a break-bearing
  run as inert scaffolding the way footnote references are handled in
  `engine/surgical-spans.js`, so edits anchor around it and it is neither split
  nor recreated.
- Alternative: give `w:br` a distinct sentinel character in the text model and
  teach the reconstruction writer to emit `w:br` for it and a paragraph break
  only for a real `'\n'`.

Either way, add `assertNoNestedParagraphs` coverage (already available from
Phase 1) and un-skip the `w:br survives an edit in an adjacent run` corpus case.

### 8.3 Untargeted paragraphs are silently deleted (F15)

```js
// body has three paragraphs; original names only the first
await applyRedlineToOxml(threeParagraphDoc, 'alpha beta gamma', 'alpha beta delta')
// -> status: 'ok', and paragraphs two and three are GONE
```

`AGENTS.md` gotcha 3 already warns that paragraph APIs are not always safe on a
full `word/document.xml`, so a caller doing this is misusing the API — but
`status: 'ok'` while deleting two paragraphs is the wrong failure mode for a
misuse, and nothing in the result distinguishes it from a clean edit.

Fix: when reconstruction is about to rebuild a container, compare the paragraphs
it matched against the paragraphs present. If it did not match all of them,
either scope the rewrite to the matched paragraphs (preferred) or return
`status: 'error'` with a new `PARTIAL_TARGET` code (Phase 3's contract). Do not
leave a path that silently drops content and reports success.

Add a fuzz shape for it: multi-paragraph body, `original` naming exactly one
paragraph, asserting the other paragraphs survive byte-identical.

**Acceptance for Phase 8:** all three `KNOWN-GAP` suppressions removed;
`grep -rn "KNOWN-GAP" tests/` returns nothing for Phase 8; fuzz sweep of ≥12,000
cases green with no known-gap counts; the Phase 1 follow-up port of
`assertNoNestedParagraphs` / `assertSectPrLast` into `core/redline-validation.js`
done as part of this phase.

---

## Execution order and dependencies

```
Phase 1  (honest oracle)        ← DONE; every later phase is verified through it
Phase 8  (reconstruction shape) ← needs 1; schedule next, output is schema-invalid today
Phase 2  (diff correctness)     ← needs 1 for the whitespace-exact assertions
Phase 3  (error contract)       ← needs 2 (introduces DIFF_TOKEN_LIMIT)
Phase 4  (silent mutation)      ← needs 1; independent of 2/3
Phase 5  (global state)         ← independent
Phase 6  (batch atomicity)      ← independent
Phase 7  (tooling)              ← 7.1/7.2 anytime; 7.3 last
```

Highest value per unit of work, if the whole plan cannot be taken on:
**Phase 8.1** (one-line-ish placement fix, clears a defect affecting essentially
all document-scoped reconstruction output), **Phase 2.1b** (leading whitespace,
also small), **Phase 2.1** (silent text loss at scale), then **Phase 4.1**
(silent revision loss).

Each phase is a separate commit with the suite green. Suggested messages:

- `test: compare round-trip text losslessly and widen the fuzz corpus` *(done)*
- `fix: keep w:sectPr last and stop emitting nested paragraphs`
- `fix: guard and widen diff token space; make diff output deterministic`
- `feat: return structured errors instead of throwing on malformed OOXML`
- `fix: stop discarding existing revisions and mutating caller text`
- `refactor: scope revision id allocation to a single document`
- `feat: make batched document operations atomic by default`
- `build: type-check index.d.ts for real, add lint and coverage`

---

## Global guardrails (carried forward, plus new)

From the previous plan — still binding:

- Never emit `w:t` inside `w:del`; never nest `w:ins`/`w:del`; never emit empty
  `w:t`/`w:delText` or empty revision wrappers.
- Source author strings through `createRevisionMetadata`; never trim or transform
  them at call sites.
- Use `RPR_SCHEMA_ORDER` (`engine/rpr-helpers.js`) for `w:rPr` ordering; `w:pPr`
  is the first child of `w:p` and `w:rPr` is its last child.

New, from this plan:

- **A function must never silently drop content it cannot map.** `charsToWords`
  discarding out-of-range tokens is the archetype: if a value cannot be handled,
  raise or report it.
- **`hasChanges: false` must guarantee the returned payload is byte-identical to
  the input.** If the library normalized anything, that is a change and must be
  reported as one.
- **The library must not rewrite caller-supplied document text** outside of the
  explicitly-requested edit. Heuristic cleanup is opt-in.
- **Never use a lossy reader as a test oracle.** Verification helpers compare
  exact bytes/characters; normalization in an assertion needs a written
  justification at the call site.
- **Wall-clock time must not influence output.** Any timeout that changes results
  rather than just aborting is a determinism bug.

---

## Verification commands

```bash
npm test                          # 24/24 as of Phase 1
npm run test:isolation
npm run check:types
npm run lint                      # added in Phase 7.2
npm run test:coverage             # added in Phase 7.3
node scripts/export-validation-fixtures.mjs
FUZZ_SEED=1 FUZZ_ITERATIONS=5000 node tests/roundtrip_fuzz_tests.mjs
grep -rn "KNOWN-GAP" tests/       # every suppression names the phase that owns it
```

The fuzz run prints a per-shape breakdown and a line per known gap, e.g.:

```
PASS: 12000 fuzz round-trip cases (base seed 20260704)
  [paragraph=4000 multiParagraph=2000 tableCell=2000 whitespace=2000 existingRevisions=2000]
  KNOWN-GAP sectPr-not-last (Phase 8): 10000 case(s)
  KNOWN-GAP leading-whitespace-dropped (Phase 2): 12 case(s)
```

A gap reporting `0 cases -- possibly fixed` means the entry is stale: confirm
and delete it.

Release-time, unchanged from `docs/VALIDATION.md`:

```bash
npm run smoke:word:diff           # Windows + desktop Word differential
bash scripts/validate-fixtures-xsd.sh
soffice --headless --convert-to pdf tmp/validation-docx/*.docx
```
