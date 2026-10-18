# Hand-off for the Word add-in (AIWordPlugin): docx-redline-js 0.10.0

Audience: an AI agent or developer working on the Word add-in that consumes `@ansonlai/docx-redline-js`.
Source of truth: `CHANGELOG.md` "0.10.0" and `docs/releases/0.10.0.md` in this repo.

## Scope decision

0.10.0 does **not** ship the proposed `@ansonlai/docx-redline-js/word` subpath. The library stays host-independent
(see `ARCHITECTURE.md`, "No Word add-in entrypoints or host-specific integration layers"). The add-in keeps its Word Flat
OPC layer (`consumer-core.js`, `word-source-identity.js`): reading `body.getOoxml()`, building the single `insertOoxml`
package, and the Word-aware revision tokens.

What the add-in's Word layer imports today:

| Import | Public? | Notes |
|---|---|---|
| `applyOperationsToDocumentXml` (`/standalone-runner`) | Yes | Documented subpath. |
| `inspectDocumentParts`, `computeDocumentPartsRevisionToken`, `areRevisionTokensEqual`, `mergeNumberingXmlBySchemaOrder`, `ensureComments*ArtifactsInZip`, `extractReplacementNodesFromOoxml`, `normalizeBodySectionOrderStandalone`, `getDefaultAuthor`, `NS_W` | Yes | Root exports. |
| `parseXml` / `serializeXml` / `createParser` / `createSerializer` (`adapters/xml-adapter.js`) | Deep import | The root exports `parseOoxml`, `serializeOoxml` and `parseOoxmlSafe`; prefer those. |
| `reconcileCommentSiblingParts` (`services/comment-thread-parts.js`), `MemoryZip` (`document/zip-archive.js`) | Deep import | The real gap: merging comment parts (including `commentsIds` / `commentsExtensible`) outside a ZIP. A parts-level comment merge is a candidate for a later release; until then these deep imports are unversioned internals. |
| `getPartSpec` (`services/package-parts.js`), `buildDocumentFragmentPackage` (`services/package-builder.js`) | Deep import | Not needed in production: hard-code the part constants; `getOoxml` always returns a full package. Test-only use is fine. |

Deep imports resolve through the package's wildcard exports, but they are not covered by semver. Pin the exact release.

## What 0.10.0 gives the add-in

1. **`existingRevisions: 'auto'`** (capability `existing-revisions-auto-v1`, contract stays 9). Per target:
   `merge-same-author` when the target has no tracked changes or only the operation author's, otherwise
   `slice-cross-author`. Unattributed changes count as another author's, and a `targetEnd` range resolves over every
   paragraph in it. Preflight reports the resolved policy. Replaces the add-in's own per-paragraph policy choice (and its
   `TRACKING_POLICY_MISMATCH` guard). Rejected-view insertion still needs explicit `slice-cross-author`.
2. **`compareDocumentInspections(expected, actual)`**: inspect the prepared package and what Word committed, then compare.
   Ignores ids, dates, rsids, paragraph ids, comment ids and order, and run boundaries. `{ matches, differences }`, at
   most 10 differences.
3. **Signature lines survive.** Underscore runs (`_____`) no longer collapse to `_` after a table reconciliation or a
   plain-text rewrite, and a caller's `structuredContent: false` keeps every inline Markdown marker literal. If the add-in
   relied on `structuredContent: false` to apply inline bold or italic, send `true`.
4. **Comment reply paraIds.** Pass `documentXml` to `applyCommentReplyToParts` (the `comment_reply` operation already
   does) and the reply's `w14:paraId` avoids body paragraph ids. Before 0.10.0 body ids were never checked.
5. **Unattributed revisions are another author's.** A target with the add-in author's changes plus a change without
   `w:author` now gets `EXISTING_REVISIONS` under `merge-same-author` (it was `UNSAFE_REVISION_NESTING` with
   `manual_resolution`), and `slice-cross-author` / `auto` edit around it. If the add-in maps
   `UNSAFE_REVISION_NESTING` to a manual-resolution message for this case, it will now see the ordinary
   `EXISTING_REVISIONS` recovery instead.

## Upgrade steps

1. Pin 0.10.0. Do not import `@ansonlai/docx-redline-js/word`; it does not exist.
2. Require `existing-revisions-auto-v1` if you send `auto`.
3. Switch Flat OPC parsing to `parseOoxmlSafe` / `serializeOoxml` where the parse-error behavior matches.
4. Keep the post-commit `compareDocumentInspections` check; treat a mismatch as "re-inspect and report", not "retry".
