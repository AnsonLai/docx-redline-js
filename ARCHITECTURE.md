# Reconciliation Core Architecture

This document describes the standalone, host-agnostic OOXML reconciliation package and how to work inside it safely.

## Scope

This repository contains only the publishable package surface:

- `adapters/`
- `core/`
- `engine/`
- `pipeline/`
- `services/`
- `orchestration/`
- `index.js`
- `standalone.js`

No Word add-in entrypoints or host-specific integration layers are part of this package.

## Goals

- Preserve Word-compatible redlines by editing OOXML directly.
- Keep core logic host-independent (no Office.js globals, no Word API calls).
- Reuse the same engine in browser, Node.js, and other JavaScript runtimes.

## Folder Layout

```text
.
├── adapters/
│   ├── config.js
│   ├── logger.js
│   └── xml-adapter.js
├── core/
├── engine/
│   └── formatting-removal.js
├── orchestration/
├── pipeline/
├── services/
│   ├── numbering-helpers.js
│   ├── revision-comment-management.js
│   ├── standalone-docx-plumbing.js
│   └── standalone-operation-runner.js
└── index.js
```

## Entry Points

- `index.js` (Root exports containing the reconciliation logic)

## Module Responsibilities

- `adapters/config.js`
  - Runtime configuration for defaults (`setDefaultAuthor`, `getDefaultAuthor`, `setPlatform`, `getPlatform`).
- `adapters/xml-adapter.js`
  - XML parser/serializer injection for browser or Node.js runtimes.
- `adapters/logger.js`
  - Runtime logger injection and shared logging methods.
- `core/*`
  - Shared types, OOXML identity helpers, target resolution, list/table targeting heuristics, and XML query helpers.
- `engine/oxml-engine.js`
  - Main reconciliation router and mode selection.
- `engine/formatting-removal.js`
  - Shared formatting removal and highlight helpers.
- `pipeline/*`
  - Ingestion, markdown preprocessing, diffing, patching, and serialization stages.
- `services/numbering-helpers.js`
  - Dynamic numbering ID allocation, numbering payload remapping, and schema-order-safe numbering merges.
- `services/standalone-docx-plumbing.js`
  - Package-level extraction/wiring/validation for `word/document.xml`, `word/numbering.xml`, and `word/comments.xml`.
- `services/revision-comment-management.js`
  - OOXML transforms for accepting/rejecting tracked changes by author/all-authors and deleting comments by author/all-authors.
- `services/standalone-operation-runner.js`
  - Host-agnostic operation bridge for `redline`, `highlight`, and `comment` workflows.
- `orchestration/*`
  - Route planning and list fallback orchestration utilities.

## End-to-End Flow

1. Caller imports from `index.js`.
2. Caller configures XML provider/logger/defaults when needed via `adapters/*`.
3. Caller invokes reconciliation APIs (`applyRedlineToOxml`, operation runner, ingestion/export helpers).
4. `engine/oxml-engine.js` routes to format, table, list, surgical, or reconstruction flows.
5. Pipeline/services return OOXML and optional package artifacts (`numberingXml`, comments payloads).
6. Optional revision/comment management transforms can accept/reject revisions or delete comments by author.
7. Caller writes resulting XML back to package/document boundaries.

## Public Surfaces

- Primary: `index.js`
Keep exports centralized through `index.js`.


## Build Output

`npm run build` generates CDN-ready ESM bundles under `dist/`:

- `dist/docx-redline-js.esm.js`
- `dist/docx-redline-js.esm.js.map`
- `dist/docx-redline-js.esm.min.js`
- `dist/docx-redline-js.esm.min.js.map`

The bundle inlines `diff-match-patch` and keeps `@xmldom/xmldom` external.

## Testing

- `npm test`
  - Runs the package test runner (`scripts/run-tests.mjs`) against all `tests/*.mjs` except setup helpers.
- `npm run test:isolation`
  - Runs boundary checks for Word API markers and dependency-graph isolation.

Use these checks before publishing or tagging.

## Fast Orientation For Contributors

Use this sequence to understand or modify behavior without reading everything:

1. Start at `index.js` to locate the exported API.
2. Follow exports into `engine/oxml-engine.js` or relevant `services/*` module.
3. For targeting bugs, inspect `core/paragraph-targeting.js`, `core/list-targeting.js`, and `core/table-targeting.js`.
4. For package wiring issues, inspect `services/standalone-docx-plumbing.js`.
5. For revision/comment cleanup behavior, inspect `services/revision-comment-management.js`.
6. For numbering/list issues, inspect `services/numbering-helpers.js` and orchestration list-fallback modules.
