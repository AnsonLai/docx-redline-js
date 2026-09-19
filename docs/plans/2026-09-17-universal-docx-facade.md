# Universal Document Facade & Zero-Node Architecture Plan

**Status:** In Progress (WP-00 through WP-03 completed; ready for WP-04)  
**Date:** 2026-09-17 (Updated 2026-09-19)  
**Target:** 0.8.0  
**Priority:** Eliminate runtime Node dependencies from complete-DOCX operations (`openDocx`) by adopting cross-runtime standards (`Uint8Array`, `fflate`, pure-JS DOM fallback, pure-JS SHA-256), enabling zero-dependency universal bundles that run in sandboxes (n8n, Cloudflare Workers, Deno, Bun, and browser runtimes).

---

## 1. Executive Summary & Problem Statement

### 1.1 Context
`@ansonlai/docx-redline-js` was originally designed with a host-independent core (`core/`, `engine/`, `pipeline/`, `services/`) that reconciles OOXML strings and DOM elements. However, handling complete `.docx` files was delegated to a separate Node-only facade in `node/` (`DocxDocument`, `openDocx`, `zip-archive.js`, `cli.js`).

### 1.2 The Problem
In real-world workflow automation platforms (e.g., **n8n Code nodes**, **Cloudflare Workers**, **AWS Lambda**, **browser extensions**, and **Deno/Bun** environments):
1. **Node Built-in Restrictions:** `node/zip-archive.js` imports `node:zlib` (`deflateRawSync`, `inflateRawSync`), and `node/docx-document.js` imports `node:crypto` (`createHash`). These fail in browser runtimes, edge workers, and restricted sandboxes like n8n Cloud where built-in Node modules are disabled.
2. **Peer Dependency Friction:** `@xmldom/xmldom` is configured as an external `peerDependency`. Non-Node environments or sandboxes cannot easily `require` or link external peer packages.
3. **Binary Incompatibility:** The Node facade relies on Node's `Buffer` class rather than standard web-interoperable `Uint8Array`.
4. **No Self-Contained Distribution:** Users cannot take a single bundle and drop it into a workflow code block or sandbox without packaging a container or deploying a separate microservice.

### 1.3 The Goal
Transform the complete-document facade from a **Node-specific layer** into a **Universal Document Facade**:
* Replace `node:zlib` with `fflate` for lightweight, zero-dependency ZIP processing.
* Internalize/bundle a pure-JS XML DOM parser as a transparent fallback to `globalThis.DOMParser`.
* Replace `Buffer` with `Uint8Array` + `TextEncoder` / `TextDecoder`.
* Replace `node:crypto` with a synchronous, pure-JS SHA-256 implementation for package revision tokens.
* Provide an optional single-file, zero-external-dependency bundle suitable for sandboxes and copy-paste environments like n8n.

---

## 2. Technical Analysis & Rationale

### 2.1 The XML Parser: DOM Requirements vs. Universal Support

#### Why Not Replace the Parser with a Regex or AST Parser?
OOXML reconciliation is inherently namespace-aware and tree-mutation-heavy:
* The engine uses W3C DOM Level 2/3 APIs: `getElementsByTagNameNS()`, `createElementNS()`, `getAttributeNS()`, `insertBefore()`, `removeChild()`, and `cloneNode(true)`.
* Parsers like `fast-xml-parser` or `xml2js` convert XML into Plain Old JavaScript Objects (POJOs). They destroy node references, parent-child linkages, namespace URIs, and attribute ordering, which would require a total rewrite of the core reconciliation engine.

#### The Dual-Mode Universal Solution
1. **`globalThis.DOMParser` as Priority:** If running in a browser, Word Web Add-in, or runtime with standard web APIs, use the host's native `DOMParser` and `XMLSerializer`.
2. **Pure-JS Fallback:** When `globalThis.DOMParser` is missing (Node.js, edge workers, n8n sandboxes), fall back automatically to `@xmldom/xmldom`.
3. **No External Peer Requirement:** Bundle or bundle-inline `@xmldom/xmldom` in standalone distributions. Because `@xmldom/xmldom` is 100% pure JavaScript with zero native C++ or Node-specific dependencies, it runs identically in any JS context.

### 2.2 ZIP & Compression: Evaluating Alternatives

The repository's custom ZIP implementation in `node/zip-archive.js` is already a clean, 53-line pure-JS ZIP container reader and writer. It only uses `node:zlib` for raw DEFLATE compression and INFLATE decompression.

| Option | Size (Gzip) | Sync Support | Pure JS? | Trade-offs & Evaluation |
| :--- | :---: | :---: | :---: | :--- |
| **`node:zlib`** (current) | 0 KB | Yes | No (Node builtin) | Fails outside standard Node.js; blocked in n8n Cloud and edge workers. |
| **`pako`** | ~45 KB | Yes | Yes | Reliable zlib port, but significantly heavier than `fflate`. |
| **`jszip`** | ~35 KB | Mostly async | Yes | Already in `devDependencies`, but async-heavy; adds needless wrapper abstractions over existing custom ZIP reader. |
| **`fflate`** (Selected) | **~8 KB** | **Yes** (`zlibSync`, `unzlibSync` / `deflateSync`, `inflateSync`) | **Yes** | **Fastest, smallest, zero dependencies.** Supports synchronous raw DEFLATE/INFLATE matching the exact needs of `zip-archive.js`. |

**Decision:** Adopt `fflate`. It drops directly into the existing ZIP archive parser/serializer with virtually zero overhead.

### 2.3 Binary Handling: `Buffer` to `Uint8Array`

* Node.js `Buffer` is a subclass of `Uint8Array`. Any code accepting `Uint8Array` seamlessly accepts `Buffer` without conversion.
* Standard `TextEncoder` and `TextDecoder` are part of `globalThis` in all modern JavaScript runtimes (Node.js >= 18, Deno, Bun, browsers, Workers).
* Replace `buffer.toString('utf8')` with `new TextDecoder().decode(bytes)`.
* Replace `Buffer.from(str, 'utf8')` with `new TextEncoder().encode(str)`.
* Keep `.toBuffer()` as a backward-compatibility convenience method on `DocxDocument` when running in Node, but make `.toUint8Array()` the canonical cross-platform representation.

### 2.4 Document Revision Hashing: Pure-JS SHA-256

* In `node/docx-document.js`, `createHash('sha256')` from `node:crypto` generates deterministic package revision tokens.
* Web Crypto (`globalThis.crypto.subtle.digest`) is asynchronous, which would force `computePackageRevisionToken` and `doc.getRevisionToken()` into async functions, causing breaking API changes across the library and test suite.
* **Decision:** Include a tiny (~40-line), dependency-free, synchronous pure-JS SHA-256 function. This keeps revision tokens 100% synchronous, deterministic, and free from `node:crypto`.

---

## 3. Target Architecture & Module Layout

### 3.1 Module Organization

```text
index.js                      --> Universal entry point (exports core + openDocx + DocxDocument)
adapters/
  ├── xml-adapter.js          --> Auto-detects DOMParser, falls back to internal xmldom
  └── config.js
core/
  └── sha256.js               --> [NEW] Dependency-free synchronous SHA-256
document/                     --> [NEW] Universal document package layer (replaces node/)
  ├── docx-document.js        --> Universal DocxDocument & openDocx using Uint8Array
  └── zip-archive.js          --> Universal ZIP using fflate
node/
  ├── cli.js                  --> Node CLI (retains fs/process handling for bin/docx-redline)
  └── index.js                --> Compatibility re-exports for @ansonlai/docx-redline-js/node
dist/
  ├── docx-redline-js.esm.js  --> Universal ESM bundle
  ├── docx-redline.bundle.js  --> [NEW] Zero-dependency single-file bundle (CJS/IIFE for sandboxes)
```

### 3.2 Backward Compatibility
* Existing code importing `@ansonlai/docx-redline-js/node` will continue to work without breaking changes via re-exports from `node/index.js`.
* `openDocx` will also be exposed directly from the root package:
  ```javascript
  import { openDocx, applyRedlineToOxml } from '@ansonlai/docx-redline-js';
  ```
* Any existing code passing a Node `Buffer` to `openDocx(buffer)` will continue to work because `Buffer instanceof Uint8Array`.

---

## 4. Work Packages (Implementation Steps)

### WP-00: Capability & Baseline Audit [COMPLETED 2026-09-19]
* **Scope:** Audit runtime boundaries, isolation checks, and test suite baselines before changes.
* **Findings & Baseline State:**
  * Test suite baseline: 114 passed out of 114 tests (`npm test`).
  * Isolation check baseline: `test:isolation` passes (`no_word_api_index_check.mjs` and `core_dependency_graph_check.mjs`).
  * Boundary violations to address: `node/zip-archive.js` uses `node:zlib` and `Buffer`; `node/docx-document.js` uses `node:crypto` (`createHash`) and `Buffer`; `adapters/xml-adapter.js` requires explicit `configureXmlProvider` in Node.
  * Peer dependency friction: `@xmldom/xmldom` is configured as optional peer dependency, necessitating migration to `dependencies`.

### WP-01: Universal XML Adapter with Transparent Fallback [COMPLETED 2026-09-19]
* **Files:** `adapters/xml-adapter.js`, `package.json`
* **Changes:**
  * Updated `adapters/xml-adapter.js` with `resolveDomParserConstructor()` and `resolveXmlSerializerConstructor()`, automatically using `globalThis.DOMParser`/`globalThis.XMLSerializer` when available and falling back transparently to `@xmldom/xmldom`.
  * Moved `@xmldom/xmldom` from `peerDependencies`/`peerDependenciesMeta` to runtime `dependencies` (`^0.9.0`) in `package.json`.
  * Preserved full backwards compatibility for `configureXmlProvider({ DOMParser, XMLSerializer })`.
  * Verified parsing and serialization work out-of-the-box without requiring manual `configureXmlProvider` calls.

### WP-02: Pure-JS Synchronous SHA-256 [COMPLETED 2026-09-19]
* **Files:** `core/sha256.js`, `services/revision-token.js`, `node/docx-document.js`, `tests/revision_token_tests.mjs`
* **Changes:**
  * Implemented pure-JS synchronous SHA-256 (FIPS 180-4) in `core/sha256.js` with zero dependencies, verified across standard NIST test vectors against `node:crypto`.
  * Updated `services/revision-token.js` (`computeRevisionTokenSync` and `computeRevisionToken`) to use pure-JS `sha256` by default when `digestFn` is not provided.
  * Eliminated `node:crypto` (`createHash`) from `node/docx-document.js` completely.
  * Added unit test assertions confirming exact deterministic hash parity between pure-JS `sha256` and `node:crypto`.

### WP-03: Universal ZIP Implementation with `fflate` [COMPLETED 2026-09-19]
* **Files:** `document/zip-archive.js`, `node/zip-archive.js`, `package.json`, `tests/no_word_api_index_check.mjs`, `tests/universal_zip_archive_tests.mjs`
* **Changes:**
  * Added `fflate` (`^0.8.3`) to `dependencies` in `package.json`.
  * Added `./document/*` to `exports` and `document/` to `files` in `package.json`.
  * Added `fflate` to `allowedExternalImports` in `tests/no_word_api_index_check.mjs`.
  * Implemented `document/zip-archive.js` using `fflate.deflateSync`/`fflate.inflateSync`, `Uint8Array`, `DataView`, `TextEncoder`, and `TextDecoder` with zero Node built-in imports.
  * Refactored `node/zip-archive.js` to delegate to `document/zip-archive.js` while maintaining backward-compatible `Buffer` returns for existing Node callers.
  * Verified full round-trip ZIP creation and extraction, corrupt archive error handling, CRC32 calculations, and `MemoryZip` interface across Node and universal suites (`115 passed out of 115 tests`).

### WP-04: Universal `DocxDocument` & `openDocx` [COMPLETED 2026-09-19]
* **Files:** `document/docx-document.js`, `node/docx-document.js`, `tests/universal_docx_document_tests.mjs`
* **Changes:**
  * Ported complete document facade to `document/docx-document.js`, standardizing binary operations on `Uint8Array`, `TextEncoder`, and `TextDecoder`.
  * Added `.toUint8Array()` as canonical serialization method alongside `.toBuffer()` backwards compatibility helper.
  * Re-exported `DocxDocument`, `openDocx`, and `computePackageRevisionToken` in `node/docx-document.js` for existing `@ansonlai/docx-redline-js/node` consumers.
  * Added `tests/universal_docx_document_tests.mjs` verifying document opening from pure `Uint8Array`, inspection, operation application, revision token calculation, and serialization.

### WP-05: Main Entry Point & Node Compatibility Layer [COMPLETED 2026-09-19]
* **Files:** `index.js`, `index.d.ts`, `document/docx-document.d.ts`, `node/index.js`
* **Changes:**
  * Re-exported `openDocx`, `DocxDocument`, and `computePackageRevisionToken` directly from the root package entrypoint (`index.js`).
  * Maintained `node/index.js` and `node/docx-document.js` as compatibility re-exports to preserve non-breaking behavior for `@ansonlai/docx-redline-js/node`.
  * Updated TypeScript declarations in `index.d.ts` and `document/docx-document.d.ts`, fully passing `tsc` and `check:types` (126 exports verified).

### WP-06: Zero-Dependency Sandbox Bundle
* **Files:** `scripts/build.mjs`, `package.json`
* **Changes:**
  * Configure `esbuild` to produce a fully self-contained bundle (`dist/docx-redline.sandbox.js` or `dist/docx-redline.bundle.cjs`).
  * Bundle `fflate`, `diff-match-patch`, and `@xmldom/xmldom` with zero external dependencies and zero Node built-ins.
  * Verify that `grep -E "node:zlib|node:crypto|require\('fs'\)"` returns zero matches in the sandbox bundle.

### WP-07: Test Suite & Verification
* **Files:** `tests/*.mjs`, `scripts/check-types.mjs`
* **Changes:**
  * Run entire test suite (`npm test`).
  * Run boundary isolation tests (`npm run test:isolation`).
  * Add an automated test verifying execution in a simulated restricted sandbox (e.g. evaluating the bundle in a context without `process`, `Buffer`, `node:zlib`, or `node:crypto`).

---

## 5. Verification & Acceptance Criteria

1. **Zero Node Built-ins in Core & Document Modules:**
   * Running `rg "node:zlib|node:crypto" adapters/ core/ engine/ pipeline/ services/ document/` returns no matches.
2. **Full Test Suite Pass:**
   * All existing tests in `tests/` pass without modification to their expectations.
3. **Sandbox Readiness:**
   * A bundled build can be executed in an isolated VM / worker / n8n Code block with only standard JavaScript globals (`globalThis`, `Uint8Array`).
4. **Lightweight Distribution:**
   * Bundle size increase with `fflate` is under 15 KB gzipped.
5. **No Breaking Changes:**
   * Node callers using `openDocx(buffer)` and `doc.toBuffer()` retain full API parity.
