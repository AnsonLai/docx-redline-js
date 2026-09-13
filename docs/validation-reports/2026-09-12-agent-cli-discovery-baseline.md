# Agent CLI Discovery Baseline

**Date:** 2026-09-12

**Scope:** WP-00 baseline for command discovery, inspection breadth, contextual
search, and compact mutation output

## Before-state

Measurements were captured from the source CLI immediately before WP-01/WP-02,
using `tests/fixtures/sample_doc_test.docx` and pretty-printed CLI JSON:

| Case | Result |
|---|---:|
| Global help | 171 bytes |
| `apply --help` | 171 bytes; identical to global help |
| `inspect --non-empty` | 72,387 bytes / 63 paragraphs |
| Broad `extract --search "the"` | 13,968 bytes / 31 paragraphs |
| Context retrieval | Search plus range: 2 commands |
| Successful one-operation receipt | 464 bytes nested plus the same 464 bytes at root |

The help payload did not describe command flags, search case behavior, or
redline/comment/restore shapes. The detailed unscoped inspection exceeded the
observed 64 KiB harness ceiling. Receipt deduplication is recorded here but is
owned by WP-04, not WP-00 through WP-02.

## WP-01/WP-02 result

The checked post-change run produced:

| Case | Result |
|---|---:|
| Global command-index help | 1,390 bytes |
| Command-specific apply help | 4,400 bytes with flags and three operation shapes |
| Bounded `inspect --non-empty` | 25,093 bytes / first 20 of 63 paragraphs |
| Bounded `extract --search "the"` | 12,855 bytes / first 20 of 31 matches |
| `extract --search "agreement" --around 2 --limit 1` | 2,178 bytes / one command / three paragraphs |

Both formerly broad responses remain valid JSON below the 48 KiB soft limit and
include stable continuation metadata. Search case variants returned identical
indexes. A paragraph that individually exceeds the soft limit is returned whole
and marked `oversizeItem` rather than truncating `exactText`.

## Reproduction

Run:

```bash
npm run benchmark:agent-cli
```

The benchmark retains the observed before-state constants and records the active
post-change results in `tmp/benchmarks/agent-cli-discovery-latest.json`. It
measures CLI response bytes, calls, paragraphs, selection metadata, and native
elapsed time. It does not estimate model tokens, provider wall time, or tool
transport latency.
