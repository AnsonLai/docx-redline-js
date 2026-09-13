import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';

function stats(text) {
    return {
        lines: text.split(/\r?\n/).length,
        words: text.trim().split(/\s+/).filter(Boolean).length
    };
}

const fastStart = await readFile(new URL('../docs/AGENT_FAST_START.md', import.meta.url), 'utf8');
const launchCard = await readFile(new URL('../AGENTS.md', import.meta.url), 'utf8');
const readme = await readFile(new URL('../README.md', import.meta.url), 'utf8');
const packageJson = JSON.parse(await readFile(new URL('../package.json', import.meta.url), 'utf8'));

const fastStartStats = stats(fastStart);
assert(fastStartStats.lines >= 40 && fastStartStats.lines <= 60, JSON.stringify(fastStartStats));
assert(fastStartStats.words <= 600, JSON.stringify(fastStartStats));
for (const contract of [
    '--operations -',
    '--profile agent',
    'completion: true',
    'error.recovery.action',
    'retryPlan.base: "original"',
    'slice-cross-author'
]) {
    assert(fastStart.includes(contract), `fast start omitted ${contract}`);
}

const launchCardStats = stats(launchCard);
assert(launchCardStats.words < 800, JSON.stringify(launchCardStats));
for (const route of [
    'docs/AGENT_FAST_START.md',
    'docs/AGENT_KNOWLEDGE_BASE.md',
    'ARCHITECTURE.md',
    'docs/TESTING.md',
    'docs/schemas/document-operations.schema.json'
]) {
    assert(launchCard.includes(route), `launch card omitted ${route}`);
}

assert(readme.includes('[docs/AGENT_FAST_START.md](./docs/AGENT_FAST_START.md)'));
for (const publishedDoc of [
    'docs/AGENT_FAST_START.md',
    'docs/AGENT_KNOWLEDGE_BASE.md',
    'docs/TESTING.md',
    'docs/schemas/document-operations.schema.json',
    'docs/validation-reports/2026-09-12-agent-protocol-rollout.md'
]) {
    assert(packageJson.files.includes(publishedDoc), `package omitted ${publishedDoc}`);
}
assert.equal(packageJson.files.includes('docs/'), false);
assert.equal(packageJson.files.some(entry => entry === 'examples/' || entry.startsWith('examples/')), false);
assert(packageJson.files.includes('!scripts/benchmark-agent-workflow.mjs'));

console.log('agent documentation contract tests passed');
