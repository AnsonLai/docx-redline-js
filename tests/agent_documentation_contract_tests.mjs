import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';

function stats(text) {
    return {
        lines: text.split(/\r?\n/).length,
        words: text.trim().split(/\s+/).filter(Boolean).length
    };
}

function firstBashCommandAfter(text, heading) {
    const start = text.indexOf(heading);
    assert(start >= 0, `missing section ${heading}`);
    const match = text.slice(start).match(/```bash\r?\n([\s\S]*?)```/);
    assert(match, `missing bash command after ${heading}`);
    return match[1].split(/\r?\n/).map(line => line.trim()).find(line => line && !line.startsWith('#'));
}

const fastStart = await readFile(new URL('../docs/AGENT_FAST_START.md', import.meta.url), 'utf8');
const launchCard = await readFile(new URL('../AGENTS.md', import.meta.url), 'utf8');
const readme = await readFile(new URL('../README.md', import.meta.url), 'utf8');
const knowledgeBase = await readFile(new URL('../docs/AGENT_KNOWLEDGE_BASE.md', import.meta.url), 'utf8');
const skillAuthoring = await readFile(new URL('../docs/SKILL_AUTHORING.md', import.meta.url), 'utf8');
const releaseNotes = await readFile(new URL('../docs/releases/0.7.0.md', import.meta.url), 'utf8');
const packageJson = JSON.parse(await readFile(new URL('../package.json', import.meta.url), 'utf8'));

const fastStartStats = stats(fastStart);
assert(fastStartStats.lines >= 40 && fastStartStats.lines <= 60, JSON.stringify(fastStartStats));
assert(fastStartStats.words <= 600, JSON.stringify(fastStartStats));
for (const contract of [
    '--operations -',
    '--profile agent',
    '--around 3',
    '--context-range 1:3',
    'selection.nextAfter',
    'humanReference',
    'docx-redline apply --help',
    'completion: true',
    'results[i].change.committed: true',
    'finalDisposition: "applied"',
    'acceptedViewMatchesCompiledText: true',
    'anchorMatchCount > 1',
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
    'docs/SKILL_AUTHORING.md',
    'ARCHITECTURE.md',
    'docs/TESTING.md',
    'docs/schemas/document-operations.schema.json'
]) {
    assert(launchCard.includes(route), `launch card omitted ${route}`);
}

assert(readme.includes('[docs/AGENT_FAST_START.md](./docs/AGENT_FAST_START.md)'));
assert(stats(skillAuthoring).words <= 1000, JSON.stringify(stats(skillAuthoring)));
for (const heading of [
    '## Safety invariants',
    '## Workflow policy choices',
    '## Transport choices',
    '## Recovery contract',
    '## Presentation rules'
]) {
    assert(skillAuthoring.includes(heading), `skill authoring omitted ${heading}`);
}
for (const capability of [
    'contract version 8',
    'agent-safety-profile-v2',
    'deduplicated-cli-receipts',
    'recovery-envelope-v1',
    'localized-replacements-v1',
    'speculative-search-apply-v1',
    'localized-change-summary-v1'
]) {
    assert(skillAuthoring.includes(capability), `skill authoring omitted ${capability}`);
}
for (const recoveryField of [
    'error.recovery.action',
    'error.recovery.requiresReinspection',
    'error.recovery.requiresUserAuthorization',
    'error.recovery.sameArgumentsSafe'
]) {
    assert(skillAuthoring.includes(recoveryField), `skill authoring omitted ${recoveryField}`);
}
assert.equal(skillAuthoring.includes('TARGET_NOT_FOUND'), false, 'skill guidance must not duplicate an error-code matrix');
assert.match(skillAuthoring, /operations file and serializer-backed stdin are peer transports/i);
assert.match(skillAuthoring, /AI Redliner.*valid visible fallback/);
for (const guidance of [launchCard, fastStart, readme, knowledgeBase, skillAuthoring]) {
    assert.match(guidance, /longer, fairly unique/i);
    assert(guidance.includes('anchorMatchCount'), 'anchor guidance omitted anchorMatchCount');
}
assert(releaseNotes.includes('"results": ['));
assert(releaseNotes.includes('"change": {'));
assert.match(releaseNotes, /excerpts come from the resolved batch-start source and compiled desired\s+text/i);

for (const [text, heading] of [
    [launchCard, '## Ordinary document edits'],
    [fastStart, '## CLI fallback'],
    [readme, '### Agent CLI'],
    [knowledgeBase, '#### Standard Workflow (Fast & Direct)'],
    [skillAuthoring, '## Ordinary generated workflow']
]) {
    const firstCommand = firstBashCommandAfter(text, heading);
    assert.match(firstCommand, /^docx-redline apply .*--find .*--replace/, `${heading}: ${firstCommand}`);
}

for (const publishedDoc of [
    'docs/AGENT_FAST_START.md',
    'docs/AGENT_KNOWLEDGE_BASE.md',
    'docs/SKILL_AUTHORING.md',
    'docs/TESTING.md',
    'docs/releases/0.7.0.md',
    'docs/schemas/document-operations.schema.json',
    'docs/validation-reports/2026-09-12-agent-protocol-rollout.md',
    'docs/validation-reports/2026-09-12-agent-cli-discovery-baseline.md',
    'docs/validation-reports/2026-09-13-agent-cli-efficiency-rollout.md',
    'docs/validation-reports/2026-09-14-localized-patching-rollout.md'
]) {
    assert(packageJson.files.includes(publishedDoc), `package omitted ${publishedDoc}`);
}
assert.equal(packageJson.files.includes('docs/'), false);
assert.equal(packageJson.files.some(entry => entry === 'examples/' || entry.startsWith('examples/')), false);
assert(packageJson.files.includes('!scripts/benchmark-agent-workflow.mjs'));
assert(packageJson.files.includes('!scripts/benchmark-localized-turn-reduction.mjs'));

console.log('agent documentation contract tests passed');
