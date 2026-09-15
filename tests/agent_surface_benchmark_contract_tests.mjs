import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';

const source = await readFile(
    new URL('../scripts/benchmark-agent-surface-simplification.mjs', import.meta.url),
    'utf8'
);

for (const policy of [
    'v0.6.2-extract-full',
    'v0.7.1-speculation-led',
    'simplified-extract-localized'
]) {
    assert(source.includes(policy), `benchmark omitted policy ${policy}`);
}

for (const scenario of [
    'literal-long-paragraph',
    'semantic-contextual-rewrite',
    'repeated-visible-literal',
    'rejected-view-restore',
    'nbsp-localized-target',
    'independent-two-target-batch'
]) {
    assert(source.includes(scenario), `benchmark omitted scenario ${scenario}`);
}

for (const clientDerivedTerm of [
    'Salary',
    'Privacy Policy',
    'Data Processing Addendum',
    'Customer Data'
]) {
    assert.equal(source.includes(clientDerivedTerm), false, `benchmark contains client-derived term ${clientDerivedTerm}`);
}

assert.match(source, /modelAndNetwork:\s*\{\s*measured: false/);
assert.match(source, /Every output is re-extracted and compared/);
assert.equal(JSON.parse(await readFile(new URL('../package.json', import.meta.url), 'utf8')).scripts['benchmark:agent-surface'], 'node scripts/benchmark-agent-surface-simplification.mjs');

console.log('agent surface benchmark contract tests passed');
