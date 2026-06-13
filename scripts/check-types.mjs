import { readFileSync } from 'fs';

const dts = readFileSync(new URL('../index.d.ts', import.meta.url), 'utf8');

const requiredSnippets = [
  'export interface RedlineOptions',
  'export interface RedlineResult',
  'export function applyRedlineToOxml',
  'export function acceptTrackedChangesInOoxml',
  'export function rejectTrackedChangesInOoxml',
  'export function deleteCommentsByAuthorInOoxml'
];

for (const snippet of requiredSnippets) {
  if (!dts.includes(snippet)) {
    throw new Error(`Missing declaration snippet: ${snippet}`);
  }
}

let balance = 0;
for (const char of dts) {
  if (char === '{') balance += 1;
  if (char === '}') balance -= 1;
  if (balance < 0) throw new Error('index.d.ts has unbalanced braces');
}
if (balance !== 0) throw new Error('index.d.ts has unbalanced braces');

console.log('PASS: index.d.ts declaration smoke check');
