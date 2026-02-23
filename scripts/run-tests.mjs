import { readdirSync } from 'fs';
import { join, dirname } from 'path';
import { execSync } from 'child_process';
import { fileURLToPath } from 'url';

const __dirname = dirname(fileURLToPath(import.meta.url));
const testDir = join(__dirname, '..', 'tests');
const testFiles = readdirSync(testDir)
  .filter(f => f.endsWith('.mjs') && f !== 'setup-xml-provider.mjs')
  .sort();

let passed = 0;
let failed = 0;
const failures = [];

for (const file of testFiles) {
  const filePath = join(testDir, file);
  process.stdout.write(`  ${file} ... `);
  try {
    execSync(`node "${filePath}"`, { stdio: 'pipe', timeout: 30000 });
    console.log('PASS');
    passed++;
  } catch (err) {
    console.log('FAIL');
    failures.push({ file, stderr: err.stderr?.toString() || err.message });
    failed++;
  }
}

console.log(`\n${passed} passed, ${failed} failed out of ${passed + failed} tests`);
if (failures.length > 0) {
  for (const f of failures) {
    console.error(`\n--- ${f.file} ---\n${f.stderr}`);
  }
  process.exit(1);
}
