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
const failOutputPattern = /(?:❌\s*(?:FAIL|FAILED|FAILURE|TEST FAILED)|\bTEST FAILED\b)/i;

for (const file of testFiles) {
  const filePath = join(testDir, file);
  process.stdout.write(`  ${file} ... `);
  try {
    const output = execSync(`node "${filePath}"`, { stdio: 'pipe', timeout: 30000 });
    const outputText = output.toString();
    if (failOutputPattern.test(outputText)) {
      throw new Error(`Test printed a failure marker while exiting successfully:\n${outputText}`);
    }
    console.log('PASS');
    passed++;
  } catch (err) {
    console.log('FAIL');
    const stdout = err.stdout?.toString() || '';
    const stderr = err.stderr?.toString() || '';
    failures.push({ file, stderr: [stdout, stderr, err.message].filter(Boolean).join('\n') });
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
