import { readFileSync } from 'node:fs';

const declarationPath = new URL('../index.d.ts', import.meta.url);
const declarationText = readFileSync(declarationPath, 'utf8');
const declaredRuntimeNames = new Set();
const declarationPattern = /^export\s+(?:declare\s+)?(?:function|class|const|let|var|enum)\s+([A-Za-z_$][\w$]*)/gm;
for (const match of declarationText.matchAll(declarationPattern)) {
    declaredRuntimeNames.add(match[1]);
}

const runtimeModule = await import('../index.js');
const runtimeNames = Object.keys(runtimeModule).sort();
const missingDeclarations = runtimeNames.filter(name => !declaredRuntimeNames.has(name));

if (missingDeclarations.length > 0) {
    throw new Error(
        `Runtime exports missing from index.d.ts:\n${missingDeclarations.map(name => `- ${name}`).join('\n')}`
    );
}

console.log(`PASS: ${runtimeNames.length} runtime exports have declarations`);
