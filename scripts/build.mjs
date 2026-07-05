import { build } from 'esbuild';
import { readFileSync } from 'fs';
import { dirname, resolve } from 'path';
import { fileURLToPath } from 'url';

const repoRoot = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const entryPoint = resolve(repoRoot, 'index.js');
const distDir = resolve(repoRoot, 'dist');
const pkg = JSON.parse(readFileSync(resolve(repoRoot, 'package.json'), 'utf8'));

// ESM bundle with diff-match-patch inlined (for CDN/browser <script type="module">)
await build({
  entryPoints: [entryPoint],
  bundle: true,
  format: 'esm',
  outfile: resolve(distDir, 'docx-redline-js.esm.js'),
  platform: 'neutral',         // no Node builtins assumed
  target: 'es2020',
  minify: false,                // keep readable for debugging
  sourcemap: true,
  banner: {
    js: `// @ansonlai/docx-redline-js v${pkg.version} — https://github.com/AnsonLai/docx-redline-js`
  },
  external: ['@xmldom/xmldom']  // never bundle the Node-only XML parser
});

// Minified version for production CDN use
await build({
  entryPoints: [entryPoint],
  bundle: true,
  format: 'esm',
  outfile: resolve(distDir, 'docx-redline-js.esm.min.js'),
  platform: 'neutral',
  target: 'es2020',
  minify: true,
  sourcemap: true,
  external: ['@xmldom/xmldom']
});

console.log('Build complete: dist/docx-redline-js.esm.js, dist/docx-redline-js.esm.min.js');
