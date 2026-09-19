import { build } from 'esbuild';
import { readFileSync } from 'fs';
import { dirname, resolve } from 'path';
import { fileURLToPath } from 'url';

const repoRoot = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const entryPoint = resolve(repoRoot, 'index.js');
const distDir = resolve(repoRoot, 'dist');
const pkg = JSON.parse(readFileSync(resolve(repoRoot, 'package.json'), 'utf8'));

// ESM bundle with diff-match-patch and fflate inlined (for bundlers / browser)
await build({
  entryPoints: [entryPoint],
  bundle: true,
  format: 'esm',
  outfile: resolve(distDir, 'docx-redline-js.esm.js'),
  platform: 'neutral',
  mainFields: ['module', 'main'],
  target: 'es2020',
  minify: false,
  sourcemap: true,
  banner: {
    js: `// @ansonlai/docx-redline-js v${pkg.version} — https://github.com/AnsonLai/docx-redline-js`
  },
  external: ['@xmldom/xmldom']
});

// Minified version for production CDN use
await build({
  entryPoints: [entryPoint],
  bundle: true,
  format: 'esm',
  outfile: resolve(distDir, 'docx-redline-js.esm.min.js'),
  platform: 'neutral',
  mainFields: ['module', 'main'],
  target: 'es2020',
  minify: true,
  sourcemap: true,
  external: ['@xmldom/xmldom']
});

// Standalone zero-dependency ESM bundle (includes @xmldom/xmldom, fflate, diff-match-patch)
await build({
  entryPoints: [entryPoint],
  bundle: true,
  format: 'esm',
  outfile: resolve(distDir, 'docx-redline.bundle.js'),
  platform: 'neutral',
  mainFields: ['module', 'main'],
  target: 'es2020',
  minify: false,
  sourcemap: true,
  banner: {
    js: `// @ansonlai/docx-redline-js standalone bundle v${pkg.version} — https://github.com/AnsonLai/docx-redline-js`
  }
});

// Standalone zero-dependency CommonJS bundle (for n8n, isolated-vm, CJS sandboxes)
await build({
  entryPoints: [entryPoint],
  bundle: true,
  format: 'cjs',
  outfile: resolve(distDir, 'docx-redline.bundle.cjs'),
  platform: 'neutral',
  mainFields: ['module', 'main'],
  target: 'es2020',
  minify: false,
  sourcemap: true,
  banner: {
    js: `// @ansonlai/docx-redline-js standalone CommonJS bundle v${pkg.version} — https://github.com/AnsonLai/docx-redline-js`
  }
});

console.log('Build complete: dist/docx-redline-js.esm.js, dist/docx-redline-js.esm.min.js, dist/docx-redline.bundle.js, dist/docx-redline.bundle.cjs');

