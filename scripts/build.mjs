import { build } from 'esbuild';
import { readFileSync } from 'fs';

const pkg = JSON.parse(readFileSync('./package.json', 'utf8'));

// ESM bundle with diff-match-patch inlined (for CDN/browser <script type="module">)
await build({
  entryPoints: ['./index.js'],
  bundle: true,
  format: 'esm',
  outfile: 'dist/docx-reconciliation.esm.js',
  platform: 'neutral',         // no Node builtins assumed
  target: 'es2020',
  minify: false,                // keep readable for debugging
  sourcemap: true,
  banner: {
    js: `// @gsd/docx-reconciliation v${pkg.version} — https://github.com/YOUR_ORG/docx-reconciliation`
  },
  external: ['@xmldom/xmldom']  // never bundle the Node-only XML parser
});

// Minified version for production CDN use
await build({
  entryPoints: ['./index.js'],
  bundle: true,
  format: 'esm',
  outfile: 'dist/docx-reconciliation.esm.min.js',
  platform: 'neutral',
  target: 'es2020',
  minify: true,
  sourcemap: true,
  external: ['@xmldom/xmldom']
});

console.log('Build complete: dist/docx-reconciliation.esm.js, dist/docx-reconciliation.esm.min.js');
