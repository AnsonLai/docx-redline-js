import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { resolve, dirname } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import vm from 'node:vm';
import { zipDocx } from '../document/zip-archive.js';

const __dirname = dirname(fileURLToPath(import.meta.url));
const repoRoot = resolve(__dirname, '..');
const distDir = resolve(repoRoot, 'dist');

const cjsBundlePath = resolve(distDir, 'docx-redline.bundle.cjs');
const esmBundlePath = resolve(distDir, 'docx-redline.bundle.js');

const textEncoder = new TextEncoder();

function createMinimalDocxUint8Array(paragraphText = 'Hello Sandbox') {
    const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
  <w:body>
    <w:p w14:paraId="00000001">
      <w:r><w:t>${paragraphText}</w:t></w:r>
    </w:p>
  </w:body>
</w:document>`;

    const contentTypesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`;

    const relsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

    const entries = new Map([
        ['[Content_Types].xml', textEncoder.encode(contentTypesXml)],
        ['_rels/.rels', textEncoder.encode(relsXml)],
        ['word/document.xml', textEncoder.encode(documentXml)]
    ]);

    return zipDocx(entries);
}

// 1. Static inspection: verify zero forbidden Node imports/calls in sandbox bundles
{
    const cjsSource = readFileSync(cjsBundlePath, 'utf8');
    const esmSource = readFileSync(esmBundlePath, 'utf8');

    const forbiddenPatterns = [
        /node:zlib/,
        /node:crypto/,
        /require\(['"]node:zlib['"]\)/,
        /require\(['"]node:crypto['"]\)/,
        /require\(['"]fs['"]\)/,
        /require\(['"]node:fs['"]\)/,
        /require\(['"]path['"]\)/,
        /require\(['"]node:path['"]\)/,
        /require\(['"]child_process['"]\)/
    ];

    for (const pattern of forbiddenPatterns) {
        assert.ok(!pattern.test(cjsSource), `CJS bundle must not match ${pattern}`);
        assert.ok(!pattern.test(esmSource), `ESM bundle must not match ${pattern}`);
    }
}

// 2. Strict CJS Sandbox Execution without Node built-ins (no process, Buffer, require)
{
    const cjsSource = readFileSync(cjsBundlePath, 'utf8');

    const sandbox = {
        Uint8Array,
        Array,
        Object,
        String,
        Number,
        Boolean,
        Map,
        Set,
        Math,
        JSON,
        TextEncoder,
        TextDecoder,
        DataView,
        ArrayBuffer,
        Error,
        TypeError,
        RangeError,
        Promise,
        console: {
            log: () => {},
            warn: () => {},
            error: () => {}
        },
        module: { exports: {} },
        exports: {}
    };

    // Explicitly verify forbidden globals do not exist in sandbox
    assert.equal(sandbox.process, undefined);
    assert.equal(sandbox.Buffer, undefined);
    assert.equal(sandbox.require, undefined);
    assert.equal(sandbox.DOMParser, undefined);
    assert.equal(sandbox.XMLSerializer, undefined);

    const context = vm.createContext(sandbox);
    vm.runInContext(cjsSource, context);

    const api = context.module.exports;
    assert.equal(typeof api.openDocx, 'function');
    assert.equal(typeof api.DocxDocument, 'function');
    assert.equal(typeof api.computePackageRevisionToken, 'function');
    assert.equal(typeof api.applyRedlineToOxml, 'function');

    // Test applyRedlineToOxml in isolated sandbox
    const oxml = '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:r><w:t>Hello sandbox</w:t></w:r></w:p>';
    const redlinePromise = api.applyRedlineToOxml(oxml, 'Hello sandbox', 'Hello isolated sandbox');
    const redlineResult = await redlinePromise;
    assert.equal(redlineResult.hasChanges, true);
    assert.ok(redlineResult.oxml.includes('w:ins'));
    assert.ok(redlineResult.oxml.includes('isolated '));

    // Test openDocx with Uint8Array in isolated sandbox
    const inputBytes = createMinimalDocxUint8Array('Hello Sandbox Document');
    const doc = api.openDocx(inputBytes);

    const inspection = doc.inspect();
    assert.equal(inspection.status, 'ok');
    assert.equal(inspection.paragraphs.length, 1);
    assert.equal(inspection.paragraphs[0].exactText, 'Hello Sandbox Document');

    const token = doc.getRevisionToken();
    assert.equal(token.scope, 'package');
    assert.equal(token.algorithm, 'sha256');
    assert.equal(typeof token.value, 'string');
    assert.equal(token.value.length, 64);

    const operations = [
        {
            type: 'replace',
            target: { text: 'Hello Sandbox Document' },
            modified: 'Modified in pure sandbox'
        }
    ];

    const applyResult = await doc.applyOperations(operations, { author: 'SandboxBot' });
    assert.equal(applyResult.status, 'ok');
    assert.equal(applyResult.written, true);
    assert.equal(applyResult.hasChanges, true);

    const outputBytes = doc.toUint8Array();
    assert.ok(outputBytes instanceof Uint8Array);
    assert.ok(!(outputBytes instanceof (globalThis.Buffer || Uint8Array.prototype.constructor)));

    // Re-open modified document in sandbox and inspect
    const reloadedDoc = api.openDocx(outputBytes);
    const reloadedInspection = reloadedDoc.inspect();
    assert.equal(reloadedInspection.status, 'ok');
    assert.equal(reloadedInspection.paragraphs.length, 1);
    assert.ok(reloadedInspection.paragraphs[0].exactText.includes('sandbox'));
}

// 3. ESM Bundle Execution
{
    const esmApi = await import(pathToFileURL(esmBundlePath).href);
    assert.equal(typeof esmApi.openDocx, 'function');
    assert.equal(typeof esmApi.DocxDocument, 'function');
    assert.equal(typeof esmApi.applyRedlineToOxml, 'function');

    const inputBytes = createMinimalDocxUint8Array('ESM Bundle Test');
    const doc = esmApi.openDocx(inputBytes);
    assert.equal(doc.inspect().paragraphs[0].exactText, 'ESM Bundle Test');

    const result = await doc.applyOperations([{
        type: 'replace',
        target: { text: 'ESM Bundle Test' },
        modified: 'ESM Bundle Verified'
    }], { author: 'ESMBot' });

    assert.equal(result.status, 'ok');
    assert.equal(result.written, true);
    assert.ok(result.uint8Array instanceof Uint8Array);
    assert.equal(doc.inspect().paragraphs[0].exactText, 'ESM Bundle Verified');
}

console.log('PASS: sandbox bundles execute cleanly in restricted sandbox without Node built-ins');
