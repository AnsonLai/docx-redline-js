import '../tests/setup-xml-provider.mjs';

import fs from 'fs';
import path from 'path';
import { spawnSync } from 'child_process';
import { fileURLToPath } from 'url';
import { buildMinimalDocx } from './lib/minimal-zip.mjs';
import { applyOperationsToDocumentXml } from '../services/standalone-operation-runner.js';
import {
    acceptTrackedChangesInOoxml,
    rejectTrackedChangesInOoxml
} from '../index.js';
import { elementsByLocalName, parseXml } from '../tests/helpers/ooxml-assertions.mjs';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const rootDir = path.resolve(__dirname, '..');
const outputDir = path.join(rootDir, 'tmp', 'multilevel-bullet-visual');
const pythonExe = path.join(rootDir, 'tmp', 'visual-qa-venv', 'Scripts', 'python.exe');

const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

const NUMBERING_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="${NS_W}" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">
  <w:abstractNum w:abstractNumId="0" w15:restartNumberingAfterBreak="0">
    <w:nsid w:val="71A1B101"/><w:multiLevelType w:val="hybridMultilevel"/><w:tmpl w:val="71A1B101"/>
    <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="upperRoman"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:tabs><w:tab w:val="num" w:pos="720"/></w:tabs><w:ind w:left="720" w:hanging="360"/></w:pPr><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/></w:rPr></w:lvl>
    <w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="(%2)"/><w:lvlJc w:val="left"/><w:pPr><w:tabs><w:tab w:val="num" w:pos="1440"/></w:tabs><w:ind w:left="1440" w:hanging="360"/></w:pPr><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/></w:rPr></w:lvl>
    <w:lvl w:ilvl="2"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/><w:lvlJc w:val="left"/><w:pPr><w:tabs><w:tab w:val="num" w:pos="2160"/></w:tabs><w:ind w:left="2160" w:hanging="360"/></w:pPr><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/></w:rPr></w:lvl>
  </w:abstractNum>
  <w:num w:numId="10"><w:abstractNumId w:val="0"/></w:num>
</w:numbering>`;

const SOURCE_DOCUMENT_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${NS_W}">
  <w:body>
    <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="32"/></w:rPr><w:t>Corporate Governance Charter</w:t></w:r></w:p>
    <w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="10"/></w:numPr></w:pPr><w:r><w:t>Corporate Governance Framework.</w:t></w:r></w:p>
    <w:p><w:pPr><w:numPr><w:ilvl w:val="1"/><w:numId w:val="10"/></w:numPr></w:pPr><w:r><w:t>Audit and Risk Oversight.</w:t></w:r></w:p>
    <w:p><w:pPr><w:numPr><w:ilvl w:val="2"/><w:numId w:val="10"/></w:numPr></w:pPr><w:r><w:t>Review internal accounting controls.</w:t></w:r></w:p>
    <w:p><w:pPr><w:numPr><w:ilvl w:val="2"/><w:numId w:val="10"/></w:numPr></w:pPr><w:r><w:t>Deprecated manual transaction logging.</w:t></w:r></w:p>
    <w:p><w:pPr><w:numPr><w:ilvl w:val="1"/><w:numId w:val="10"/></w:numPr></w:pPr><w:r><w:t>Executive Compensation.</w:t></w:r></w:p>
    <w:p><w:pPr><w:numPr><w:ilvl w:val="2"/><w:numId w:val="10"/></w:numPr></w:pPr><w:r><w:t>Review executive incentive structures.</w:t></w:r></w:p>
    <w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="10"/></w:numPr></w:pPr><w:r><w:t>Regulatory Compliance Protocol.</w:t></w:r></w:p>
    <w:sectPr/>
  </w:body>
</w:document>`;

async function main() {
    fs.mkdirSync(outputDir, { recursive: true });

    console.log('=== Multi-Level Bullet Visual Test ===');
    console.log('Building 3-level list document and executing concurrent multi-level operations...');

    // Multi-level operations:
    // 1. Level 0 edit: Corporate Governance Framework -> Corporate Governance and Enterprise Risk Framework
    // 2. Level 2 insert: under Audit and Risk Oversight, insert "Conduct annual cybersecurity vulnerability audit."
    // 3. Level 2 edit: Review internal accounting controls -> Review and audit internal accounting controls
    // 4. Level 2 consolidation/edit: Deprecated manual transaction logging -> Maintain automated electronic audit trails
    // 5. Level 1 edit: Executive Compensation -> Executive Compensation and Performance Oversight
    const operations = [
        {
            type: 'redline',
            target: 'Corporate Governance Framework.',
            modified: 'Corporate Governance and Enterprise Risk Framework.'
        },
        {
            type: 'list-change',
            target: 'Audit and Risk Oversight.',
            modified: '1. Audit and Risk Oversight.\n  - Conduct annual cybersecurity vulnerability audit.'
        },
        {
            type: 'redline',
            target: 'Review internal accounting controls.',
            modified: 'Review and audit internal accounting controls.'
        },
        {
            type: 'redline',
            target: 'Deprecated manual transaction logging.',
            modified: 'Maintain automated electronic audit trails.'
        },
        {
            type: 'redline',
            target: 'Executive Compensation.',
            modified: 'Executive Compensation and Performance Oversight.'
        }
    ];

    const result = await applyOperationsToDocumentXml(SOURCE_DOCUMENT_XML, operations, 'VisualTester');
    if (!result.hasChanges) {
        throw new Error('Operations did not report changes');
    }

    // Generate the 3 OOXML states
    const allMarkupXml = result.documentXml;
    const acceptAllXml = acceptTrackedChangesInOoxml(allMarkupXml, { author: 'VisualTester' }).oxml;
    const _rejectAllXml = rejectTrackedChangesInOoxml(allMarkupXml, { author: 'VisualTester' }).oxml;

    // Verify structural invariants in OOXML
    const acceptedDoc = parseXml(acceptAllXml);
    const acceptedParagraphs = elementsByLocalName(acceptedDoc, 'p');
    // Heading + 8 list items = 9 paragraphs
    console.log(`Accepted paragraphs: ${acceptedParagraphs.length}`);

    const parts = { numberingXml: NUMBERING_XML };
    const docxPath = path.join(outputDir, 'multilevel-bullet-edits.docx');
    fs.writeFileSync(docxPath, buildMinimalDocx(allMarkupXml, parts));
    console.log(`Generated DOCX fixture: ${docxPath}`);

    // Render 3 views using Word COM via PowerShell
    console.log('Rendering 3 views in Microsoft Word COM (ExportAsFixedFormat)...');
    const psScriptPath = path.join(__dirname, 'render-multilevel-bullet-visual.ps1');
    const psResult = spawnSync('powershell', [
        '-NoProfile',
        '-ExecutionPolicy',
        'Bypass',
        '-File',
        psScriptPath,
        '-DocxPath',
        docxPath,
        '-OutputDir',
        outputDir
    ], {
        encoding: 'utf8',
        cwd: rootDir
    });

    if (psResult.status !== 0) {
        console.error('Word COM rendering failed:', psResult.stderr || psResult.stdout);
        throw new Error('Word COM rendering failed');
    }
    console.log(psResult.stdout.trim());

    // Render high-res PNGs and contact sheet using PyMuPDF
    console.log('Converting PDFs to 150 DPI page images using PyMuPDF...');
    const pyScript = `
import os
import fitz
from PIL import Image

output_dir = r"${outputDir}"
views = ["allMarkup", "acceptAll", "rejectAll"]
rendered_images = []

for view in views:
    pdf_path = os.path.join(output_dir, f"multilevel-bullet--{view}.pdf")
    doc = fitz.open(pdf_path)
    page = doc.load_page(0)
    pix = page.get_pixmap(dpi=150)
    img_path = os.path.join(output_dir, f"multilevel-bullet--{view}--page1.png")
    pix.save(img_path)
    rendered_images.append(img_path)
    print(f"Rendered {img_path} ({pix.width}x{pix.height})")
    doc.close()

# Create side-by-side contact sheet
images = [Image.open(p) for p in rendered_images]
total_w = sum(img.width for img in images) + 40
max_h = max(img.height for img in images) + 20
sheet = Image.new("RGB", (total_w, max_h), (240, 240, 240))
x = 10
for img in images:
    sheet.paste(img, (x, 10))
    x += img.width + 10

sheet_path = os.path.join(output_dir, "multilevel-bullet--contact-sheet.png")
sheet.save(sheet_path)
print(f"Created contact sheet: {sheet_path} ({sheet.width}x{sheet.height})")
`;

    const pyResult = spawnSync(pythonExe, ['-c', pyScript], {
        encoding: 'utf8',
        cwd: rootDir
    });

    if (pyResult.status !== 0) {
        console.error('PyMuPDF conversion failed:', pyResult.stderr || pyResult.stdout);
        throw new Error('PyMuPDF conversion failed');
    }
    console.log(pyResult.stdout.trim());

    console.log('\n✅ Visual test execution complete! Images ready for multimodal inspection in:', outputDir);
}

main().catch(err => {
    console.error('Fatal error:', err);
    process.exit(1);
});
