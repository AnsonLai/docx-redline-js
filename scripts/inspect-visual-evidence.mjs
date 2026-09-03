import { existsSync, mkdirSync, readFileSync } from 'fs';
import { dirname, join, resolve } from 'path';
import { fileURLToPath } from 'url';
import { execFileSync } from 'child_process';

const repoRoot = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const pythonVenvExe = join(repoRoot, 'tmp', 'visual-qa-venv', 'Scripts', 'python.exe');

export function loadManifest(manifestPath) {
    if (!existsSync(manifestPath)) {
        return null;
    }
    let content = readFileSync(manifestPath, 'utf8');
    if (content.charCodeAt(0) === 0xFEFF) {
        content = content.slice(1);
    }
    return JSON.parse(content);
}

export function inspectManifest(manifest, baseDir) {
    const results = {
        totalCases: 0,
        renderedCases: 0,
        passedCases: 0,
        anomalies: [],
        caseSummaries: []
    };

    if (!manifest || !Array.isArray(manifest.cases)) {
        return results;
    }

    results.totalCases = manifest.cases.length;

    for (const testCase of manifest.cases) {
        const caseName = testCase.name || testCase.identity;
        const views = testCase.views || {};
        let caseRendered = true;
        let caseValid = true;
        const viewDetails = {};

        for (const viewName of ['allMarkup', 'acceptAll', 'rejectAll']) {
            const viewInfo = views[viewName];
            if (!viewInfo || viewInfo.status !== 'rendered') {
                caseRendered = false;
                caseValid = false;
                continue;
            }

            const pdfFileName = viewInfo.pdf;
            const pdfPath = join(baseDir, pdfFileName);
            const exists = existsSync(pdfPath);
            const pages = viewInfo.pages;
            const bytes = viewInfo.bytes;

            viewDetails[viewName] = { exists, pages, bytes };

            if (!exists) {
                results.anomalies.push({
                    caseName,
                    view: viewName,
                    reason: `Missing PDF file: ${pdfFileName}`
                });
                caseValid = false;
            } else if (!bytes || bytes < 1000) {
                results.anomalies.push({
                    caseName,
                    view: viewName,
                    reason: `Suspiciously small PDF size: ${bytes} bytes`
                });
                caseValid = false;
            } else if (!pages || pages < 1) {
                results.anomalies.push({
                    caseName,
                    view: viewName,
                    reason: `Invalid page count: ${pages}`
                });
                caseValid = false;
            }
        }

        if (caseRendered) {
            results.renderedCases++;
        }

        // Layout sanity check between acceptAll and rejectAll
        const acceptPages = viewDetails.acceptAll?.pages;
        const rejectPages = viewDetails.rejectAll?.pages;
        if (acceptPages !== undefined && rejectPages !== undefined) {
            const pageDiff = Math.abs(acceptPages - rejectPages);
            if (pageDiff > 2 && acceptPages > 0 && rejectPages > 0) {
                results.anomalies.push({
                    caseName,
                    reason: `Unusual page count disparity: acceptAll=${acceptPages} vs rejectAll=${rejectPages}`
                });
                caseValid = false;
            }
        }

        if (caseValid) {
            results.passedCases++;
        }

        results.caseSummaries.push({
            name: caseName,
            status: caseValid ? 'valid' : 'flagged',
            rendered: caseRendered,
            views: viewDetails
        });
    }

    return results;
}

export function generateContactSheetWithPyMuPdf(pdfPath, outputPngPath, maxPages = 20) {
    if (!existsSync(pythonVenvExe)) {
        return false;
    }

    const pythonScript = `
import sys
import fitz # PyMuPDF
from PIL import Image

pdf_path = sys.argv[1]
output_png = sys.argv[2]
max_pages = int(sys.argv[3]) if len(sys.argv) > 3 else 20

doc = fitz.open(pdf_path)
page_count = min(len(doc), max_pages)

# Render pages as pixmaps
images = []
for i in range(page_count):
    page = doc.load_page(i)
    pix = page.get_pixmap(dpi=72)
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    images.append(img)

doc.close()

if not images:
    sys.exit(1)

# Tile into contact sheet grid (up to 5 columns)
cols = min(5, len(images))
rows = (len(images) + cols - 1) // cols
page_w, page_h = images[0].size
sheet = Image.new("RGB", (cols * page_w, rows * page_h), color=(240, 240, 240))

for idx, img in enumerate(images):
    c = idx % cols
    r = idx // cols
    sheet.paste(img, (c * page_w, r * page_h))

sheet.save(output_png)
print("Saved contact sheet:", output_png)
`;

    try {
        execFileSync(pythonVenvExe, ['-c', pythonScript, pdfPath, outputPngPath, String(maxPages)], {
            stdio: 'pipe',
            encoding: 'utf8'
        });
        return true;
    } catch (e) {
        return false;
    }
}

function runCli() {
    console.log('=== Word Visual Evidence Inspection ===\n');

    const syntheticDir = join(repoRoot, 'tmp', 'word-visual-review', 'rendered');
    const syntheticManifestPath = join(syntheticDir, 'manifest.json');
    const corpusDir = join(repoRoot, 'tmp', 'superdoc-word-visual-review', 'rendered');
    const corpusManifestPath = join(corpusDir, 'manifest.json');

    let totalCases = 0;
    let totalRendered = 0;
    let totalValid = 0;
    let totalAnomalies = 0;

    // Check synthetic lane
    const syntheticManifest = loadManifest(syntheticManifestPath);
    if (syntheticManifest) {
        const syntheticResults = inspectManifest(syntheticManifest, syntheticDir);
        console.log(`Synthetic Word Visual Evidence (${syntheticDir}):`);
        console.log(`  Word Version: ${syntheticManifest.word?.version || 'unknown'} (Build: ${syntheticManifest.word?.build || 'unknown'})`);
        console.log(`  Total Cases: ${syntheticResults.totalCases}`);
        console.log(`  Rendered Cases: ${syntheticResults.renderedCases}`);
        console.log(`  Valid Cases: ${syntheticResults.passedCases}`);
        if (syntheticResults.anomalies.length > 0) {
            console.log('  Anomalies:');
            for (const a of syntheticResults.anomalies) {
                console.log(`    - [${a.caseName}] ${a.view ? a.view + ': ' : ''}${a.reason}`);
            }
        }
        console.log('');
        totalCases += syntheticResults.totalCases;
        totalRendered += syntheticResults.renderedCases;
        totalValid += syntheticResults.passedCases;
        totalAnomalies += syntheticResults.anomalies.length;
    } else {
        console.log(`Synthetic lane manifest not found at ${syntheticManifestPath}\n  Run: npm run test:word:visual\n`);
    }

    // Check corpus lane
    const corpusManifest = loadManifest(corpusManifestPath);
    if (corpusManifest) {
        const corpusResults = inspectManifest(corpusManifest, corpusDir);
        console.log(`SuperDoc Real-Document Visual Evidence (${corpusDir}):`);
        console.log(`  Word Version: ${corpusManifest.word?.version || 'unknown'} (Build: ${corpusManifest.word?.build || 'unknown'})`);
        console.log(`  Total Cases: ${corpusResults.totalCases}`);
        console.log(`  Rendered Cases: ${corpusResults.renderedCases}`);
        console.log(`  Valid Cases: ${corpusResults.passedCases}`);
        if (corpusResults.anomalies.length > 0) {
            console.log('  Anomalies:');
            for (const a of corpusResults.anomalies) {
                console.log(`    - [${a.caseName}] ${a.view ? a.view + ': ' : ''}${a.reason}`);
            }
        }
        console.log('');
        totalCases += corpusResults.totalCases;
        totalRendered += corpusResults.renderedCases;
        totalValid += corpusResults.passedCases;
        totalAnomalies += corpusResults.anomalies.length;
    } else {
        console.log(`SuperDoc corpus manifest not found at ${corpusManifestPath}\n  Run: npm run test:corpus:word:visual\n`);
    }

    console.log(`Summary: ${totalRendered}/${totalCases} rendered, ${totalValid} valid, ${totalAnomalies} anomalies.\n`);

    if (process.argv.includes('--contact-sheets')) {
        console.log('Generating visual review contact sheets...');
        const sheetsDir = join(repoRoot, 'tmp', 'word-visual-review', 'inspected-sheets');
        mkdirSync(sheetsDir, { recursive: true });

        let generated = 0;
        if (syntheticManifest) {
            for (const c of syntheticManifest.cases.slice(0, 10)) {
                for (const [viewName, viewInfo] of Object.entries(c.views)) {
                    const pdfPath = join(syntheticDir, viewInfo.pdf);
                    const outPath = join(sheetsDir, `synthetic--${c.name}--${viewName}.png`);
                    if (existsSync(pdfPath)) {
                        const success = generateContactSheetWithPyMuPdf(pdfPath, outPath, 10);
                        if (success) generated++;
                    }
                }
            }
        }
        if (corpusManifest) {
            for (const c of corpusManifest.cases.slice(0, 10)) {
                for (const [viewName, viewInfo] of Object.entries(c.views)) {
                    const pdfPath = join(corpusDir, viewInfo.pdf);
                    const outPath = join(sheetsDir, `superdoc--${c.scenarioKey || c.name}--${viewName}.png`);
                    if (existsSync(pdfPath)) {
                        const success = generateContactSheetWithPyMuPdf(pdfPath, outPath, 10);
                        if (success) generated++;
                    }
                }
            }
        }
        console.log(`Generated ${generated} visual contact sheets in ${sheetsDir}`);
    }
}

const isCli = process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url);
if (isCli) {
    runCli();
}
