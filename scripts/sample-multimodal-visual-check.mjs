import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'fs';
import { dirname, join, resolve } from 'path';
import { fileURLToPath } from 'url';
import { execFileSync } from 'child_process';
import { generateContactSheetWithPyMuPdf, loadManifest } from './inspect-visual-evidence.mjs';

const repoRoot = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const pythonVenvExe = join(repoRoot, 'tmp', 'visual-qa-venv', 'Scripts', 'python.exe');

export function renderPdfPageToImage(pdfPath, outputPngPath, pageIndex = 0, dpi = 150) {
    if (!existsSync(pythonVenvExe) || !existsSync(pdfPath)) {
        return false;
    }

    const pythonScript = `
import sys
import fitz

pdf_path = sys.argv[1]
output_png = sys.argv[2]
page_idx = int(sys.argv[3]) if len(sys.argv) > 3 else 0
dpi = int(sys.argv[4]) if len(sys.argv) > 4 else 150

doc = fitz.open(pdf_path)
if page_idx < 0 or page_idx >= len(doc):
    page_idx = 0
page = doc.load_page(page_idx)
pix = page.get_pixmap(dpi=dpi)
pix.save(output_png)
doc.close()
`;

    try {
        execFileSync(pythonVenvExe, ['-c', pythonScript, pdfPath, outputPngPath, String(pageIndex), String(dpi)], {
            stdio: 'pipe',
            encoding: 'utf8'
        });
        return true;
    } catch {
        return false;
    }
}

export function sampleScenarios(scenarios, count = 3, seed = null) {
    const list = [...scenarios];
    if (seed !== null) {
        // Simple seeded pseudo-random shuffle (mulberry32)
        let s = seed;
        const random = () => {
            s |= 0;
            s = (s + 0x6D2B79F5) | 0;
            let t = Math.imul(s ^ (s >>> 15), 1 | s);
            t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
            return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
        };
        for (let i = list.length - 1; i > 0; i--) {
            const j = Math.floor(random() * (i + 1));
            [list[i], list[j]] = [list[j], list[i]];
        }
    } else {
        // Standard shuffle
        for (let i = list.length - 1; i > 0; i--) {
            const j = Math.floor(Math.random() * (i + 1));
            [list[i], list[j]] = [list[j], list[i]];
        }
    }
    return list.slice(0, count);
}

function parseCliArgs() {
    const args = process.argv.slice(2);
    const options = {
        count: 3,
        seed: null,
        cases: null
    };

    for (const arg of args) {
        if (arg.startsWith('--count=')) {
            options.count = parseInt(arg.split('=')[1], 10) || 3;
        } else if (arg.startsWith('--seed=')) {
            options.seed = parseInt(arg.split('=')[1], 10) || 12345;
        } else if (arg.startsWith('--cases=')) {
            options.cases = arg.split('=')[1].split(',').map(s => s.trim()).filter(Boolean);
        }
    }

    return options;
}

export function buildMultimodalSpotCheckSamples(options = {}) {
    const corpusDir = join(repoRoot, 'tmp', 'superdoc-word-visual-review', 'rendered');
    const corpusManifestPath = join(corpusDir, 'manifest.json');
    const manifest = loadManifest(corpusManifestPath);

    if (!manifest || !Array.isArray(manifest.cases)) {
        throw new Error(`Corpus visual manifest not found at ${corpusManifestPath}. Run: npm run test:corpus:word:visual`);
    }

    const scenariosPath = join(repoRoot, 'tests', 'corpus', 'superdoc-word-scenarios.json');
    const scenariosData = JSON.parse(readFileSync(scenariosPath, 'utf8'));
    const scenarioMap = new Map();
    for (const sc of scenariosData.scenarios) {
        if (sc.name) scenarioMap.set(sc.name, sc);
    }

    // Filter or sample cases
    let selectedCases = [];
    if (options.cases && options.cases.length > 0) {
        selectedCases = manifest.cases.filter(c =>
            options.cases.some(target => c.identity?.includes(target) || c.name?.includes(target) || c.scenarioKey?.includes(target))
        );
    } else {
        selectedCases = sampleScenarios(manifest.cases, options.count || 3, options.seed);
    }

    if (selectedCases.length === 0) {
        throw new Error('No matching corpus visual cases found to sample.');
    }

    const outputDir = join(repoRoot, 'tmp', 'multimodal-visual-spot-checks');
    const imagesDir = join(outputDir, 'images');
    mkdirSync(imagesDir, { recursive: true });

    const sampleManifest = {
        createdAt: new Date().toISOString(),
        sampleCount: selectedCases.length,
        cases: []
    };

    let promptMarkdown = '# Multimodal LLM Visual Inspection Prompt Bundle\n\n';
    promptMarkdown += '> **Role:** You are an expert document layout quality assessor evaluating Microsoft Word tracked change renderings.\n';
    promptMarkdown += '> **Task:** Visually inspect the provided side-by-side renders across three views (`allMarkup`, `acceptAll`, `rejectAll`) to detect visual layout, table, list, or typographic defects.\n\n';

    for (let i = 0; i < selectedCases.length; i++) {
        const testCase = selectedCases[i];
        const scenarioKey = testCase.scenarioKey || testCase.name;
        const caseRecord = {
            index: i + 1,
            scenarioKey,
            category: testCase.category,
            shape: testCase.shape,
            pages: testCase.views?.allMarkup?.pages || 1,
            images: {}
        };

        promptMarkdown += `## Sample ${i + 1}: \`${scenarioKey}\`\n`;
        promptMarkdown += `- **Category:** ${testCase.category} | **Shape:** ${testCase.shape}\n`;
        promptMarkdown += `- **Document Length:** ${caseRecord.pages} page(s)\n\n`;

        for (const viewName of ['allMarkup', 'acceptAll', 'rejectAll']) {
            const viewInfo = testCase.views?.[viewName];
            if (!viewInfo || !viewInfo.pdf) continue;

            const pdfPath = join(corpusDir, viewInfo.pdf);
            const highResImageName = `${scenarioKey}--${viewName}--page1.png`;
            const highResImagePath = join(imagesDir, highResImageName);
            const sheetImageName = `${scenarioKey}--${viewName}--sheet.png`;
            const sheetImagePath = join(imagesDir, sheetImageName);

            // Render high-res page 1
            renderPdfPageToImage(pdfPath, highResImagePath, 0, 150);
            // Render contact sheet (up to 6 pages)
            generateContactSheetWithPyMuPdf(pdfPath, sheetImagePath, 6);

            caseRecord.images[viewName] = {
                highRes: highResImagePath,
                sheet: sheetImagePath
            };

            promptMarkdown += `### View: \`${viewName}\`\n`;
            promptMarkdown += `- High-Res Page 1: \`${highResImagePath}\`\n`;
            promptMarkdown += `- Contact Sheet: \`${sheetImagePath}\`\n\n`;
        }

        promptMarkdown += '### Visual Inspection Questions:\n';
        promptMarkdown += '1. **Markup Visibility & Isolation (`allMarkup`):** Are tracked insertions (underlined/colored) and deletions (strikethrough) clearly visible and cleanly localized, without wrapping or clipping adjacent text?\n';
        promptMarkdown += '2. **Accepted State Correctness (`acceptAll`):** Does the document render cleanly without any leftover deletion markers or awkward spacing?\n';
        promptMarkdown += '3. **Rejected State Fidelity (`rejectAll`):** Does the page restore the exact original layout, fonts, and numbering without ghost bullets or shifted margins?\n';
        promptMarkdown += '4. **Structural & Table Integrity:** Did table column widths, grid lines, background fills, or list indents remain stable across all three views?\n\n';
        promptMarkdown += '---\n\n';

        sampleManifest.cases.push(caseRecord);
    }

    const manifestOutPath = join(outputDir, 'sample-manifest.json');
    const promptOutPath = join(outputDir, 'multimodal-prompt.md');

    writeFileSync(manifestOutPath, JSON.stringify(sampleManifest, null, 2), 'utf8');
    writeFileSync(promptOutPath, promptMarkdown, 'utf8');

    return {
        sampleManifest,
        outputDir,
        manifestOutPath,
        promptOutPath
    };
}

async function runCli() {
    console.log('=== Multimodal Visual Spot Check Sample Generator ===\n');
    const options = parseCliArgs();

    try {
        const result = buildMultimodalSpotCheckSamples(options);
        console.log(`Generated ${result.sampleManifest.sampleCount} multimodal inspection samples:`);
        for (const c of result.sampleManifest.cases) {
            console.log(`  - [Sample ${c.index}] ${c.scenarioKey} (${c.category}, ${c.shape}, ${c.pages} pages)`);
        }
        console.log(`\nManifest: ${result.manifestOutPath}`);
        console.log(`Prompt bundle: ${result.promptOutPath}`);
    } catch (err) {
        console.error('Error generating multimodal samples:', err.message);
        process.exit(1);
    }
}

const isCli = process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url);
if (isCli) {
    runCli();
}
