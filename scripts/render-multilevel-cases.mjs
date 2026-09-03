import fs from 'fs';
import { spawnSync } from 'child_process';
import { resolve } from 'path';

const cases = [
    '37-administrative-administrative-list-change-board-agenda-child',
    '40-administrative-administrative-list-change-ppg-agenda-addition'
];

const fixtureDir = 'tmp/superdoc-word-fixtures';
const visualDir = 'tmp/multilevel-bullet-visual';
const brainArtifactDir = 'C:/Users/Phara/.gemini/antigravity-ide/brain/ee12439b-cade-4f45-ab6e-a2a7b1bf1610';

if (!fs.existsSync(visualDir)) {
    fs.mkdirSync(visualDir, { recursive: true });
}

for (const baseName of cases) {
    console.log('\n=== Rendering Word COM for', baseName, '===');
    const docxPath = `${fixtureDir}/${baseName}.docx`;
    const acceptedDocx = `${fixtureDir}/${baseName}.accepted.docx`;
    const rejectedDocx = `${fixtureDir}/${baseName}.rejected.docx`;

    const psScript = `
    $ErrorActionPreference = 'Stop'
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    try {
        Write-Output 'Rendering allMarkup...'
        $doc = $word.Documents.OpenNoRepairDialog((Resolve-Path '${docxPath}').Path)
        $doc.ShowRevisions = $true
        $doc.PrintRevisions = $true
        $doc.ExportAsFixedFormat((Resolve-Path '${visualDir}').Path + '/${baseName}--allMarkup.pdf', 17, $false, 0, 0, 1, 1, 7, $true, $true, 0, $true, $true, $false)
        $doc.Close(0)

        Write-Output 'Rendering acceptAll...'
        $docAcc = $word.Documents.OpenNoRepairDialog((Resolve-Path '${acceptedDocx}').Path)
        $docAcc.ExportAsFixedFormat((Resolve-Path '${visualDir}').Path + '/${baseName}--acceptAll.pdf', 17, $false, 0, 0, 1, 1, 0, $true, $true, 0, $true, $true, $false)
        $docAcc.Close(0)

        Write-Output 'Rendering rejectAll...'
        $docRej = $word.Documents.OpenNoRepairDialog((Resolve-Path '${rejectedDocx}').Path)
        $docRej.ExportAsFixedFormat((Resolve-Path '${visualDir}').Path + '/${baseName}--rejectAll.pdf', 17, $false, 0, 0, 1, 1, 0, $true, $true, 0, $true, $true, $false)
        $docRej.Close(0)
        Write-Output 'SUCCESS: Rendered all 3 views for ${baseName}.'
    } finally {
        $word.Quit()
    }
    `;

    const psRes = spawnSync('powershell', ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-Command', psScript], { encoding: 'utf8' });
    console.log(psRes.stdout);
    if (psRes.stderr) console.error(psRes.stderr);

    const pyScript = `
import pymupdf
import shutil
views = ['allMarkup', 'acceptAll', 'rejectAll']
base = '${baseName}'
vis_dir = '${visualDir}'
art_dir = '${brainArtifactDir}'

for v in views:
    pdf_file = f"{vis_dir}/{base}--{v}.pdf"
    doc = pymupdf.open(pdf_file)
    print(f"{base} {v} page count: {len(doc)}")
    for i, page in enumerate(doc):
        pix = page.get_pixmap(dpi=150)
        out_png = f"{vis_dir}/{base}--{v}-p{i+1}.png"
        pix.save(out_png)
        art_png = f"{art_dir}/{base}--{v}-p{i+1}.png"
        shutil.copyfile(out_png, art_png)
        print(f"  Rendered {out_png} -> {art_png}")
`;

    const pyRes = spawnSync('tmp/visual-qa-venv/Scripts/python.exe', ['-c', pyScript], { encoding: 'utf8' });
    console.log(pyRes.stdout);
    if (pyRes.stderr) console.error(pyRes.stderr);
}
