import { spawnSync } from 'child_process';

const baseName = '48-administrative-administrative-list-change-dash-agenda-child';
const fixtureDir = 'tmp/superdoc-word-fixtures';
const visualDir = 'tmp/multilevel-bullet-visual';

const docxPath = `${fixtureDir}/${baseName}.docx`;
const acceptedDocx = `${fixtureDir}/${baseName}.accepted.docx`;
const rejectedDocx = `${fixtureDir}/${baseName}.rejected.docx`;

console.log('Rendering Word COM for', baseName);

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
    Write-Output 'SUCCESS: Rendered all 3 views.'
} finally {
    $word.Quit()
}
`;

const psRes = spawnSync('powershell', ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-Command', psScript], { encoding: 'utf8' });
console.log(psRes.stdout);
if (psRes.stderr) console.error(psRes.stderr);

// Render PDF pages to PNG using python pymupdf
const pyScript = `
import pymupdf
views = ['allMarkup', 'acceptAll', 'rejectAll']
base = '${baseName}'
vis_dir = '${visualDir}'

for v in views:
    pdf_file = f"{vis_dir}/{base}--{v}.pdf"
    doc = pymupdf.open(pdf_file)
    print(f"{v} page count: {len(doc)}")
    for i, page in enumerate(doc):
        pix = page.get_pixmap(dpi=150)
        out_png = f"{vis_dir}/{base}--{v}-p{i+1}.png"
        pix.save(out_png)
        print(f"  Rendered {out_png} ({pix.width}x{pix.height})")
`;

const pyRes = spawnSync('tmp/visual-qa-venv/Scripts/python.exe', ['-c', pyScript], { encoding: 'utf8' });
console.log(pyRes.stdout);
if (pyRes.stderr) console.error(pyRes.stderr);
