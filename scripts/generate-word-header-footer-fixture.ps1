# Generates tests/fixtures/word-authored/header-footer.docx using real Word: default + first-page
# header/footer with a PAGE field, and a second section. Ground truth for header/footer part
# content types, relationship types and sectPr references.
$ErrorActionPreference = 'Stop'
$outDir = Join-Path $PSScriptRoot '..\tests\fixtures\word-authored'
New-Item -ItemType Directory -Path $outDir -Force | Out-Null
$outPath = [System.IO.Path]::GetFullPath((Join-Path $outDir 'header-footer.docx'))
if (Test-Path $outPath) { Remove-Item $outPath -Force }

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
    $doc = $word.Documents.Add()
    $doc.TrackRevisions = $false
    $doc.Range(0, 0).Text = "Body paragraph one.`rBody paragraph two."
    $sec = $doc.Sections.Item(1)
    $sec.PageSetup.DifferentFirstPageHeaderFooter = -1
    $sec.Headers.Item(1).Range.Text = "CONFIDENTIAL - Acme Master Agreement"      # wdHeaderFooterPrimary
    $sec.Headers.Item(2).Range.Text = "First page header"                          # wdHeaderFooterFirstPage
    $sec.Footers.Item(1).Range.Text = "Page "
    $footerRange = $sec.Footers.Item(1).Range
    $footerRange.Collapse(0)
    $doc.Fields.Add($footerRange, 33) | Out-Null                                   # wdFieldPage
    $sec.Footers.Item(2).Range.Text = "First page footer"

    [object]$path = $outPath
    [object]$fmt = 16
    $doc.SaveAs2([ref]$path, [ref]$fmt)
    $doc.Close($false)
    Write-Host "Saved $outPath"
} finally { $word.Quit() }
