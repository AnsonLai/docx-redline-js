# Generates tests/fixtures/word-authored/multi-section-headers.docx using real Word: two sections where the second
# section has its own header (not linked to previous) but shares the first section's footer.
$ErrorActionPreference = 'Stop'
$outDir = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\tests\fixtures\word-authored'))
New-Item -ItemType Directory -Path $outDir -Force | Out-Null
$path = [string](Join-Path $outDir 'multi-section-headers.docx')
if (Test-Path $path) { Remove-Item $path -Force }
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
    $doc = $word.Documents.Add()
    $doc.Range(0, 0).Text = "Section one body."
    $doc.Sections.Item(1).Headers.Item(1).Range.Text = "Section one header"
    $doc.Sections.Item(1).Footers.Item(1).Range.Text = "Shared footer"
    $end = $doc.Content
    $end.InsertParagraphAfter()
    $doc.Paragraphs.Item($doc.Paragraphs.Count).Range.InsertBreak(2)   # wdSectionBreakNextPage
    $doc.Paragraphs.Item($doc.Paragraphs.Count).Range.InsertBefore("Section two body.")
    $sec2 = $doc.Sections.Item(2)
    $sec2.Headers.Item(1).LinkToPrevious = $false
    $sec2.Headers.Item(1).Range.Text = "Section two header"
    [object]$p = $path; [object]$f = 16
    $doc.SaveAs2([ref]$p, [ref]$f)
    $doc.Close($false)
    Write-Host "Saved $path"
} finally { $word.Quit() }
