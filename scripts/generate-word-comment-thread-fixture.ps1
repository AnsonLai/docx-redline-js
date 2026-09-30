# Generates tests/fixtures/word-authored/threaded-comments.docx using real Word.
# The fixture is the oracle for package content types: tests compare our output against
# what Word itself wrote, so the expected values are not derived from our own constants.
$ErrorActionPreference = 'Stop'
$outDir = Join-Path $PSScriptRoot '..\tests\fixtures\word-authored'
New-Item -ItemType Directory -Path $outDir -Force | Out-Null
$outPath = [System.IO.Path]::GetFullPath((Join-Path $outDir 'threaded-comments.docx'))
if (Test-Path $outPath) { Remove-Item $outPath -Force }

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
    $doc = $word.Documents.Add()
    $doc.TrackRevisions = $false
    $doc.Range(0, 0).Text = "The Supplier shall deliver the goods within thirty days.`rPayment is due upon receipt of invoice.`rThis agreement is governed by local law."

    # Thread 1: comment with a reply (open)
    $c1 = $doc.Comments.Add($doc.Paragraphs.Item(1).Range, "Can we shorten the delivery window?")
    $c1.Author = "Counterparty"
    $r1 = $c1.Replies.Add($c1.Scope, "We can offer twenty days.")
    $r1.Author = "Internal"

    # Thread 2: resolved comment
    $c2 = $doc.Comments.Add($doc.Paragraphs.Item(2).Range, "Confirm payment terms.")
    $c2.Author = "Counterparty"
    $c2.Done = $true

    [object]$path = $outPath
    [object]$fmt = 16
    $doc.SaveAs2([ref]$path, [ref]$fmt)
    $doc.Close($false)
    Write-Host "Saved $outPath"
} finally {
    $word.Quit()
}
