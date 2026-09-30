# Generates tests/fixtures/word-authored/comment-edge-cases.docx using real Word: comments on middle words,
# across two paragraphs, inside a table cell, on text with another author's tracked insertion, and on a hyperlink.
$ErrorActionPreference = 'Stop'
$outDir = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\tests\fixtures\word-authored'))
New-Item -ItemType Directory -Path $outDir -Force | Out-Null
$path = [string](Join-Path $outDir 'comment-edge-cases.docx')
if (Test-Path $path) { Remove-Item $path -Force }
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
    $doc = $word.Documents.Add()
    $doc.TrackRevisions = $false
    $doc.Range(0, 0).Text = "Alpha beta gamma delta epsilon.`rThe counterparty proposed tenure.`rSpanning first paragraph.`rSpanning second paragraph.`rLink here: example.com now.`rTable follows."

    # 1: comment on the middle words only
    $r = $doc.Paragraphs.Item(1).Range
    $mid = $doc.Range($r.Start + 6, $r.Start + 16)   # "beta gamma"
    $c = $doc.Comments.Add($mid, "Middle words comment"); $c.Author = "Reviewer"

    # 2: another author's tracked insertion inside a commented paragraph
    $word.UserName = "Other Author"
    $doc.TrackRevisions = $true
    $p2 = $doc.Paragraphs.Item(2).Range
    $ins = $doc.Range($p2.End - 8, $p2.End - 8)      # before "tenure."
    $ins.InsertAfter("long-term ")
    $doc.TrackRevisions = $false
    $word.UserName = "Reviewer"
    $c = $doc.Comments.Add($doc.Paragraphs.Item(2).Range, "Comment over a tracked insertion"); $c.Author = "Reviewer"

    # 3: one comment spanning two paragraphs
    $span = $doc.Range($doc.Paragraphs.Item(3).Range.Start, $doc.Paragraphs.Item(4).Range.End - 1)
    $c = $doc.Comments.Add($span, "Spans two paragraphs"); $c.Author = "Reviewer"

    # 4: comment on a hyperlink
    $p5 = $doc.Paragraphs.Item(5).Range
    $linkRange = $doc.Range($p5.Start + 11, $p5.Start + 22)   # "example.com"
    $doc.Hyperlinks.Add($linkRange, "https://example.com") | Out-Null
    $p5 = $doc.Paragraphs.Item(5).Range
    $c = $doc.Comments.Add($p5, "On a paragraph with a hyperlink"); $c.Author = "Reviewer"

    # 5: a table with a commented cell
    $end = $doc.Content; $end.InsertParagraphAfter()
    $tableRange = $doc.Paragraphs.Item($doc.Paragraphs.Count).Range
    $table = $doc.Tables.Add($tableRange, 2, 2)
    $table.Cell(1, 1).Range.Text = "Cell one"
    $table.Cell(1, 2).Range.Text = "Cell two"
    $table.Cell(2, 1).Range.Text = "Cell three"
    $table.Cell(2, 2).Range.Text = "Cell four"
    $cellText = $table.Cell(1, 2).Range
    $cellRange = $doc.Range($cellText.Start, $cellText.End - 1)
    $c = $doc.Comments.Add($cellRange, "Comment in a table cell"); $c.Author = "Reviewer"

    [object]$p = $path; [object]$f = 16
    $doc.SaveAs2([ref]$p, [ref]$f)
    $doc.Close($false)
    Write-Host "Saved $path"
} finally { $word.Quit() }
