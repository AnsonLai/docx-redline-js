# Generates real-Word fixtures for comment thread behavior (tests/fixtures/word-authored/):
#   multi-paragraph-thread.docx  a three-paragraph comment with one reply. Word keys commentsExtended /
#                                commentsIds on the LAST paragraph of the comment.
#   resolved-threads.docx        thread 1: root + two replies, resolved by setting Done on the ROOT.
#                                thread 2: root + reply, resolved by setting Done on the REPLY.
#                                Word marks the WHOLE thread done in both cases.
#   reply-to-resolved.docx       a reply added to an already-resolved thread. Word writes the reply done too.
$ErrorActionPreference = 'Stop'
$outDir = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\tests\fixtures\word-authored'))
New-Item -ItemType Directory -Path $outDir -Force | Out-Null
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0

function Save-Doc($doc, $name) {
    $path = [string](Join-Path $outDir $name)
    if (Test-Path $path) { Remove-Item $path -Force }
    [object]$p = $path; [object]$f = 16
    $doc.SaveAs2([ref]$p, [ref]$f)
    $doc.Close($false)
    Write-Host "Saved $path"
}

try {
    $doc = $word.Documents.Add()
    $doc.Range(0, 0).Text = "Alpha paragraph."
    $c = $doc.Comments.Add($doc.Paragraphs.Item(1).Range, "first para of comment`rsecond para of comment`rthird para")
    $c.Author = "Counterparty"
    $r = $c.Replies.Add($c.Scope, "a reply")
    $r.Author = "Internal"
    Save-Doc $doc 'multi-paragraph-thread.docx'

    $doc = $word.Documents.Add()
    $doc.Range(0, 0).Text = "Alpha paragraph.`rBeta paragraph."
    $c1 = $doc.Comments.Add($doc.Paragraphs.Item(1).Range, "root one"); $c1.Author = "A"
    $r1 = $c1.Replies.Add($c1.Scope, "reply one"); $r1.Author = "B"
    $r2 = $c1.Replies.Add($c1.Scope, "reply two"); $r2.Author = "C"
    $c1.Done = $true
    $c2 = $doc.Comments.Add($doc.Paragraphs.Item(2).Range, "root two"); $c2.Author = "A"
    $rr = $c2.Replies.Add($c2.Scope, "reply to two"); $rr.Author = "B"
    $rr.Done = $true
    Save-Doc $doc 'resolved-threads.docx'

    $doc = $word.Documents.Add()
    $doc.Range(0, 0).Text = "Alpha paragraph."
    $c = $doc.Comments.Add($doc.Paragraphs.Item(1).Range, "root"); $c.Author = "A"
    $c.Done = $true
    $r = $c.Replies.Add($c.Scope, "late reply"); $r.Author = "B"
    Save-Doc $doc 'reply-to-resolved.docx'
} finally { $word.Quit() }
