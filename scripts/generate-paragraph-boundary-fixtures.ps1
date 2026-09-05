$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem

$fixturesDir = "C:\Users\Phara\Desktop\Projects\Docx Redline JS\tests\fixtures\paragraph-boundaries"
if (-not (Test-Path $fixturesDir)) {
    New-Item -ItemType Directory -Path $fixturesDir -Force | Out-Null
}

function Extract-DocumentXml($docxPath, $xmlPath) {
    $zip = [System.IO.Compression.ZipFile]::OpenRead($docxPath)
    try {
        $entry = $zip.GetEntry("word/document.xml")
        if ($entry) {
            $reader = New-Object System.IO.StreamReader($entry.Open(), [System.Text.Encoding]::UTF8)
            try {
                $xml = $reader.ReadToEnd()
                [System.IO.File]::WriteAllText($xmlPath, $xml, [System.Text.Encoding]::UTF8)
            } finally {
                $reader.Dispose()
            }
        }
    } finally {
        $zip.Dispose()
    }
}

function Save-Triple($doc, [string]$baseName) {
    $pendingDocx = [string](Join-Path $fixturesDir "$baseName-pending.docx")
    $acceptedDocx = [string](Join-Path $fixturesDir "$baseName-accepted.docx")
    $rejectedDocx = [string](Join-Path $fixturesDir "$baseName-rejected.docx")
    
    [object]$pRef = $pendingDocx
    [object]$aRef = $acceptedDocx
    [object]$rRef = $rejectedDocx
    [object]$fmt = 16
    
    # Save pending
    $doc.SaveAs2([ref]$pRef, [ref]$fmt)
    
    # Accept All
    $doc.Revisions.AcceptAll()
    $doc.SaveAs2([ref]$aRef, [ref]$fmt)
    $doc.Close($false)
    
    # Reopen pending and Reject All
    $reopened = $global:word.Documents.Open([ref]$pRef)
    $reopened.Revisions.RejectAll()
    $reopened.SaveAs2([ref]$rRef, [ref]$fmt)
    $reopened.Close($false)
    
    # Extract XML
    Extract-DocumentXml $pendingDocx (Join-Path $fixturesDir "$baseName-pending.xml")
    Extract-DocumentXml $acceptedDocx (Join-Path $fixturesDir "$baseName-accepted.xml")
    Extract-DocumentXml $rejectedDocx (Join-Path $fixturesDir "$baseName-rejected.xml")
    Write-Host "Saved: $baseName"
}

$global:word = New-Object -ComObject Word.Application
$global:word.Visible = $false
$global:word.DisplayAlerts = 0

try {
    # 1. split-middle
    Write-Host "1. split-middle"
    $doc = $global:word.Documents.Add()
    $doc.TrackRevisions = $false
    $doc.Range(0, 0).Text = "Sentence one. Sentence two."
    $doc.TrackRevisions = $true
    $splitPoint = $doc.Range(14, 14)
    $splitPoint.InsertParagraph()
    Save-Triple $doc "split-middle"

    # 2. delete-boundary
    Write-Host "2. delete-boundary"
    $doc = $global:word.Documents.Add()
    $doc.TrackRevisions = $false
    $r = $doc.Range(0, 0)
    $r.Text = "Paragraph one."
    $r.InsertParagraphAfter()
    $doc.Paragraphs.Item(2).Range.Text = "Paragraph two."
    $doc.TrackRevisions = $true
    $p1 = $doc.Paragraphs.Item(1).Range
    $delPoint = $doc.Range($p1.End - 1, $p1.End)
    $delPoint.Delete() | Out-Null
    Save-Triple $doc "delete-boundary"

    # 3. delete-middle-paragraph
    Write-Host "3. delete-middle-paragraph"
    $doc = $global:word.Documents.Add()
    $doc.TrackRevisions = $false
    $r = $doc.Range(0, 0)
    $r.Text = "Paragraph one."
    $r.InsertParagraphAfter()
    $doc.Paragraphs.Item(2).Range.Text = "Paragraph two."
    $doc.Paragraphs.Item(2).Range.InsertParagraphAfter()
    $doc.Paragraphs.Item(3).Range.Text = "Paragraph three."
    $doc.TrackRevisions = $true
    $p2 = $doc.Paragraphs.Item(2).Range
    $p2.Delete() | Out-Null
    Save-Triple $doc "delete-middle-paragraph"

    # 4. insert-blank-paragraph
    Write-Host "4. insert-blank-paragraph"
    $doc = $global:word.Documents.Add()
    $doc.TrackRevisions = $false
    $r = $doc.Range(0, 0)
    $r.Text = "Paragraph one."
    $r.InsertParagraphAfter()
    $doc.Paragraphs.Item(2).Range.Text = "Paragraph two."
    $doc.TrackRevisions = $true
    $p1 = $doc.Paragraphs.Item(1).Range
    $insPoint = $doc.Range($p1.End - 1, $p1.End - 1)
    $insPoint.InsertParagraph()
    Save-Triple $doc "insert-blank-paragraph"

    # 5. different-styles-boundary
    Write-Host "5. different-styles-boundary"
    $doc = $global:word.Documents.Add()
    $doc.TrackRevisions = $false
    $p1 = $doc.Paragraphs.Item(1).Range
    $p1.Text = "Heading title"
    $p1.Style = "Heading 1"
    $p1.InsertParagraphAfter()
    $p2 = $doc.Paragraphs.Item(2).Range
    $p2.Text = "Normal body text."
    $p2.Style = "Normal"
    $doc.TrackRevisions = $true
    $p1End = $doc.Paragraphs.Item(1).Range.End
    $delPoint = $doc.Range($p1End - 1, $p1End)
    $delPoint.Delete() | Out-Null
    Save-Triple $doc "different-styles-boundary"

    # 6. different-list-levels-boundary
    Write-Host "6. different-list-levels-boundary"
    $doc = $global:word.Documents.Add()
    $doc.TrackRevisions = $false
    $p1 = $doc.Paragraphs.Item(1).Range
    $p1.Text = "List item 1"
    $p1.InsertParagraphAfter()
    $p2 = $doc.Paragraphs.Item(2).Range
    $p2.Text = "List item 2"
    $doc.Paragraphs.Item(1).Range.ListFormat.ApplyBulletDefault()
    $doc.Paragraphs.Item(2).Range.ListFormat.ApplyBulletDefault()
    $doc.Paragraphs.Item(2).Range.ListFormat.ListIndent()
    $doc.TrackRevisions = $true
    $p1End = $doc.Paragraphs.Item(1).Range.End
    $delPoint = $doc.Range($p1End - 1, $p1End)
    $delPoint.Delete() | Out-Null
    Save-Triple $doc "different-list-levels-boundary"

    # 7. section-break-boundary
    Write-Host "7. section-break-boundary"
    $doc = $global:word.Documents.Add()
    $doc.TrackRevisions = $false
    $p1 = $doc.Paragraphs.Item(1).Range
    $p1.Text = "Section one text."
    $p1.InsertBreak(2) # wdSectionBreakNextPage
    $p2 = $doc.Paragraphs.Item(2).Range
    $p2.Text = "Section two text."
    $doc.TrackRevisions = $true
    $p1End = $doc.Paragraphs.Item(1).Range.End
    $delPoint = $doc.Range($p1End - 1, $p1End)
    $delPoint.Delete() | Out-Null
    Save-Triple $doc "section-break-boundary"

    # 8. table-cell-boundary
    Write-Host "8. table-cell-boundary"
    $doc = $global:word.Documents.Add()
    $doc.TrackRevisions = $false
    $table = $doc.Tables.Add($doc.Range(0, 0), 1, 1)
    $cell = $table.Cell(1, 1)
    $cell.Range.Text = "Cell paragraph 1"
    $cell.Range.InsertParagraphAfter()
    $p2 = $doc.Paragraphs.Item(2).Range
    $p2.Text = "Cell paragraph 2"
    $doc.TrackRevisions = $true
    $p1End = $doc.Paragraphs.Item(1).Range.End
    $delPoint = $doc.Range($p1End - 1, $p1End)
    $delPoint.Delete() | Out-Null
    Save-Triple $doc "table-cell-boundary"

    # 9. adjacent-bookmark-comment
    Write-Host "9. adjacent-bookmark-comment"
    $doc = $global:word.Documents.Add()
    $doc.TrackRevisions = $false
    $p1 = $doc.Paragraphs.Item(1).Range
    $p1.Text = "Paragraph with bookmark."
    $p1.InsertParagraphAfter()
    $p2 = $doc.Paragraphs.Item(2).Range
    $p2.Text = "Second paragraph."
    $bmRange = $doc.Range(0, 9)
    $doc.Bookmarks.Add("TestBookmark", $bmRange) | Out-Null
    $doc.Comments.Add($bmRange, "Comment on first paragraph") | Out-Null
    $doc.TrackRevisions = $true
    $p1End = $doc.Paragraphs.Item(1).Range.End
    $delPoint = $doc.Range($p1End - 1, $p1End)
    $delPoint.Delete() | Out-Null
    Save-Triple $doc "adjacent-bookmark-comment"

    # 10. multi-author-boundary
    Write-Host "10. multi-author-boundary"
    $doc = $global:word.Documents.Add()
    $doc.TrackRevisions = $false
    $doc.Range(0, 0).Text = "Author one text. Author two text."
    $doc.TrackRevisions = $true
    $global:word.UserName = "AuthorOne"
    $doc.Range(16, 16).InsertParagraph()
    $global:word.UserName = "AuthorTwo"
    $doc.Paragraphs.Item(2).Range.InsertBefore("Updated ")
    Save-Triple $doc "multi-author-boundary"

    Write-Host "All 10 Word paragraph boundary golden fixtures generated successfully!"
} finally {
    $global:word.Quit()
}
