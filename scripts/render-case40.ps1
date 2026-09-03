$ErrorActionPreference = 'Stop'
$fixtureDir = (Resolve-Path 'tmp/superdoc-word-fixtures').Path
$visualDir = (Resolve-Path 'tmp/multilevel-bullet-visual').Path
$baseName = '40-administrative-administrative-list-change-ppg-agenda-addition'

function Export-Docx([string]$docxName, [string]$pdfName, [bool]$isMarkup) {
    Write-Output "Exporting $docxName to $pdfName..."
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    try {
        $doc = $word.Documents.Open("$fixtureDir/$docxName")
        if ($isMarkup) {
            $doc.ShowRevisions = $true
            $doc.PrintRevisions = $true
            $doc.ExportAsFixedFormat("$visualDir/$pdfName", 17, $false, 0, 0, 1, 1, 7, $true, $true, 0, $true, $true, $false)
        } else {
            $doc.ExportAsFixedFormat("$visualDir/$pdfName", 17, $false, 0, 0, 1, 1, 0, $true, $true, 0, $true, $true, $false)
        }
        Start-Sleep -Milliseconds 800
        try { $doc.Close(0) } catch {}
    } finally {
        try { $word.Quit(0) } catch {}
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    }
}

if (!(Test-Path "$visualDir/$baseName--allMarkup.pdf")) {
    Export-Docx "$baseName.docx" "$baseName--allMarkup.pdf" $true
    Start-Sleep -Seconds 1
}
Export-Docx "$baseName.accepted.docx" "$baseName--acceptAll.pdf" $false
Start-Sleep -Seconds 1
Export-Docx "$baseName.rejected.docx" "$baseName--rejectAll.pdf" $false
Write-Output 'All views exported successfully!'
