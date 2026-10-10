param(
    [string]$FixturesDir = "tests/fixtures/word-revision-semantics"
)

# Run with pwsh (PowerShell 7): Windows PowerShell 5.1 hangs in SaveAs2 over COM here.
# Records what desktop Word does on Reject All / Accept All for each paragraph-mark
# revision shape in cases.json, and writes expectations.json (checked in). Run
# `node scripts/generate-word-revision-semantics-matrix.mjs` first.

$ErrorActionPreference = 'Stop'
$FixturesDir = (Resolve-Path $FixturesDir).Path
$cases = Get-Content (Join-Path $FixturesDir 'cases.json') -Raw -Encoding UTF8 | ConvertFrom-Json
$work = Join-Path ([System.IO.Path]::GetTempPath()) ("word-rev-sem-" + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $work | Out-Null

Add-Type -AssemblyName System.IO.Compression.FileSystem

function Convert-Text([string]$t) {
    $t = $t -replace [string][char]7, ''
    $t = $t -replace [string][char]11, "`n"
    $t = $t -replace "`r", ''
    return $t
}

function Get-XmlParagraphCount([string]$path) {
    $zip = [System.IO.Compression.ZipFile]::OpenRead($path)
    try {
        $entry = $zip.GetEntry('word/document.xml')
        $reader = New-Object System.IO.StreamReader($entry.Open())
        $xml = $reader.ReadToEnd(); $reader.Close()
        return ([regex]::Matches($xml, '<w:p[ >/]')).Count
    } finally { $zip.Dispose() }
}

function Get-Snapshot($doc) {
    $items = @()
    $count = $doc.Paragraphs.Count
    for ($i = 1; $i -le $count; $i++) {
        $p = $doc.Paragraphs.Item($i)
        $r = $p.Range
        $raw = [string]$r.Text
        $inTable = [bool]$r.Information(12)
        # End-of-row marks surface as paragraphs with text "\r\a" and no cell; skip them.
        if ($inTable -and $raw -eq ("`r" + [char]7)) {
            $cellCount = 0
            try { $cellCount = $r.Cells.Count } catch { $cellCount = 0 }
            if ($cellCount -eq 0) { continue }
        }
        $items += [ordered]@{
            text = (Convert-Text $raw)
            style = [string]$p.Style.NameLocal
            listType = [int]$r.ListFormat.ListType
            listString = [string]$r.ListFormat.ListString
            inTable = $inTable
        }
    }
    return ,$items
}

function Invoke-Mode($word, [string]$src, [string]$mode, [string]$name) {
    $copy = Join-Path $work "$name-$mode.docx"
    Copy-Item $src $copy
    $doc = $word.Documents.Open($copy, $false, $false)
    try {
        $doc.TrackRevisions = $false
        $before = $doc.Revisions.Count
        if ($mode -eq 'reject') { $doc.Revisions.RejectAll() } else { $doc.Revisions.AcceptAll() }
        $snap = Get-Snapshot $doc
        $out = Join-Path $work "$name-$mode-out.docx"
        $doc.SaveAs2($out, 16)
    } finally { $doc.Close($false) }
    return @{ paragraphs = $snap; xmlParagraphs = (Get-XmlParagraphCount $out); revisionsBefore = $before }
}

$word = New-Object -ComObject Word.Application
try {
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $results = @()
    foreach ($case in $cases) {
        $src = Join-Path $FixturesDir ($case.name + '.docx')
        $rej = Invoke-Mode $word $src 'reject' $case.name
        $acc = Invoke-Mode $word $src 'accept' $case.name
        Write-Output ("{0}: revisions={1} reject={2}p accept={3}p" -f $case.name, $rej.revisionsBefore, $rej.paragraphs.Count, $acc.paragraphs.Count)
        $results += [ordered]@{
            name = $case.name
            description = $case.description
            bodyXml = $case.bodyXml
            revisionsBefore = $rej.revisionsBefore
            reject = $rej.paragraphs
            accept = $acc.paragraphs
            rejectXmlParagraphs = $rej.xmlParagraphs
            acceptXmlParagraphs = $acc.xmlParagraphs
        }
    }
    $doc = [ordered]@{
        wordVersion = [string]$word.Version
        wordBuild = [string]$word.Build
        generatedAt = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
        cases = $results
    }
    $json = $doc | ConvertTo-Json -Depth 8
    [System.IO.File]::WriteAllText((Join-Path $FixturesDir 'expectations.json'), $json + "`n", (New-Object System.Text.UTF8Encoding($false)))
} finally {
    try { $word.Quit(0) } catch {}
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word)
    Remove-Item $work -Recurse -Force -ErrorAction SilentlyContinue
}
