param(
    [string]$FixturesDir = "tmp/superdoc-word-fixtures",
    [string]$SourcesDir = "tmp/superdoc-corpus"
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

function Get-UntouchedPartHashes([string]$path, [string]$changedPart) {
    $hashes = @{}
    $zip = [IO.Compression.ZipFile]::OpenRead($path)
    try {
        foreach ($entry in $zip.Entries) {
            if ($entry.FullName -eq $changedPart) { continue }
            $stream = $entry.Open()
            $sha = [Security.Cryptography.SHA256]::Create()
            try {
                $hashes[$entry.FullName] = [BitConverter]::ToString($sha.ComputeHash($stream)).Replace('-', '').ToLowerInvariant()
            }
            finally {
                $sha.Dispose()
                $stream.Dispose()
            }
        }
    }
    finally { $zip.Dispose() }
    return $hashes
}

$resolvedFixtures = (Resolve-Path -LiteralPath $FixturesDir).Path
$resolvedSources = (Resolve-Path -LiteralPath $SourcesDir).Path
$suite = Get-Content -LiteralPath (Join-Path $resolvedFixtures 'suite.json') -Raw -Encoding UTF8 | ConvertFrom-Json

foreach ($case in $suite.cases) {
    $sourcePath = Join-Path $resolvedSources "$($case.sourceId).docx"
    $outputPath = Join-Path $resolvedFixtures "$($case.name).docx"
    $xmlPath = Join-Path $resolvedFixtures "$($case.name).document.xml"
    $revisionPart = if ($case.revisionPart) { [string]$case.revisionPart } else { 'word/document.xml' }
    $before = Get-UntouchedPartHashes $sourcePath $revisionPart

    Copy-Item -LiteralPath $sourcePath -Destination $outputPath -Force
    $zip = [IO.Compression.ZipFile]::Open($outputPath, [IO.Compression.ZipArchiveMode]::Update)
    try {
        $oldEntry = $zip.GetEntry($revisionPart)
        if ($null -eq $oldEntry) { throw "Missing $revisionPart in $sourcePath" }
        $oldEntry.Delete()
        $newEntry = $zip.CreateEntry($revisionPart, [IO.Compression.CompressionLevel]::Optimal)
        $stream = $newEntry.Open()
        try {
            $bytes = [Text.UTF8Encoding]::new($false).GetBytes((Get-Content -LiteralPath $xmlPath -Raw -Encoding UTF8))
            $stream.Write($bytes, 0, $bytes.Length)
        }
        finally { $stream.Dispose() }
    }
    finally { $zip.Dispose() }

    $after = Get-UntouchedPartHashes $outputPath $revisionPart
    if ($before.Count -ne $after.Count) { throw "$($case.name): package part count changed" }
    foreach ($name in $before.Keys) {
        if ($after[$name] -ne $before[$name]) { throw "$($case.name): untouched package part changed: $name" }
    }
    Write-Output "Packaged $($case.name) (verified $($before.Count) untouched parts)"
}
