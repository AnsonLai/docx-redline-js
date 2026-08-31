param(
    [string]$FixturesDir = "tmp/validation-docx",
    [string]$SourcesDir = ""
)

# Differential validation against desktop Microsoft Word (the authoritative
# OOXML consumer). For each generated fixture, Word itself accepts all
# revisions and then rejects all revisions, and the resulting document text
# is compared to the expected outcomes recorded at generation time. This
# makes Word an independent oracle for the redline engine instead of
# verifying the library against its own accept/reject transforms.
#
# Usage:
#   node scripts/export-validation-fixtures.mjs
#   npm run smoke:word:diff

$ErrorActionPreference = 'Stop'

function Get-ComparableText([string]$text, [string]$fidelity = 'exact') {
    if ($null -eq $text) { return '' }
    $text = $text -replace [string][char]7, ''   # table cell markers
    $text = $text -replace [string][char]2, ''   # footnote/endnote reference markers
    $text = $text -replace "`r`n", "`n"
    $text = $text -replace "`r", "`n"
    $text = $text -replace "`n+$", ''           # Word's terminal paragraph mark
    if ($fidelity -eq 'normalized') {
        return ($text -replace '\s+', ' ').Trim()
    }
    return $text
}

function Test-ContainsAssertions([string]$text, $required, $absent, [string]$phase, [string]$name) {
    $passed = $true
    foreach ($value in @($required)) {
        if (-not $text.Contains([string]$value)) {
            Write-Output "FAIL ${name}: ${phase} is missing expected text: $value"
            $passed = $false
        }
    }
    foreach ($value in @($absent)) {
        if ($text.Contains([string]$value)) {
            Write-Output "FAIL ${name}: ${phase} still contains superseded text: $value"
            $passed = $false
        }
    }
    return $passed
}

function Test-UntouchedPartHashes([string]$docxPath, $expectedHashes, [string]$name) {
    $properties = @($expectedHashes.PSObject.Properties)
    if ($properties.Count -eq 0) { return $true }

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $archive = [System.IO.Compression.ZipFile]::OpenRead($docxPath)
    $sha256 = [System.Security.Cryptography.SHA256]::Create()
    try {
        foreach ($property in $properties) {
            $entry = $archive.GetEntry([string]$property.Name)
            if ($null -eq $entry) {
                Write-Output "FAIL ${name}: package is missing untouched part $($property.Name)"
                return $false
            }
            $stream = $entry.Open()
            try {
                $actual = ([System.BitConverter]::ToString($sha256.ComputeHash($stream))).Replace('-', '').ToLowerInvariant()
            }
            finally {
                $stream.Dispose()
            }
            if ($actual -ne [string]$property.Value) {
                Write-Output "FAIL ${name}: untouched part hash changed for $($property.Name)"
                return $false
            }
        }
    }
    finally {
        $sha256.Dispose()
        $archive.Dispose()
    }
    return $true
}

$resolvedDir = Resolve-Path -LiteralPath $FixturesDir -ErrorAction SilentlyContinue
if (-not $resolvedDir) {
    Write-Error "Fixtures directory '$FixturesDir' not found. Run: node scripts/export-validation-fixtures.mjs"
    exit 1
}

$suitePath = Join-Path $resolvedDir 'suite.json'
if (Test-Path -LiteralPath $suitePath) {
    $suite = Get-Content -LiteralPath $suitePath -Raw -Encoding UTF8 | ConvertFrom-Json
    $expectations = @($suite.cases | ForEach-Object {
        $caseName = if ($_ -is [string]) { [string]$_ } else { [string]$_.name }
        Get-Item -LiteralPath (Join-Path $resolvedDir "$caseName.expected.json") -ErrorAction Stop
    })
}
else {
    $expectations = @(Get-ChildItem -LiteralPath $resolvedDir -Filter '*.expected.json' | Sort-Object Name)
}
if ($expectations.Count -eq 0) {
    Write-Error "No *.expected.json fixtures in '$resolvedDir'. Run: node scripts/export-validation-fixtures.mjs"
    exit 1
}

$word = $null
$failures = 0
$results = @()

function Open-FixtureDocument($word, [string]$path) {
    # OpenNoRepairDialog prevents malformed real-world packages from blocking
    # a headless run behind an invisible repair prompt. It still leaves the
    # document writable for accept/reject; nothing is ever saved.
    return $word.Documents.OpenNoRepairDialog($path)
}

try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0  # wdAlertsNone

    foreach ($expectationFile in $expectations) {
        $name = $expectationFile.BaseName -replace '\.expected$', ''
        $docxPath = Join-Path $resolvedDir "$name.docx"
        if (-not (Test-Path -LiteralPath $docxPath)) {
            Write-Warning "SKIP ${name}: no matching .docx"
            continue
        }

        # -Encoding UTF8 is required: Windows PowerShell 5.1 otherwise reads
        # BOM-less UTF-8 sidecars as ANSI and garbles non-ASCII expectations.
        $expected = Get-Content -LiteralPath $expectationFile.FullName -Raw -Encoding UTF8 | ConvertFrom-Json
        $fidelity = if ($expected.textFidelity) { [string]$expected.textFidelity } else { 'exact' }
        $assertionMode = if ($expected.assertionMode) { [string]$expected.assertionMode } else { 'exact' }
        if ($assertionMode -eq 'exact') {
            $expectedAccepted = Get-ComparableText $expected.expectedAcceptedText $fidelity
            $expectedRejected = Get-ComparableText $expected.expectedRejectedText $fidelity
        }
        elseif ($assertionMode -eq 'word-source-exact') {
            if (-not $SourcesDir) { throw "${name}: word-source-exact requires -SourcesDir" }
            $sourcePath = Join-Path (Resolve-Path -LiteralPath $SourcesDir) "$($expected.sourceId).docx"
            $sourceDocument = $null
            try {
                $sourceDocument = Open-FixtureDocument $word ([string]$sourcePath)
                $expectedRejected = Get-ComparableText $sourceDocument.Content.Text $fidelity
            }
            finally {
                if ($null -ne $sourceDocument) { $sourceDocument.Close(0) | Out-Null }
            }
            $targetCount = [regex]::Matches($expectedRejected, [regex]::Escape([string]$expected.originalTarget)).Count
            if ($targetCount -ne 1) {
                throw "${name}: expected target occurs $targetCount times in Word's source text"
            }
            $expectedAccepted = $expectedRejected.Replace([string]$expected.originalTarget, [string]$expected.modifiedTarget)
        }
        $caseFailed = $false
        $document = $null

        if ($expected.untouchedPartSha256 -and -not (Test-UntouchedPartHashes $docxPath $expected.untouchedPartSha256 $name)) {
            $failures++
            $results += "FAIL $name"
            continue
        }

        # Phase 1: open cleanly, revisions present, accept-all matches intent.
        try {
            $document = Open-FixtureDocument $word ([string]$docxPath)
            $revisionCount = $document.Revisions.Count
            if ($revisionCount -lt 1) {
                Write-Output "FAIL ${name}: Word sees no tracked revisions"
                $caseFailed = $true
            }
            else {
                $document.AcceptAllRevisions()
                $acceptedText = Get-ComparableText $document.Content.Text $fidelity
                if ($assertionMode -eq 'contains') {
                    if (-not (Test-ContainsAssertions $acceptedText $expected.expectedAcceptedContains $expected.expectedAcceptedAbsent 'accept-all' $name)) {
                        $caseFailed = $true
                    }
                }
                elseif ($acceptedText -ne $expectedAccepted) {
                    Write-Output "FAIL ${name}: accept-all mismatch"
                    Write-Output "  expected: $expectedAccepted"
                    Write-Output "  actual:   $acceptedText"
                    $caseFailed = $true
                }
            }
        }
        catch {
            Write-Output "FAIL ${name}: Word could not open/accept: $($_.Exception.Message)"
            $caseFailed = $true
        }
        finally {
            if ($null -ne $document) { $document.Close(0) | Out-Null; $document = $null }
        }

        # Phase 2: fresh open, reject-all restores the original text.
        if (-not $caseFailed) {
            try {
                $document = Open-FixtureDocument $word ([string]$docxPath)
                $document.RejectAllRevisions()
                $rejectedText = Get-ComparableText $document.Content.Text $fidelity
                if ($assertionMode -eq 'contains') {
                    if (-not (Test-ContainsAssertions $rejectedText $expected.expectedRejectedContains $expected.expectedRejectedAbsent 'reject-all' $name)) {
                        $caseFailed = $true
                    }
                }
                elseif ($rejectedText -ne $expectedRejected) {
                    Write-Output "FAIL ${name}: reject-all mismatch"
                    Write-Output "  expected: $expectedRejected"
                    Write-Output "  actual:   $rejectedText"
                    $caseFailed = $true
                }
            }
            catch {
                Write-Output "FAIL ${name}: Word could not open/reject: $($_.Exception.Message)"
                $caseFailed = $true
            }
            finally {
                if ($null -ne $document) { $document.Close(0) | Out-Null; $document = $null }
            }
        }

        if ($caseFailed) {
            $failures++
            $results += "FAIL $name"
        }
        else {
            Write-Output "PASS ${name} (revisions: $revisionCount)"
            $results += "PASS $name"
        }
    }
}
finally {
    if ($null -ne $word) { $word.Quit() | Out-Null }
}

Write-Output ""
Write-Output "Word differential: $($results.Count - $failures)/$($results.Count) fixtures passed."
if ($failures -gt 0) { exit 1 }
