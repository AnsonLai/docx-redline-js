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

# Messages go straight to stdout: Write-Output inside a function that also returns a boolean would make the
# return value a non-empty array, and `-not` of that array is always false (failures would be silently ignored).
function Test-ContainsAssertions([string]$text, $required, $absent, [string]$phase, [string]$name) {
    $passed = $true
    foreach ($value in @($required)) {
        if (-not $text.Contains([string]$value)) {
            [Console]::Out.WriteLine("FAIL ${name}: ${phase} is missing expected text: $value")
            $passed = $false
        }
    }
    foreach ($value in @($absent)) {
        if ($text.Contains([string]$value)) {
            [Console]::Out.WriteLine("FAIL ${name}: ${phase} still contains superseded text: $value")
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

function Get-ScopeText($document, [string]$scope, [string]$fidelity) {
    if ($scope -ne 'headers') {
        return Get-ComparableText $document.Content.Text $fidelity
    }
    $parts = @()
    foreach ($section in @($document.Sections)) {
        foreach ($headerType in 1..3) {
            $header = $section.Headers.Item($headerType)
            if ($header.Exists) {
                $parts += Get-ComparableText $header.Range.Text $fidelity
            }
        }
    }
    return ($parts -join "`n")
}

# Paragraph structure, not just text: an untracked paragraph-mark change can leave the
# Reject All text identical while adding or removing an (empty) paragraph.
function Get-ScopeParagraphCount($document, [string]$scope) {
    if ($scope -ne 'headers') { return $document.Paragraphs.Count }
    $count = 0
    foreach ($section in @($document.Sections)) {
        foreach ($headerType in 1..3) {
            $header = $section.Headers.Item($headerType)
            if ($header.Exists) { $count += $header.Range.Paragraphs.Count }
        }
    }
    return $count
}

# List membership, not just text: Range.Text never contains list labels, so a numbering binding
# (w:numPr in or out of a w:pPrChange) is only visible through ListFormat.ListString per paragraph.
function Test-ListStrings($document, $expectedStrings, [string]$phase, [string]$name) {
    if ($null -eq $expectedStrings) { return $true }
    $actual = @(foreach ($paragraph in @($document.Paragraphs)) { [string]$paragraph.Range.ListFormat.ListString })
    $expectedList = @($expectedStrings | ForEach-Object { [string]$_ })
    if (($actual -join '|') -ne ($expectedList -join '|')) {
        [Console]::Out.WriteLine("FAIL ${name}: $phase list labels differ")
        [Console]::Out.WriteLine("  expected: $($expectedList -join ' | ')")
        [Console]::Out.WriteLine("  actual:   $($actual -join ' | ')")
        return $false
    }
    return $true
}

# List identity, not just labels: a start value can make two separate lists read A. and B., so each paragraph's
# ListFormat.List.ListParagraphs.Count says how many paragraphs share its list (0 = not a list item).
function Test-ListGroupSizes($document, $expectedSizes, [string]$phase, [string]$name) {
    if ($null -eq $expectedSizes) { return $true }
    $actual = @(foreach ($paragraph in @($document.Paragraphs)) {
        $list = $paragraph.Range.ListFormat.List
        if ($null -eq $list) { 0 } else { [int]$list.ListParagraphs.Count }
    })
    $expectedList = @($expectedSizes | ForEach-Object { [int]$_ })
    if (($actual -join '|') -ne ($expectedList -join '|')) {
        [Console]::Out.WriteLine("FAIL ${name}: $phase list group sizes differ")
        [Console]::Out.WriteLine("  expected: $($expectedList -join ' | ')")
        [Console]::Out.WriteLine("  actual:   $($actual -join ' | ')")
        return $false
    }
    return $true
}

# Use the Revisions collection, which matches the ribbon's "Accept/Reject All Changes"
# (CommandBars.ExecuteMso and WordBasic agree). Document.AcceptAllRevisions() /
# RejectAllRevisions() can leave a revision behind: on a document whose last body paragraph
# has a tracked-inserted mark, RejectAllRevisions() keeps that mark and an extra empty
# paragraph (verified on Word 16.0.20430, 2026-10-10).
function Invoke-AllRevisions($document, [string]$scope, [string]$action) {
    $ranges = if ($scope -ne 'headers') { @($document.Content) } else {
        @(foreach ($section in @($document.Sections)) {
            foreach ($headerType in 1..3) {
                $header = $section.Headers.Item($headerType)
                if ($header.Exists) { $header.Range }
            }
        })
    }
    if ($scope -ne 'headers') {
        if ($action -eq 'accept') { $document.Revisions.AcceptAll() } else { $document.Revisions.RejectAll() }
    }
    else {
        foreach ($range in $ranges) {
            if ($action -eq 'accept') { $range.Revisions.AcceptAll() } else { $range.Revisions.RejectAll() }
        }
    }
    return (Get-ScopeRevisionCount $document $scope)
}

function Get-ScopeRevisionCount($document, [string]$scope) {
    if ($scope -ne 'headers') { return $document.Revisions.Count }
    $count = 0
    foreach ($section in @($document.Sections)) {
        foreach ($headerType in 1..3) {
            $header = $section.Headers.Item($headerType)
            if ($header.Exists) { $count += $header.Range.Revisions.Count }
        }
    }
    return $count
}

function Close-WordDocumentSafely($doc) {
    if ($null -eq $doc) { return }
    for ($attempt = 1; $attempt -le 5; $attempt++) {
        try {
            $doc.Close(0) | Out-Null
            break
        }
        catch [System.Runtime.InteropServices.COMException] {
            $hr = $_.Exception.HResult
            # 0x80010001 = RPC_E_CALL_REJECTED (-2147418111)
            if ($hr -eq -2147418111 -and $attempt -lt 5) {
                Start-Sleep -Milliseconds 200
                continue
            }
            break
        }
        catch {
            break
        }
    }
    try {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) | Out-Null
    } catch {}
}

function Stop-WordSafely($application) {
    if ($null -eq $application) { return }
    for ($attempt = 1; $attempt -le 5; $attempt++) {
        try {
            $application.Quit(0) | Out-Null
            break
        }
        catch [System.Runtime.InteropServices.COMException] {
            $hr = $_.Exception.HResult
            # 0x80010001 = RPC_E_CALL_REJECTED (-2147418111)
            # 0x80010108 = RPC_E_DISCONNECTED (-2147417848)
            if ($hr -eq -2147417848) {
                break
            }
            if ($hr -eq -2147418111 -and $attempt -lt 5) {
                Start-Sleep -Milliseconds 300
                continue
            }
            break
        }
        catch {
            break
        }
    }
    try {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($application) | Out-Null
    } catch {}
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
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
        $assertionScope = if ($expected.assertionScope) { [string]$expected.assertionScope } else { 'document' }
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
                $expectedRejected = Get-ScopeText $sourceDocument $assertionScope $fidelity
            }
            finally {
                Close-WordDocumentSafely $sourceDocument
                $sourceDocument = $null
            }
            $replacements = if ($expected.replacements) { @($expected.replacements) } else { @(@{
                originalTarget = $expected.originalTarget
                modifiedTarget = $expected.modifiedTarget
            }) }
            $expectedAccepted = $expectedRejected
            foreach ($replacement in $replacements) {
                $targetCount = [regex]::Matches($expectedAccepted, [regex]::Escape([string]$replacement.originalTarget)).Count
                if ($targetCount -ne 1) {
                    throw "${name}: expected target '$($replacement.originalTarget)' occurs $targetCount times in Word's source $assertionScope text"
                }
                $expectedAccepted = $expectedAccepted.Replace([string]$replacement.originalTarget, [string]$replacement.modifiedTarget)
            }
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
            $revisionCount = Get-ScopeRevisionCount $document $assertionScope
            if ($revisionCount -lt 1) {
                Write-Output "FAIL ${name}: Word sees no tracked revisions"
                $caseFailed = $true
            }
            else {
                $leftAfterAccept = Invoke-AllRevisions $document $assertionScope 'accept'
                if ($leftAfterAccept -ne 0) {
                    Write-Output "FAIL ${name}: $leftAfterAccept revision(s) remain after accept-all"
                    $caseFailed = $true
                }
                $acceptedText = Get-ScopeText $document $assertionScope $fidelity
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
                if (-not $caseFailed -and -not (Test-ListStrings $document $expected.expectedAcceptedListStrings 'accept-all' $name)) {
                    $caseFailed = $true
                }
                if (-not $caseFailed -and -not (Test-ListGroupSizes $document $expected.expectedAcceptedListGroupSizes 'accept-all' $name)) {
                    $caseFailed = $true
                }
            }
        }
        catch {
            Write-Output "FAIL ${name}: Word could not open/accept: $($_.Exception.Message)"
            $caseFailed = $true
        }
        finally {
            Close-WordDocumentSafely $document
            $document = $null
        }

        # Phase 2: fresh open, reject-all restores the original text.
        if (-not $caseFailed) {
            try {
                $document = Open-FixtureDocument $word ([string]$docxPath)
                $leftAfterReject = Invoke-AllRevisions $document $assertionScope 'reject'
                if ($leftAfterReject -ne 0) {
                    Write-Output "FAIL ${name}: $leftAfterReject revision(s) remain after reject-all"
                    $caseFailed = $true
                }
                $rejectedText = Get-ScopeText $document $assertionScope $fidelity
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
                if (-not $caseFailed -and -not (Test-ListStrings $document $expected.expectedRejectedListStrings 'reject-all' $name)) {
                    $caseFailed = $true
                }
                $sourceDocxPath = Join-Path $resolvedDir "$name.source.docx"
                if (-not $caseFailed -and (Test-Path -LiteralPath $sourceDocxPath)) {
                    $rejectedParagraphs = Get-ScopeParagraphCount $document $assertionScope
                    $sourceCountDocument = $null
                    try {
                        $sourceCountDocument = Open-FixtureDocument $word ([string]$sourceDocxPath)
                        $sourceParagraphs = Get-ScopeParagraphCount $sourceCountDocument $assertionScope
                    }
                    finally {
                        Close-WordDocumentSafely $sourceCountDocument
                        $sourceCountDocument = $null
                    }
                    if ($rejectedParagraphs -ne $sourceParagraphs) {
                        Write-Output "FAIL ${name}: reject-all paragraph count $rejectedParagraphs, source has $sourceParagraphs"
                        $caseFailed = $true
                    }
                }
            }
            catch {
                Write-Output "FAIL ${name}: Word could not open/reject: $($_.Exception.Message)"
                $caseFailed = $true
            }
            finally {
                Close-WordDocumentSafely $document
                $document = $null
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
    Stop-WordSafely $word
    $word = $null
}

Write-Output ""
Write-Output "Word differential: $($results.Count - $failures)/$($results.Count) fixtures passed."
if ($failures -gt 0) { exit 1 }
