param(
    [string]$FixturesDir = "",
    [string]$OutputDir = "tmp/superdoc-word-visual-review/rendered",
    [string[]]$Case = @()
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
if (-not $FixturesDir) { $FixturesDir = "tmp/superdoc-word-visual-review/fixtures-$PID" }
$fixturesPath = [System.IO.Path]::GetFullPath((Join-Path $repoRoot $FixturesDir))
$outputPath = [System.IO.Path]::GetFullPath((Join-Path $repoRoot $OutputDir))
$manifestPath = Join-Path $outputPath 'manifest.json'
New-Item -ItemType Directory -Force -Path $fixturesPath, $outputPath | Out-Null

& (Join-Path $PSScriptRoot 'word-com-corpus-suite.ps1') -FixturesDir $fixturesPath
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

$prepareArgs = @((Join-Path $PSScriptRoot 'prepare-corpus-word-visual-review.mjs'), "--output=$manifestPath")
foreach ($caseName in $Case) { $prepareArgs += "--case=$caseName" }
& node @prepareArgs
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

$manifest = Get-Content -LiteralPath $manifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
$word = $null
$rpcDisconnected = -2147417848 # 0x80010108 RPC_E_DISCONNECTED

function Test-RpcDisconnected($exception) {
    return $exception -is [Runtime.InteropServices.COMException] -and $exception.HResult -eq $rpcDisconnected
}

function Start-WordApplication {
    $application = New-Object -ComObject Word.Application
    $application.Visible = $false
    $application.DisplayAlerts = 0
    return $application
}

function Stop-WordApplicationSafely($application) {
    if ($null -eq $application) { return }
    try { $application.Quit() | Out-Null }
    catch {
        if (-not (Test-RpcDisconnected $_.Exception)) { throw }
        Write-Warning 'Word disconnected while quitting; the COM process was already gone.'
    }
}

function Close-WordDocumentSafely($document) {
    if ($null -eq $document) { return }
    try { $document.Close(0) | Out-Null }
    catch {
        if (-not (Test-RpcDisconnected $_.Exception)) { throw }
        # ExportAsFixedFormat occasionally tears down Word's document proxy
        # after successfully writing a PDF. Do not let cleanup hide the valid
        # render; the next operation will restart Word if needed.
        Write-Warning 'Word disconnected after export while closing the document.'
    }
}

function Export-View($word, [string]$docxPath, [string]$pdfPath, [string]$viewName) {
    $document = $null
    try {
        $document = $word.Documents.OpenNoRepairDialog($docxPath)
        if ($viewName -eq 'acceptAll') { $document.AcceptAllRevisions(); $exportItem = 0 }
        elseif ($viewName -eq 'rejectAll') { $document.RejectAllRevisions(); $exportItem = 0 }
        else { $document.ShowRevisions = $true; $document.PrintRevisions = $true; $exportItem = 7 }
        $pages = $document.ComputeStatistics(2)
        $document.ExportAsFixedFormat($pdfPath, 17, $false, 0, 0, 1, 1, $exportItem, $true, $true, 0, $true, $true, $false)
        $pdf = Get-Item -LiteralPath $pdfPath -ErrorAction Stop
        if ($pdf.Length -lt 1000 -or $pages -lt 1) { throw "Word produced invalid visual evidence" }
        return @{ pages = [int]$pages; bytes = [long]$pdf.Length }
    }
    finally { Close-WordDocumentSafely $document }
}

try {
    $word = Start-WordApplication
    $manifest.word.version = [string]$word.Version
    try { $manifest.word.build = [string]$word.Build } catch { $manifest.word.build = $null }
    foreach ($entry in $manifest.cases) {
        $docxPath = Join-Path $fixturesPath "$($entry.name).docx"
        foreach ($viewProperty in $entry.views.PSObject.Properties) {
            $pdfPath = Join-Path $outputPath ([string]$viewProperty.Value.pdf)
            $render = $null
            for ($attempt = 1; $attempt -le 2; $attempt++) {
                try {
                    $render = Export-View $word $docxPath $pdfPath ([string]$viewProperty.Name)
                    break
                }
                catch {
                    if (-not (Test-RpcDisconnected $_.Exception) -or $attempt -eq 2) { throw }
                    Write-Warning "Word disconnected while rendering $($entry.identity) $($viewProperty.Name); restarting and retrying once."
                    Stop-WordApplicationSafely $word
                    $word = Start-WordApplication
                }
            }
            $viewProperty.Value.status = 'rendered'
            $viewProperty.Value.pages = $render.pages
            $viewProperty.Value.bytes = $render.bytes
        }
        $entry.renderStatus = 'rendered'
        Write-Output "RENDERED $($entry.identity) (markup, accept-all, reject-all)"
    }
}
finally {
    Stop-WordApplicationSafely $word
    $manifest | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $manifestPath -Encoding UTF8
}

Write-Output "Corpus visual evidence: $($manifest.cases.Count) cases, $($manifest.cases.Count * 3) PDFs."
Write-Output "Manifest: $manifestPath"

# Keep the self-contained comparison dashboard synchronized with the exact
# process-isolated DOCX fixtures that Word just validated and rendered.
& node (Join-Path $PSScriptRoot 'build-test-dashboard.mjs') --corpus-fixtures-dir $fixturesPath
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
Write-Output "Dashboard refreshed from: $fixturesPath"
