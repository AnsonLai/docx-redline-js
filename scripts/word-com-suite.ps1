param(
    [string]$OutputDir = "tmp/word-validation"
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
$exportScript = Join-Path $PSScriptRoot 'export-validation-fixtures.mjs'
$differentialScript = Join-Path $PSScriptRoot 'word-com-differential.ps1'
Push-Location $repoRoot
try {
    # Clean up any orphaned background Word processes from prior aborted runs
    Get-Process WINWORD -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowHandle -eq 0 } | Stop-Process -Force -ErrorAction SilentlyContinue

    & node $exportScript --output-dir $OutputDir
    if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

    & powershell -NoProfile -ExecutionPolicy Bypass -File $differentialScript -FixturesDir $OutputDir
    exit $LASTEXITCODE
}
finally {
    Pop-Location
}
