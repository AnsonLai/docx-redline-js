param(
    [string]$FixturesDir = "tmp/superdoc-word-fixtures",
    [string]$SourcesDir = "tmp/superdoc-corpus"
)

$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
$fixturesPath = if ([System.IO.Path]::IsPathRooted($FixturesDir)) {
    [System.IO.Path]::GetFullPath($FixturesDir)
} else {
    [System.IO.Path]::GetFullPath((Join-Path $repoRoot $FixturesDir))
}
$sourcesPath = if ([System.IO.Path]::IsPathRooted($SourcesDir)) {
    [System.IO.Path]::GetFullPath($SourcesDir)
} else {
    [System.IO.Path]::GetFullPath((Join-Path $repoRoot $SourcesDir))
}
$manifestPath = Join-Path $repoRoot 'tests\corpus\superdoc-english-legal-administrative.json'
$manifest = Get-Content -LiteralPath $manifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
$fetchArgs = @((Join-Path $repoRoot 'scripts\fetch-superdoc-corpus.mjs'))
foreach ($document in $manifest.documents) {
    $fetchArgs += '--id'
    $fetchArgs += [string]$document.id
}

& node @fetchArgs
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

& node (Join-Path $repoRoot 'scripts\prepare-superdoc-word-corpus.mjs') `
    --input-dir $sourcesPath `
    --output-dir $fixturesPath
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

& (Join-Path $repoRoot 'scripts\package-superdoc-word-fixtures.ps1') `
    -FixturesDir $fixturesPath `
    -SourcesDir $sourcesPath
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

& (Join-Path $repoRoot 'scripts\word-com-differential.ps1') `
    -FixturesDir $fixturesPath `
    -SourcesDir $sourcesPath
exit $LASTEXITCODE
