$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
$manifestPath = Join-Path $repoRoot 'tests\corpus\superdoc-english-legal-administrative.json'
$manifest = Get-Content -LiteralPath $manifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
$fetchArgs = @((Join-Path $repoRoot 'scripts\fetch-superdoc-corpus.mjs'))
foreach ($document in $manifest.documents) {
    $fetchArgs += '--id'
    $fetchArgs += [string]$document.id
}

& node @fetchArgs
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

& node (Join-Path $repoRoot 'scripts\prepare-superdoc-word-corpus.mjs')
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

& (Join-Path $repoRoot 'scripts\package-superdoc-word-fixtures.ps1') `
    -FixturesDir (Join-Path $repoRoot 'tmp\superdoc-word-fixtures') `
    -SourcesDir (Join-Path $repoRoot 'tmp\superdoc-corpus')
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

& (Join-Path $repoRoot 'scripts\word-com-differential.ps1') `
    -FixturesDir (Join-Path $repoRoot 'tmp\superdoc-word-fixtures') `
    -SourcesDir (Join-Path $repoRoot 'tmp\superdoc-corpus')
exit $LASTEXITCODE
