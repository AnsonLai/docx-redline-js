<#
Opens every .docx in tmp/word-package-integrity (built by generate-word-package-integrity-cases.mjs)
in real Word with OpenAndRepair disabled. Word throws on a package it considers unreadable, which is
exactly what the "unreadable content" repair prompt reports.

  expect-open-*  must open cleanly.
  expect-fail-*  are controls that reproduce a known-bad package and must be REFUSED.
                 If Word opens a control, this check cannot detect the bug and fails.
#>
param([string]$Dir = (Join-Path $PSScriptRoot '..\tmp\word-package-integrity'))
$ErrorActionPreference = 'Stop'
$Dir = [System.IO.Path]::GetFullPath($Dir)
$files = Get-ChildItem -LiteralPath $Dir -Filter *.docx | Sort-Object Name
if (-not $files) { Write-Error "No cases in $Dir. Run: node scripts/generate-word-package-integrity-cases.mjs"; exit 2 }

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
$missing = [Type]::Missing
$failures = 0
try {
    foreach ($file in $files) {
        $opened = $false; $detail = ''
        $doc = $null
        try {
            # FileName, ConfirmConversions, ReadOnly, AddToRecentFiles, PasswordDocument, PasswordTemplate, Revert,
            # WritePasswordDocument, WritePasswordTemplate, Format, Encoding, Visible, OpenAndRepair
            $doc = $word.Documents.Open([string]$file.FullName, $false, $true, $false, $missing, $missing, $false, $missing, $missing, $missing, $missing, $false, $false)
            $opened = $true
            $detail = "comments=$($doc.Comments.Count)"
            $expectFile = [System.IO.Path]::ChangeExtension($file.FullName, '.expect.json')
            if (Test-Path -LiteralPath $expectFile) {
                $expect = Get-Content -LiteralPath $expectFile -Raw | ConvertFrom-Json
                $n = $doc.Comments.Count
                $checks = @{}
                if ($null -ne $expect.commentsDone) {
                    $actual = @(); for ($i = 1; $i -le $n; $i++) { $actual += [bool]$doc.Comments.Item($i).Done }
                    $checks['done'] = @(($actual -join ','), (@($expect.commentsDone) -join ','))
                }
                if ($null -ne $expect.ancestors) {
                    $actual = @()
                    for ($i = 1; $i -le $n; $i++) {
                        $anc = $doc.Comments.Item($i).Ancestor
                        $idx = 0
                        if ($anc) { for ($j = 1; $j -le $n; $j++) { $c = $doc.Comments.Item($j); if ($c.Author -eq $anc.Author -and $c.Range.Text -eq $anc.Range.Text) { $idx = $j; break } } }
                        $actual += $idx
                    }
                    $checks['ancestors'] = @(($actual -join ','), (@($expect.ancestors) -join ','))
                }
                if ($null -ne $expect.texts) {
                    $actual = @(); for ($i = 1; $i -le $n; $i++) { $actual += $doc.Comments.Item($i).Range.Text }
                    $checks['texts'] = @(($actual -join '|'), (@($expect.texts) -join '|'))
                }
                if ($expect.probe) {
                    $rows = @()
                    for ($i = 1; $i -le $n; $i++) {
                        $c = $doc.Comments.Item($i); $anc = $c.Ancestor; $idx = 0
                        if ($anc) { for ($j = 1; $j -le $n; $j++) { $o = $doc.Comments.Item($j); if ($o.Author -eq $anc.Author -and $o.Range.Text -eq $anc.Range.Text) { $idx = $j; break } } }
                        $rows += "#$i[$($c.Author): $($c.Range.Text) | anc=$idx done=$([bool]$c.Done) scope='$($c.Scope.Text)']"
                    }
                    $detail += " PROBE " + ($rows -join ' ; ')
                }
                foreach ($key in $checks.Keys) {
                    $detail += " $key=[$($checks[$key][0])]"
                    if ($checks[$key][0] -ne $checks[$key][1]) { $opened = $false; $detail += " EXPECTED $key=[$($checks[$key][1])]" }
                }
            }
        } catch { $detail = $_.Exception.Message.Trim() }
        finally { if ($doc) { $doc.Close($false) | Out-Null } }

        $expectOpen = $file.Name.StartsWith('expect-open-')
        if ($opened -eq $expectOpen) { Write-Host "PASS  $($file.Name)  ($(if ($opened) { 'opened' } else { 'refused as expected' }); $detail)" }
        else { Write-Host "FAIL  $($file.Name)  ($(if ($opened) { 'opened but should be refused' } else { 'refused but should open' }); $detail)"; $failures++ }
    }
} finally { $word.Quit() | Out-Null }
if ($failures) { Write-Host "$failures case(s) failed"; exit 1 }
Write-Host "All Word package-integrity cases passed"
