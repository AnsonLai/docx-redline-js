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
        } catch { $detail = $_.Exception.Message.Trim() }
        finally { if ($doc) { $doc.Close($false) | Out-Null } }

        $expectOpen = $file.Name.StartsWith('expect-open-')
        if ($opened -eq $expectOpen) { Write-Host "PASS  $($file.Name)  ($(if ($opened) { 'opened' } else { 'refused as expected' }); $detail)" }
        else { Write-Host "FAIL  $($file.Name)  ($(if ($opened) { 'opened but should be refused' } else { 'refused but should open' }); $detail)"; $failures++ }
    }
} finally { $word.Quit() | Out-Null }
if ($failures) { Write-Host "$failures case(s) failed"; exit 1 }
Write-Host "All Word package-integrity cases passed"
