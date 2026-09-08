$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem

$repoRoot = Split-Path -Parent $PSScriptRoot
$fixturesDir = Join-Path $repoRoot "tests\fixtures\cross-author-slicing"
if (-not (Test-Path $fixturesDir)) {
    New-Item -ItemType Directory -Path $fixturesDir -Force | Out-Null
}

function Extract-DocumentXml($docxPath, $xmlPath) {
    $zip = [System.IO.Compression.ZipFile]::OpenRead($docxPath)
    try {
        $entry = $zip.GetEntry("word/document.xml")
        if ($entry) {
            $reader = New-Object System.IO.StreamReader($entry.Open(), [System.Text.Encoding]::UTF8)
            try {
                $xml = $reader.ReadToEnd()
                [System.IO.File]::WriteAllText($xmlPath, $xml, [System.Text.Encoding]::UTF8)
            } finally {
                $reader.Dispose()
            }
        }
    } finally {
        $zip.Dispose()
    }
}

function Save-Triple($doc, [string]$baseName) {
    $pendingDocx = [string](Join-Path $fixturesDir "$baseName-pending.docx")
    $acceptedDocx = [string](Join-Path $fixturesDir "$baseName-accepted.docx")
    $rejectedDocx = [string](Join-Path $fixturesDir "$baseName-rejected.docx")
    
    [object]$pRef = $pendingDocx
    [object]$aRef = $acceptedDocx
    [object]$rRef = $rejectedDocx
    [object]$fmt = 16
    
    # Save pending
    $doc.SaveAs2([ref]$pRef, [ref]$fmt)
    
    # Accept All
    $doc.Revisions.AcceptAll()
    $doc.SaveAs2([ref]$aRef, [ref]$fmt)
    $doc.Close($false)
    
    # Reopen pending and Reject All
    $reopened = $global:word.Documents.Open([ref]$pRef)
    $reopened.Revisions.RejectAll()
    $reopened.SaveAs2([ref]$rRef, [ref]$fmt)
    $reopened.Close($false)
    
    # Extract XML
    Extract-DocumentXml $pendingDocx (Join-Path $fixturesDir "$baseName-pending.xml")
    Extract-DocumentXml $acceptedDocx (Join-Path $fixturesDir "$baseName-accepted.xml")
    Extract-DocumentXml $rejectedDocx (Join-Path $fixturesDir "$baseName-rejected.xml")
    Write-Host "Saved: $baseName"
}

function Find-RequiredText($doc, [string]$text, [string]$scenario) {
    $foundPos = $doc.Content.Text.IndexOf($text)
    if ($foundPos -lt 0) {
        throw "Scenario '$scenario' could not find required text '$text'."
    }
    return $foundPos
}

$origUserName = $null
$origUserInitials = $null

try {
    $global:word = New-Object -ComObject Word.Application
    $global:word.Visible = $false
    $global:word.DisplayAlerts = 0
    $origUserName = $global:word.UserName
    $origUserInitials = $global:word.UserInitials

    # -------------------------------------------------------------------------
    # Scenario 1: insert-interior
    # Author A (Barry) inserts: "amended by this Agreement."
    # Author B (Anson) inserts "MASTER " before "Agreement"
    # -------------------------------------------------------------------------
    Write-Host "1. insert-interior"
    $doc = $global:word.Documents.Add()
    $doc.TrackRevisions = $false
    $doc.Range(0, 0).Text = "Contract terms "
    
    # Author A insertion
    $global:word.UserName = "Barry Plasteras"
    $global:word.UserInitials = "BP"
    $doc.TrackRevisions = $true
    $endR = $doc.Range($doc.Content.End - 1, $doc.Content.End - 1)
    $endR.Text = "amended by this Agreement."
    
    # Author B insertion inside Author A's insertion
    $global:word.UserName = "Anson Lai"
    $global:word.UserInitials = "AL"
    # Find "Agreement" and insert "MASTER " before it
    $targetText = "Agreement."
    $foundPos = Find-RequiredText $doc $targetText "insert-interior"
    $insRange = $doc.Range($foundPos, $foundPos)
    $insRange.Text = "MASTER "
    Save-Triple $doc "insert-interior"

    # -------------------------------------------------------------------------
    # Scenario 2: delete-interior
    # Author A (Barry) inserts: "The Services will process the Input to generate outputs for Customer."
    # Author B (Anson) deletes: "generate "
    # -------------------------------------------------------------------------
    Write-Host "2. delete-interior"
    $doc = $global:word.Documents.Add()
    $doc.TrackRevisions = $false
    $doc.Range(0, 0).Text = "Background. "
    
    # Author A insertion
    $global:word.UserName = "Barry Plasteras"
    $global:word.UserInitials = "BP"
    $doc.TrackRevisions = $true
    $endR = $doc.Range($doc.Content.End - 1, $doc.Content.End - 1)
    $endR.Text = "The Services will process the Input to generate outputs for Customer."
    
    # Author B deletion inside Author A's insertion
    $global:word.UserName = "Anson Lai"
    $global:word.UserInitials = "AL"
    $delWord = "generate "
    $foundPos = Find-RequiredText $doc $delWord "delete-interior"
    $delRange = $doc.Range($foundPos, $foundPos + $delWord.Length)
    $delRange.Delete() | Out-Null
    Save-Triple $doc "delete-interior"

    # -------------------------------------------------------------------------
    # Scenario 3: delete-boundary-start
    # Author A (Barry) inserts: "Notwithstanding the foregoing, the NDA remains in effect."
    # Author B (Anson) deletes: "Notwithstanding the foregoing, "
    # -------------------------------------------------------------------------
    Write-Host "3. delete-boundary-start"
    $doc = $global:word.Documents.Add()
    $doc.TrackRevisions = $false
    $doc.Range(0, 0).Text = "Section 1. "
    
    # Author A insertion
    $global:word.UserName = "Barry Plasteras"
    $global:word.UserInitials = "BP"
    $doc.TrackRevisions = $true
    $endR = $doc.Range($doc.Content.End - 1, $doc.Content.End - 1)
    $endR.Text = "Notwithstanding the foregoing, the NDA remains in effect."
    
    # Author B deletion at start of insertion
    $global:word.UserName = "Anson Lai"
    $global:word.UserInitials = "AL"
    $delWord = "Notwithstanding the foregoing, "
    $foundPos = Find-RequiredText $doc $delWord "delete-boundary-start"
    $delRange = $doc.Range($foundPos, $foundPos + $delWord.Length)
    $delRange.Delete() | Out-Null
    Save-Triple $doc "delete-boundary-start"

    # -------------------------------------------------------------------------
    # Scenario 4: delete-boundary-end
    # Author A (Barry) inserts: "subject to Section 2.8 and applicable law."
    # Author B (Anson) deletes: " and applicable law."
    # -------------------------------------------------------------------------
    Write-Host "4. delete-boundary-end"
    $doc = $global:word.Documents.Add()
    $doc.TrackRevisions = $false
    $doc.Range(0, 0).Text = "Compliance: "
    
    # Author A insertion
    $global:word.UserName = "Barry Plasteras"
    $global:word.UserInitials = "BP"
    $doc.TrackRevisions = $true
    $endR = $doc.Range($doc.Content.End - 1, $doc.Content.End - 1)
    $endR.Text = "subject to Section 2.8 and applicable law."
    
    # Author B deletion at end of insertion
    $global:word.UserName = "Anson Lai"
    $global:word.UserInitials = "AL"
    $delWord = " and applicable law."
    $foundPos = Find-RequiredText $doc $delWord "delete-boundary-end"
    $delRange = $doc.Range($foundPos, $foundPos + $delWord.Length)
    $delRange.Delete() | Out-Null
    Save-Triple $doc "delete-boundary-end"

    # -------------------------------------------------------------------------
    # Scenario 5: delete-straddle-baseline-insertion
    # Baseline: "Baseline start "
    # Author A (Barry) inserts: "inserted finish."
    # Author B (Anson) deletes: "start inserted" (straddling baseline and insertion)
    # -------------------------------------------------------------------------
    Write-Host "5. delete-straddle-baseline-insertion"
    $doc = $global:word.Documents.Add()
    $doc.TrackRevisions = $false
    $doc.Range(0, 0).Text = "Baseline start "
    
    # Author A insertion
    $global:word.UserName = "Barry Plasteras"
    $global:word.UserInitials = "BP"
    $doc.TrackRevisions = $true
    $endR = $doc.Range($doc.Content.End - 1, $doc.Content.End - 1)
    $endR.Text = "inserted finish."
    
    # Author B deletion straddling baseline and insertion
    $global:word.UserName = "Anson Lai"
    $global:word.UserInitials = "AL"
    $delWord = "start inserted"
    $foundPos = Find-RequiredText $doc $delWord "delete-straddle-baseline-insertion"
    $delRange = $doc.Range($foundPos, $foundPos + $delWord.Length)
    $delRange.Delete() | Out-Null
    Save-Triple $doc "delete-straddle-baseline-insertion"

    # -------------------------------------------------------------------------
    # Scenario 6: multi-author-stacked
    # Baseline: "Provision "
    # Author A (Barry) inserts: "first draft of the proposal with initial metrics."
    # Author B (Anson) deletes: "of the proposal "
    # Author C (Chris) deletes: "initial " from the remaining text
    # -------------------------------------------------------------------------
    Write-Host "6. multi-author-stacked"
    $doc = $global:word.Documents.Add()
    $doc.TrackRevisions = $false
    $doc.Range(0, 0).Text = "Provision "
    
    # Author A insertion
    $global:word.UserName = "Barry Plasteras"
    $global:word.UserInitials = "BP"
    $doc.TrackRevisions = $true
    $endR = $doc.Range($doc.Content.End - 1, $doc.Content.End - 1)
    $endR.Text = "first draft of the proposal with initial metrics."
    
    # Author B deletion inside Author A's insertion
    $global:word.UserName = "Anson Lai"
    $global:word.UserInitials = "AL"
    $delWordB = "of the proposal "
    $foundPosB = Find-RequiredText $doc $delWordB "multi-author-stacked (Author B)"
    $delRangeB = $doc.Range($foundPosB, $foundPosB + $delWordB.Length)
    $delRangeB.Delete() | Out-Null
    
    # Author C deletion inside Author A's remaining insertion
    $global:word.UserName = "Chris Davis"
    $global:word.UserInitials = "CD"
    $delWordC = "initial "
    $foundPosC = Find-RequiredText $doc $delWordC "multi-author-stacked (Author C)"
    $delRangeC = $doc.Range($foundPosC, $foundPosC + $delWordC.Length)
    $delRangeC.Delete() | Out-Null
    Save-Triple $doc "multi-author-stacked"

    Write-Host "All 6 fixture triples successfully generated!"
}
finally {
    if ($null -ne $origUserName -and $null -ne $global:word) {
        $global:word.UserName = $origUserName
        $global:word.UserInitials = $origUserInitials
    }
    if ($null -ne $global:word) {
        $global:word.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($global:word) | Out-Null
    }
}
