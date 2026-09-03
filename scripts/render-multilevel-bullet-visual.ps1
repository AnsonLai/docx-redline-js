param(
    [string]$DocxPath,
    [string]$OutputDir
)

$ErrorActionPreference = 'Stop'
$resolvedDocx = (Resolve-Path -LiteralPath $DocxPath).Path
$resolvedOut = (Resolve-Path -LiteralPath $OutputDir).Path

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0

try {
    foreach ($view in @('allMarkup', 'acceptAll', 'rejectAll')) {
        $pdfPath = Join-Path $resolvedOut "multilevel-bullet--$view.pdf"
        $doc = $word.Documents.OpenNoRepairDialog($resolvedDocx)
        try {
            if ($view -eq 'acceptAll') { $doc.AcceptAllRevisions(); $exportItem = 0 }
            elseif ($view -eq 'rejectAll') { $doc.RejectAllRevisions(); $exportItem = 0 }
            else { $doc.ShowRevisions = $true; $doc.PrintRevisions = $true; $exportItem = 7 }
            $pages = $doc.ComputeStatistics(2)
            $doc.ExportAsFixedFormat($pdfPath, 17, $false, 0, 0, 1, 1, $exportItem, $true, $true, 0, $true, $true, $false)
            $item = Get-Item -LiteralPath $pdfPath
            Write-Output "Rendered $view -> $pdfPath (pages=$pages, bytes=$($item.Length))"
        } finally {
            $doc.Close(0)
        }
    }
} finally {
    $word.Quit()
}
