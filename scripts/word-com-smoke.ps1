param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$Path
)

$resolved = Resolve-Path -LiteralPath $Path -ErrorAction Stop
$word = $null
$document = $null

try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false

    $document = $word.Documents.Open(
        [string]$resolved,
        $false,
        $true,
        $false,
        [Type]::Missing,
        [Type]::Missing,
        $false,
        [Type]::Missing,
        [Type]::Missing,
        [Type]::Missing,
        [Type]::Missing,
        $false,
        $false,
        $false,
        $false,
        $false
    )

    Write-Output "Opened: $resolved"
    Write-Output "Revisions: $($document.Revisions.Count)"
    Write-Output "PASS: Word opened the document without throwing."
}
catch {
    Write-Error "FAIL: Word could not open '$resolved'. $($_.Exception.Message)"
    exit 1
}
finally {
    if ($document -ne $null) {
        $document.Close($false) | Out-Null
    }
    if ($word -ne $null) {
        $word.Quit() | Out-Null
    }
}
