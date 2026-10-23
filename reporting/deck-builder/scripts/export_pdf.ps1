param(
    [Parameter(Mandatory=$true)][string]$Docx,
    [Parameter(Mandatory=$true)][string]$Out
)
# docx -> vector PDF via WPS COM (Kwps.Application). No UI is shown.
$ErrorActionPreference = 'Stop'
$dir = Split-Path -Parent $Out
if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Force -Path $dir | Out-Null }
$app = $null; $doc = $null
try {
    $app = New-Object -ComObject Kwps.Application
    $app.Visible = $false
    $doc = $app.Documents.Open($Docx, $false, $true)
    Start-Sleep -Milliseconds 800
    $doc.ExportAsFixedFormat($Out, 17)   # 17 = PDF
    Write-Output "OK: $Out ($([math]::Round((Get-Item $Out).Length/1MB,2)) MB)"
}
finally {
    if ($doc) { try { $doc.Close(0) } catch {} }
    if ($app) { try { $app.Quit() } catch {} }
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null
}
