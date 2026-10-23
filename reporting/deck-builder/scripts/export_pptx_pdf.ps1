param(
  [Parameter(Mandatory=$true)][string]$Pptx,
  [Parameter(Mandatory=$true)][string]$Out
)
# 坑（10-19 实测，技能官验收记录）：
# 1. 打印 "OK: ..." 后本脚本可能卡在 COM 清理（Quit/ReleaseComObject）几分钟不退出——
#    导出已成功，自动化调用请设超时，看到 OK 即成功，不必等进程退出。
# 2. 用户自己开着 WPS 时慎跑：COM 可能附着已开实例，$app.Quit() 有连带关掉用户实例的风险。
# 3. -Pptx/-Out 传相对路径时 WPS Open 会报 HRESULT E_FAIL——脚本已在入口统一转绝对路径，
#    传什么都行；绕过本脚本直接调 COM 的话必须自己传绝对路径。
$ErrorActionPreference = "Stop"
$pptxAbs = (Convert-Path -LiteralPath $Pptx)
$outParent = Split-Path -Parent $Out
if (-not $outParent) { $outParent = (Get-Location).Path }
$outAbs = Join-Path (Convert-Path -LiteralPath $outParent) (Split-Path -Leaf $Out)
if (Test-Path -LiteralPath $outAbs) { Remove-Item -LiteralPath $outAbs -Force }
$app = New-Object -ComObject Kwpp.Application
try {
  try { $app.Visible = $false } catch {}
  $pres = $app.Presentations.Open($pptxAbs, $false, $true, $false)
  $pres.SaveAs($outAbs, 32)   # 32 = ppSaveAsPDF（WPS 实测可用；ExportAsFixedFormat 会 E_FAIL）
  $pres.Close()
  Write-Output ("OK: {0} ({1:N2} MB)" -f $outAbs, ((Get-Item -LiteralPath $outAbs).Length / 1MB))
}
finally {
  $app.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null
}
