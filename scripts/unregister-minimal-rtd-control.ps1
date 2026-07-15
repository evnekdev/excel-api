[CmdletBinding(SupportsShouldProcess)]
param()
$ErrorActionPreference = 'Stop'
$classes = 'HKCU:\Software\Classes'
foreach ($key in @("$classes\ExcelApi.ControlRtd", "$classes\CLSID\{F370A35B-7251-49E7-9FB2-3D6655FD1778}")) {
    if ((Test-Path -LiteralPath $key) -and $PSCmdlet.ShouldProcess($key, 'remove test-only per-user COM registration')) { Remove-Item -LiteralPath $key -Recurse -Force }
}
[pscustomobject]@{status='unregistered'; scope='HKCU'; prog_id='ExcelApi.ControlRtd'} | ConvertTo-Json
