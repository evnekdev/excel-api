[CmdletBinding(SupportsShouldProcess)]
param([string]$Path = 'target/rtd-control/ExcelApi.ControlRtd.dll', [switch]$ValidateOnly)

$ErrorActionPreference = 'Stop'
$dll = (Resolve-Path -LiteralPath $Path).Path
$progId = 'ExcelApi.ControlRtd'
$clsid = '{F370A35B-7251-49E7-9FB2-3D6655FD1778}'
$className = 'ControlRtd'
$runtime = 'v4.0.30319'
$classes = 'HKCU:\Software\Classes'
$keys = @("$classes\$progId", "$classes\$progId\CLSID", "$classes\CLSID\$clsid", "$classes\CLSID\$clsid\ProgID", "$classes\CLSID\$clsid\InprocServer32")
if ($ValidateOnly) { [pscustomobject]@{status='validated-no-write'; dll=$dll; keys=$keys} | ConvertTo-Json; return }
if ($PSCmdlet.ShouldProcess($progId, 'register test-only .NET RTD control per user')) {
    foreach ($key in $keys) { New-Item -Path $key -Force | Out-Null }
    Set-Item -LiteralPath "$classes\$progId" -Value 'Excel API test-only RTD control'
    Set-Item -LiteralPath "$classes\$progId\CLSID" -Value $clsid
    Set-Item -LiteralPath "$classes\CLSID\$clsid" -Value 'Excel API test-only RTD control'
    Set-Item -LiteralPath "$classes\CLSID\$clsid\ProgID" -Value $progId
    $server = "$classes\CLSID\$clsid\InprocServer32"
    Set-Item -LiteralPath $server -Value 'mscoree.dll'
    New-ItemProperty -LiteralPath $server -Name ThreadingModel -Value Both -PropertyType String -Force | Out-Null
    New-ItemProperty -LiteralPath $server -Name Class -Value $className -PropertyType String -Force | Out-Null
    New-ItemProperty -LiteralPath $server -Name Assembly -Value 'ExcelApi.ControlRtd, Version=0.0.0.0, Culture=neutral, PublicKeyToken=null' -PropertyType String -Force | Out-Null
    New-ItemProperty -LiteralPath $server -Name RuntimeVersion -Value $runtime -PropertyType String -Force | Out-Null
    New-ItemProperty -LiteralPath $server -Name CodeBase -Value ([uri]$dll).AbsoluteUri -PropertyType String -Force | Out-Null
}
[pscustomobject]@{status='registered'; scope='HKCU'; dll=$dll; prog_id=$progId; clsid=$clsid; threading_model='Both'} | ConvertTo-Json
