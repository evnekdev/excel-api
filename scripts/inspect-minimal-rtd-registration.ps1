[CmdletBinding()]
param([string]$ExpectedPath = 'target/release/excel_api_minimal_rtd.dll')

$ErrorActionPreference = 'Stop'
$dll = (Resolve-Path -LiteralPath $ExpectedPath).Path
$progId = 'ExcelApi.MinimalRtd'
$clsid = '{DC738FE5-30EE-40E8-A8C2-3D16F217C52D}'
$classes = 'HKCU:\Software\Classes'
$progKey = Get-Item -LiteralPath "$classes\$progId"
$progClsid = Get-Item -LiteralPath "$classes\$progId\CLSID"
$classKey = Get-Item -LiteralPath "$classes\CLSID\$clsid"
$classProg = Get-Item -LiteralPath "$classes\CLSID\$clsid\ProgID"
$programmable = Get-Item -LiteralPath "$classes\CLSID\$clsid\Programmable"
$serverKey = Get-Item -LiteralPath "$classes\CLSID\$clsid\InprocServer32"
$actualPath = [string]$serverKey.GetValue('')
$threadingModel = [string]$serverKey.GetValue('ThreadingModel')
if ([string]$progClsid.GetValue('') -ne $clsid) { throw 'ProgID CLSID mapping is incorrect' }
if ([string]$classProg.GetValue('') -ne $progId) { throw 'CLSID ProgID mapping is incorrect' }
if (-not [string]::Equals($actualPath, $dll, [StringComparison]::OrdinalIgnoreCase)) { throw 'InprocServer32 path is incorrect' }
if ($threadingModel -ne 'Apartment') { throw 'ThreadingModel is not Apartment' }
[PSCustomObject]@{ status = 'valid'; scope = 'HKCU'; dll = $actualPath; prog_id = $progId; clsid = $clsid; threading_model = $threadingModel; programmable = $null -ne $programmable; descriptions = @([string]$progKey.GetValue(''), [string]$classKey.GetValue('')) } | ConvertTo-Json -Depth 3
