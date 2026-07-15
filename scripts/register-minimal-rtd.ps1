[CmdletBinding(SupportsShouldProcess)]
param(
    [string]$Path = 'target/release/excel_api_minimal_rtd.dll',
    [switch]$ValidateOnly
)

$ErrorActionPreference = 'Stop'
$dll = (Resolve-Path -LiteralPath $Path).Path
if ([IO.Path]::GetExtension($dll) -ne '.dll') { throw 'RTD server path must name a DLL' }

$progId = 'ExcelApi.MinimalRtd'
$clsid = '{DC738FE5-30EE-40E8-A8C2-3D16F217C52D}'
$classes = 'HKCU:\Software\Classes'
$changes = @(
    "$classes\$progId",
    "$classes\$progId\CLSID",
    "$classes\CLSID\$clsid",
    "$classes\CLSID\$clsid\ProgID",
    "$classes\CLSID\$clsid\Programmable",
    "$classes\CLSID\$clsid\InprocServer32"
)
if ($ValidateOnly) {
    [PSCustomObject]@{ status = 'validated-no-write'; dll = $dll; prog_id = $progId; clsid = $clsid; keys = $changes } | ConvertTo-Json -Depth 3
    return
}

if ($PSCmdlet.ShouldProcess($progId, 'register per-user COM RTD server')) {
    foreach ($key in $changes) { New-Item -Path $key -Force | Out-Null }
    Set-Item -LiteralPath "$classes\$progId" -Value 'Excel API minimal RTD compatibility prototype'
    Set-Item -LiteralPath "$classes\$progId\CLSID" -Value $clsid
    Set-Item -LiteralPath "$classes\CLSID\$clsid" -Value 'Excel API minimal RTD compatibility prototype'
    Set-Item -LiteralPath "$classes\CLSID\$clsid\ProgID" -Value $progId
    Set-Item -LiteralPath "$classes\CLSID\$clsid\InprocServer32" -Value $dll
    New-ItemProperty -LiteralPath "$classes\CLSID\$clsid\InprocServer32" `
        -Name ThreadingModel -Value Apartment -PropertyType String -Force | Out-Null
}

[PSCustomObject]@{ status = 'registered'; scope = 'HKCU'; dll = $dll; prog_id = $progId; clsid = $clsid; threading_model = 'Apartment'; keys = $changes } | ConvertTo-Json -Depth 3
