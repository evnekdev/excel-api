[CmdletBinding(SupportsShouldProcess)]
param(
    [string]$Path = 'target/release/excel_api_minimal_rtd.dll',
    [switch]$ValidateOnly
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'minimal-rtd-validation-helpers.ps1')
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

$conflicts = @()
foreach ($scope in @('HKCU:\Software\Classes', 'HKLM:\Software\Classes', 'Registry::HKEY_LOCAL_MACHINE\Software\Classes\Wow6432Node')) {
    $existingServer = Get-Item -LiteralPath "$scope\CLSID\$clsid\InprocServer32" -ErrorAction SilentlyContinue
    $existingProg = Get-Item -LiteralPath "$scope\CLSID\$clsid\ProgID" -ErrorAction SilentlyContinue
    if ($null -ne $existingServer -or $null -ne $existingProg) {
        $registeredPath = if ($null -ne $existingServer) { [string]$existingServer.GetValue('') } else { $null }
        $registeredProgId = if ($null -ne $existingProg) { [string]$existingProg.GetValue('') } else { $null }
        if (Test-MinimalRtdRegistrationConflict -RegisteredPath $registeredPath -ExpectedPath $dll -RegisteredProgId $registeredProgId -ExpectedProgId $progId) {
            $conflicts += [pscustomobject]@{ scope=$scope; server_path=$registeredPath; prog_id=$registeredProgId }
        }
    }
}
if ($conflicts.Count -ne 0) {
    $conflicts | ConvertTo-Json -Depth 3 | Write-Output
    throw 'conflicting existing RTD registration detected; no registry key was modified'
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
