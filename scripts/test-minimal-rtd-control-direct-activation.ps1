[CmdletBinding()]
param([string]$DllPath = 'target/rtd-control/ExcelApi.ControlRtd.dll')
$ErrorActionPreference = 'Stop'
$dll = (Resolve-Path -LiteralPath $DllPath).Path
$server = $null
try {
    powershell -File scripts/register-minimal-rtd-control.ps1 -Path $dll | Out-Null
    if ($LASTEXITCODE -ne 0) { throw "control registration failed with exit code $LASTEXITCODE" }
    $server = New-Object -ComObject ExcelApi.ControlRtd
    $heartbeat = $server.Heartbeat()
    $server.ServerTerminate()
    if ([Runtime.InteropServices.Marshal]::IsComObject($server)) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($server) }
    $server = $null
    [pscustomobject]@{status='passed'; heartbeat=$heartbeat; prog_id='ExcelApi.ControlRtd'; dll=$dll} | ConvertTo-Json
} finally {
    if ($null -ne $server -and [Runtime.InteropServices.Marshal]::IsComObject($server)) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($server) }
    powershell -File scripts/unregister-minimal-rtd-control.ps1 | Out-Null
}
