[CmdletBinding()]
param(
    [string]$DllPath = 'target/release/excel_api_minimal_rtd.dll',
    [string]$DiagnosticsPath = 'target/rtd-direct-activation.jsonl'
)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'minimal-rtd-validation-helpers.ps1')
$dll = (Resolve-Path -LiteralPath $DllPath).Path
$diagnostics = [IO.Path]::GetFullPath($DiagnosticsPath)
New-Item -ItemType Directory -Path (Split-Path -Parent $diagnostics) -Force | Out-Null
Remove-Item -LiteralPath $diagnostics -Force -ErrorAction SilentlyContinue
$old = [Environment]::GetEnvironmentVariable('EXCEL_API_MINIMAL_RTD_DIAGNOSTICS', 'Process')
$server = $null
try {
    [Environment]::SetEnvironmentVariable('EXCEL_API_MINIMAL_RTD_DIAGNOSTICS', $diagnostics, 'Process')
    powershell -File scripts/register-minimal-rtd.ps1 -Path $dll | Out-Null
    if ($LASTEXITCODE -ne 0) { throw "RTD registration failed with exit code $LASTEXITCODE" }
    powershell -File scripts/inspect-minimal-rtd-registration.ps1 -ExpectedPath $dll | Out-Null
    if ($LASTEXITCODE -ne 0) { throw "RTD registration inspection failed with exit code $LASTEXITCODE" }
    $server = New-Object -ComObject ExcelApi.MinimalRtd
    $heartbeat = $server.Heartbeat()
    $server.ServerTerminate()
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($server)
    $server = $null
    $events = @(Get-Content -LiteralPath $diagnostics | ForEach-Object { $_ | ConvertFrom-Json })
    $stage = Get-MinimalRtdActivationStage -DllLoaded $true -Events $events
    if ($stage -ne 'object_created') { throw "direct activation did not reach object creation: $stage" }
    [pscustomobject]@{status='passed'; heartbeat=$heartbeat; activation_stage=$stage; methods=@($events.method | Sort-Object -Unique); diagnostics=$diagnostics} | ConvertTo-Json -Depth 5
} finally {
    if ($null -ne $server) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($server) }
    powershell -File scripts/unregister-minimal-rtd.ps1 | Out-Null
    [Environment]::SetEnvironmentVariable('EXCEL_API_MINIMAL_RTD_DIAGNOSTICS', $old, 'Process')
}
