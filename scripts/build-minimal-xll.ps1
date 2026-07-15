[CmdletBinding()]
param(
    [ValidateSet('debug', 'release')]
    [string]$Profile = 'release',
    [switch]$XlcOnTimeResearch
)

$ErrorActionPreference = 'Stop'
$workspace = Split-Path -Parent $PSScriptRoot
$cargoArgs = @('build', '-p', 'minimal-xll')
if ($Profile -eq 'release') { $cargoArgs += '--release' }
if ($XlcOnTimeResearch) { $cargoArgs += @('--features', 'xlcontime-research') }

Push-Location $workspace
try {
    & cargo @cargoArgs
    if ($LASTEXITCODE -ne 0) { throw "cargo build failed with exit code $LASTEXITCODE" }

    $dll = Join-Path $workspace "target/$Profile/minimal_xll.dll"
    $xllName = if ($XlcOnTimeResearch) { 'minimal_xll_ontime_research.xll' } else { 'minimal_xll.xll' }
    $xll = Join-Path $workspace "target/$Profile/$xllName"
    if (-not (Test-Path -LiteralPath $dll)) { throw "expected cdylib was not produced: $dll" }
    Copy-Item -LiteralPath $dll -Destination $xll -Force
    Write-Output $xll
}
finally {
    Pop-Location
}
