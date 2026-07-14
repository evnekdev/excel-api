[CmdletBinding()]
param(
    [ValidateSet('debug', 'release')]
    [string]$Profile = 'release'
)

$ErrorActionPreference = 'Stop'
$workspace = Split-Path -Parent $PSScriptRoot
$cargoArgs = @('build', '-p', 'minimal-xll')
if ($Profile -eq 'release') { $cargoArgs += '--release' }

Push-Location $workspace
try {
    & cargo @cargoArgs
    if ($LASTEXITCODE -ne 0) { throw "cargo build failed with exit code $LASTEXITCODE" }

    $dll = Join-Path $workspace "target/$Profile/minimal_xll.dll"
    $xll = Join-Path $workspace "target/$Profile/minimal_xll.xll"
    if (-not (Test-Path -LiteralPath $dll)) { throw "expected cdylib was not produced: $dll" }
    Copy-Item -LiteralPath $dll -Destination $xll -Force
    Write-Output $xll
}
finally {
    Pop-Location
}
