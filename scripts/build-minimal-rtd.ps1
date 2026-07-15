[CmdletBinding()]
param(
    [ValidateSet('debug', 'release')]
    [string]$Profile = 'release'
)

$ErrorActionPreference = 'Stop'
$arguments = @('build', '-p', 'excel-api-minimal-rtd')
if ($Profile -eq 'release') { $arguments += '--release' }
& cargo @arguments
if ($LASTEXITCODE -ne 0) { throw "cargo build failed with exit code $LASTEXITCODE" }

$dll = Join-Path (Resolve-Path '.').Path "target/$Profile/excel_api_minimal_rtd.dll"
if (-not (Test-Path -LiteralPath $dll -PathType Leaf)) {
    throw "RTD prototype DLL was not produced: $dll"
}

[PSCustomObject]@{
    package = 'excel-api-minimal-rtd'
    target = 'x86_64-pc-windows-msvc'
    bitness = 64
    profile = $Profile
    dll = $dll
    prog_id = 'ExcelApi.MinimalRtd'
    clsid = '{DC738FE5-30EE-40E8-A8C2-3D16F217C52D}'
} | ConvertTo-Json
