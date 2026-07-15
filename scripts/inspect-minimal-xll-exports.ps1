[CmdletBinding()]
param(
    [string]$Path = "target/release/minimal_xll.xll",
    [string]$ReportPath = "target/release/minimal_xll.exports.txt"
)

$ErrorActionPreference = 'Stop'
$required = @(
    'rust_add', 'rust_echo', 'rust_array_echo', 'rust_reference_kind',
    'rust_option_kind', 'rust_async_double', 'excel_api_calculation_canceled',
    'excel_api_calculation_ended', 'xlAutoOpen', 'xlAutoClose', 'xlAutoAdd',
    'xlAutoRemove', 'xlAddInManagerInfo12', 'xlAutoFree12', 'SetExcel12EntryPt'
)

if (-not (Test-Path -LiteralPath $Path)) {
    throw "XLL does not exist: $Path"
}

function Resolve-Dumpbin {
    $command = Get-Command dumpbin.exe -ErrorAction SilentlyContinue
    if ($null -ne $command) { return $command.Source }

    $installerRoot = ${env:ProgramFiles(x86)}
    if ([string]::IsNullOrEmpty($installerRoot)) {
        throw 'ProgramFiles(x86) is unavailable and dumpbin.exe is not on PATH'
    }
    $vswhere = Join-Path $installerRoot 'Microsoft Visual Studio\Installer\vswhere.exe'
    if (-not (Test-Path -LiteralPath $vswhere)) {
        throw "dumpbin.exe is not on PATH and vswhere.exe was not found: $vswhere"
    }

    $installation = & $vswhere -latest -products * `
        -requires Microsoft.VisualStudio.Component.VC.Tools.x86.x64 `
        -property installationPath
    if ($LASTEXITCODE -ne 0 -or [string]::IsNullOrWhiteSpace($installation)) {
        throw 'vswhere.exe could not locate Visual Studio C++ tools'
    }

    $toolsRoot = Join-Path $installation.Trim() 'VC\Tools\MSVC'
    $candidate = Get-ChildItem -LiteralPath $toolsRoot -Directory |
        Sort-Object Name -Descending |
        ForEach-Object { Join-Path $_.FullName 'bin\Hostx64\x64\dumpbin.exe' } |
        Where-Object { Test-Path -LiteralPath $_ } |
        Select-Object -First 1
    if ([string]::IsNullOrWhiteSpace($candidate)) {
        throw "Visual Studio C++ tools contain no x64 dumpbin.exe under: $toolsRoot"
    }
    return $candidate
}

$dumpbin = Resolve-Dumpbin
$report = (& $dumpbin /exports $Path) -join [Environment]::NewLine
if ($LASTEXITCODE -ne 0) { throw "dumpbin failed with exit code $LASTEXITCODE" }
$report | Set-Content -LiteralPath $ReportPath
foreach ($symbol in $required) {
    if ($report -notmatch "(?m)\b$([regex]::Escape($symbol))\b") {
        throw "required export is missing: $symbol"
    }
}
Write-Output "PASS: all $($required.Count) required exports are present."
