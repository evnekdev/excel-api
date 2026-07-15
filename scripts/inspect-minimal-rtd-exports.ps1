[CmdletBinding()]
param(
    [string]$Path = 'target/release/excel_api_minimal_rtd.dll',
    [string]$ReportPath = 'target/release/excel_api_minimal_rtd.exports.txt'
)

$ErrorActionPreference = 'Stop'

function Resolve-Dumpbin {
    $command = Get-Command dumpbin.exe -ErrorAction SilentlyContinue
    if ($null -ne $command) { return $command.Source }
    $vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio\Installer\vswhere.exe'
    if (-not (Test-Path -LiteralPath $vswhere)) { throw 'dumpbin.exe and vswhere.exe are unavailable' }
    $installation = & $vswhere -latest -products * `
        -requires Microsoft.VisualStudio.Component.VC.Tools.x86.x64 `
        -property installationPath
    if ($LASTEXITCODE -ne 0 -or [string]::IsNullOrWhiteSpace($installation)) {
        throw 'Visual Studio C++ tools were not found'
    }
    $toolsRoot = Join-Path $installation.Trim() 'VC\Tools\MSVC'
    $candidate = Get-ChildItem -LiteralPath $toolsRoot -Directory |
        Sort-Object Name -Descending |
        ForEach-Object { Join-Path $_.FullName 'bin\Hostx64\x64\dumpbin.exe' } |
        Where-Object { Test-Path -LiteralPath $_ } |
        Select-Object -First 1
    if ([string]::IsNullOrWhiteSpace($candidate)) { throw 'x64 dumpbin.exe was not found' }
    return $candidate
}

$resolved = (Resolve-Path -LiteralPath $Path).Path
$dumpbin = Resolve-Dumpbin
$headers = (& $dumpbin /headers $resolved) -join [Environment]::NewLine
if ($LASTEXITCODE -ne 0 -or $headers -notmatch '(?im)^\s*8664 machine \(x64\)') {
    throw 'RTD prototype is not an x64 PE image'
}
$report = (& $dumpbin /exports $resolved) -join [Environment]::NewLine
if ($LASTEXITCODE -ne 0) { throw "dumpbin /exports failed with exit code $LASTEXITCODE" }
$report | Set-Content -LiteralPath $ReportPath

$exports = @($report -split "`r?`n" | ForEach-Object {
    if ($_ -match '^\s+\d+\s+[0-9A-F]+\s+[0-9A-F]+\s+(\S+)(?:\s+=.*)?$') { $Matches[1] }
})
$expected = @('DllCanUnloadNow', 'DllGetClassObject')
$unexpected = @($exports | Where-Object { $_ -notin $expected })
$missing = @($expected | Where-Object { $_ -notin $exports })
if ($missing.Count -ne 0 -or $unexpected.Count -ne 0 -or $exports.Count -ne 2) {
    throw "unexpected COM export surface; missing=$($missing -join ',') unexpected=$($unexpected -join ',')"
}
$forbidden = @('xlAutoOpen', 'xlAutoClose', 'SetExcel12EntryPt', 'Excel12', 'Excel12v')
foreach ($symbol in $forbidden) {
    if ($report -match "(?m)\b$([regex]::Escape($symbol))\b") {
        throw "XLL/backend symbol appears in RTD DLL: $symbol"
    }
}
Write-Output 'PASS: x64 RTD DLL exports exactly DllCanUnloadNow and DllGetClassObject.'
