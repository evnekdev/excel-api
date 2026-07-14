[CmdletBinding()]
param(
    [string]$Path = "target/release/minimal_xll.xll",
    [string]$ReportPath = "target/release/minimal_xll.exports.txt"
)

$ErrorActionPreference = 'Stop'
$required = @(
    'rust_add', 'rust_echo', 'rust_array_echo', 'rust_reference_kind',
    'rust_option_kind', 'xlAutoOpen', 'xlAutoClose', 'xlAutoAdd',
    'xlAutoRemove', 'xlAddInManagerInfo12', 'xlAutoFree12', 'SetExcel12EntryPt'
)

if (-not (Test-Path -LiteralPath $Path)) {
    throw "XLL does not exist: $Path"
}

$report = (& dumpbin /exports $Path) -join [Environment]::NewLine
if ($LASTEXITCODE -ne 0) { throw "dumpbin failed with exit code $LASTEXITCODE" }
$report | Set-Content -LiteralPath $ReportPath
foreach ($symbol in $required) {
    if ($report -notmatch "(?m)\b$([regex]::Escape($symbol))\b") {
        throw "required export is missing: $symbol"
    }
}
Write-Output "PASS: all $($required.Count) required M8 exports are present."
