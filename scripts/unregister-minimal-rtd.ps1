[CmdletBinding(SupportsShouldProcess)]
param([switch]$ValidateOnly)

$ErrorActionPreference = 'Stop'
$progId = 'ExcelApi.MinimalRtd'
$clsid = '{DC738FE5-30EE-40E8-A8C2-3D16F217C52D}'
$classes = 'HKCU:\Software\Classes'
$targets = @("$classes\$progId", "$classes\CLSID\$clsid")
if ($ValidateOnly) {
    [PSCustomObject]@{ status = 'validated-no-write'; targets = $targets } | ConvertTo-Json -Depth 2
    return
}
foreach ($target in $targets) {
    if ((Test-Path -LiteralPath $target) -and $PSCmdlet.ShouldProcess($target, 'remove prototype COM registration')) {
        Remove-Item -LiteralPath $target -Recurse -Force
    }
}
[PSCustomObject]@{ status = 'unregistered'; targets = $targets } | ConvertTo-Json -Depth 2
