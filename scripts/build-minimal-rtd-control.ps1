[CmdletBinding()]
param([string]$OutputDirectory = 'target/rtd-control')

$ErrorActionPreference = 'Stop'
$csc = "$env:WINDIR\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
$piaCandidates = @(
    "$env:WINDIR\assembly\GAC_MSIL\Microsoft.Office.Interop.Excel\15.0.0.0__71e9bce111e9429c\Microsoft.Office.Interop.Excel.dll",
    "$env:ProgramFiles\Microsoft Office\root\Office16\Microsoft.Office.Interop.Excel.dll"
)
$pia = $piaCandidates | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
if (-not (Test-Path -LiteralPath $csc)) { throw '64-bit .NET Framework C# compiler was not found' }
if ([string]::IsNullOrWhiteSpace($pia)) { throw 'installed Microsoft Excel PIA was not found' }
$output = [IO.Path]::GetFullPath($OutputDirectory)
$source = (Resolve-Path -LiteralPath 'examples/minimal-rtd-control-server/ControlRtd.cs').Path
New-Item -ItemType Directory -Path $output -Force | Out-Null
$dll = Join-Path $output 'ExcelApi.ControlRtd.dll'
& $csc /nologo /target:library /platform:x64 /optimize+ "/out:$dll" "/link:$pia" $source
if ($LASTEXITCODE -ne 0) { throw "control RTD compilation failed with exit code $LASTEXITCODE" }
[PSCustomObject]@{ status='built'; dll=$dll; compiler=$csc; office_pia=$pia; prog_id='ExcelApi.ControlRtd'; clsid='{F370A35B-7251-49E7-9FB2-3D6655FD1778}' } | ConvertTo-Json
