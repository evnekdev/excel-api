[CmdletBinding()]
param([string]$DllPath = 'target/release/excel_api_minimal_rtd.dll')

$ErrorActionPreference = 'Stop'
$dll = (Resolve-Path -LiteralPath $DllPath).Path
$excel = Get-ItemProperty 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe' -ErrorAction SilentlyContinue
$excelPath = if ($null -ne $excel) { [string]$excel.'(default)' } else { $null }

$policyRoots = @(
    'Registry::HKEY_CURRENT_USER\Software\Policies\Microsoft\Office\16.0\Excel',
    'Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Office\16.0\Excel',
    'Registry::HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Excel',
    'Registry::HKEY_LOCAL_MACHINE\Software\Microsoft\Office\16.0\Excel'
)
$policyNames = @('DisableRTD', 'DisableAllAddins', 'RequireAddinSig', 'VBAWarnings', 'AutomationSecurity', 'OpenInProtectedView')
$policies = @()
foreach ($root in $policyRoots) {
    foreach ($suffix in @('', '\Security', '\Options', '\ProtectedView', '\Resiliency')) {
        $path = "$root$suffix"
        $item = Get-ItemProperty -LiteralPath $path -ErrorAction SilentlyContinue
        if ($null -ne $item) {
            foreach ($name in $policyNames) {
                if ($null -ne $item.PSObject.Properties[$name]) {
                    $policies += [pscustomobject]@{ path=$path; name=$name; value=$item.$name }
                }
            }
            if ($suffix -eq '\Resiliency' -and $null -ne $item.PSObject.Properties['DisabledItems']) {
                $policies += [pscustomobject]@{ path=$path; name='DisabledItems'; value='present-redacted' }
            }
        }
    }
}

$signature = Get-AuthenticodeSignature -LiteralPath $dll
$dependencies = $null
try {
    $vswhere = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe"
    $installation = (& $vswhere -latest -products * -requires Microsoft.VisualStudio.Component.VC.Tools.x86.x64 -property installationPath | Select-Object -First 1)
    $dumpbin = Get-ChildItem -LiteralPath (Join-Path $installation 'VC\Tools\MSVC') -Directory |
        Sort-Object Name -Descending |
        ForEach-Object { Join-Path $_.FullName 'bin\Hostx64\x64\dumpbin.exe' } |
        Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
    if ($dumpbin) { $dependencies = @(& $dumpbin /dependents $dll | Where-Object { $_ -match '^\s+[A-Za-z0-9_.-]+\.dll\s*$' } | ForEach-Object { $_.Trim() }) }
} catch { $dependencies = @('inspection-unavailable') }

$loaderSource = @'
using System;
using System.Runtime.InteropServices;
public static class RtdLoaderProbe {
  [DllImport("kernel32", SetLastError=true, CharSet=CharSet.Unicode)] public static extern IntPtr LoadLibraryEx(string path, IntPtr file, uint flags);
  [DllImport("kernel32", SetLastError=true)] public static extern bool FreeLibrary(IntPtr module);
}
'@
if (-not ('RtdLoaderProbe' -as [type])) { Add-Type -TypeDefinition $loaderSource }
$module = [RtdLoaderProbe]::LoadLibraryEx($dll, [IntPtr]::Zero, 0x00001100)
$loader = if ($module -eq [IntPtr]::Zero) { [ordered]@{status='failed'; win32_error=[Runtime.InteropServices.Marshal]::GetLastWin32Error()} } else { [void][RtdLoaderProbe]::FreeLibrary($module); [ordered]@{status='passed'; win32_error=0} }

[pscustomobject]@{
    status = if ($loader.status -eq 'passed') { 'inspected' } else { 'loader-failed' }
    excel = if ($excelPath) { [ordered]@{ path=$excelPath; version=(Get-Item -LiteralPath $excelPath).VersionInfo.FileVersion; architecture='64-bit host required and verified by live harness' } } else { $null }
    rtd_dll = [ordered]@{ path=$dll; signature_status=[string]$signature.Status; signer=if($signature.SignerCertificate){$signature.SignerCertificate.Subject}else{$null}; dependencies=$dependencies; load_library_ex=$loader }
    policy_observations = $policies
    policy_note = 'read-only inspection; absent keys do not prove that endpoint controls permit RTD activation'
} | ConvertTo-Json -Depth 7
