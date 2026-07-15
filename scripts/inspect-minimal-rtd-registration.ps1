[CmdletBinding()]
param([string]$ExpectedPath = 'target/release/excel_api_minimal_rtd.dll')

$ErrorActionPreference = 'Stop'
$dll = (Resolve-Path -LiteralPath $ExpectedPath).Path
$progId = 'ExcelApi.MinimalRtd'
$clsid = '{DC738FE5-30EE-40E8-A8C2-3D16F217C52D}'

function Read-ComView([Microsoft.Win32.RegistryHive]$Hive, [Microsoft.Win32.RegistryView]$View, [string]$Root) {
    $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey($Hive, $View)
    try {
        $prog = $base.OpenSubKey("$Root$progId\CLSID")
        $class = $base.OpenSubKey("$Root" + "CLSID\$clsid\ProgID")
        $server = $base.OpenSubKey("$Root" + "CLSID\$clsid\InprocServer32")
        [pscustomobject]@{
            hive = [string]$Hive
            view = [string]$View
            present = ($null -ne $prog -or $null -ne $class -or $null -ne $server)
            progid_clsid = if ($null -ne $prog) { [string]$prog.GetValue('') } else { $null }
            clsid_progid = if ($null -ne $class) { [string]$class.GetValue('') } else { $null }
            server_path = if ($null -ne $server) { [string]$server.GetValue('') } else { $null }
            threading_model = if ($null -ne $server) { [string]$server.GetValue('ThreadingModel') } else { $null }
        }
        if ($null -ne $prog) { $prog.Dispose() }
        if ($null -ne $class) { $class.Dispose() }
        if ($null -ne $server) { $server.Dispose() }
    } finally { $base.Dispose() }
}

$views = @(
    Read-ComView CurrentUser Registry64 'Software\Classes\'
    Read-ComView LocalMachine Registry64 'Software\Classes\'
    Read-ComView ClassesRoot Registry64 ''
    Read-ComView CurrentUser Registry32 'Software\Classes\'
    Read-ComView LocalMachine Registry32 'Software\Classes\'
    Read-ComView ClassesRoot Registry32 ''
)
$hkcu64 = $views | Where-Object { $_.hive -eq 'CurrentUser' -and $_.view -eq 'Registry64' }
$hkcr64 = $views | Where-Object { $_.hive -eq 'ClassesRoot' -and $_.view -eq 'Registry64' }
if ($hkcu64.progid_clsid -ne $clsid) { throw '64-bit HKCU ProgID to CLSID mapping is incorrect' }
if ($hkcu64.clsid_progid -ne $progId) { throw '64-bit HKCU CLSID to ProgID mapping is incorrect' }
if (-not [string]::Equals($hkcu64.server_path, $dll, [StringComparison]::OrdinalIgnoreCase)) { throw '64-bit HKCU InprocServer32 path is incorrect' }
if ($hkcu64.threading_model -ne 'Apartment') { throw '64-bit HKCU ThreadingModel is not Apartment' }
if (-not [string]::Equals($hkcr64.server_path, $dll, [StringComparison]::OrdinalIgnoreCase)) { throw 'effective 64-bit HKCR server path does not match HKCU registration' }

$conflicts = @($views | Where-Object {
    $_.present -and -not ($_.hive -eq 'CurrentUser' -and $_.view -eq 'Registry64') -and
    ((-not [string]::IsNullOrWhiteSpace($_.server_path) -and -not [string]::Equals($_.server_path, $dll, [StringComparison]::OrdinalIgnoreCase)) -or
     (-not [string]::IsNullOrWhiteSpace($_.progid_clsid) -and $_.progid_clsid -ne $clsid) -or
     (-not [string]::IsNullOrWhiteSpace($_.clsid_progid) -and $_.clsid_progid -ne $progId))
})

$bytes = [IO.File]::ReadAllBytes($dll)
$peOffset = [BitConverter]::ToInt32($bytes, 0x3c)
$machine = [BitConverter]::ToUInt16($bytes, $peOffset + 4)
if ($machine -ne 0x8664) { throw ('RTD DLL is not PE32+ x64; machine=0x{0:X4}' -f $machine) }

[PSCustomObject]@{
    status = if ($conflicts.Count -eq 0) { 'valid' } else { 'conflict-detected' }
    expected = [ordered]@{ dll=$dll; path_exists=(Test-Path -LiteralPath $dll -PathType Leaf); machine='AMD64 (0x8664)'; prog_id=$progId; clsid=$clsid; threading_model='Apartment' }
    views = $views
    conflicts = $conflicts
} | ConvertTo-Json -Depth 6
if ($conflicts.Count -ne 0) { throw 'conflicting COM registrations were detected; no machine-wide keys were modified' }
