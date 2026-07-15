[CmdletBinding()]
param(
    [string]$DllPath = 'target/release/excel_api_minimal_rtd.dll',
    [string]$OutputDirectory = 'target/rtd-validation',
    [int]$ObservationSeconds = 4,
    [int]$ProcessTimeoutSeconds = 90,
    [ValidateSet('Rust', 'Control')]
    [string]$Server = 'Rust',
    [string]$RunDirectory,
    [switch]$Worker,
    [switch]$ValidateOnly
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'minimal-rtd-validation-helpers.ps1')
if ($ObservationSeconds -lt 2 -or $ObservationSeconds -gt 60) {
    throw 'ObservationSeconds must be between 2 and 60'
}
if ($ProcessTimeoutSeconds -lt 15 -or $ProcessTimeoutSeconds -gt 600) {
    throw 'ProcessTimeoutSeconds must be between 15 and 600'
}
$dll = (Resolve-Path -LiteralPath $DllPath).Path
$root = Join-Path (Resolve-Path -LiteralPath '.').Path $OutputDirectory
$run = if ([string]::IsNullOrWhiteSpace($RunDirectory)) {
    Join-Path $root (Get-Date -Format 'yyyyMMddTHHmmssZ')
} else {
    [IO.Path]::GetFullPath($RunDirectory)
}
$diagnostics = Join-Path $run 'server-events.jsonl'
$summaryPath = Join-Path $run 'summary.json'
$coordinationPath = Join-Path $run 'owned-process.json'
$isControl = $Server -eq 'Control'
$progId = if ($isControl) { 'ExcelApi.ControlRtd' } else { 'ExcelApi.MinimalRtd' }
$clsid = if ($isControl) { '{F370A35B-7251-49E7-9FB2-3D6655FD1778}' } else { '{DC738FE5-30EE-40E8-A8C2-3D16F217C52D}' }
$diagnosticVariable = if ($isControl) { 'EXCEL_API_CONTROL_RTD_DIAGNOSTICS' } else { 'EXCEL_API_MINIMAL_RTD_DIAGNOSTICS' }
$settings = [ordered]@{
    server = $Server
    prog_id = $progId
    clsid = $clsid
    dll = $dll
    observation_seconds = $ObservationSeconds
    formula = "=RTD(`"$progId`",,`"COUNTER`")"
    formula_matrix = @(Get-MinimalRtdFormulaMatrix -ProgId $progId)
}

function Quote-ProcessArgument([string]$Value) { '"' + ($Value -replace '(\\*)"', '$1$1\"' -replace '(\\+)$', '$1$1') + '"' }
function Start-RtdWorker([string]$Shell, [string[]]$Arguments, [string]$WorkingDirectory) {
    $start = New-Object Diagnostics.ProcessStartInfo
    $start.FileName = $Shell
    $start.WorkingDirectory = $WorkingDirectory
    $start.Arguments = ($Arguments | ForEach-Object { Quote-ProcessArgument $_ }) -join ' '
    $start.UseShellExecute = $false
    $start.RedirectStandardOutput = $true
    $start.RedirectStandardError = $true
    $start.CreateNoWindow = $true
    # Preserve the inherited environment block without enumerating its
    # case-insensitive PATH/Path aliases.
    $process = New-Object Diagnostics.Process
    $process.StartInfo = $start
    if (-not $process.Start()) { throw 'failed to start direct RTD worker' }
    [pscustomobject]@{ process=$process; output=$process.StandardOutput.ReadToEndAsync(); error=$process.StandardError.ReadToEndAsync() }
}
if ($ValidateOnly) {
    if ($isControl) { powershell -File scripts/register-minimal-rtd-control.ps1 -Path $dll -ValidateOnly | Out-Null }
    else { powershell -File scripts/register-minimal-rtd.ps1 -Path $dll -ValidateOnly | Out-Null }
    powershell -File scripts/test-minimal-rtd-validation-helpers.ps1 | Out-Null
    [PSCustomObject]@{ status = 'validated-no-excel'; settings = $settings } | ConvertTo-Json -Depth 4
    return
}

if (-not $Worker) {
    $stale = @()
    if (Test-Path -LiteralPath $root) {
        foreach ($ownedFile in Get-ChildItem -LiteralPath $root -Filter 'owned-process.json' -File -Recurse -ErrorAction SilentlyContinue) {
            try {
                $prior = Get-Content -LiteralPath $ownedFile.FullName -Raw | ConvertFrom-Json
                $ownedProcess = Get-Process -Id ([int]$prior.excel_pid) -ErrorAction SilentlyContinue
                if ($null -ne $ownedProcess) {
                    $stale += [pscustomobject]@{ pid=$ownedProcess.Id; relationship='recorded-owned-excel'; record=$ownedFile.FullName; start_time_utc=try{$ownedProcess.StartTime.ToUniversalTime().ToString('o')}catch{'access-denied'} }
                }
                foreach ($child in @(Get-CimInstance Win32_Process -Filter "ParentProcessId=$([int]$prior.excel_pid)" -ErrorAction SilentlyContinue)) {
                    $stale += [pscustomobject]@{ pid=[int]$child.ProcessId; parent_pid=[int]$child.ParentProcessId; relationship='verified-direct-descendant'; image_path=[string]$child.ExecutablePath; command_line=if($child.CommandLine){[string]$child.CommandLine}else{'unavailable'}; record=$ownedFile.FullName }
                }
            } catch { }
        }
    }
    if ($stale.Count -ne 0) {
        [pscustomobject]@{ status='blocked'; classification='stale-owned-test-processes'; cleanup='manual-admin-or-reboot-required'; stale_processes=$stale } | ConvertTo-Json -Depth 6
        return
    }
    New-Item -ItemType Directory -Path $run -Force | Out-Null
    $stdout = Join-Path $run 'worker.stdout.txt'
    $stderr = Join-Path $run 'worker.stderr.txt'
    $arguments = @(
        '-NoProfile', '-File', $PSCommandPath, '-Worker',
        '-DllPath', $dll, '-OutputDirectory', $OutputDirectory,
        '-ObservationSeconds', $ObservationSeconds,
        '-ProcessTimeoutSeconds', $ProcessTimeoutSeconds,
        '-Server', $Server,
        '-RunDirectory', $run
    )
    $worker = Start-RtdWorker (Get-Process -Id $PID).Path $arguments (Resolve-Path '.').Path
    $process = $worker.process
    if (-not $process.WaitForExit($ProcessTimeoutSeconds * 1000)) {
        $owned = $null
        if (Test-Path -LiteralPath $coordinationPath) {
            $owned = Get-Content -LiteralPath $coordinationPath -Raw | ConvertFrom-Json
        }
        Stop-Process -Id $process.Id -Force -ErrorAction SilentlyContinue
        $excelCleanup = [ordered]@{
            owned = 'owned-pid-unavailable'
            direct_excel_children = @()
        }
        if ($null -ne $owned -and $owned.excel_pid -gt 0) {
            $candidate = Get-Process -Id $owned.excel_pid -ErrorAction SilentlyContinue
            if ($null -ne $candidate -and $candidate.ProcessName -eq 'EXCEL' -and
                $candidate.StartTime.ToUniversalTime().ToString('o') -eq $owned.excel_start_time_utc) {
                Stop-Process -Id $candidate.Id -Force
                $excelCleanup.owned = 'exact-owned-pid-terminated'
            } else {
                $excelCleanup.owned = 'owned-pid-start-mismatch-not-terminated'
            }
            Start-Sleep -Milliseconds 250
            $children = @(Get-CimInstance Win32_Process -Filter "ParentProcessId=$($owned.excel_pid)" -ErrorAction SilentlyContinue |
                Where-Object { $_.Name -eq 'EXCEL.EXE' })
            foreach ($child in $children) {
                $childResult = [ordered]@{
                    pid = [int]$child.ProcessId
                    parent_pid = [int]$child.ParentProcessId
                    image_path = [string]$child.ExecutablePath
                    start_time = [string]$child.CreationDate
                    command_line = if ($child.CommandLine) { [string]$child.CommandLine } else { 'unavailable' }
                    result = 'not-terminated'
                }
                try {
                    Stop-Process -Id $child.ProcessId -Force -ErrorAction Stop
                    $childResult.result = 'verified-direct-child-terminated'
                } catch {
                    $childResult.result = 'verified-direct-child-termination-denied'
                }
                $excelCleanup.direct_excel_children += $childResult
            }
        }
        if ($isControl) { powershell -File scripts/unregister-minimal-rtd-control.ps1 | Out-Null }
        else { powershell -File scripts/unregister-minimal-rtd.ps1 | Out-Null }
        $timeoutSummary = [ordered]@{
            status = 'blocked'
            classification = 'host-blocked-plain-workbooks-add-timeout'
            effective_timeout_seconds = $ProcessTimeoutSeconds
            worker_pid = $process.Id
            owned_process = $owned
            cleanup = $excelCleanup
            completed_utc = [datetime]::UtcNow.ToString('o')
        }
        $timeoutSummary | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $summaryPath
        Write-Output ($timeoutSummary | ConvertTo-Json -Depth 6)
        Write-Output "summary=$summaryPath"
        return
    }
    $worker.output.Result | Set-Content -LiteralPath $stdout
    $worker.error.Result | Set-Content -LiteralPath $stderr
    if (Test-Path -LiteralPath $stdout) { Get-Content -LiteralPath $stdout }
    $process.Refresh()
    if ($process.ExitCode -ne 0 -and (Test-Path -LiteralPath $stderr)) {
        Get-Content -LiteralPath $stderr -ErrorAction SilentlyContinue
        throw "RTD validation worker exited with code $($process.ExitCode)"
    }
    return
}

New-Item -ItemType Directory -Path $run -Force | Out-Null
$summary = [ordered]@{
    status = 'failed'
    classification = 'not-run'
    settings = $settings
    started_utc = [datetime]::UtcNow.ToString('o')
    excel = $null
    workbook_preflight = $false
    registration = 'not-run'
    initial_value = $null
    updated_value = $null
    duplicate_value = $null
    reconnect_value = $null
    diagnostics = $null
    activation_stage = 'dll_not_loaded'
    module_evidence = $null
    formula_results = @()
    cleanup = [ordered]@{ workbook_closed = $false; excel_quit = $false; process_exited = $false; unregistered = $false }
}
$excel = $null
$workbook = $null
$excelPid = 0
$registered = $false
$oldDiagnostic = [Environment]::GetEnvironmentVariable($diagnosticVariable, 'Process')
try {
    [Environment]::SetEnvironmentVariable($diagnosticVariable, $diagnostics, 'Process')
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $hwnd = [Int64]$excel.Hwnd
    if (-not ('RtdValidation.NativeMethods' -as [type])) {
        Add-Type -TypeDefinition @'
namespace RtdValidation {
  public static class NativeMethods {
    [System.Runtime.InteropServices.DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(System.IntPtr hWnd, out uint processId);
  }
}
'@
    }
    [uint32]$ownedProcessIdRaw = 0
    [void][RtdValidation.NativeMethods]::GetWindowThreadProcessId(
        [IntPtr]$hwnd,
        [ref]$ownedProcessIdRaw
    )
    $excelPid = [int]$ownedProcessIdRaw
    $process = Get-Process -Id $excelPid
    if ($process.ProcessName -ne 'EXCEL') { throw 'owned COM process is not EXCEL.EXE' }
    $summary.excel = [ordered]@{
        pid = $excelPid
        hwnd = $hwnd
        start_time_utc = $process.StartTime.ToUniversalTime().ToString('o')
        version = [string]$excel.Version
        build = [string]$excel.Build
        architecture = '64-bit (required; DLL/export inspection is authoritative)'
    }
    [ordered]@{
        worker_pid = $PID
        excel_pid = $excelPid
        excel_hwnd = $hwnd
        excel_start_time_utc = $process.StartTime.ToUniversalTime().ToString('o')
    } | ConvertTo-Json | Set-Content -LiteralPath $coordinationPath
    try {
        $workbook = $excel.Workbooks.Add()
        $summary.workbook_preflight = $true
    } catch {
        $summary.status = 'blocked'
        $summary.classification = 'host-blocked-plain-workbooks-add'
        $summary.error = [string]$_.Exception.Message
        return
    }

    if ($isControl) {
        powershell -File scripts/register-minimal-rtd-control.ps1 -Path $dll | Out-Null
    } else {
        powershell -File scripts/register-minimal-rtd.ps1 -Path $dll | Out-Null
        powershell -File scripts/inspect-minimal-rtd-registration.ps1 -ExpectedPath $dll | Out-Null
    }
    $registered = $true
    $summary.registration = 'valid-per-user'

    $sheet = $workbook.Worksheets.Item(1)
    $listSeparator = [string]$excel.International(5)
    $formulaMatrix = @(Get-MinimalRtdFormulaMatrix -ProgId $progId -ListSeparator $listSeparator)
    for ($index = 0; $index -lt $formulaMatrix.Count; $index++) {
        $cell = $sheet.Cells.Item($index + 1, 1)
        $entry = $formulaMatrix[$index]
        if ($entry.property -eq 'FormulaLocal') { $cell.FormulaLocal = $entry.formula } else { $cell.Formula = $entry.formula }
        $summary.formula_results += [ordered]@{ name=$entry.name; requested=$entry.formula; property=$entry.property; formula=[string]$cell.Formula; formula_local=[string]$cell.FormulaLocal; value=$null }
    }
    $sheet.Range('A6').Formula = $settings.formula
    Start-Sleep -Seconds $ObservationSeconds
    for ($index = 0; $index -lt $formulaMatrix.Count; $index++) { $summary.formula_results[$index].value = $sheet.Cells.Item($index + 1, 1).Value2 }
    $summary.initial_value = $sheet.Range('A1').Value2
    $summary.duplicate_value = $sheet.Range('A6').Value2
    try {
        $ownedExcel = Get-Process -Id $excelPid -ErrorAction Stop
        $moduleMatch = @($ownedExcel.Modules | Where-Object { [string]::Equals($_.FileName, $dll, [StringComparison]::OrdinalIgnoreCase) })
        $summary.module_evidence = [ordered]@{ status='available'; exact_dll_loaded=($moduleMatch.Count -gt 0); matches=@($moduleMatch | ForEach-Object { [ordered]@{module=$_.ModuleName; path=$_.FileName} }) }
    } catch {
        $summary.module_evidence = [ordered]@{ status='access-denied'; exact_dll_loaded=$false; error=[string]$_.Exception.Message }
    }
    Start-Sleep -Seconds $ObservationSeconds
    $summary.updated_value = $sheet.Range('A1').Value2
    $sheet.Range('A1:A6').ClearContents()
    Start-Sleep -Seconds 1
    $sheet.Range('A1').Formula = $settings.formula
    Start-Sleep -Seconds $ObservationSeconds
    $summary.reconnect_value = $sheet.Range('A1').Value2
    if ($null -eq $summary.initial_value -or $null -eq $summary.updated_value -or
        [double]$summary.updated_value -le [double]$summary.initial_value -or
        $null -eq $summary.reconnect_value) {
        throw 'RTD formula did not produce the expected observable counter progression'
    }
    $summary.status = 'passed'
    $summary.classification = 'rtd-live-compatible'
} catch {
    $summary.error = [string]$_.Exception.Message
} finally {
    if ($null -ne $workbook) {
        try { $workbook.Close($false); $summary.cleanup.workbook_closed = $true } catch {}
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($workbook)
    }
    if ($null -ne $excel) {
        try { $excel.Quit(); $summary.cleanup.excel_quit = $true } catch {}
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
    }
    if ($excelPid -ne 0) {
        try {
            Wait-Process -Id $excelPid -Timeout 15 -ErrorAction Stop
            $summary.cleanup.process_exited = $true
        } catch {
            $summary.cleanup.process_exited = $null -eq (Get-Process -Id $excelPid -ErrorAction SilentlyContinue)
        }
    }
    if ($registered) {
        try {
            if ($isControl) { powershell -File scripts/unregister-minimal-rtd-control.ps1 | Out-Null }
            else { powershell -File scripts/unregister-minimal-rtd.ps1 | Out-Null }
            $summary.cleanup.unregistered = $true
        } catch {}
    }
    [Environment]::SetEnvironmentVariable($diagnosticVariable, $oldDiagnostic, 'Process')
    $events = @()
    if (Test-Path -LiteralPath $diagnostics) {
        $events = @(Get-Content -LiteralPath $diagnostics | ForEach-Object { $_ | ConvertFrom-Json })
        $summary.diagnostics = [ordered]@{
            count = $events.Count
            methods = @($events.method | Sort-Object -Unique)
            thread_ids = @($events.thread_id | Sort-Object -Unique)
            apartments = @($events.apartment | Sort-Object -Unique)
        }
    }
    if ($null -eq $summary.module_evidence) { $summary.module_evidence = [ordered]@{ status='not-observed'; exact_dll_loaded=$false } }
    $summary.activation_stage = Get-MinimalRtdActivationStage -DllLoaded ([bool]$summary.module_evidence.exact_dll_loaded) -Events $events
    $summary.completed_utc = [datetime]::UtcNow.ToString('o')
    $summary | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $summaryPath
    Write-Output ($summary | ConvertTo-Json -Depth 8)
    Write-Output "summary=$summaryPath"
}
