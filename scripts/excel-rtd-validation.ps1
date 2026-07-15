[CmdletBinding()]
param(
    [string]$DllPath = 'target/release/excel_api_minimal_rtd.dll',
    [string]$OutputDirectory = 'target/rtd-validation',
    [int]$ObservationSeconds = 4,
    [int]$ProcessTimeoutSeconds = 90,
    [string]$RunDirectory,
    [switch]$Worker,
    [switch]$ValidateOnly
)

$ErrorActionPreference = 'Stop'
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
$settings = [ordered]@{
    prog_id = 'ExcelApi.MinimalRtd'
    clsid = '{DC738FE5-30EE-40E8-A8C2-3D16F217C52D}'
    dll = $dll
    observation_seconds = $ObservationSeconds
    formula = '=RTD("ExcelApi.MinimalRtd","","COUNTER")'
}
if ($ValidateOnly) {
    powershell -File scripts/register-minimal-rtd.ps1 -Path $dll -ValidateOnly | Out-Null
    [PSCustomObject]@{ status = 'validated-no-excel'; settings = $settings } | ConvertTo-Json -Depth 4
    return
}

if (-not $Worker) {
    New-Item -ItemType Directory -Path $run -Force | Out-Null
    $stdout = Join-Path $run 'worker.stdout.txt'
    $stderr = Join-Path $run 'worker.stderr.txt'
    $arguments = @(
        '-NoProfile', '-File', $PSCommandPath, '-Worker',
        '-DllPath', $dll, '-OutputDirectory', $OutputDirectory,
        '-ObservationSeconds', $ObservationSeconds,
        '-ProcessTimeoutSeconds', $ProcessTimeoutSeconds,
        '-RunDirectory', $run
    )
    $process = Start-Process -FilePath powershell.exe -ArgumentList $arguments `
        -WindowStyle Hidden -RedirectStandardOutput $stdout `
        -RedirectStandardError $stderr -PassThru
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
                $childResult = [ordered]@{ pid = [int]$child.ProcessId; result = 'not-terminated' }
                try {
                    Stop-Process -Id $child.ProcessId -Force -ErrorAction Stop
                    $childResult.result = 'verified-direct-child-terminated'
                } catch {
                    $childResult.result = 'verified-direct-child-termination-denied'
                }
                $excelCleanup.direct_excel_children += $childResult
            }
        }
        powershell -File scripts/unregister-minimal-rtd.ps1 | Out-Null
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
    cleanup = [ordered]@{ workbook_closed = $false; excel_quit = $false; process_exited = $false; unregistered = $false }
}
$excel = $null
$workbook = $null
$excelPid = 0
$registered = $false
$oldDiagnostic = [Environment]::GetEnvironmentVariable('EXCEL_API_MINIMAL_RTD_DIAGNOSTICS', 'Process')
try {
    [Environment]::SetEnvironmentVariable('EXCEL_API_MINIMAL_RTD_DIAGNOSTICS', $diagnostics, 'Process')
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

    powershell -File scripts/register-minimal-rtd.ps1 -Path $dll | Out-Null
    powershell -File scripts/inspect-minimal-rtd-registration.ps1 -ExpectedPath $dll | Out-Null
    $registered = $true
    $summary.registration = 'valid-per-user'

    $sheet = $workbook.Worksheets.Item(1)
    $sheet.Range('A1').Formula = $settings.formula
    $sheet.Range('A2').Formula = $settings.formula
    Start-Sleep -Seconds $ObservationSeconds
    $summary.initial_value = $sheet.Range('A1').Value2
    $summary.duplicate_value = $sheet.Range('A2').Value2
    Start-Sleep -Seconds $ObservationSeconds
    $summary.updated_value = $sheet.Range('A1').Value2
    $sheet.Range('A1:A2').ClearContents()
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
            powershell -File scripts/unregister-minimal-rtd.ps1 | Out-Null
            $summary.cleanup.unregistered = $true
        } catch {}
    }
    [Environment]::SetEnvironmentVariable('EXCEL_API_MINIMAL_RTD_DIAGNOSTICS', $oldDiagnostic, 'Process')
    if (Test-Path -LiteralPath $diagnostics) {
        $events = @(Get-Content -LiteralPath $diagnostics | ForEach-Object { $_ | ConvertFrom-Json })
        $summary.diagnostics = [ordered]@{
            count = $events.Count
            methods = @($events.method | Sort-Object -Unique)
            thread_ids = @($events.thread_id | Sort-Object -Unique)
            apartments = @($events.apartment | Sort-Object -Unique)
        }
    }
    $summary.completed_utc = [datetime]::UtcNow.ToString('o')
    $summary | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $summaryPath
    Write-Output ($summary | ConvertTo-Json -Depth 8)
    Write-Output "summary=$summaryPath"
}
