[CmdletBinding()]
param(
    [switch]$ValidateOnly,
    [string]$XllPath = "target/release/minimal_xll_ontime_research.xll",
    [string]$OutputDirectory = "target/ontime-validation"
)

$ErrorActionPreference = 'Stop'

function Assert-SourceContract {
    $header = Get-Content -LiteralPath 'tools/Excel2013XLLSDK/INCLUDE/XLCALL.H' -Raw
    if ($header -notmatch '#define\s+xlcOnTime\s+\(148\s*\|\s*xlCommand\)') {
        throw 'checked-in XLCALL.H does not contain the expected xlcOnTime definition'
    }
    if ($header -notmatch '#define\s+xlfNow\s+74') {
        throw 'checked-in XLCALL.H does not contain the expected xlfNow definition'
    }
    $source = Get-Content -LiteralPath 'examples/minimal-xll/src/lib.rs' -Raw
    foreach ($name in @('RUST.ONTIME.SCHEDULE', 'RUST.ONTIME.CALLBACK', 'RUST.ONTIME.CANCEL', 'RUST.ONTIME.STATUS', 'RUST.ONTIME.DUMP')) {
        if ($source -notmatch [regex]::Escape($name)) { throw "missing experimental registration: $name" }
    }
    $manifest = Get-Content -LiteralPath 'examples/minimal-xll/Cargo.toml' -Raw
    if ($manifest -notmatch 'xlcontime-research\s*=') {
        throw 'research XLL feature is not declared'
    }
    if ($source -notmatch '#\[cfg\(feature = "xlcontime-research"\)\]') {
        throw 'research registrations are not feature-gated'
    }
    Write-Output 'PASS: xlcOnTime validation sources and checked-in constants are present.'
}

Assert-SourceContract
if ($ValidateOnly) { exit 0 }

& powershell -File scripts/build-minimal-xll.ps1 -Profile release -XlcOnTimeResearch
if ($LASTEXITCODE -ne 0) { throw "research XLL build failed with exit code $LASTEXITCODE" }
$resolvedXll = (Resolve-Path -LiteralPath $XllPath).Path
New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$summaryPath = Join-Path $OutputDirectory "xlcOnTime-$stamp.json"
$statusPath = [IO.Path]::GetFullPath((Join-Path $OutputDirectory 'excel-api-xlcontime-status.json'))
$enableMarker = [IO.Path]::GetFullPath((Join-Path $OutputDirectory 'excel-api-xlcontime-enable.marker'))
Remove-Item -LiteralPath $statusPath -ErrorAction SilentlyContinue

if (-not ('ExcelOnTime.NativeMethods' -as [type])) {
    Add-Type -TypeDefinition @'
namespace ExcelOnTime {
    using System;
    using System.Runtime.InteropServices;
    public static class NativeMethods {
        [DllImport("user32.dll")]
        public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
    }
}
'@
}

function Read-OnTimeStatus([string]$StatusPath) {
    $raw = Get-Content -LiteralPath $StatusPath -Raw
    return $raw | ConvertFrom-Json
}

function Stop-OwnedExcelProcess([Diagnostics.Process]$OwnedProcess) {
    $current = Get-Process -Id $OwnedProcess.Id -ErrorAction SilentlyContinue
    if ($null -eq $current) { return $true }
    if ($current.ProcessName -ne 'EXCEL') {
        throw "refusing to stop PID $($OwnedProcess.Id): process is not EXCEL.EXE"
    }
    if ($current.StartTime.ToUniversalTime() -ne $OwnedProcess.StartTime.ToUniversalTime()) {
        throw "refusing to stop PID $($OwnedProcess.Id): start time changed"
    }
    Stop-Process -Id $OwnedProcess.Id -Force
    Wait-Process -Id $OwnedProcess.Id -Timeout 20 -ErrorAction SilentlyContinue
    return -not [bool](Get-Process -Id $OwnedProcess.Id -ErrorAction SilentlyContinue)
}

function Read-XlmSecurity {
    $paths = @(
        'HKCU:\Software\Policies\Microsoft\Office\16.0\Excel\Security',
        'HKCU:\Software\Microsoft\Office\16.0\Excel\Security'
    )
    $records = @()
    foreach ($path in $paths) {
        if (Test-Path -LiteralPath $path) {
            $value = Get-ItemProperty -LiteralPath $path
            $records += [pscustomobject]@{
                scope = $path
                xl4MacroOff = $value.XL4MacroOff
                excel4MacroWarnings = $value.Excel4MacroWarnings
            }
        }
    }
    return $records
}

$excel = $null
$ownedProcess = $null
$terminateOwnedProcess = $false
$summary = [ordered]@{
    classification = 'failed'
    startedUtc = (Get-Date).ToUniversalTime().ToString('o')
    command = "powershell -File scripts/excel-ontime-validation.ps1"
    xll = $resolvedXll
    security = Read-XlmSecurity
    excel = $null
    plainWorkbookAdd = $null
    postXllWorkbookAdd = $null
    registerXll = $false
    bootstrapAttempted = $false
    bootstrapSucceeded = $false
    unsafeToUnload = $false
    commandInvocation = 'not-attempted'
    twoArgumentSchedule = $false
    latestTimeSchedule = $false
    callbacksObserved = 0
    callbackOnMainThread = $false
    macroContextCallSucceeded = $false
    cancelBeforeExecution = $false
    repeatedCancelRejected = $false
    cancellationBeforeUnload = $false
    cleanupAction = 'not-attempted'
    processExited = $false
    status = $null
    pendingManualMatrix = @(
        'XLM setting enabled versus disabled',
        'trusted versus ordinary XLL location',
        'signed versus unsigned XLL',
        'recalculation delay and worksheet-function overlap',
        'cell edit and modal-dialog delay',
        'latest-time expiry',
        'cut/copy, undo, flicker, latency, and idle CPU',
        'Add-in Manager unload/reload before scheduled time'
    )
    errors = @()
}

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $hwnd = [int64]$excel.Hwnd
    [uint32]$ownedExcelPid = 0
    [void][ExcelOnTime.NativeMethods]::GetWindowThreadProcessId([IntPtr]$hwnd, [ref]$ownedExcelPid)
    $ownedProcess = Get-Process -Id $ownedExcelPid
    if ($ownedProcess.ProcessName -ne 'EXCEL') { throw "owned PID $ownedExcelPid is not EXCEL.EXE" }
    Set-Content -LiteralPath $enableMarker -Value ([string]$ownedExcelPid)
    $summary.excel = [ordered]@{
        version = [string]$excel.Version
        build = [string]$excel.Build
        operatingSystem = [string]$excel.OperatingSystem
        architecture = if ([Environment]::Is64BitProcess) { '64-bit process' } else { '32-bit process' }
        hwnd = $hwnd
        pid = [int]$ownedExcelPid
        processStartUtc = $ownedProcess.StartTime.ToUniversalTime().ToString('o')
        mainThreadId = $null
    }

    try {
        $book = $excel.Workbooks.Add()
        $book.Close($false)
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book)
        $summary.plainWorkbookAdd = 'passed'
    } catch {
        $summary.plainWorkbookAdd = "failed: $($_.Exception.Message)"
    }

    $summary.registerXll = [bool]$excel.RegisterXLL($resolvedXll)
    if (-not $summary.registerXll) { throw 'RegisterXLL returned FALSE' }

    $statusDeadline = (Get-Date).AddSeconds(5)
    while (-not (Test-Path -LiteralPath $statusPath) -and (Get-Date) -lt $statusDeadline) {
        Start-Sleep -Milliseconds 100
    }
    if (-not (Test-Path -LiteralPath $statusPath)) {
        throw 'the XLL bootstrap produced no diagnostics file'
    }

    try {
        $book = $excel.Workbooks.Add()
        $book.Close($false)
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book)
        $summary.postXllWorkbookAdd = 'passed'
    } catch {
        $summary.postXllWorkbookAdd = "failed: $($_.Exception.Message)"
    }

    $initial = Read-OnTimeStatus $statusPath
    $summary.bootstrapAttempted = [bool]$initial.bootstrap_attempted
    $summary.bootstrapSucceeded = [bool]$initial.bootstrap_succeeded
    $summary.unsafeToUnload = [bool]$initial.unsafe_to_unload
    if (-not $summary.bootstrapAttempted) {
        throw 'research marker was present but bootstrap did not report an attempt'
    }
    if (-not $summary.bootstrapSucceeded) {
        $terminateOwnedProcess = $summary.unsafeToUnload -or ([int]$initial.pending_count -gt 0)
        throw "research bootstrap failed: $($initial.bootstrap_failure)"
    }
    $summary.excel.mainThreadId = [uint64]$initial.main_thread_id
    $summary.twoArgumentSchedule = [bool]($initial.events | Where-Object { $_.kind -eq 'schedule' -and $_.form -eq 'two' -and $_.raw_code -eq 0 -and $_.result -eq 'bool:true' })
    $summary.latestTimeSchedule = [bool]($initial.events | Where-Object { $_.kind -eq 'schedule' -and $_.form -eq 'latest' -and $_.raw_code -eq 0 -and $_.result -eq 'bool:true' })
    $summary.cancelBeforeExecution = [bool]($initial.events | Where-Object { $_.kind -eq 'cancel' -and $_.raw_code -eq 0 -and $_.result -eq 'bool:true' })
    $summary.repeatedCancelRejected = [bool]($initial.events | Where-Object { $_.kind -eq 'cancel-missing' -and $_.raw_code -eq 0 -and $_.result -eq 'error:15' })
    try {
        [void]$excel.Run('RUST.ONTIME.DUMP')
        $summary.commandInvocation = 'passed'
    } catch {
        $summary.commandInvocation = "blocked: $($_.Exception.Message)"
    }

    $deadline = (Get-Date).AddSeconds(20)
    do {
        Start-Sleep -Milliseconds 250
        $status = Read-OnTimeStatus $statusPath
    } while ([int]$status.callback_count -lt 2 -and (Get-Date) -lt $deadline)
    $summary.callbacksObserved = [int]$status.callback_count
    $callbackEvents = @($status.events | Where-Object kind -eq 'callback')
    $summary.callbackOnMainThread = [bool]($callbackEvents.Count -ge 2 -and @($callbackEvents | Where-Object { [uint64]$_.thread_id -ne [uint64]$status.main_thread_id }).Count -eq 0)
    $summary.macroContextCallSucceeded = [bool]($callbackEvents.Count -ge 2 -and @($callbackEvents | Where-Object { $_.result -notlike 'macro-context-xlAbort:*' }).Count -eq 0)

    $summary.status = Read-OnTimeStatus $statusPath
    try {
        while ([int]$summary.status.pending_count -gt 0) {
            [void]$excel.Run('RUST.ONTIME.CANCEL')
            [void]$excel.Run('RUST.ONTIME.DUMP')
            $summary.status = Read-OnTimeStatus $statusPath
        }
        $summary.cancellationBeforeUnload = $true
    } catch {
        $summary.errors += "cancellation before unload failed: $($_.Exception.Message)"
        if (Test-Path -LiteralPath $statusPath) {
            $summary.status = Read-OnTimeStatus $statusPath
        }
        $terminateOwnedProcess = ([int]$summary.status.pending_count -gt 0) -or [bool]$summary.status.unsafe_to_unload
        throw 'pending callbacks could not be proved canceled; isolated Excel will be terminated'
    }
    $excel.Quit()
    $summary.cleanupAction = 'Excel.Quit after pending count reached zero'
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
    $excel = $null
    Start-Sleep -Milliseconds 250
    if (Test-Path -LiteralPath $statusPath) {
        $summary.status = Read-OnTimeStatus $statusPath
        $summary.unsafeToUnload = [bool]$summary.status.unsafe_to_unload
    }
    try {
        Wait-Process -Id $ownedExcelPid -Timeout 20 -ErrorAction Stop
    } catch {
        $summary.errors += "owned Excel PID did not exit after Quit: $ownedExcelPid"
    }
    $summary.processExited = -not [bool](Get-Process -Id $ownedExcelPid -ErrorAction SilentlyContinue)

    $required = @(
        $summary.registerXll,
        $summary.bootstrapAttempted,
        $summary.bootstrapSucceeded,
        $summary.twoArgumentSchedule,
        $summary.latestTimeSchedule,
        ($summary.callbacksObserved -ge 2),
        $summary.callbackOnMainThread,
        $summary.macroContextCallSucceeded,
        $summary.cancelBeforeExecution,
        $summary.repeatedCancelRejected,
        $summary.cancellationBeforeUnload,
        $summary.processExited
    )
    if ($required -notcontains $false) {
        $summary.classification = if ($summary.plainWorkbookAdd -eq 'passed' -and $summary.postXllWorkbookAdd -eq 'passed') { 'core-pass-workbook-pass-manual-matrix-pending' } else { 'core-pass-workbook-blocked-manual-matrix-pending' }
    }
} catch {
    $summary.errors += $_.Exception.ToString()
} finally {
    if ($null -ne $excel) {
        if (-not $terminateOwnedProcess) {
            try {
                $excel.Quit()
                $summary.cleanupAction = 'Excel.Quit'
            } catch {}
        }
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
        $excel = $null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    if ($null -ne $ownedProcess) {
        if ($terminateOwnedProcess) {
            try {
                $summary.processExited = Stop-OwnedExcelProcess $ownedProcess
                $summary.cleanupAction = 'terminated exact owned Excel PID after unsafe/inconclusive cancellation'
            } catch {
                $summary.errors += $_.Exception.ToString()
            }
        } else {
            try { Wait-Process -Id $ownedProcess.Id -Timeout 15 -ErrorAction SilentlyContinue } catch {}
        }
        $summary.processExited = -not [bool](Get-Process -Id $ownedProcess.Id -ErrorAction SilentlyContinue)
    }
    if ($summary.classification -eq 'failed' -and
        $summary.registerXll -and
        $summary.plainWorkbookAdd -like 'failed:*' -and
        $null -eq $summary.status) {
        $summary.classification = 'blocked-host'
    }
    if ($summary.registerXll -and -not $summary.bootstrapSucceeded) {
        $summary.classification = if ($terminateOwnedProcess) { 'unsafe-inconclusive' } else { 'bootstrap-failed' }
    }
    $summary.completedUtc = (Get-Date).ToUniversalTime().ToString('o')
    $summary | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $summaryPath -Encoding UTF8
    Remove-Item -LiteralPath $enableMarker -ErrorAction SilentlyContinue
    Write-Output "RESULT: $($summary.classification)"
    Write-Output "SUMMARY: $summaryPath"
}

if ($summary.classification -notlike 'core-pass-*') { exit 1 }
