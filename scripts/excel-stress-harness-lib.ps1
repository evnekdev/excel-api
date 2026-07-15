Set-StrictMode -Version Latest

function Get-HarnessSettings([string]$Mode, [int]$RecalculationCount, [int]$LifecycleCycles, [int]$ProcessTimeoutSeconds) {
    $defaults = if ($Mode -eq 'Soak') {
        @{ recalculations = 2000; lifecycle_cycles = 25; formula_rows = 1024; timeout_seconds = 1800 }
    }
    else {
        @{ recalculations = 25; lifecycle_cycles = 2; formula_rows = 128; timeout_seconds = 180 }
    }
    [ordered]@{
        recalculations = if ($RecalculationCount -gt 0) { $RecalculationCount } else { $defaults.recalculations }
        lifecycle_cycles = if ($LifecycleCycles -gt 0) { $LifecycleCycles } else { $defaults.lifecycle_cycles }
        formula_rows = $defaults.formula_rows
        timeout_seconds = if ($ProcessTimeoutSeconds -gt 0) { $ProcessTimeoutSeconds } else { $defaults.timeout_seconds }
    }
}

function Test-OwnedProcessIdentity([object]$Identity, [object]$Process) {
    if ($null -eq $Identity -or $null -eq $Process) { return $false }
    if ([int]$Identity.excel_pid -ne [int]$Process.Id) { return $false }
    if ([string]$Process.ProcessName -ine 'EXCEL') { return $false }
    $expected = ([datetime]$Identity.excel_start_time_utc).ToUniversalTime()
    $actual = ([datetime]$Process.StartTime).ToUniversalTime()
    return [math]::Abs(($actual - $expected).TotalMilliseconds) -lt 1000
}

function Select-OwnedExcelProcess([object]$Identity, [object[]]$Processes) {
    if ($null -eq $Identity) { return $null }
    foreach ($process in @($Processes)) {
        if (Test-OwnedProcessIdentity $Identity $process) { return $process }
    }
    return $null
}

function Get-TimeoutCleanupPlan([int]$WorkerPid, [object]$Identity, [object[]]$Processes) {
    $owned = Select-OwnedExcelProcess $Identity $Processes
    [ordered]@{
        worker_pid = $WorkerPid
        excel_pid = if ($null -ne $owned) { [int]$owned.Id } else { $null }
        excel_identity_verified = $null -ne $owned
        unrelated_pids = @($Processes | Where-Object { $null -eq $owned -or $_.Id -ne $owned.Id } | ForEach-Object { [int]$_.Id })
    }
}

function Test-ExcelCrashEvent([object]$Event, [int]$ExcelPid) {
    if ($null -eq $Event -or [string]::IsNullOrEmpty([string]$Event.Message)) { return $false }
    return [string]$Event.Message -match '(?i)\bEXCEL\.EXE\b'
}

function Measure-ProcessSamples([object[]]$Samples) {
    $live = @($Samples | Where-Object { -not $_.exited })
    if ($live.Count -eq 0) { return [ordered]@{ sample_count = 0 } }
    $summary = [ordered]@{ sample_count = $live.Count; first = $live[0]; last = $live[-1] }
    foreach ($field in @('working_set_bytes', 'private_memory_bytes', 'handle_count', 'thread_count')) {
        $values = @($live | ForEach-Object { [long]($_.$field) })
        $measure = $values | Measure-Object -Minimum -Maximum
        $summary[$field] = [ordered]@{
            first = $values[0]
            last = $values[-1]
            minimum = [long]$measure.Minimum
            maximum = [long]$measure.Maximum
            delta = [long]($values[-1] - $values[0])
        }
    }
    $summary
}

function Get-PreflightClassification([bool]$PlainSucceeded, [bool]$XllSucceeded) {
    if (-not $PlainSucceeded) { return 'plain-com-failed' }
    if (-not $XllSucceeded) { return 'post-xll-failed' }
    return 'passed'
}

function Test-HarnessResultSchema([object]$Result) {
    foreach ($name in @('cycle', 'status', 'worker_pid', 'excel', 'preflight', 'process_samples', 'process_trend', 'failure_events')) {
        if ($Result -is [System.Collections.IDictionary]) {
            if (-not $Result.Contains($name)) { return $false }
        }
        elseif ($null -eq $Result.PSObject.Properties[$name]) { return $false }
    }
    foreach ($name in @('pid', 'window_handle', 'start_time_utc', 'version', 'build')) {
        if ($Result.excel -is [System.Collections.IDictionary]) {
            if (-not $Result.excel.Contains($name)) { return $false }
        }
        elseif ($null -eq $Result.excel.PSObject.Properties[$name]) { return $false }
    }
    return $true
}
