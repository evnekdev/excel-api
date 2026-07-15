[CmdletBinding()]
param(
    [ValidateSet('Smoke', 'Soak')] [string]$Mode = 'Smoke',
    [string]$XllPath,
    [string]$OutputDirectory,
    [int]$RecalculationCount = 0,
    [int]$LifecycleCycles = 0,
    [ValidateRange(0, 7200)] [int]$ProcessTimeoutSeconds = 0,
    [switch]$ValidateOnly,
    [switch]$Preflight,
    [switch]$Worker,
    [int]$Cycle = 0,
    [string]$CoordinationPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'excel-stress-harness-lib.ps1')

function Release-ComObject([object]$Value) {
    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Value)
    }
}

function Write-Json([object]$Value, [string]$Path) {
    $Value | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $Path -Encoding utf8
}

function Initialize-WindowProcessApi {
    if ($null -eq ('ExcelStressNativeMethods' -as [type])) {
        Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class ExcelStressNativeMethods {
    [DllImport("user32.dll", SetLastError = true)]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
}
'@
    }
}

function Get-ExcelIdentity([object]$Excel) {
    Initialize-WindowProcessApi
    $hwnd = [intptr]([long]$Excel.Hwnd)
    if ($hwnd -eq [intptr]::Zero) { throw 'Excel.Application.Hwnd returned zero' }
    [uint32]$excelProcessId = 0
    $thread = [ExcelStressNativeMethods]::GetWindowThreadProcessId($hwnd, [ref]$excelProcessId)
    if ($thread -eq 0 -or $excelProcessId -eq 0) { throw "GetWindowThreadProcessId failed for Excel HWND $hwnd" }
    $process = Get-Process -Id $excelProcessId -ErrorAction Stop
    if ($process.ProcessName -ine 'EXCEL') { throw "HWND resolved to $($process.ProcessName), not EXCEL.EXE" }
    [ordered]@{
        worker_pid = $PID
        excel_pid = [int]$excelProcessId
        excel_window_handle = [long]$hwnd
        excel_start_time_utc = $process.StartTime.ToUniversalTime().ToString('o')
        excel_version = [string]$Excel.Version
        excel_build = [string]$Excel.Build
    }
}

function Get-OwnedProcess([object]$Identity) {
    try { $process = Get-Process -Id ([int]$Identity.excel_pid) -ErrorAction Stop } catch { return $null }
    if (Test-OwnedProcessIdentity $Identity $process) { return $process }
    return $null
}

function Get-ProcessSample([object]$Identity, [string]$Stage) {
    $process = Get-OwnedProcess $Identity
    if ($null -eq $process) {
        return [ordered]@{ timestamp_utc = [datetime]::UtcNow.ToString('o'); stage = $Stage; exited = $true }
    }
    [ordered]@{
        timestamp_utc = [datetime]::UtcNow.ToString('o')
        stage = $Stage
        exited = $false
        working_set_bytes = [long]$process.WorkingSet64
        private_memory_bytes = [long]$process.PrivateMemorySize64
        handle_count = [int]$process.HandleCount
        thread_count = [int]$process.Threads.Count
    }
}

function Test-DirectoryWritable([string]$Path) {
    $record = [ordered]@{ path = $Path; exists = $false; writable = $false }
    if ([string]::IsNullOrEmpty($Path)) { return $record }
    $record.exists = Test-Path -LiteralPath $Path -PathType Container
    if (-not $record.exists) { return $record }
    $probe = Join-Path $Path ("excel-stress-{0}.tmp" -f [guid]::NewGuid().ToString('N'))
    try { [IO.File]::WriteAllText($probe, 'probe'); $record.writable = $true }
    catch { $record.error = $_.Exception.Message }
    finally { if (Test-Path -LiteralPath $probe) { Remove-Item -LiteralPath $probe -Force } }
    $record
}

function Get-EnvironmentDiagnostics([object]$Excel, [object]$Identity) {
    $process = Get-OwnedProcess $Identity
    $commandLine = $null
    try { $commandLine = (Get-CimInstance Win32_Process -Filter "ProcessId=$($Identity.excel_pid)" -ErrorAction Stop).CommandLine } catch {}
    $defaultPath = $null; $startupPath = $null
    try { $defaultPath = [string]$Excel.DefaultFilePath } catch {}
    try { $startupPath = [string]$Excel.StartupPath } catch {}
    $root = [IO.Path]::GetPathRoot($env:TEMP); $drive = $null
    try { $drive = Get-PSDrive -Name $root.Substring(0, 1) -ErrorAction Stop } catch {}
    $integritySid = $null
    try { $integritySid = ([Security.Principal.WindowsIdentity]::GetCurrent().Groups | Where-Object { $_.Value -like 'S-1-16-*' } | Select-Object -First 1).Value } catch {}
    [ordered]@{
        excel_version = [string]$Excel.Version
        excel_build = [string]$Excel.Build
        excel_operating_system = [string]$Excel.OperatingSystem
        excel_process_path = if ($null -ne $process) { $process.Path } else { $null }
        excel_session_id = if ($null -ne $process) { $process.SessionId } else { $null }
        current_user = [Security.Principal.WindowsIdentity]::GetCurrent().Name
        user_sid = [Security.Principal.WindowsIdentity]::GetCurrent().User.Value
        integrity_level_sid = $integritySid
        command_line = $commandLine
        automation_switch = $null -ne $commandLine -and $commandLine -match '(?i)/(automation|embedding)'
        safe_mode_switch = $null -ne $commandLine -and $commandLine -match '(?i)/safe\b'
        default_file_path = $defaultPath
        startup_path = $startupPath
        temp = Test-DirectoryWritable $env:TEMP
        tmp = Test-DirectoryWritable $env:TMP
        disk_free_bytes = if ($null -ne $drive) { [long]$drive.Free } else { $null }
        existing_workbooks = [int]$Excel.Workbooks.Count
        powershell_is_64_bit = [Environment]::Is64BitProcess
    }
}

function Get-ExcelFailureEvents([datetime]$Since, [datetime]$Until, [object]$Identity) {
    try {
        $events = @(Get-WinEvent -FilterHashtable @{ LogName = 'Application'; StartTime = $Since; EndTime = $Until } -ErrorAction Stop |
            Where-Object { $_.ProviderName -in @('Application Error', 'Windows Error Reporting') -and (Test-ExcelCrashEvent $_ ([int]$Identity.excel_pid)) } |
            Select-Object -First 20)
        @($events | ForEach-Object {
            [ordered]@{
                time_utc = $_.TimeCreated.ToUniversalTime().ToString('o')
                provider = $_.ProviderName
                event_id = $_.Id
                excel_pid = [int]$Identity.excel_pid
                pid_correlated = $_.Message -match ("(?i)(process id[^0-9a-f]*(0x)?{0:x}|\b{0}\b)" -f [int]$Identity.excel_pid)
                message = $_.Message
            }
        })
    }
    catch {
        if ($_.FullyQualifiedErrorId -like 'NoMatchingEventsFound*') { return @() }
        @([ordered]@{ unavailable = $true; reason = $_.Exception.Message })
    }
}

function Assert-Equal([object]$Actual, [object]$Expected, [string]$Name) {
    if ($Actual -cne $Expected) { throw "${Name}: expected '$Expected', got '$Actual'" }
}

function Assert-Error([object]$Actual, [string]$Name) {
    if ($Actual -isnot [Runtime.InteropServices.ErrorWrapper] -and "$Actual" -notmatch '^#') { throw "${Name}: expected an Excel error, got '$Actual'" }
}

function Invoke-Worker([string]$ArtifactDirectory, [System.Collections.IDictionary]$Settings) {
    $started = [datetime]::UtcNow
    $result = [ordered]@{ cycle = $Cycle; status = 'failed'; worker_pid = $PID; started_utc = $started.ToString('o'); process_samples = @(); probes = [ordered]@{} }
    $excel = $null; $books = $null; $book = $null; $plainBook = $null; $sheets = $null; $sheet = $null; $mtr = $null
    $identity = $null; $stage = 'start'; $registered = $false
    try {
        $stage = 'create Excel COM application'; $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false; $excel.DisplayAlerts = $false; $excel.AskToUpdateLinks = $false
        $identity = Get-ExcelIdentity $excel
        Write-Json $identity $CoordinationPath
        $result.excel = [ordered]@{
            pid = $identity.excel_pid; window_handle = $identity.excel_window_handle; start_time_utc = $identity.excel_start_time_utc
            version = $identity.excel_version; build = $identity.excel_build
        }
        $result.process_samples += Get-ProcessSample $identity 'excel-started'
        $result.environment = Get-EnvironmentDiagnostics $excel $identity
        $books = $excel.Workbooks

        $plainSucceeded = $false; $plainError = $null
        try { $stage = 'plain COM workbook preflight'; $plainBook = $books.Add(); $plainSucceeded = $true; [void]$plainBook.Close($false); $plainBook = $null }
        catch { $plainError = $_.Exception.ToString() }

        $stage = 'register XLL'; $registered = [bool]$excel.RegisterXLL($XllPath)
        if (-not $registered) { throw "Excel rejected RegisterXLL($XllPath)" }
        $result.process_samples += Get-ProcessSample $identity 'xll-registered'

        $xllSucceeded = $false; $xllError = $null
        try { $stage = 'post-XLL workbook preflight'; $book = $books.Add(); $xllSucceeded = $true }
        catch { $xllError = $_.Exception.ToString() }
        $result.preflight = [ordered]@{
            plain_com_workbook = [ordered]@{ succeeded = $plainSucceeded; error = $plainError }
            xll_loaded_workbook = [ordered]@{ succeeded = $xllSucceeded; error = $xllError }
            classification = Get-PreflightClassification $plainSucceeded $xllSucceeded
            registration_succeeded = $registered
        }
        if (-not $plainSucceeded -or -not $xllSucceeded) { throw "workbook preflight failed: $($result.preflight.classification)" }
        if ($Preflight) { $result.status = 'passed'; return }

        $stage = 'function setup'; $sheets = $book.Worksheets; $sheet = $sheets.Item(1)
        $unicode = [string]([char]0x0041) + [char]0x00E9 + [char]0x6C34 + [char]0xD83D + [char]0xDE00
        $sheet.Range('A1').Formula2 = '=RUST.ADD(2,3)'
        $sheet.Range('A2').Formula2 = ('=RUST.ECHO("' + $unicode + '")')
        $sheet.Range('A3').Value2 = $unicode; $sheet.Range('A4').Formula2 = '=RUST.ECHO(A3)'
        $sheet.Range('A5').Formula2 = '=RUST.REFERENCE.KIND(A1:B1)'; $sheet.Range('A6').Formula2 = '=RUST.OPTION.KIND()'
        $sheet.Range('A7').Formula2 = '=RUST.OPTION.KIND(H1)'; $sheet.Range('A8').Formula2 = '=RUST.ECHO(1)'
        $sheet.Range('A9').Formula2 = '=RUST.ASYNC.DOUBLE(21)'
        $sheet.Range('B1').Value2 = 'text'; $sheet.Range('B2').Value2 = $true; $sheet.Range('B3').Formula2 = '=NA()'
        $sheet.Range('D1').Formula2 = '=RUST.ARRAY.ECHO(A1:B3)'
        $sheet.Range('F1').Formula2 = '=RUST.ADD(ROW(),1)'; $sheet.Range("F1:F$($Settings.formula_rows)").FillDown()
        $mtr = $excel.MultiThreadedCalculation; $mtr.Enabled = $true
        $result.mtr_enabled = [bool]$mtr.Enabled; $result.mtr_threads = [int]$mtr.ThreadCount
        $result.process_samples += Get-ProcessSample $identity 'workbook-configured'

        $stage = 'recalculate'; $sampleEvery = [math]::Max(1, [math]::Ceiling($Settings.recalculations / 20))
        for ($i = 1; $i -le $Settings.recalculations; $i++) {
            $excel.CalculateFullRebuild()
            if ($i -eq 1 -or $i -eq $Settings.recalculations -or $i % $sampleEvery -eq 0) { $result.process_samples += Get-ProcessSample $identity "recalculation-$i" }
        }
        $asyncDeadline = [DateTime]::UtcNow.AddSeconds(30)
        while ([DateTime]::UtcNow -lt $asyncDeadline -and $sheet.Range('A9').Value2 -ne 42) {
            Start-Sleep -Milliseconds 25
        }
        Assert-Equal $sheet.Range('A1').Value2 5 'RUST.ADD scalar'
        Assert-Equal $sheet.Range('A2').Value2 $unicode 'RUST.ECHO direct formula string'; Assert-Equal $sheet.Range('A4').Value2 $unicode 'RUST.ECHO cell string'
        if ($sheet.Range('A5').Value2 -notin @('SRef', 'Ref')) { throw 'RUST.REFERENCE.KIND did not receive a reference' }
        Assert-Equal $sheet.Range('A6').Value2 'missing' 'RUST.OPTION.KIND omitted argument'
        if ($sheet.Range('A7').Value2 -notin @('nil', 'value')) { throw 'RUST.OPTION.KIND blank argument was not nil or value' }
        Assert-Error $sheet.Range('A8').Value2 'RUST.ECHO controlled fallback'
        Assert-Equal $sheet.Range('A9').Value2 42 'RUST.ASYNC.DOUBLE completion'
        Assert-Equal $sheet.Range('D1').Value2 5 'array number'; Assert-Equal $sheet.Range('E1').Value2 'text' 'array text'
        Assert-Equal $sheet.Range('D2').Value2 $unicode 'array UTF-16'; Assert-Equal $sheet.Range('E2').Value2 $true 'array Boolean'
        Assert-Error $sheet.Range('D3').Value2 'array error'; Assert-Equal $sheet.Range("F$($Settings.formula_rows)").Value2 ($Settings.formula_rows + 1) 'MTR last formula'
        $stage = 'invoke command'; [void]$excel.Run('RUST.PING.COMMAND')
        $result.probes = [ordered]@{
            scalar = $sheet.Range('A1').Value2; unicode_formula = $sheet.Range('A2').Value2; unicode_cell = $sheet.Range('A4').Value2
            u_reference = $sheet.Range('A5').Value2; missing = $sheet.Range('A6').Value2; nil_or_value = $sheet.Range('A7').Value2
            controlled_error = $sheet.Range('A8').Text; array_error = $sheet.Range('D3').Text
            async_double = $sheet.Range('A9').Value2
            mtr_last = $sheet.Range("F$($Settings.formula_rows)").Value2; command_invoked = 'RUST.PING.COMMAND'
        }
        $stage = 'save workbook'; $book.SaveAs((Join-Path $ArtifactDirectory "cycle-$Cycle.xlsx"))
        $result.process_samples += Get-ProcessSample $identity 'before-unregister'
        $result.status = 'passed'
    }
    catch { $result.error = "stage=$stage; $($_.Exception.ToString())" }
    finally {
        if ($null -ne $book) { try { [void]$book.Close($false) } catch {} }
        if ($registered -and $null -ne $excel) { try { $result.unregistration_succeeded = [bool]$excel.UnregisterXLL($XllPath) } catch { $result.unregistration_succeeded = $false } }
        if ($null -ne $excel) { try { [void]$excel.Quit() } catch {} }
        Release-ComObject $mtr; Release-ComObject $sheet; Release-ComObject $sheets; Release-ComObject $plainBook; Release-ComObject $book; Release-ComObject $books; Release-ComObject $excel
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
        if ($null -ne $identity) {
            $deadline = [datetime]::UtcNow.AddSeconds(10)
            while ([datetime]::UtcNow -lt $deadline -and $null -ne (Get-OwnedProcess $identity)) { Start-Sleep -Milliseconds 100 }
            $result.process_samples += Get-ProcessSample $identity 'after-exit'
            $result.failure_events = Get-ExcelFailureEvents $started ([datetime]::UtcNow) $identity
        }
        else { $result.failure_events = @() }
        $result.process_trend = Measure-ProcessSamples $result.process_samples
        $result.finished_utc = [datetime]::UtcNow.ToString('o')
        $result.elapsed_ms = [int](([datetime]::UtcNow - $started).TotalMilliseconds)
        $result.schema_valid = Test-HarnessResultSchema ([pscustomobject]$result)
        if (-not $result.schema_valid -and $result.status -eq 'passed') { $result.status = 'failed'; $result.error = 'result schema validation failed' }
        Write-Json $result (Join-Path $ArtifactDirectory "cycle-$Cycle.json")
    }
    if ($result.status -ne 'passed') { throw $result.error }
}

function Quote-ProcessArgument([string]$Value) { '"' + ($Value -replace '(\\*)"', '$1$1\"' -replace '(\\+)$', '$1$1') + '"' }

function Start-DirectWorker([string]$Shell, [string[]]$Arguments, [string]$WorkingDirectory) {
    $start = New-Object Diagnostics.ProcessStartInfo
    $start.FileName = $Shell; $start.WorkingDirectory = $WorkingDirectory
    $start.Arguments = ($Arguments | ForEach-Object { Quote-ProcessArgument $_ }) -join ' '
    $start.UseShellExecute = $false; $start.RedirectStandardOutput = $true; $start.RedirectStandardError = $true; $start.CreateNoWindow = $true
    # Leave EnvironmentVariables untouched so .NET inherits the exact shell
    # environment block without enumerating duplicate PATH/Path keys.
    $process = New-Object Diagnostics.Process; $process.StartInfo = $start
    if (-not $process.Start()) { throw 'failed to start direct stress worker' }
    [pscustomobject]@{ process = $process; output = $process.StandardOutput.ReadToEndAsync(); error = $process.StandardError.ReadToEndAsync() }
}

function Stop-OwnedExcel([string]$CoordinationFile) {
    if (-not (Test-Path -LiteralPath $CoordinationFile)) { return [ordered]@{ killed = $false; reason = 'coordination-file-missing' } }
    try { $identity = Get-Content -Raw -LiteralPath $CoordinationFile | ConvertFrom-Json } catch { return [ordered]@{ killed = $false; reason = 'coordination-file-invalid' } }
    $owned = Get-OwnedProcess $identity
    if ($null -eq $owned) { return [ordered]@{ killed = $false; reason = 'owned-excel-not-found-or-start-time-mismatch' } }
    Stop-Process -Id $owned.Id -Force
    [ordered]@{ killed = $true; excel_pid = $owned.Id; start_time_verified = $true }
}

function Invoke-Harness {
    $settings = Get-HarnessSettings $Mode $RecalculationCount $LifecycleCycles $ProcessTimeoutSeconds
    $run = Join-Path $OutputDirectory ("{0}-{1}" -f $(if ($Preflight) { 'preflight' } else { $Mode.ToLowerInvariant() }), (Get-Date -Format 'yyyyMMddTHHmmssZ'))
    New-Item -ItemType Directory -Force -Path $run | Out-Null
    $summary = [ordered]@{ mode = $Mode; preflight_only = [bool]$Preflight; xll = $XllPath; settings = $settings; effective_timeout_seconds = $settings.timeout_seconds; started_utc = [datetime]::UtcNow.ToString('o'); cycles = @(); status = 'failed' }
    try {
        $cycles = if ($Preflight) { 1 } else { $settings.lifecycle_cycles }
        for ($number = 1; $number -le $cycles; $number++) {
            $coord = Join-Path $run "cycle-$number.ownership.json"; $out = Join-Path $run "cycle-$number.worker.out.log"; $err = Join-Path $run "cycle-$number.worker.err.log"
            $shell = (Get-Process -Id $PID).Path
            $args = @('-Sta', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $PSCommandPath, '-Worker', '-Cycle', $number, '-Mode', $Mode, '-XllPath', $XllPath, '-OutputDirectory', $run, '-CoordinationPath', $coord, '-RecalculationCount', $settings.recalculations, '-LifecycleCycles', 1, '-ProcessTimeoutSeconds', $settings.timeout_seconds)
            if ($Preflight) { $args += '-Preflight' }
            $worker = Start-DirectWorker $shell $args $workspace
            if (-not $worker.process.WaitForExit($settings.timeout_seconds * 1000)) {
                $worker.process.Kill(); $worker.process.WaitForExit()
                $summary.timeout_cleanup = Stop-OwnedExcel $coord
                throw "cycle $number timed out after $($settings.timeout_seconds) seconds"
            }
            Set-Content -LiteralPath $out -Value $worker.output.Result -Encoding utf8; Set-Content -LiteralPath $err -Value $worker.error.Result -Encoding utf8
            $resultPath = Join-Path $run "cycle-$number.json"
            if ($worker.process.ExitCode -ne 0 -or -not (Test-Path -LiteralPath $resultPath)) { throw "cycle $number failed; see $err" }
            $cycleResult = Get-Content -Raw -LiteralPath $resultPath | ConvertFrom-Json
            if ([int]$cycleResult.worker_pid -ne $worker.process.Id) { throw "worker PID mismatch for cycle $number" }
            $cycleResult | Add-Member -NotePropertyName worker_exit_code -NotePropertyValue $worker.process.ExitCode
            $summary.cycles += $cycleResult
        }
        $summary.status = 'passed'
    }
    catch { $summary.error = $_.Exception.ToString() }
    finally {
        $summary.finished_utc = [datetime]::UtcNow.ToString('o')
        $summary.process_trends = @($summary.cycles | ForEach-Object { $_.process_trend })
        Write-Json $summary (Join-Path $run 'summary.json')
    }
    Write-Output "Excel stress artifacts: $run"
    if ($summary.status -ne 'passed') { throw $summary.error }
}

$workspace = Split-Path -Parent $PSScriptRoot
if ([string]::IsNullOrEmpty($OutputDirectory)) { $OutputDirectory = Join-Path $workspace 'target/excel-stress' }
if ([string]::IsNullOrEmpty($XllPath)) { $XllPath = Join-Path $workspace 'target/release/minimal_xll.xll' }
$XllPath = [IO.Path]::GetFullPath($XllPath)
$settings = Get-HarnessSettings $Mode $RecalculationCount $LifecycleCycles $ProcessTimeoutSeconds
if (-not (Test-Path -LiteralPath $XllPath) -and -not $ValidateOnly) { throw "XLL not found: $XllPath" }
if ($ValidateOnly) {
    & (Join-Path $PSScriptRoot 'test-excel-stress-harness.ps1')
    [ordered]@{ mode = $Mode; xll = $XllPath; settings = $settings; effective_timeout_seconds = $settings.timeout_seconds; required_fields_validated = $true; status = 'validated-no-excel' } | ConvertTo-Json -Depth 5
    exit 0
}
if ($Worker) { Invoke-Worker $OutputDirectory $settings } else { Invoke-Harness }
