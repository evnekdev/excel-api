[CmdletBinding()]
param(
    [ValidateSet('Smoke', 'Soak')]
    [string]$Mode = 'Smoke',
    [string]$XllPath,
    [string]$OutputDirectory,
    [int]$RecalculationCount = 0,
    [int]$LifecycleCycles = 0,
    [ValidateRange(30, 3600)] [int]$ProcessTimeoutSeconds = 180,
    [switch]$ValidateOnly,
    [switch]$Worker,
    [int]$Cycle = 0
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Release-ComObject([object]$Value) {
    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Value)
    }
}

function Write-Json([object]$Value, [string]$Path) {
    $Value | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $Path -Encoding utf8
}

function Get-ExcelProcesses {
    @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | ForEach-Object {
        [ordered]@{ id = $_.Id; started = $_.StartTime.ToUniversalTime().ToString('o'); working_set_bytes = $_.WorkingSet64; private_memory_bytes = $_.PrivateMemorySize64; handles = $_.HandleCount }
    })
}

function Get-FailureEvents([datetime]$Since) {
    try {
        @(Get-WinEvent -FilterHashtable @{ LogName = 'Application'; StartTime = $Since } -ErrorAction Stop |
            Where-Object { $_.ProviderName -in @('Application Error', 'Windows Error Reporting') } |
            Select-Object -First 20 | ForEach-Object {
                [ordered]@{ time = $_.TimeCreated.ToUniversalTime().ToString('o'); provider = $_.ProviderName; id = $_.Id; message = $_.Message }
            })
    }
    catch { @([ordered]@{ unavailable = $true; reason = $_.Exception.Message }) }
}

function Assert-Equal([object]$Actual, [object]$Expected, [string]$Name) {
    if ($Actual -cne $Expected) { throw "${Name}: expected '$Expected', got '$Actual'" }
}

function Assert-Error([object]$Actual, [string]$Name) {
    if ($Actual -isnot [System.Runtime.InteropServices.ErrorWrapper] -and "$Actual" -notmatch '^#') { throw "${Name}: expected an Excel error, got '$Actual'" }
}

function Get-Settings {
    $defaults = if ($Mode -eq 'Smoke') { @{ recalculations = 25; lifecycle_cycles = 2; formula_rows = 128 } } else { @{ recalculations = 2000; lifecycle_cycles = 25; formula_rows = 1024 } }
    [ordered]@{
        recalculations = if ($RecalculationCount -gt 0) { $RecalculationCount } else { $defaults.recalculations }
        lifecycle_cycles = if ($LifecycleCycles -gt 0) { $LifecycleCycles } else { $defaults.lifecycle_cycles }
        formula_rows = $defaults.formula_rows
    }
}

function Start-StressWorker([string]$Shell, [string[]]$Arguments) {
    # Start-Job avoids Start-Process's PATH/Path environment bug in Windows
    # PowerShell while still giving the parent a separately killable worker.
    Start-Job -ArgumentList $Shell, $Arguments -ScriptBlock {
        param($workerShell, $workerArguments)
        & $workerShell @workerArguments
        [int]$LASTEXITCODE
    }
}

function Invoke-Worker([string]$ArtifactDirectory, [System.Collections.IDictionary]$Settings) {
    $started = [datetime]::UtcNow
    $result = [ordered]@{
        cycle = $Cycle
        started_utc = $started.ToString('o')
        status = 'failed'
        before_processes = Get-ExcelProcesses
        diagnostics = [ordered]@{ status = 'unavailable'; reason = 'The sample XLL has no COM-readable diagnostic snapshot; controlled errors and Windows crash evidence are recorded instead.' }
        probes = [ordered]@{}
    }
    $excel = $null; $books = $null; $book = $null; $sheets = $null; $sheet = $null; $stage = 'start'
    try {
        $stage = 'create Excel COM application'; $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false; $excel.DisplayAlerts = $false; $excel.AskToUpdateLinks = $false
        $stage = 'register XLL'; if (-not $excel.RegisterXLL($XllPath)) { throw "Excel rejected RegisterXLL($XllPath)" }
        $stage = 'create workbook'; $books = $excel.Workbooks; $book = $books.Add(); $sheets = $book.Worksheets; $sheet = $sheets.Item(1)
        $unicode = [string]([char]0x0041) + [char]0x00E9 + [char]0x6C34 + [char]0xD83D + [char]0xDE00
        $sheet.Range('A1').Formula2 = '=RUST.ADD(2,3)'
        $sheet.Range('A2').Formula2 = ('=RUST.ECHO("' + $unicode + '")')
        $sheet.Range('A3').Value2 = $unicode
        $sheet.Range('A4').Formula2 = '=RUST.ECHO(A3)'
        $sheet.Range('A5').Formula2 = '=RUST.REFERENCE.KIND(A1:B1)'
        $sheet.Range('A6').Formula2 = '=RUST.OPTION.KIND()'
        $sheet.Range('A7').Formula2 = '=RUST.OPTION.KIND(H1)'
        $sheet.Range('A8').Formula2 = '=RUST.ECHO(1)'
        $sheet.Range('B1').Value2 = 'text'; $sheet.Range('B2').Value2 = $true; $sheet.Range('B3').Formula2 = '=NA()'
        $sheet.Range('D1').Formula2 = '=RUST.ARRAY.ECHO(A1:B3)'
        $sheet.Range('F1').Formula2 = '=RUST.ADD(ROW(),1)'
        $sheet.Range("F1:F$($Settings.formula_rows)").FillDown()
        $mtr = $excel.MultiThreadedCalculation; $mtr.Enabled = $true
        $result.mtr_enabled = [bool]$mtr.Enabled; $result.mtr_threads = [int]$mtr.ThreadCount
        $stage = 'recalculate'; for ($i = 0; $i -lt $Settings.recalculations; $i++) { $excel.CalculateFullRebuild() }
        Assert-Equal $sheet.Range('A1').Value2 5 'RUST.ADD scalar'
        Assert-Equal $sheet.Range('A2').Value2 $unicode 'RUST.ECHO direct formula string'
        Assert-Equal $sheet.Range('A4').Value2 $unicode 'RUST.ECHO direct cell string'
        if ($sheet.Range('A5').Value2 -notin @('SRef', 'Ref')) { throw 'RUST.REFERENCE.KIND did not receive a reference' }
        Assert-Equal $sheet.Range('A6').Value2 'missing' 'RUST.OPTION.KIND omitted argument'
        if ($sheet.Range('A7').Value2 -notin @('nil', 'value')) { throw 'RUST.OPTION.KIND blank argument was not nil or value' }
        Assert-Error $sheet.Range('A8').Value2 'RUST.ECHO controlled conversion fallback'
        Assert-Equal $sheet.Range('D1').Value2 5 'RUST.ARRAY.ECHO scalar array value'
        Assert-Equal $sheet.Range('E1').Value2 'text' 'RUST.ARRAY.ECHO string array value'
        Assert-Equal $sheet.Range('D2').Value2 $unicode 'RUST.ARRAY.ECHO UTF-16 array value'
        Assert-Equal $sheet.Range('E2').Value2 $true 'RUST.ARRAY.ECHO Boolean array value'
        Assert-Error $sheet.Range('D3').Value2 'RUST.ARRAY.ECHO error array value'
        Assert-Equal $sheet.Range("F$($Settings.formula_rows)").Value2 ($Settings.formula_rows + 1) 'MTR repeated formula'
        $stage = 'invoke command'; [void]$excel.Run('RUST.PING.COMMAND')
        $result.probes = [ordered]@{
            scalar = $sheet.Range('A1').Value2; unicode_formula = $sheet.Range('A2').Value2; unicode_cell = $sheet.Range('A4').Value2
            q_array = @($sheet.Range('D1').Value2, $sheet.Range('E1').Value2, $sheet.Range('D2').Value2, $sheet.Range('E2').Value2)
            u_reference = $sheet.Range('A5').Value2; missing = $sheet.Range('A6').Value2; nil_or_value = $sheet.Range('A7').Value2
            controlled_error = $sheet.Range('A8').Text; error_array = $sheet.Range('D3').Text; mtr_last = $sheet.Range("F$($Settings.formula_rows)").Value2; command = 'RUST.PING.COMMAND'
        }
        $result.excel = [ordered]@{ version = $excel.Version; build = $excel.Build; operating_system = $excel.OperatingSystem }
        $stage = 'save workbook'; $book.SaveAs((Join-Path $ArtifactDirectory "cycle-$Cycle.xlsx"))
        $stage = 'unregister XLL'; if (-not $excel.UnregisterXLL($XllPath)) { throw "Excel rejected UnregisterXLL($XllPath)" }
        [void]$book.Close($false); $book = $null; [void]$excel.Quit(); $excel = $null
        $result.status = 'passed'
    }
    catch { $result.error = "stage=$stage; $($_.Exception.ToString())" }
    finally {
        if ($null -ne $book) { [void]$book.Close($false) }; if ($null -ne $excel) { [void]$excel.Quit() }
        Release-ComObject $sheet; Release-ComObject $sheets; Release-ComObject $book; Release-ComObject $books; Release-ComObject $excel
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
        $result.finished_utc = [datetime]::UtcNow.ToString('o'); $result.elapsed_ms = [int](([datetime]::UtcNow - $started).TotalMilliseconds)
        $result.after_processes = Get-ExcelProcesses; $result.failure_events = Get-FailureEvents $started
        Write-Json $result (Join-Path $ArtifactDirectory "cycle-$Cycle.json")
    }
    if ($result.status -ne 'passed') { throw $result.error }
}

function Invoke-Harness {
    $settings = Get-Settings
    $run = Join-Path $OutputDirectory ("{0}-{1}" -f $Mode.ToLowerInvariant(), (Get-Date -Format 'yyyyMMddTHHmmssZ'))
    New-Item -ItemType Directory -Force -Path $run | Out-Null
    $summary = [ordered]@{ mode = $Mode; xll = $XllPath; settings = $settings; started_utc = [datetime]::UtcNow.ToString('o'); cycles = @(); status = 'failed' }
    try {
        for ($number = 1; $number -le $settings.lifecycle_cycles; $number++) {
            $before = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
            $workerLog = Join-Path $run "cycle-$number.worker.out.log"; $workerErrorLog = Join-Path $run "cycle-$number.worker.err.log"; $shell = (Get-Process -Id $PID).Path
            $arguments = @('-Sta', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $PSCommandPath, '-Worker', '-Cycle', $number, '-Mode', $Mode, '-XllPath', $XllPath, '-OutputDirectory', $run, '-RecalculationCount', $settings.recalculations, '-LifecycleCycles', 1, '-ProcessTimeoutSeconds', $ProcessTimeoutSeconds)
            $worker = Start-StressWorker $shell $arguments
            if ($null -eq (Wait-Job -Job $worker -Timeout $ProcessTimeoutSeconds)) {
                Stop-Job -Job $worker -ErrorAction SilentlyContinue
                @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | Where-Object { $_.Id -notin $before }) | Stop-Process -Force
                throw "cycle $number timed out after $ProcessTimeoutSeconds seconds; worker and its new Excel processes were terminated"
            }
            $workerOutput = @(Receive-Job -Job $worker)
            Set-Content -LiteralPath $workerLog -Value $workerOutput -Encoding utf8
            Set-Content -LiteralPath $workerErrorLog -Value $worker.ChildJobs[0].JobStateInfo.Reason -Encoding utf8
            $workerExitCode = if ($workerOutput.Count -gt 0) { [int]$workerOutput[-1] } else { 1 }
            Remove-Job -Job $worker -Force
            $resultPath = Join-Path $run "cycle-$number.json"
            if ($workerExitCode -ne 0 -or -not (Test-Path -LiteralPath $resultPath)) { throw "cycle $number failed; see $workerLog" }
            $cycleResult = Get-Content -Raw -LiteralPath $resultPath | ConvertFrom-Json
            $cycleResult | Add-Member -NotePropertyName worker_exit_code -NotePropertyValue $workerExitCode
            $summary.cycles += $cycleResult
        }
        $summary.status = 'passed'
    }
    catch { $summary.error = $_.Exception.ToString() }
    finally { $summary.finished_utc = [datetime]::UtcNow.ToString('o'); Write-Json $summary (Join-Path $run 'summary.json') }
    Write-Output "Excel stress artifacts: $run"
    if ($summary.status -ne 'passed') { throw $summary.error }
}

$workspace = Split-Path -Parent $PSScriptRoot
if ([string]::IsNullOrEmpty($OutputDirectory)) { $OutputDirectory = Join-Path $workspace 'target/excel-stress' }
if ([string]::IsNullOrEmpty($XllPath)) { $XllPath = Join-Path $workspace 'target/release/minimal_xll.xll' }
$XllPath = [IO.Path]::GetFullPath($XllPath)
if (-not (Test-Path -LiteralPath $XllPath) -and -not $ValidateOnly) { throw "XLL not found: $XllPath. Build it with scripts/build-minimal-xll.ps1 first." }
if ($ValidateOnly) { [ordered]@{ mode = $Mode; xll = $XllPath; settings = Get-Settings; timeout_seconds = $ProcessTimeoutSeconds; status = 'validated-no-excel' } | ConvertTo-Json; exit 0 }
if ($Worker) { Invoke-Worker $OutputDirectory (Get-Settings) } else { Invoke-Harness }
