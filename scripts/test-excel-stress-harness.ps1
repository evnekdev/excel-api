$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'excel-stress-harness-lib.ps1')

function Assert([bool]$Condition, [string]$Message) {
    if (-not $Condition) { throw $Message }
}

$smoke = Get-HarnessSettings 'Smoke' 0 0 0
$soak = Get-HarnessSettings 'Soak' 0 0 0
Assert ($smoke.timeout_seconds -eq 180) 'Smoke timeout default'
Assert ($soak.timeout_seconds -eq 1800) 'Soak timeout default'
Assert ((Get-HarnessSettings 'Soak' 0 0 333).timeout_seconds -eq 333) 'Timeout override'

$start = [datetime]'2026-01-01T00:00:00Z'
$identity = [pscustomobject]@{ excel_pid = 41; excel_start_time_utc = $start.ToString('o') }
$owned = [pscustomobject]@{ Id = 41; ProcessName = 'EXCEL'; StartTime = $start }
$unrelated = [pscustomobject]@{ Id = 42; ProcessName = 'EXCEL'; StartTime = $start.AddSeconds(1) }
$reused = [pscustomobject]@{ Id = 41; ProcessName = 'EXCEL'; StartTime = $start.AddMinutes(1) }
Assert (Test-OwnedProcessIdentity $identity $owned) 'Exact PID identity selection'
Assert (-not (Test-OwnedProcessIdentity $identity $reused)) 'PID reuse rejection'
Assert ((Select-OwnedExcelProcess $identity @($unrelated, $owned)).Id -eq 41) 'Unrelated process exclusion'
Assert ($null -eq (Select-OwnedExcelProcess $null @($unrelated))) 'Missing PID-file behavior'
$plan = Get-TimeoutCleanupPlan 99 $identity @($unrelated, $owned)
Assert ($plan.worker_pid -eq 99 -and $plan.excel_pid -eq 41) 'Timeout cleanup selection'
Assert ($plan.unrelated_pids -contains 42) 'Timeout cleanup retains unrelated PID'

Assert (Test-ExcelCrashEvent ([pscustomobject]@{ Message = 'Faulting application name: EXCEL.EXE, process id: 0x29' }) 41) 'Excel event selection'
Assert (-not (Test-ExcelCrashEvent ([pscustomobject]@{ Message = 'Faulting application name: OTHER.EXE' }) 41)) 'Non-Excel event rejection'

$samples = @(
    [pscustomobject]@{ exited = $false; working_set_bytes = 10; private_memory_bytes = 20; handle_count = 3; thread_count = 4 },
    [pscustomobject]@{ exited = $false; working_set_bytes = 15; private_memory_bytes = 18; handle_count = 5; thread_count = 4 }
)
$trend = Measure-ProcessSamples $samples
Assert ($trend.working_set_bytes.delta -eq 5 -and $trend.handle_count.maximum -eq 5) 'Process sample aggregation'
Assert ((Get-PreflightClassification $false $false) -eq 'plain-com-failed') 'Plain preflight classification'
Assert ((Get-PreflightClassification $true $false) -eq 'post-xll-failed') 'XLL preflight classification'
Assert ((Get-PreflightClassification $true $true) -eq 'passed') 'Successful preflight classification'

$schema = [pscustomobject]@{
    cycle = 1; status = 'passed'; worker_pid = 99
    excel = [pscustomobject]@{ pid = 41; window_handle = 1; start_time_utc = $start; version = '16'; build = '1' }
    preflight = [pscustomobject]@{}; process_samples = @(); process_trend = [pscustomobject]@{}; failure_events = @()
}
Assert (Test-HarnessResultSchema $schema) 'Result JSON schema'

Write-Output 'PASS: 15 Excel stress harness helper assertions.'
