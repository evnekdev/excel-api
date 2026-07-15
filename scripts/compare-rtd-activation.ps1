[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$RustSummary,
    [Parameter(Mandatory)][string]$ControlSummary
)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'minimal-rtd-validation-helpers.ps1')
$rust = Get-Content -LiteralPath $RustSummary -Raw | ConvertFrom-Json
$control = Get-Content -LiteralPath $ControlSummary -Raw | ConvertFrom-Json
$controlStage = if ($control.status -eq 'passed') { 'lifecycle_succeeded' } elseif ($control.activation_stage -in @('server_start_reached','topic_connected','refresh_observed')) { 'server_start_reached' } else { 'failed_before_server_start' }
$rustStage = if ($rust.status -eq 'passed') { 'lifecycle_succeeded' } else { [string]$rust.activation_stage }
[pscustomobject]@{
    rust_stage = $rustStage
    control_stage = $controlStage
    decision = Get-MinimalRtdDecision -RustStage $rustStage -ControlStage $controlStage
} | ConvertTo-Json
