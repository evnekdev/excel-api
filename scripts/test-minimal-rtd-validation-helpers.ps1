[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'minimal-rtd-validation-helpers.ps1')

function Assert-Equal($Expected, $Actual, [string]$Message) {
    if ($Expected -ne $Actual) { throw "$Message expected=[$Expected] actual=[$Actual]" }
}

$comma = @(Get-MinimalRtdFormulaMatrix -ListSeparator ',')
Assert-Equal 5 $comma.Count 'formula count'
Assert-Equal '=RTD("ExcelApi.MinimalRtd",,"COUNTER")' $comma[0].formula 'omitted server formula'
Assert-Equal '=RTD("ExcelApi.MinimalRtd","","COUNTER")' $comma[1].formula 'explicit server formula'
$semicolon = @(Get-MinimalRtdFormulaMatrix -ListSeparator ';')
Assert-Equal '=RTD("ExcelApi.MinimalRtd";"";"COUNTER")' $semicolon[4].formula 'local separator formula'

Assert-Equal 'dll_not_loaded' (Get-MinimalRtdActivationStage -DllLoaded $false) 'no load stage'
Assert-Equal 'dll_loaded_no_class_factory' (Get-MinimalRtdActivationStage -DllLoaded $true) 'module-only stage'
Assert-Equal 'class_factory_reached' (Get-MinimalRtdActivationStage -DllLoaded $true -Events @([pscustomobject]@{method='DllGetClassObject_enter'})) 'factory stage'
Assert-Equal 'refresh_observed' (Get-MinimalRtdActivationStage -DllLoaded $true -Events @([pscustomobject]@{method='RefreshData_enter'})) 'refresh stage'

Assert-Equal 'A' (Get-MinimalRtdDecision -RustStage not_run -ControlStage failed_before_server_start) 'decision A'
Assert-Equal 'B' (Get-MinimalRtdDecision -RustStage dll_not_loaded -ControlStage server_start_reached) 'decision B'
Assert-Equal 'C' (Get-MinimalRtdDecision -RustStage server_start_reached -ControlStage server_start_reached) 'decision C'
Assert-Equal 'D' (Get-MinimalRtdDecision -RustStage topic_connected -ControlStage lifecycle_succeeded) 'decision D'
Assert-Equal 'E' (Get-MinimalRtdDecision -RustStage lifecycle_succeeded -ControlStage lifecycle_succeeded) 'decision E'
Assert-Equal 'unclassified' (Get-MinimalRtdDecision -RustStage not_run -ControlStage not_run) 'blocked decision'

Assert-Equal $false (Test-MinimalRtdRegistrationConflict -RegisteredPath $null -ExpectedPath 'x' -RegisteredProgId $null -ExpectedProgId 'p') 'absent registration'
Assert-Equal $false (Test-MinimalRtdRegistrationConflict -RegisteredPath 'C:\x.dll' -ExpectedPath 'c:\X.dll' -RegisteredProgId 'P' -ExpectedProgId 'p') 'matching registration'
Assert-Equal $true (Test-MinimalRtdRegistrationConflict -RegisteredPath 'C:\old.dll' -ExpectedPath 'C:\x.dll' -RegisteredProgId 'P' -ExpectedProgId 'P') 'path conflict'

[PSCustomObject]@{ status = 'passed'; assertions = 18 } | ConvertTo-Json
