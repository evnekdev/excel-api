Set-StrictMode -Version Latest

function Get-MinimalRtdFormulaMatrix {
    [CmdletBinding()]
    param(
        [string]$ProgId = 'ExcelApi.MinimalRtd',
        [ValidateNotNullOrEmpty()][string]$ListSeparator = ','
    )

    $local = $ListSeparator
    @(
        [PSCustomObject]@{ name = 'invariant-omitted-server'; property = 'Formula'; formula = "=RTD(`"$ProgId`",,`"COUNTER`")" }
        [PSCustomObject]@{ name = 'invariant-explicit-empty-server'; property = 'Formula'; formula = "=RTD(`"$ProgId`",`"`",`"COUNTER`")" }
        [PSCustomObject]@{ name = 'lower-case-progid'; property = 'Formula'; formula = "=RTD(`"$($ProgId.ToLowerInvariant())`",,`"COUNTER`")" }
        [PSCustomObject]@{ name = 'upper-case-progid'; property = 'Formula'; formula = "=RTD(`"$($ProgId.ToUpperInvariant())`",,`"COUNTER`")" }
        [PSCustomObject]@{ name = 'local-explicit-empty-server'; property = 'FormulaLocal'; formula = "=RTD(`"$ProgId`"$local`"`"$local`"COUNTER`")" }
    )
}

function Get-MinimalRtdActivationStage {
    [CmdletBinding()]
    param(
        [bool]$DllLoaded,
        [object[]]$Events = @()
    )

    $methods = @($Events | ForEach-Object { [string]$_.method })
    if ($methods -contains 'RefreshData_enter' -or $methods -contains 'RefreshData') { return 'refresh_observed' }
    if ($methods -contains 'ConnectData_enter' -or $methods -contains 'ConnectData') { return 'topic_connected' }
    if ($methods -contains 'ServerStart_enter' -or $methods -contains 'ServerStart') { return 'server_start_reached' }
    if ($methods -contains 'CreateInstance_exit') { return 'object_created' }
    if ($methods -contains 'DllGetClassObject_enter') { return 'class_factory_reached' }
    if ($DllLoaded) { return 'dll_loaded_no_class_factory' }
    'dll_not_loaded'
}

function Get-MinimalRtdDecision {
    [CmdletBinding()]
    param(
        [ValidateSet('not_run', 'dll_not_loaded', 'dll_loaded_no_class_factory', 'class_factory_reached', 'object_created', 'server_start_reached', 'topic_connected', 'refresh_observed', 'lifecycle_succeeded')]
        [string]$RustStage,
        [ValidateSet('not_run', 'failed_before_server_start', 'server_start_reached', 'lifecycle_succeeded')]
        [string]$ControlStage
    )

    if ($ControlStage -eq 'failed_before_server_start') { return 'A' }
    if ($ControlStage -in @('server_start_reached', 'lifecycle_succeeded') -and
        $RustStage -in @('dll_not_loaded', 'dll_loaded_no_class_factory', 'class_factory_reached')) { return 'B' }
    if ($RustStage -in @('object_created', 'server_start_reached')) { return 'C' }
    if ($RustStage -eq 'topic_connected') { return 'D' }
    if ($RustStage -eq 'lifecycle_succeeded') { return 'E' }
    'unclassified'
}

function Test-MinimalRtdRegistrationConflict {
    [CmdletBinding()]
    param(
        [AllowNull()][string]$RegisteredPath,
        [string]$ExpectedPath,
        [AllowNull()][string]$RegisteredProgId,
        [string]$ExpectedProgId
    )

    if ([string]::IsNullOrWhiteSpace($RegisteredPath) -and [string]::IsNullOrWhiteSpace($RegisteredProgId)) {
        return $false
    }
    -not ([string]::Equals($RegisteredPath, $ExpectedPath, [StringComparison]::OrdinalIgnoreCase) -and
        [string]::Equals($RegisteredProgId, $ExpectedProgId, [StringComparison]::OrdinalIgnoreCase))
}
