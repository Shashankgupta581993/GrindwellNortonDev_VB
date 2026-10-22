param([Parameter(Mandatory = $true)][string]$AssemblyPath)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$assembly = [Reflection.Assembly]::LoadFrom((Resolve-Path -LiteralPath $AssemblyPath))
$collector = $assembly.GetType('GrindwellNortonDev_VB.SchedulerDebugCollector', $true)
$names = @('GN_SCHEDULER_DIAGNOSTICS_OFF', 'GN_SCHEDULER_DEBUG', 'GN_SCHEDULER_METRICS')
$saved = @{}
foreach ($name in $names) { $saved[$name] = [Environment]::GetEnvironmentVariable($name, 'Process') }
try {
    # Change only this short-lived test process, never the user's saved settings.
    [Environment]::SetEnvironmentVariable('GN_SCHEDULER_DEBUG', '1', 'Process')
    [Environment]::SetEnvironmentVariable('GN_SCHEDULER_METRICS', '1', 'Process')
    [Environment]::SetEnvironmentVariable('GN_SCHEDULER_DIAGNOSTICS_OFF', '0', 'Process')
    $before = [Activator]::CreateInstance($collector)
    foreach ($value in @('1', 'true', 'ON', 'yes')) {
        [Environment]::SetEnvironmentVariable('GN_SCHEDULER_DIAGNOSTICS_OFF', $value, 'Process')
        $disabled = [Activator]::CreateInstance($collector)
        if ($disabled.Enabled -or $disabled.MetricsEnabled) { throw "Kill switch failed: $value" }
    }
    [Environment]::SetEnvironmentVariable('GN_SCHEDULER_DIAGNOSTICS_OFF', '0', 'Process')
    $after = [Activator]::CreateInstance($collector)
    if ($before.Enabled -ne $after.Enabled -or $before.MetricsEnabled -ne $after.MetricsEnabled) {
        throw 'Disabling the kill switch did not restore existing configuration behaviour.'
    }
    Write-Output 'PASS: diagnostics-off overrides debug/metrics and is reversible; persistent settings unchanged.'
} finally {
    foreach ($name in $names) { [Environment]::SetEnvironmentVariable($name, $saved[$name], 'Process') }
}
