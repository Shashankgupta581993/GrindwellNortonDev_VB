param([Parameter(Mandatory=$true)][string]$AssemblyPath, [string]$EnabledChannels='')
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
$assembly=[Reflection.Assembly]::LoadFrom((Resolve-Path -LiteralPath $AssemblyPath))
$prefix='GrindwellNortonDev_VB.'
$flags=[Reflection.BindingFlags]'Static,Instance,Public,NonPublic'
$settings=$assembly.GetType($prefix+'SchedulerDebugSettings',$true)
$channels=$assembly.GetType($prefix+'SchedulerDebugChannel',$true)
$type=$assembly.GetType($prefix+'SchedulerDebugCollector',$true)
$expected=@($EnabledChannels.Split(',') | Where-Object { $_ -and $_ -ne '__none__' })
$collector=[Activator]::CreateInstance($type)
$collector.Enabled=$true
$collector.MetricsEnabled=$true
foreach($name in [Enum]::GetNames($channels)) {
    $actual=$settings.GetMethod('IsEnabled',$flags).Invoke($null,@([Enum]::Parse($channels,$name)))
    if($actual -ne ($expected -contains $name)) { throw "Effective switch failed: $name" }
}
$entries=@(
    @('FieldMap','AddFieldMap','DebugFieldMapRow','FieldMapRows','02_FieldMap.csv'),
    @('WipDiagnostics','AddWipDiagnostic','WipDiagnosticRow','WipDiagnostics','05_WipChainDiagnostics.csv'),
    @('StageEligibility','AddStageEligibility','StageEligibilityRow','StageEligibilityRows','06_StageEligibilityDiagnostics.csv'),
    @('OptimizerCandidateTrace','TraceCandidateStep','OptimizerCandidateTraceRow','OptimizerCandidateTraceRows','07_OptimizerCandidateTrace.csv'),
    @('BatchPlanTrace','TracePlanStep','BatchTunnelSwkPlanTraceRow','BatchTunnelSwkPlanTraceRows','08_BatchTunnelSwkPlanTrace.csv'),
    @('ScheduleAttemptTrace','TraceScheduleAttempt','ScheduleAttemptTraceRow','ScheduleAttemptTraceRows','09_ScheduleAttemptTrace.csv'),
    @('ResourceValidation','AddResourceValidation','ResourceValidationRow','ResourceValidationRows','10_ResourceValidation.csv'),
    @('ActionMetrics','AddActionMetrics','SchedulerActionMetricsRow','ActionMetricsRows','12_ActionMetrics.csv'))
foreach($entry in $entries) {
    $row=[Activator]::CreateInstance($assembly.GetType($prefix+$entry[2],$true))
    $type.GetMethod($entry[1]).Invoke($collector,@($row)) | Out-Null
    if($type.GetField($entry[3]).GetValue($collector).Count -ne [int]($expected -contains $entry[0])) { throw "Collection gate failed: $($entry[0])" }
}
$needsSnapshot=@('OperationSnapshots','WipDiagnostics','StageEligibility','OptimizerCandidateTrace','ScheduleAttemptTrace') | Where-Object { $expected -contains $_ }
if($type.GetProperty('NeedsSnapshot',$flags).GetValue($collector) -ne [bool]$needsSnapshot) { throw 'Snapshot dependency failed.' }
$stages=$assembly.GetType($prefix+'SchedulerStageDiagnostics',$true)
if(-not $needsSnapshot -and $expected -notcontains 'FieldMap') {
    if($stages.GetMethod('BuildOrderOperationSnapshot').Invoke($null,@($null,$collector)).Count -ne 0) { throw 'Disabled snapshot accessed board.' }
}
foreach($entry in @(@('WipDiagnostics','DiagnoseWip'),@('StageEligibility','DiagnoseAll'))) {
    if($expected -notcontains $entry[0]) { $stages.GetMethod($entry[1]).Invoke($null,@($null,$collector)) | Out-Null }
}
$folder=Join-Path ([IO.Path]::GetTempPath()) ('gn_settings_'+[Guid]::NewGuid().ToString('N'))
$previous=[Environment]::CurrentDirectory
try {
    [IO.Directory]::CreateDirectory($folder) | Out-Null
    [Environment]::CurrentDirectory=$folder
    $collector.ExportFolder=Join-Path $folder 'export'
    $collector.ExportAll($null)
    $files=@{}
    foreach($entry in $entries) { $files[$entry[0]]=$entry[4] }
    $files.OperationSnapshots='03_OrderOperationSnapshot.csv'
    $files.ReasonSummary='11_ReasonCodeSummary.csv'
    $files.ProblemStatementExport='13_GptProblemStatement.json'
    foreach($name in $files.Keys) {
        if((Test-Path -LiteralPath (Join-Path $collector.ExportFolder $files[$name])) -ne ($expected -contains $name)) { throw "Export gate failed: $name" }
    }
    if($expected -contains 'ProblemStatementExport') {
        $manifest=Get-Content -Raw -LiteralPath (Join-Path $collector.ExportFolder $files.ProblemStatementExport) | ConvertFrom-Json
        foreach($file in $manifest.files) { if(-not (Test-Path -LiteralPath (Join-Path $collector.ExportFolder $file))) { throw "Missing manifest file: $file" } }
    }
    foreach($entry in @(@('firingOptimizer_vf','ExportPlanToCsv','BatchDebugCsv'),@('firingOptimizer_vf','ExportFiringCandidateDebug','BatchDebugCsv'),@('swkOptimizer_vf','ExportSwkPlanToCsv','SwkDebugCsv'))) {
        if($expected -notcontains $entry[2]) {
            $exporter=$assembly.GetType($prefix+$entry[0],$true)
            $exporter.GetMethod($entry[1]).Invoke([Activator]::CreateInstance($exporter),@($null,[string](Join-Path $folder 'forbidden'))) | Out-Null
        }
    }
    if(Test-Path -LiteralPath (Join-Path $folder 'forbidden')) { throw 'Disabled standalone export wrote files.' }
    $legacy=$assembly.GetType($prefix+'AlgoSeq3',$true)
    $legacy.GetMethod('InitLogging',$flags).Invoke([Activator]::CreateInstance($legacy),@()) | Out-Null
    $logs=@(Get-ChildItem -LiteralPath $folder -Recurse -Filter '*.csv' | Where-Object { $_.Directory.Name -eq 'AlgoSeqLogs' })
    if(@($logs | Where-Object Name -Like 'DecisionLog*').Count -ne [int]($expected -contains 'AlgoSeq3DecisionLog')) { throw 'Legacy decision gate failed.' }
    if(@($logs | Where-Object Name -Like 'PlannedSchedule*').Count -ne [int]($expected -contains 'AlgoSeq3ScheduleLog')) { throw 'Legacy schedule gate failed.' }
    if($expected.Count -eq 0) {
        if(Test-Path -LiteralPath $collector.ExportFolder) { throw 'Disabled diagnostics created a directory.' }
        $tool=$assembly.GetType($prefix+'SchedulerDiagnosticsTool',$true)
        if($tool.GetMethod('ExportSchedulingDiagnostics').Invoke([Activator]::CreateInstance($tool),@($null)) -ne 0) { throw 'Disabled diagnostic tool accessed COM.' }
    }
    Write-Output "PASS: code settings, collection/export gates and dependencies [$EnabledChannels]."
} finally {
    [Environment]::CurrentDirectory=$previous
    $resolved=[IO.Path]::GetFullPath($folder)
    if($resolved.StartsWith([IO.Path]::GetFullPath([IO.Path]::GetTempPath()),[StringComparison]::OrdinalIgnoreCase)) { Remove-Item -LiteralPath $resolved -Recurse -Force }
}
