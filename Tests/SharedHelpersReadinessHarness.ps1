param(
    [string]$AssemblyPath = ""
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($AssemblyPath)) {
    $AssemblyPath = Join-Path $root "obj\x64\Debug\GrindwellNortonDev_VB.dll"
}

if (-not (Test-Path -LiteralPath $AssemblyPath)) {
    throw "Build output not found: $AssemblyPath. Run MSBuild first or pass -AssemblyPath."
}

$assembly = [Reflection.Assembly]::LoadFrom((Resolve-Path -LiteralPath $AssemblyPath))
$helpers = $assembly.GetType("GrindwellNortonDev_VB.SharedHelpers", $true)
$failures = 0

function Invoke-SharedHelper {
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [AllowEmptyCollection()][object[]]$Arguments = @()
    )

    $methods = @($helpers.GetMethods([Reflection.BindingFlags]::Public -bor [Reflection.BindingFlags]::Static) |
        Where-Object { $_.Name -eq $Name -and $_.GetParameters().Count -eq $Arguments.Count })

    if ($Name -eq "IsCompletedOrActualizedOp") {
        $methods = @($methods | Where-Object {
            $_.GetParameters()[0].ParameterType.FullName -like "System.Collections.Generic.IDictionary*"
        })
    }

    if ($methods.Count -ne 1) {
        throw "Could not resolve helper method '$Name' with $($Arguments.Count) argument(s). Matches=$($methods.Count)."
    }

    return ,($methods[0].Invoke($null, $Arguments))
}

function Assert-True {
    param([bool]$Value, [string]$Message)
    if ($Value) { return }
    $script:failures += 1
    Write-Error "FAIL: $Message" -ErrorAction Continue
}

function Assert-False {
    param([bool]$Value, [string]$Message)
    Assert-True -Value (-not $Value) -Message $Message
}

function Assert-Equal {
    param([int]$Expected, [int]$Actual, [string]$Message)
    if ($Expected -eq $Actual) { return }
    $script:failures += 1
    Write-Error "FAIL: $Message Expected=$Expected Actual=$Actual" -ErrorAction Continue
}

function Args1 {
    param($A)
    $args = New-Object 'object[]' 1
    $args[0] = $A
    return ,$args
}

function Args2 {
    param($A, $B)
    $args = New-Object 'object[]' 2
    $args[0] = $A
    $args[1] = $B
    return ,$args
}

function New-Routing {
    return ,(Invoke-SharedHelper -Name "BuildRoutingSchema" -Arguments @())
}

function Add-Operation {
    param(
        [Parameter(Mandatory = $true)]$Table,
        [Parameter(Mandatory = $true)][int]$Record,
        [Parameter(Mandatory = $true)][string]$OrderNo,
        [Parameter(Mandatory = $true)][int]$OpNo,
        [bool]$Released = $false,
        [DateTime]$ReleaseTime = [DateTime]::MinValue,
        [bool]$Completed = $false,
        [bool]$Actualized = $false,
        [DateTime]$ActualEndTime = [DateTime]::MinValue,
        [int]$KilnType = 0,
        [int]$ParentRecord = 1
    )

    $row = $Table.NewRow()
    $row["OrdersID"] = $Record
    $row["Order No"] = $OrderNo
    $row["Operation Number"] = $OpNo
    $row["parent_record"] = $ParentRecord
    $row["Kiln Type"] = $KilnType
    $row["is_scheduled"] = $false
    $row["source_is_completed"] = $Completed
    $row["opcenter_use_actual"] = $Actualized
    if ($ActualEndTime -ne [DateTime]::MinValue) {
        $row["actual_end_time"] = $ActualEndTime
    }
    $row["operation_effective_completed"] = $Completed
    $row["operation_releases_next"] = $Released
    if ($ReleaseTime -ne [DateTime]::MinValue) {
        $row["operation_release_time"] = $ReleaseTime
    }
    [void]$Table.Rows.Add($row)
}

function Get-Readiness {
    param($Table)
    return ,(Invoke-SharedHelper -Name "BuildFiringReadinessByOrder" -Arguments (Args1 $Table))
}

function Get-LatestReleaseBoundaries {
    param($Table)
    return ,(Invoke-SharedHelper -Name "BuildLatestReleaseBoundaries" -Arguments (Args1 $Table))
}

function Test-DirectPressingToLoadingReadiness {
    $dt = New-Routing
    Add-Operation $dt 1 "A" 200 -Released $true -ReleaseTime ([DateTime]"2026-07-09T08:00:00") -Completed $true
    Add-Operation $dt 2 "A" 290
    Add-Operation $dt 3 "A" 300

    $readiness = Get-Readiness $dt
    Assert-True $readiness.ContainsKey("A") "200 -> 290 direct route should produce firing readiness."
    Assert-Equal 200 $readiness["A"].LastReleaseOpNo "Direct route readiness should use op 200."
    Assert-False $readiness["A"].LoadingAlreadyReleased "Direct route should still require loading."
}

function Test-DryingCompletedReadiness {
    $dt = New-Routing
    Add-Operation $dt 10 "B" 200 -Released $true -ReleaseTime ([DateTime]"2026-07-09T07:00:00") -Completed $true
    Add-Operation $dt 11 "B" 260 -Released $true -ReleaseTime ([DateTime]"2026-07-09T09:00:00") -Completed $true
    Add-Operation $dt 12 "B" 290
    Add-Operation $dt 13 "B" 300

    $readiness = Get-Readiness $dt
    Assert-True $readiness.ContainsKey("B") "Completed drying route should produce firing readiness."
    Assert-Equal 260 $readiness["B"].LastReleaseOpNo "Drying route readiness should use latest released pre-290 op."
}

function Test-LoadingCompletedReadiness {
    $dt = New-Routing
    Add-Operation $dt 20 "C" 200 -Released $true -ReleaseTime ([DateTime]"2026-07-09T06:00:00") -Completed $true
    Add-Operation $dt 21 "C" 290 -Released $true -ReleaseTime ([DateTime]"2026-07-09T10:00:00") -Completed $true
    Add-Operation $dt 22 "C" 300

    $readiness = Get-Readiness $dt
    Assert-True $readiness.ContainsKey("C") "Completed loading route should produce firing readiness."
    Assert-Equal 290 $readiness["C"].LastReleaseOpNo "Loading readiness should use op 290."
    Assert-True $readiness["C"].LoadingAlreadyReleased "Completed loading should not require loading time again."
}

function Test-CompletedPostFiringSkipPrimitive {
    $dt = New-Routing
    Add-Operation $dt 30 "D" 310 -Released $true -ReleaseTime ([DateTime]"2026-07-09T11:00:00") -Completed $true
    Add-Operation $dt 31 "D" 320
    $rows = Invoke-SharedHelper -Name "BuildOperationRowIndex" -Arguments (Args1 $dt)

    Assert-True (Invoke-SharedHelper -Name "IsCompletedOrActualizedOp" -Arguments (Args2 $rows 30)) "Completed op 310 should be skipped."
    Assert-False (Invoke-SharedHelper -Name "IsCompletedOrActualizedOp" -Arguments (Args2 $rows 31)) "Incomplete op 320 should remain schedulable."
}

function Test-CompletedFourHundredOnwardSkipPrimitive {
    $dt = New-Routing
    Add-Operation $dt 40 "E" 400 -Released $true -ReleaseTime ([DateTime]"2026-07-09T11:00:00") -Completed $true
    Add-Operation $dt 41 "E" 410 -Released $true -ReleaseTime ([DateTime]"2026-07-09T12:00:00") -Completed $true
    Add-Operation $dt 42 "E" 420 -Released $true -ReleaseTime ([DateTime]"2026-07-09T13:00:00") -Completed $true
    Add-Operation $dt 43 "E" 430
    $rows = Invoke-SharedHelper -Name "BuildOperationRowIndex" -Arguments (Args1 $dt)

    Assert-True (Invoke-SharedHelper -Name "IsCompletedOrActualizedOp" -Arguments (Args2 $rows 40)) "Completed op 400 should be skipped."
    Assert-True (Invoke-SharedHelper -Name "IsCompletedOrActualizedOp" -Arguments (Args2 $rows 41)) "Completed op 410 should be skipped."
    Assert-True (Invoke-SharedHelper -Name "IsCompletedOrActualizedOp" -Arguments (Args2 $rows 42)) "Completed op 420 should be skipped."
    Assert-False (Invoke-SharedHelper -Name "IsCompletedOrActualizedOp" -Arguments (Args2 $rows 43)) "Incomplete op 430 should remain schedulable."
}

function Test-ActualizedOperationSkipPrimitive {
    $dt = New-Routing
    Add-Operation $dt 50 "F" 320 -Actualized $true -ActualEndTime ([DateTime]"2026-07-09T14:00:00")
    Add-Operation $dt 51 "F" 390
    $rows = Invoke-SharedHelper -Name "BuildOperationRowIndex" -Arguments (Args1 $dt)

    Assert-True (Invoke-SharedHelper -Name "IsCompletedOrActualizedOp" -Arguments (Args2 $rows 50)) "Actualized op 320 should be skipped."
    Assert-False (Invoke-SharedHelper -Name "IsCompletedOrActualizedOp" -Arguments (Args2 $rows 51)) "Incomplete op 390 should remain schedulable."
}

function Test-LatestReleaseBoundarySelectsCompleted320 {
    $dt = New-Routing
    Add-Operation $dt 60 "G" 310 -Released $true -ReleaseTime ([DateTime]"2026-07-09T10:00:00") -Completed $true -ParentRecord 60
    Add-Operation $dt 61 "G" 320 -Released $true -ReleaseTime ([DateTime]"2026-07-09T11:00:00") -Completed $true -ParentRecord 60
    Add-Operation $dt 62 "G" 390 -ParentRecord 60

    $boundaries = Get-LatestReleaseBoundaries $dt
    Assert-Equal 1 $boundaries.Count "Latest release boundaries should group by effective parent."
    Assert-Equal 320 ([int]$boundaries[0].OpNo) "Completed 320 should be the latest release boundary."
}

function Test-LatestReleaseBoundaryAllows395ToRelease400 {
    $dt = New-Routing
    Add-Operation $dt 70 "H" 395 -Released $true -ReleaseTime ([DateTime]"2026-07-09T15:00:00") -Completed $true -ParentRecord 70
    Add-Operation $dt 71 "H" 400 -ParentRecord 70
    Add-Operation $dt 72 "H" 500 -ParentRecord 70

    $boundaries = Get-LatestReleaseBoundaries $dt
    Assert-Equal 1 $boundaries.Count "Completed 395 should create one latest release boundary."
    Assert-Equal 395 ([int]$boundaries[0].OpNo) "Completed 395 should seed 400+ scheduling."
}

function Test-LatestReleaseBoundaryAllows490ToRelease500 {
    $dt = New-Routing
    Add-Operation $dt 80 "I" 395 -Released $true -ReleaseTime ([DateTime]"2026-07-09T15:00:00") -Completed $true -ParentRecord 80
    Add-Operation $dt 81 "I" 490 -Released $true -ReleaseTime ([DateTime]"2026-07-09T16:00:00") -Completed $true -ParentRecord 80
    Add-Operation $dt 82 "I" 500 -ParentRecord 80

    $boundaries = Get-LatestReleaseBoundaries $dt
    Assert-Equal 1 $boundaries.Count "Completed 490 should supersede completed 395 in the same order."
    Assert-Equal 490 ([int]$boundaries[0].OpNo) "Completed 490 should seed pending 500."
}

function Test-BoundaryIgnoresUnreleasedRows {
    $dt = New-Routing
    Add-Operation $dt 90 "J" 310 -Released $true -ReleaseTime ([DateTime]"2026-07-09T10:00:00") -Completed $true -ParentRecord 90
    Add-Operation $dt 91 "J" 320 -ParentRecord 90

    $boundaries = Get-LatestReleaseBoundaries $dt
    Assert-Equal 1 $boundaries.Count "Unreleased pending operations must not become boundaries."
    Assert-Equal 310 ([int]$boundaries[0].OpNo) "Boundary should remain at completed 310."
}

function Test-FiringStageConfigStopsAt310 {
    $source = Get-Content -Raw -LiteralPath (Join-Path $root "AlgoSeq4.vb")
    $batchFiring = [regex]::Match($source, "(?s)Public Function runFiring\(.*?^\s*End Function", "Multiline").Value
    $swkFiring = [regex]::Match($source, "(?s)Public Function runSWKFiring\(.*?^\s*End Function", "Multiline").Value
    $tunnelFiring = [regex]::Match($source, "(?s)Public Function runFiring2\(.*?^\s*End Function", "Multiline").Value

    Assert-True ($source -match "Public Function runFiringFollowOn") "Separate 320-395 scheduler entrypoint should exist."
    Assert-False ($batchFiring -match "New Integer\(\) \{ULDBICK, PREINSPC, KILNACK\}") "Batch firing should not directly schedule 320-395 resources."
    Assert-False ($swkFiring -match "New Integer\(\) \{ULDSW, PREINSPC, KILNACK\}") "SWK firing should not directly schedule 320-395 resources."
    Assert-False ($tunnelFiring -match "New Integer\(\) \{ULDPTK, FTDSD20, PREINSPC, KILNACK\}") "Tunnel firing should not directly schedule 320-395 resources."
    Assert-False ($tunnelFiring -match "ScheduleTunnelReleasedContinuations\(preactor,") "Tunnel firing should not call the old direct continuation scheduler."
}

function Test-TunnelContinuationUsesOrderLevelKilnType {
    $dt = New-Routing
    Add-Operation $dt 100 "T1" 300 -Released $true -ReleaseTime ([DateTime]"2026-07-09T08:00:00") -Completed $true -KilnType 2 -ParentRecord 100
    Add-Operation $dt 101 "T1" 310 -Released $true -ReleaseTime ([DateTime]"2026-07-09T10:00:00") -Completed $true -ParentRecord 100
    Add-Operation $dt 102 "T1" 320 -ParentRecord 100

    $candidates = Invoke-SharedHelper -Name "BuildTunnelReleasedContinuations" -Arguments (Args1 $dt)
    Assert-Equal 1 $candidates.Count "Tunnel continuation should select one latest released operation."
    Assert-Equal 310 ([int]$candidates[0].OpNo) "Completed 310 should be selected even when only 300 has Kiln Type = 2."
    Assert-Equal 1 ([int]$candidates[0].StartStageIndex) "Completed 310 should resume at the 320 stage."
}

function Test-TunnelContinuationSkipsCompleted310And320To390Stage {
    $dt = New-Routing
    Add-Operation $dt 110 "T2" 300 -Released $true -ReleaseTime ([DateTime]"2026-07-09T08:00:00") -Completed $true -KilnType 2 -ParentRecord 110
    Add-Operation $dt 111 "T2" 310 -Released $true -ReleaseTime ([DateTime]"2026-07-09T10:00:00") -Completed $true -ParentRecord 110
    Add-Operation $dt 112 "T2" 320 -Released $true -ReleaseTime ([DateTime]"2026-07-09T11:00:00") -Completed $true -ParentRecord 110
    Add-Operation $dt 113 "T2" 390 -ParentRecord 110

    $candidates = Invoke-SharedHelper -Name "BuildTunnelReleasedContinuations" -Arguments (Args1 $dt)
    Assert-Equal 1 $candidates.Count "Tunnel continuation should keep the latest released stage per order."
    Assert-Equal 320 ([int]$candidates[0].OpNo) "Completed 320 should supersede completed 310."
    Assert-Equal 2 ([int]$candidates[0].StartStageIndex) "Completed 320 should resume at the 390 stage."
}

function Test-TunnelContinuationAllowsMissing320 {
    $dt = New-Routing
    Add-Operation $dt 120 "T3" 300 -Released $true -ReleaseTime ([DateTime]"2026-07-09T08:00:00") -Completed $true -KilnType 2 -ParentRecord 120
    Add-Operation $dt 121 "T3" 310 -Released $true -ReleaseTime ([DateTime]"2026-07-09T10:00:00") -Completed $true -ParentRecord 120
    Add-Operation $dt 122 "T3" 390 -ParentRecord 120

    $candidates = Invoke-SharedHelper -Name "BuildTunnelReleasedContinuations" -Arguments (Args1 $dt)
    Assert-Equal 1 $candidates.Count "Missing 320 should not block selecting completed 310 as the release point."
    Assert-Equal 310 ([int]$candidates[0].OpNo) "Completed 310 remains the release point when 320 does not exist."
    Assert-Equal 1 ([int]$candidates[0].StartStageIndex) "Stage walker should start at optional 320 and advance to 390 when 320 is absent."
}

function Test-TunnelContinuationRejectsNonTunnelOrder {
    $dt = New-Routing
    Add-Operation $dt 130 "N1" 310 -Released $true -ReleaseTime ([DateTime]"2026-07-09T10:00:00") -Completed $true -ParentRecord 130
    Add-Operation $dt 131 "N1" 320 -ParentRecord 130

    $candidates = Invoke-SharedHelper -Name "BuildTunnelReleasedContinuations" -Arguments (Args1 $dt)
    Assert-Equal 0 $candidates.Count "Non-tunnel orders must not enter tunnel continuation."
}

function Test-MetricsOnlyCollectorAndExport {
    $originalMetricsSetting = $env:GN_SCHEDULER_METRICS
    $originalDebugSetting = $env:GN_SCHEDULER_DEBUG
    $testFolder = Join-Path ([IO.Path]::GetTempPath()) ("gn_scheduler_metrics_" + [Guid]::NewGuid().ToString("N"))

    try {
        $env:GN_SCHEDULER_METRICS = "true"
        $env:GN_SCHEDULER_DEBUG = "false"

        $collectorType = $assembly.GetType("GrindwellNortonDev_VB.SchedulerDebugCollector", $true)
        $metricsType = $assembly.GetType("GrindwellNortonDev_VB.SchedulerActionMetricsRow", $true)
        $collector = [Activator]::CreateInstance($collectorType)
        $metrics = [Activator]::CreateInstance($metricsType)

        Assert-False ([bool]$collector.Enabled) "Metrics-only mode should not enable detailed tracing."
        Assert-True ([bool]$collector.MetricsEnabled) "Metrics-only mode should enable aggregate metrics."

        $metrics.ActionName = "HarnessAction"
        $metrics.StartedAt = [DateTime]"2026-07-29T10:00:00"
        $metrics.FinishedAt = [DateTime]"2026-07-29T10:00:01"
        $metrics.ElapsedMilliseconds = 1000
        $metrics.RecordsScanned = 12
        $collector.AddActionMetrics($metrics)

        Assert-Equal 1 $collector.ActionMetricsRows.Count "Collector should retain one action metrics row."
        Assert-True (-not [string]::IsNullOrWhiteSpace($metrics.RunId)) "Collector should assign the run ID."
        Assert-True ($metrics.RunId -eq $collector.RunId) "Metrics row should use the collector run ID."

        $collector.ExportFolder = $testFolder
        $collector.ExportAll($null)

        Assert-True (Test-Path -LiteralPath (Join-Path $testFolder "00_RunSummary.md")) "Metrics-only export should write the run summary."
        Assert-True (Test-Path -LiteralPath (Join-Path $testFolder "12_ActionMetrics.csv")) "Metrics-only export should write action metrics."
        Assert-False (Test-Path -LiteralPath (Join-Path $testFolder "01_FieldMap.csv")) "Metrics-only export should not write detailed trace files."
    }
    finally {
        $env:GN_SCHEDULER_METRICS = $originalMetricsSetting
        $env:GN_SCHEDULER_DEBUG = $originalDebugSetting

        $resolvedTempRoot = [IO.Path]::GetFullPath([IO.Path]::GetTempPath())
        $resolvedTestFolder = [IO.Path]::GetFullPath($testFolder)
        if ((Test-Path -LiteralPath $resolvedTestFolder) -and $resolvedTestFolder.StartsWith($resolvedTempRoot, [StringComparison]::OrdinalIgnoreCase)) {
            Remove-Item -LiteralPath $resolvedTestFolder -Recurse -Force
        }
    }
}

function Test-RunScopedLookupCacheSourceGuards {
    $cacheSource = Get-Content -Raw -LiteralPath (Join-Path $root "Debugging\SchedulerDebugFieldCache.vb")
    $algoSource = Get-Content -Raw -LiteralPath (Join-Path $root "AlgoSeq4.vb")
    $postFiringSource = Get-Content -Raw -LiteralPath (Join-Path $root "PostFiringScheduler.vb")
    $runCacheSource = [regex]::Match($cacheSource, "(?s)Friend NotInheritable Class SchedulerRunLookupCache.*?End Class").Value

    Assert-True ($cacheSource -match "Friend NotInheritable Class SchedulerRunLookupCache") "Run-scoped lookup cache should remain assembly-internal and non-inheritable."
    Assert-False ($runCacheSource -match "(?m)^\s*Private Shared ReadOnly _") "Run-scoped lookup values must not be stored in shared mutable dictionaries."
    Assert-False ($runCacheSource -match "Private ReadOnly _preactor|Private ReadOnly _planningboard") "Run-scoped lookup cache must not retain Opcenter interface instances."
    Assert-True ($runCacheSource -match "(?s)Friend Sub MarkOperationPlaced\(recordNo As Integer\).*?_scheduledStates\(recordNo\) = True.*?_operationTimes\.Remove\(recordNo\)") "Placement should update scheduled state and invalidate operation times."
    Assert-True ($postFiringSource -match "forceRefresh:=True") "The immediate pre-placement scheduled-state recheck must bypass cached state."
    Assert-True ($postFiringSource -match "Friend Property LookupCache As SchedulerRunLookupCache") "Post-firing queue construction and traversal should share one action-local cache."
    Assert-True ($algoSource -match "lookupCache\.GetOperationTimes\(planningboard,\s*nextOpIndex\)") "runFix should read the approved noncontiguous next operation by nextOpIndex."
}

function Test-OperationRowIndexReuseAndInvalidation {
    $dt = New-Routing
    Add-Operation $dt 200 "INDEX-A" 200

    $first = Invoke-SharedHelper -Name "BuildOperationRowIndex" -Arguments (Args1 $dt)
    $second = Invoke-SharedHelper -Name "BuildOperationRowIndex" -Arguments (Args1 $dt)

    Assert-True ([object]::ReferenceEquals($first, $second)) "Unchanged routing table should reuse its operation-row index."
    Assert-Equal 1 $second.Count "Cached operation-row index should retain the original row."

    Add-Operation $dt 201 "INDEX-A" 260
    $afterRowAdded = Invoke-SharedHelper -Name "BuildOperationRowIndex" -Arguments (Args1 $dt)

    Assert-False ([object]::ReferenceEquals($second, $afterRowAdded)) "Routing row-count changes should rebuild the operation-row index."
    Assert-Equal 2 $afterRowAdded.Count "Rebuilt operation-row index should include the added row."

    $copy = $dt.Copy()
    $copyIndex = Invoke-SharedHelper -Name "BuildOperationRowIndex" -Arguments (Args1 $copy)
    Assert-False ([object]::ReferenceEquals($afterRowAdded, $copyIndex)) "Copied routing tables must not reuse an index containing source-table DataRows."
    Assert-True ([object]::ReferenceEquals($copy, $copyIndex[200].Table)) "Copied routing-table index should point to copied DataRows."
}

function Test-ReadOrderTableSourceGuards {
    $source = Get-Content -Raw -LiteralPath (Join-Path $root "SharedHelpers.vb")
    $snapshot = [regex]::Match($source, "(?s)Private Function ReadOrderTableCore\(.*?^\s*End Function", "Multiline").Value

    Assert-True ($source -match "Public Function readOrderTable\(ByVal preactor As IPreactor\) As DataTable") "The public readOrderTable signature must remain unchanged."
    Assert-True ($snapshot -match "(?s)dt\.BeginLoadData\(\).*?Finally\s+dt\.EndLoadData\(\)") "EndLoadData should execute from a Finally block."
    Assert-True ($snapshot -match 'TryGetFieldNumber\(preactor,\s*ordersTable,\s*"Toggle Attribute 1",\s*cache\)') "Toggle Attribute 1 must remain the completion source."
    Assert-True ($snapshot -match "(?s)Actual End Time.*?Actual End.*?Actual Finish Time.*?Actual End Date.*?End Time Actual") "Actual-end fallback order must remain intact."
    Assert-True ($snapshot -match "(?s)Actual Start Time.*?Actual Start.*?Actual Start Date.*?Start Time Actual") "Actual-start fallback order must remain intact."
    Assert-True ($snapshot -match "scheduledByRecord\(rec\)\s*=\s*cache\.IsOperationScheduled") "Scheduled state should be captured once into the snapshot array."
    Assert-True ($snapshot -match "SetOperationRowIndex\(dt, operationRows\)") "Snapshot construction should publish its already-built operation-row index."
}

function Test-RoutingSchemaCompatibility {
    $dt = New-Routing
    $requiredColumns = @(
        "OrdersID", "Order No", "Operation Number", "Operation Name",
        "Resource Group", "Required Resource", "Quantity", "Due Date",
        "Batch Time", "Tonnage", "Cycle Type", "Volume Occupancy",
        "Kiln Type", "MTS/MTO", "Wheel Dia", "Wheel thickness",
        "Pressing earliest start", "Pressing Due date", "firing due date",
        "scheduled_start_time", "scheduled_end_time", "is_scheduled",
        "parent_record", "prev_op_is_scheduled", "source_is_completed",
        "opcenter_use_actual", "actual_start_time", "actual_end_time",
        "operation_effective_completed", "operation_releases_next",
        "operation_release_time", "wip_ready_time", "wip_score",
        "wip_status", "wip_reject_reason"
    )

    foreach ($columnName in $requiredColumns) {
        Assert-True $dt.Columns.Contains($columnName) "Routing schema should retain required column '$columnName'."
    }

    foreach ($column in $dt.Columns) {
        Assert-True ($column.DataType -eq [object]) "Batch 4 must not introduce typed routing columns. Column=$($column.ColumnName)."
    }
}

function Invoke-AlgoSeq4PrivateMethod {
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [AllowEmptyCollection()][object[]]$Arguments = @()
    )

    $algoType = $assembly.GetType("GrindwellNortonDev_VB.AlgoSeq4", $true)
    $method = $algoType.GetMethod(
        $Name,
        [Reflection.BindingFlags]::Instance -bor [Reflection.BindingFlags]::NonPublic
    )
    if ($null -eq $method) {
        throw "Could not resolve private AlgoSeq4 method '$Name'."
    }

    $instance = [Activator]::CreateInstance($algoType)
    return ,($method.Invoke($instance, $Arguments))
}

function Test-FiringFollowOnKilnTypeLookup {
    $dt = New-Routing
    Add-Operation $dt 300 "K-A" 200 -KilnType 1 -ParentRecord 300
    Add-Operation $dt 301 "K-A" 300 -KilnType 2 -ParentRecord 300
    Add-Operation $dt 302 "K-A" 310 -ParentRecord 300

    Add-Operation $dt 400 "K-B" 200 -KilnType 3 -ParentRecord 400
    Add-Operation $dt 401 "K-B" 310 -KilnType 2 -ParentRecord 400

    $lookup = Invoke-AlgoSeq4PrivateMethod -Name "BuildFiringFollowOnKilnTypeLookup" -Arguments (Args1 $dt)

    Assert-Equal 2 $lookup.Count "Kiln-type lookup should contain one entry per parent group."
    Assert-Equal 2 ([int]$lookup["P:300"]) "Operation 300 kiln type should override an earlier positive fallback."
    Assert-Equal 3 ([int]$lookup["P:400"]) "Orders without operation 300 should retain the first positive kiln type."
}

function Test-FiringFollowOnOptimizationSourceGuards {
    $source = Get-Content -Raw -LiteralPath (Join-Path $root "AlgoSeq4.vb")
    $action = [regex]::Match($source, "(?s)Public Function runFiringFollowOn\(.*?^\s*End Function", "Multiline").Value

    Assert-True ($action -match "(?s)OrderByDescending\(Function\(x\) x\.WipScore\).*?ThenBy\(Function\(x\) x\.ReleaseTime\).*?ThenBy\(Function\(x\) x\.ParentRecord\).*?ThenBy\(Function\(x\) x\.OpNo\)") "Follow-on boundary ranking order must remain WIP, release time, parent, operation number."
    Assert-Equal 1 ([regex]::Matches($action, "BuildOperationRowIndex\(").Count) "Follow-on action should obtain the operation-row index once."
    Assert-Equal 1 ([regex]::Matches($action, 'GetResourceNumber\(planningboard, "PREINSPC"\)').Count) "PREINSPC should resolve once."
    Assert-Equal 1 ([regex]::Matches($action, 'GetResourceNumber\(planningboard, "KILNACK"\)').Count) "KILNACK should resolve once."
    Assert-Equal 1 ([regex]::Matches($action, 'GetResourceNumber\(planningboard, "FTDSD20"\)').Count) "FTDSD20 should resolve once."
    Assert-False ($action -match "GetBoundaryKilnType") "Follow-on action should not scan routingDt for every boundary."
    Assert-False ($action -match "Enumerable\.Repeat") "Follow-on stage-name arrays should be allocated once per action."
    Assert-True ($action.IndexOf("batchStageResources", [StringComparison]::Ordinal) -lt $action.IndexOf("For Each boundary", [StringComparison]::Ordinal)) "Batch stage arrays should be created outside the boundary loop."
    Assert-True ($action.IndexOf("tunnelStageResources", [StringComparison]::Ordinal) -lt $action.IndexOf("For Each boundary", [StringComparison]::Ordinal)) "Tunnel stage arrays should be created outside the boundary loop."
    Assert-True ($action.IndexOf("swkStageResources", [StringComparison]::Ordinal) -lt $action.IndexOf("For Each boundary", [StringComparison]::Ordinal)) "SWK stage arrays should be created outside the boundary loop."
    Assert-True ($action -match "minOpNoInclusive:=FOLLOWON_START_OP_NO") "Follow-on lower operation bound should remain 320."
    Assert-True ($action -match "maxOpNoExclusive:=FOLLOWON_END_EXCLUSIVE_OP_NO") "Follow-on upper operation bound should remain exclusive 400."
    Assert-False ($source -match "_schedulerDebug\.OperationSnapshots\.FirstOrDefault") "Debug snapshot lookup should use the collector index."
}

function Test-RunFixOptimizationSourceGuards {
    $source = Get-Content -Raw -LiteralPath (Join-Path $root "AlgoSeq4.vb")
    $action = [regex]::Match($source, "(?s)Public Function runFix\(.*?^\s*End Function", "Multiline").Value

    Assert-True ($action -match "For i As Integer = 1 To recordCount") "runFix should retain one ascending record-order scan."
    Assert-True ($action -match "(?s)Dim recOpNo As Integer =\s*lookupCache\.ReadOperationNumber\(preactor,\s*ordersTable,\s*opNoField,\s*i\)") "runFix should read the scanned record's operation number once."
    Assert-True ($action -match "If recOpNo = 260 Then") "runFix should retain operation 260 handling."
    Assert-True ($action -match "ElseIf recOpNo = 200 Then") "runFix should retain operation 200 handling."
    Assert-False ($action -match "GetOperationTimes\(i\s*\+\s*1\)") "runFix must not assume the next operation is the next physical record."
    Assert-True ($action -match "(?s)GetOperationTimes\(planningboard,\s*nextOpIndex\)") "runFix should obtain operation 290 timing using nextOpIndex."
    Assert-Equal 1 ([regex]::Matches($action, "BackTestOpOnResource\(").Count) "runFix should retain exactly one backward feasibility API call site."
    Assert-True ($action -match "(?s)BackTestOpOnResource\(i,\s*dryer,\s*nextOpStartTime\)") "Operation 260 should retain BackTestOpOnResource with the dryer and next-operation start."
    Assert-False ($action -match "(?<!Back)TestOperationOnResource\(") "runFix should not add forward feasibility tests."
    Assert-True ($action -match "Dim targetTime As DateTime = operationStart\.AddDays\(-1\)") "Predecessor operations should retain the shared one-day-earlier target time."
    Assert-True ($action -match "resources\.FirstOrDefault\(\)") "Predecessor placement should retain first-resource selection."
    Assert-True ($action -match "firstResourceByOperation\.TryGetValue") "Predecessor first-resource results should be cached within the action."
    Assert-True ($action -match "(?s)For i As Integer = 1 To recordCount\s*Try.*?Catch ex As Exception.*?Next") "Per-record exceptions should not stop the remaining scan."
    Assert-True ($action -match "(?s)Finally.*?CompleteActionMetrics.*?FinishSchedulerDebug.*?preactor\.DestroyStatus\(\)") "runFix cleanup should remain in Finally."
}

function Test-Schedule400OptimizationSourceGuards {
    $source = Get-Content -Raw -LiteralPath (Join-Path $root "AlgoSeq4.vb")
    $action = [regex]::Match($source, "(?s)Public Function Schedule400\(.*?^\s*End Function", "Multiline").Value

    Assert-True (-not [string]::IsNullOrWhiteSpace($action)) "Schedule400 action source should be available for regression guards."
    Assert-True ($action -match "For opRec As Integer = 1 To recordCount") "Schedule400 should retain one ascending record-order scan."
    Assert-Equal 1 ([regex]::Matches($action, "planningboard\.TerminatorTime").Count) "Schedule400 should read terminator time once."
    Assert-Equal 1 ([regex]::Matches($action, "preactor\.ReadFieldInt\(").Count) "Schedule400 should read each scanned operation number directly once."
    Assert-False ($action -match "ReadOperationNumber\(") "Schedule400 should not retain one-use operation numbers in the run cache."
    Assert-True ($action.IndexOf("If opNoValue < POSTFIRING_START_OP_NO Then Continue For", [StringComparison]::Ordinal) -lt $action.IndexOf("planningboard.FindResources(opRec)", [StringComparison]::Ordinal)) "Records below operation 400 must be skipped before resource lookup."
    Assert-Equal 1 ([regex]::Matches($action, "planningboard\.FindResources\(opRec\)").Count) "Schedule400 should call FindResources once per qualifying record."
    Assert-Equal 1 ([regex]::Matches($action, "For Each resRec As Integer In resources").Count) "Schedule400 should enumerate alternate resources once."
    Assert-Equal 1 ([regex]::Matches($action, "planningboard\.TestOperationOnResource\(").Count) "Schedule400 should keep one feasibility call site inside resource enumeration."
    Assert-True ($action -match "(?s)If Not hasBestChangeStart OrElse\s*changeStart < bestChangeStart Then") "Schedule400 should retain strict earliest-ChangeStart selection and first-resource tie behavior."
    Assert-True ($action -match "(?s)PutOperationOnResource\(opRec,\s*bestResRec,\s*bestChangeStart\)") "Schedule400 should place the selected record and resource at the winning ChangeStart."
    Assert-False ($action -match "MarkOperationPlaced") "Schedule400 should not update an invocation cache it never reads after placement."
    Assert-False ($action -match "afterFiring|PostFiringScheduler|WipScore|ReleaseBoundary") "Schedule400 must remain independent of afterFiring ranking and boundary logic."
    Assert-True ($action -match "(?s)Finally.*?CompleteActionMetrics.*?FinishSchedulerDebug.*?preactor\.DestroyStatus\(\)") "Schedule400 cleanup should remain in Finally."
}

function Test-PostFiringOptimizationSourceGuards {
    $algoSource = Get-Content -Raw -LiteralPath (Join-Path $root "AlgoSeq4.vb")
    $schedulerSource = Get-Content -Raw -LiteralPath (Join-Path $root "PostFiringScheduler.vb")
    $action = [regex]::Match($algoSource, "(?s)Public Function afterFiring\(.*?^\s*End Function", "Multiline").Value
    $buildQueue = [regex]::Match($schedulerSource, "(?s)Public Function BuildQueue\(.*?^\s*End Function", "Multiline").Value
    $scheduleQueue = [regex]::Match($schedulerSource, "(?s)Public Function ScheduleQueue\(.*?^\s*End Function", "Multiline").Value
    $findNext = [regex]::Match($schedulerSource, "(?s)Private Function FindNextPostFiringCandidate\(.*?^\s*End Function", "Multiline").Value

    Assert-Equal 1 ([regex]::Matches($action, "ReadOrderTableWithCache\(").Count) "afterFiring should build routing data once."
    Assert-Equal 1 ([regex]::Matches($action, "postFiring\.BuildQueue\(").Count) "afterFiring should build one post-firing queue."
    Assert-Equal 1 ([regex]::Matches($action, "postFiring\.ScheduleQueue\(").Count) "afterFiring should schedule the ranked queue once."
    Assert-False ($action -match "Schedule400") "afterFiring must remain independent from Schedule400."
    Assert-True ($action -match "(?s)Finally.*?CompleteActionMetrics.*?FinishSchedulerDebug.*?preactor\.DestroyStatus\(\)") "afterFiring cleanup should remain in Finally."

    Assert-Equal 1 ([regex]::Matches($buildQueue, "BuildOperationRowIndex\(").Count) "Post-firing queue construction should obtain the operation-row index once."
    Assert-Equal 1 ([regex]::Matches($buildQueue, "BuildLatestReleaseBoundaries\(").Count) "Post-firing queue construction should build release boundaries once."
    Assert-True ($buildQueue -match "New List\(Of QueueItem\)\(boundaries\.Count\)") "Post-firing queue capacity should use the known boundary count."
    Assert-False ($buildQueue -match "cache\.ReadOperationNumber\(") "Queue construction should reuse the operation number returned by boundary traversal."
    Assert-True ($findNext -match "ByRef selectedOpNo As Integer") "Boundary traversal should return its selected operation number without a duplicate lookup."
    Assert-True ($findNext -match "(?s)selectedOpNo = nextOpNo\s*Return nextOpRec") "Boundary traversal should return the live operation number paired with its selected record."
    Assert-True ($buildQueue -match "(?s)OrderByDescending\(Function\(x\) x\.WipScore\).*?ThenBy\(Function\(x\) x\.KilnAckEndTime\).*?ThenBy\(Function\(x\) x\.DueDate\).*?ThenBy\(Function\(x\) x\.Priority\).*?ThenBy\(Function\(x\) x\.ParentRecord\).*?ThenBy\(Function\(x\) x\.NextOpNo\)") "Post-firing ranking must remain WIP, release time, due date, priority, parent, operation number."
    Assert-True ($buildQueue -match "ReadDueDate\(preactor, ordersTable, dueDateField, nextOpRec\)") "Queue ranking should retain the live due-date read."
    Assert-True ($buildQueue -match "ReadPriority\(preactor, ordersTable, priorityField, nextOpRec\)") "Queue ranking should retain the live priority read."
    Assert-True ($buildQueue -match "\.OperationRows = rowByOpRec") "Every queue item should reuse the single operation-row dictionary."

    Assert-True ($scheduleQueue -match "(?s)For Each item As QueueItem In queue\s*Try.*?Catch ex As Exception.*?Next") "Post-firing failures should remain isolated per queued order."
    Assert-True ($scheduleQueue -match "cache\.IsOperationScheduled\(planningboard, opRec\)") "Post-firing traversal should retain already-scheduled skipping."
    Assert-True ($scheduleQueue -match "SharedHelpers\.IsCompletedOrActualizedOp\(item\.OperationRows, opRec\)") "Post-firing traversal should use the shared operation-row index for completed operations."
    Assert-True ($scheduleQueue -match "SharedHelpers\.GetOperationReleaseTime\(item\.OperationRows, opRec\)") "Completed-operation traversal should retain indexed release-time lookup."
    Assert-Equal 1 ([regex]::Matches($scheduleQueue, "planningboard\.FindResources\(opRec\)").Count) "Post-firing scheduling should find resources once per pending operation."
    Assert-Equal 1 ([regex]::Matches($scheduleQueue, "For Each resRec As Integer In resources").Count) "Post-firing scheduling should enumerate alternate resources once."
    Assert-True ($scheduleQueue -match "(?s)If Not hasBestTimes OrElse\s*changeStart < bestChangeStart Then") "Post-firing scheduling should retain strict earliest-ChangeStart selection and first-resource ties."
    Assert-True ($scheduleQueue -match "(?s)PutOperationOnResource\(opRec,\s*bestResRec,\s*bestChangeStart\)") "Post-firing placement should use the selected resource and ChangeStart."
    Assert-True ($scheduleQueue -match "forceRefresh:=True") "Post-firing placement should retain the immediate live scheduled-state recheck."
    Assert-True ($scheduleQueue -match "cache\.GetNextOperation\(planningboard, opRec, 1\)") "Post-firing routing traversal should retain action-local next-operation caching."
    Assert-True ($scheduleQueue -match "Dim debugEnabled As Boolean") "Disabled detailed debugging should avoid row-level trace helper work."

    $requiredReasonCodes = @(
        "POSTFIRING_NO_RELEASE_BOUNDARY",
        "POSTFIRING_NO_NEXT_OPERATION",
        "DATA_MISSING_OPERATION_NUMBER",
        "POSTFIRING_NEXT_ALREADY_SCHEDULED",
        "POSTFIRING_NEXT_COMPLETED",
        "POSTFIRING_NEXT_BELOW_RANGE_BLOCKED",
        "POSTFIRING_NO_FEASIBLE_RESOURCE",
        "SCHEDULE_EXCEPTION_THROWN",
        "SCHEDULE_RESULT_NOT_SCHEDULED"
    )
    foreach ($reasonCode in $requiredReasonCodes) {
        Assert-True ($schedulerSource -match [regex]::Escape($reasonCode)) "Post-firing reason code '$reasonCode' should remain available."
    }
}

function Test-OptimizerSettingsCatalog {
    $catalogType = $assembly.GetType(
        "GrindwellNortonDev_VB.SharedHelpers+OptimizerSettingsCatalog",
        $true
    )
    $getDefinitions = $catalogType.GetMethod(
        "GetPointerDefinitions",
        [Reflection.BindingFlags]::Static -bor [Reflection.BindingFlags]::NonPublic
    )
    Assert-True ($null -ne $getDefinitions) "Optimizer settings catalog should expose its assembly-internal pointer definitions."

    $definitions = @($getDefinitions.Invoke($null, @()))
    Assert-Equal 20 $definitions.Count "Optimizer settings catalog should inventory every current live pointer and hardcoded optimizer input."

    $definitionType = $definitions[0].GetType()
    $propertyFlags = [Reflection.BindingFlags]::Instance -bor [Reflection.BindingFlags]::NonPublic
    $keyProperty = $definitionType.GetProperty("Key", $propertyFlags)
    $displayNameProperty = $definitionType.GetProperty("DisplayName", $propertyFlags)
    $recordProperty = $definitionType.GetProperty("RecordNumber", $propertyFlags)
    $fieldTypeProperty = $definitionType.GetProperty("FieldType", $propertyFlags)
    $unitProperty = $definitionType.GetProperty("UnitAndConversion", $propertyFlags)
    $rangeProperty = $definitionType.GetProperty("ValidRange", $propertyFlags)
    $fallbackProperty = $definitionType.GetProperty("FallbackBehavior", $propertyFlags)
    $consumersProperty = $definitionType.GetProperty("ConsumingActions", $propertyFlags)
    $effectiveProperty = $definitionType.GetProperty("CurrentEffectiveValue", $propertyFlags)

    $byKey = @{}
    foreach ($definition in $definitions) {
        $key = [string]$keyProperty.GetValue($definition)
        $byKey[$key] = $definition
        Assert-True (-not [string]::IsNullOrWhiteSpace([string]$displayNameProperty.GetValue($definition))) "Catalog entry '$key' should include a display-name placeholder."
        Assert-True (-not [string]::IsNullOrWhiteSpace([string]$fieldTypeProperty.GetValue($definition))) "Catalog entry '$key' should include its field type."
        Assert-True (-not [string]::IsNullOrWhiteSpace([string]$unitProperty.GetValue($definition))) "Catalog entry '$key' should include units and conversion."
        Assert-True (-not [string]::IsNullOrWhiteSpace([string]$rangeProperty.GetValue($definition))) "Catalog entry '$key' should include a valid range."
        Assert-True (-not [string]::IsNullOrWhiteSpace([string]$fallbackProperty.GetValue($definition))) "Catalog entry '$key' should include fallback behavior."
        Assert-True (-not [string]::IsNullOrWhiteSpace([string]$consumersProperty.GetValue($definition))) "Catalog entry '$key' should include consumers."
        Assert-True (-not [string]::IsNullOrWhiteSpace([string]$effectiveProperty.GetValue($definition))) "Catalog entry '$key' should include its current effective value."
    }

    $expectedRecords = @{
        BatchMinOccupancy = 1
        BatchMaxOccupancy = 2
        BatchAllowUnderfilledTailConfigured = 3
        BatchStartDelayMinutesConfigured = 4
        TunnelCartsPerDay = 5
        TunnelTotalCarts = 7
        TunnelMinOccupancyPreferred = 8
        TunnelDryingToFiringBuffer = 8
        TunnelMaxOccupancy = 9
        BatchMaxBatchesPerDay = 0
        ResourceAvailableFrom = 0
        SwkMinTonnage = 0
        PressApproachingDays = 0
    }
    foreach ($entry in $expectedRecords.GetEnumerator()) {
        Assert-True $byKey.ContainsKey($entry.Key) "Catalog should contain '$($entry.Key)'."
        if ($byKey.ContainsKey($entry.Key)) {
            Assert-Equal ([int]$entry.Value) ([int]$recordProperty.GetValue($byKey[$entry.Key])) "Catalog record number should preserve '$($entry.Key)'."
        }
    }
}

function Test-OptimizerSettingsCatalogSourceGuards {
    $algoSource = Get-Content -Raw -LiteralPath (Join-Path $root "AlgoSeq4.vb")
    $helpersSource = Get-Content -Raw -LiteralPath (Join-Path $root "SharedHelpers.vb")
    $runFiring = [regex]::Match($algoSource, "(?s)Public Function runFiring\(.*?^\s*End Function", "Multiline").Value
    $runSwk = [regex]::Match($algoSource, "(?s)Public Function runSWKFiring\(.*?^\s*End Function", "Multiline").Value
    $runTunnel = [regex]::Match($algoSource, "(?s)Public Function runFiring2\(.*?^\s*End Function", "Multiline").Value
    $runPress = [regex]::Match($algoSource, "(?s)Public Function runPress\(.*?^\s*End Function", "Multiline").Value

    Assert-True ($helpersSource -match "Friend NotInheritable Class OptimizerSettingsCatalog") "Optimizer pointer catalog should remain assembly-internal and centralized in SharedHelpers."
    Assert-True ($helpersSource -match "USER REPLACEMENT INSTRUCTIONS") "Catalog should tell the user exactly how to replace pointers."
    Assert-False ($algoSource -match 'GetFormatNumber\("GN Optimizer Settings"\)') "AlgoSeq4 should not scatter the optimizer settings format name."
    Assert-False ($algoSource -match '(allowUnderfilledTail|batchStartDelayMins|maxBatchesPerDayGlobal|dailyBatchLimit|swkResourceName|approachingDays|prioritizePrevOpFirst):=\s*(True|False|[-0-9.]+|"[^"]+")') "Optimizer calls should not retain scattered hardcoded named arguments."

    Assert-True ($runFiring -match "BatchMinOccupancyRecordNumber") "Batch minimum occupancy should use its centralized record pointer."
    Assert-True ($runFiring -match "BatchMaxOccupancyRecordNumber") "Batch maximum occupancy should use its centralized record pointer."
    Assert-True ($runFiring -match "BatchAllowUnderfilledTailRecordNumber") "The existing batch toggle read should remain."
    Assert-True ($runFiring -match "BatchStartDelayMinutesRecordNumber") "The existing batch delay read should remain."
    Assert-True ($runFiring -match "BatchEffectiveAllowUnderfilledTail") "Batch optimizer should retain effective True through the catalog."
    Assert-True ($runFiring -match "BatchEffectiveStartDelayMinutes") "Batch optimizer should retain effective 60 minutes through the catalog."
    Assert-True ($runFiring -match "BatchEffectiveMaxBatchesPerDay") "Batch optimizer should retain effective two batches per day through the catalog."

    Assert-True ($runSwk -match "SwkEffectiveMinTonnage") "SWK minimum tonnage should use the centralized effective value."
    Assert-True ($runSwk -match "SwkEffectiveMaxTonnage") "SWK maximum tonnage should use the centralized effective value."
    Assert-True ($runSwk -match "SwkEffectiveDailyBatchLimit") "SWK daily limit should use the centralized effective value."
    Assert-True ($runSwk -match "SwkEffectiveStartDelayMinutes") "SWK delay should use the centralized effective value."
    Assert-True ($runSwk -match "SwkEffectiveAllowUnderfilledTail") "SWK tail behavior should use the centralized effective value."

    Assert-True ($runTunnel -match "TunnelCartsPerDayRecordNumber") "Tunnel carts per day should use its centralized record pointer."
    Assert-True ($runTunnel -match "TunnelTotalCartsRecordNumber") "Tunnel total carts should use its centralized record pointer."
    Assert-True ($runTunnel -match "TunnelMinOccupancyPreferredRecordNumber") "Tunnel minimum occupancy should use record 8 through the catalog."
    Assert-True ($runTunnel -match "TunnelDryingBufferMinutesRecordNumber") "Tunnel buffer should retain its separate record-8 pointer."
    Assert-True ($runTunnel -match "TunnelMaxOccupancyRecordNumber") "Tunnel maximum occupancy should use its centralized record pointer."
    Assert-True ($runTunnel -match "(?s)TunnelDryingBufferMinutesRecordNumber\).*?/\s*SharedHelpers\.OptimizerSettingsCatalog\.MinutesPerHour") "Tunnel buffer should retain minutes-to-hours division."

    Assert-True ($runPress -match "PressEffectiveApproachingDays") "Press approaching days should retain effective two days through the catalog."
    Assert-True ($runPress -match "PressEffectivePrioritizePreviousOperation") "Press previous-operation priority should retain effective True through the catalog."
}

function Test-Batch12HardeningSourceGuards {
    $algo2Source = Get-Content -Raw -LiteralPath (Join-Path $root "AlgoSeq2.vb")
    $loginSource = Get-Content -Raw -LiteralPath (Join-Path $root "LoginForm.vb")
    $kilnSource = Get-Content -Raw -LiteralPath (Join-Path $root "KilnAvailabilityWindow.vb")
    $dbSource = Get-Content -Raw -LiteralPath (Join-Path $root "DBConnection.vb")
    $sessionSource = Get-Content -Raw -LiteralPath (Join-Path $root "DbSession.vb")
    $repositorySource = Get-Content -Raw -LiteralPath (Join-Path $root "MetadataExplorerRepository.vb")
    $explorerSource = Get-Content -Raw -LiteralPath (Join-Path $root "DataViewExplorerForm.vb")
    $routingReaderSource = Get-Content -Raw -LiteralPath (Join-Path $root "Class[newlogic].vb")
    $algo3Source = Get-Content -Raw -LiteralPath (Join-Path $root "AlgoSeq3.vb")
    $csvActionSource = Get-Content -Raw -LiteralPath (Join-Path $root "CustomAction[csv].vb")

    Assert-False ($algo2Source -match "Dim (dueDateField|priorityField|SequenceMode|ResRec|ResRecs)\b") "The five warning-only AlgoSeq2 locals should stay removed."
    Assert-True ($loginSource -match "System\.Windows\.Forms\.DialogResult\.OK") "Login DialogResult should remain explicitly enum-qualified."

    $constructor = [regex]::Match($kilnSource, "(?s)Public Sub New\(\).*?End Sub").Value
    $onOpen = [regex]::Match($kilnSource, "(?s)Public Function OnOpen\(.*?End Function").Value
    Assert-True ($constructor -match "InitializeComponent\(\)") "Kiln availability construction should initialize designer components."
    Assert-Equal 1 ([regex]::Matches($constructor, "BuildTableUi\(\)").Count) "Kiln availability UI should be built once by the constructor."
    Assert-False ($onOpen -match "BuildTableUi\(\)") "Kiln OnOpen should not duplicate UI construction."

    Assert-True ($dbSource -match "(?s)Using connection As New SqlConnection.*?Using command As New SqlCommand.*?Using reader As SqlDataReader") "Calendar-state database objects should be disposed through nested Using blocks."
    Assert-True ($sessionSource -match "New SqlConnectionStringBuilder") "Metadata connection strings should be built with SqlConnectionStringBuilder."
    Assert-False ($sessionSource -match 'ConnectionString\s*=\s*baseConn\s*&') "Metadata passwords should not be concatenated into connection strings."

    Assert-True ($repositorySource -match "ValidateMetadataTable") "Metadata table loads should validate the selected identifier."
    Assert-True ($repositorySource -match "@schemaName") "Metadata schema validation should be parameterized."
    Assert-True ($repositorySource -match "@tableName") "Metadata table validation should be parameterized."
    Assert-True ($repositorySource -match "Using builder As New SqlCommandBuilder") "Metadata command builders should be deterministically disposed."
    Assert-True ($explorerSource -match "_allowedTables") "The metadata explorer should retain the loaded identifier allow-list."
    Assert-True ($explorerSource -match "Handles MyBase\.FormClosed") "The metadata explorer should clean up database objects when closed."
    Assert-True ($explorerSource -match "_currentAdapter\.Dispose\(\)") "The metadata explorer should dispose its current adapter."

    $pathSources = $routingReaderSource + $algo3Source + $csvActionSource
    Assert-False ($pathSources -match '[A-Za-z]:\\') "CSV source files should not contain machine-specific absolute Windows paths."
    Assert-False ($pathSources -match "/mnt/data/") "CSV source files should not contain sandbox-specific fixed paths."
    Assert-True ($routingReaderSource -match 'AppSettings\("RoutingCsvPath"\)') "The standalone routing reader should use the configured routing path."
    Assert-True ($algo3Source -match 'ParseShellString\("\{PATH\}"\)') "The legacy AlgoSeq3 action should resolve relative routing paths through Opcenter."
    Assert-True ($csvActionSource -match "VitrifiedCyclePlannerBaseFolder") "The CSV planner base folder should be configurable."
}

Test-DirectPressingToLoadingReadiness
Test-DryingCompletedReadiness
Test-LoadingCompletedReadiness
Test-CompletedPostFiringSkipPrimitive
Test-CompletedFourHundredOnwardSkipPrimitive
Test-ActualizedOperationSkipPrimitive
Test-LatestReleaseBoundarySelectsCompleted320
Test-LatestReleaseBoundaryAllows395ToRelease400
Test-LatestReleaseBoundaryAllows490ToRelease500
Test-BoundaryIgnoresUnreleasedRows
Test-FiringStageConfigStopsAt310
Test-TunnelContinuationUsesOrderLevelKilnType
Test-TunnelContinuationSkipsCompleted310And320To390Stage
Test-TunnelContinuationAllowsMissing320
Test-TunnelContinuationRejectsNonTunnelOrder
Test-MetricsOnlyCollectorAndExport
Test-RunScopedLookupCacheSourceGuards
Test-OperationRowIndexReuseAndInvalidation
Test-ReadOrderTableSourceGuards
Test-RoutingSchemaCompatibility
Test-FiringFollowOnKilnTypeLookup
Test-FiringFollowOnOptimizationSourceGuards
Test-RunFixOptimizationSourceGuards
Test-Schedule400OptimizationSourceGuards
Test-PostFiringOptimizationSourceGuards
Test-OptimizerSettingsCatalog
Test-OptimizerSettingsCatalogSourceGuards
Test-Batch12HardeningSourceGuards

if ($failures -gt 0) {
    throw "$failures SharedHelpers readiness harness assertion(s) failed."
}

Write-Host "SharedHelpers readiness harness passed."
