param(
    [string]$BaselineAssembly = 'bin/OptimizationBaseline/GrindwellNortonDev_VB.dll',
    [string]$OptimizedAssembly = 'Tests/obj/DebuggerMatrix/AllOn/bin/GrindwellNortonDev_VB.dll'
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
Get-ChildItem -LiteralPath (Join-Path $root 'bin/Debug') -Filter 'Preactor*.dll' | ForEach-Object {
    [Reflection.Assembly]::LoadFrom($_.FullName) | Out-Null
}
$baseline = [Reflection.Assembly]::LoadFile((Resolve-Path -LiteralPath (Join-Path $root $BaselineAssembly)))
$optimized = [Reflection.Assembly]::LoadFile((Resolve-Path -LiteralPath (Join-Path $root $OptimizedAssembly)))
$prefix = 'GrindwellNortonDev_VB.'
Add-Type -ReferencedAssemblies System.Data -TypeDefinition @'
using System;
using System.Data;
using System.Collections.Generic;
public class ExportFixture {
    public static readonly List<string> Reads = new List<string>();
    public static bool ThrowOnSecond;
    public string First { get { Reads.Add("First"); return "comma,quote\"\r\nline"; } }
    public object Second { get { Reads.Add("Second"); if (ThrowOnSecond) throw new InvalidOperationException("fixture"); return null; } }
    public DateTime When { get { Reads.Add("When"); return new DateTime(2026,11,3,8,9,10); } }
}
public class MutatingCell {
    public DataRow Row;
    public override string ToString() { Row[1] = "changed"; return "first"; }
}
public static class AllocationMeter {
    static AllocationMeter() { AppDomain.MonitoringIsEnabled = true; }
    public static long Read() { return AppDomain.CurrentDomain.MonitoringTotalAllocatedMemorySize; }
}
'@
function Get-PublicSurface($assembly) {
    @($assembly.GetExportedTypes() | ForEach-Object {
        $type = $_
        'TYPE ' + $type.FullName
        $type.GetCustomAttributesData() | ForEach-Object { $type.FullName + ' ATTRIBUTE ' + $_.ToString() }
        $type.GetMembers([Reflection.BindingFlags]'Public,Static,Instance,DeclaredOnly') | ForEach-Object {
            $member = $_
            $type.FullName + ' ' + $member.MemberType + ' ' + $member.ToString()
        }
    } | Sort-Object) -join "`n"
}
if ((Get-PublicSurface $baseline) -cne (Get-PublicSurface $optimized)) { throw 'Public API changed.' }
$folder = Join-Path ([IO.Path]::GetTempPath()) ('gn_optimization_' + [Guid]::NewGuid().ToString('N'))
$previous = [Environment]::CurrentDirectory
[IO.Directory]::CreateDirectory($folder) | Out-Null
function Export-Csv($assembly, $path, $rows) {
    $method = $assembly.GetType($prefix + 'SchedulerDebugExporter').GetMethod('WriteCsv').MakeGenericMethod([ExportFixture])
    $method.Invoke($null, @([string]$path, $rows.PSObject.BaseObject)) | Out-Null
}
function New-DiagnosticFixture($assembly, [int]$count) {
    $rowType = $assembly.GetType($prefix + 'OperationSnapshot')
    $listType = [Collections.Generic.List``1].MakeGenericType($rowType)
    $rows = [Activator]::CreateInstance($listType)
    for ($i = 0; $i -lt $count; $i++) {
        $row = [Activator]::CreateInstance($rowType)
        $row.RecordNo = $i + 1
        $row.ParentRecordNo = ($i - ($i % 3)) / 3 - 2
        $row.OperationNumber = @(290,300,310)[$i % 3]
        $row.KilnType = @('Batch','Tunnel','SWK')[($i - ($i % 3)) / 3 % 3]
        $row.DueDate = [DateTime]'2026-11-03'
        $row.CycleType = if ($i % 17) { '150 VT' } else { '' }
        $row.RequiredResource = 'KILN'
        $row.VolumeOccupancy = if ($i % 19) { 10 } else { 0 }
        $row.IsScheduled = $i % 23 -eq 0
        $rows.Add($row)
    }
    return ,$rows
}
function Diagnose($assembly, $rows) {
    $collector = [Activator]::CreateInstance($assembly.GetType($prefix + 'SchedulerDebugCollector'))
    $collector.Enabled = $true
    $stages = $assembly.GetType($prefix + 'SchedulerStageDiagnostics')
    $arguments = New-Object object[] 2
    $arguments[0] = $rows
    $arguments[1] = $collector
    $stages.GetMethod('DiagnoseAll').Invoke($null, $arguments) | Out-Null
    return $collector
}
function New-SwkFixture($assembly, [int]$batchCount) {
    $plan = [Activator]::CreateInstance($assembly.GetType($prefix + 'swkOptimizer_vf+SwkBatchPlan'))
    for ($batch = $batchCount; $batch -gt 0; $batch--) {
        $plan.BatchStartByBatchNo[$batch] = [DateTime]'2026-11-03T08:00:00'
        $plan.BatchEndByBatchNo[$batch] = [DateTime]'2026-11-04T08:00:00'
        $plan.ResourceByBatchNo[$batch] = 'SWK,"quoted"'
        $plan.CycleByBatchNo[$batch] = 'HT024'
        $plan.TotalTonnageByBatchNo[$batch] = 12.5
        $plan.LateCountByBatchNo[$batch] = 0
        $plan.UnderfilledByBatchNo[$batch] = $false
        for ($item = 0; $item -lt 10; $item++) {
            $record = $batch * 10 + $item
            $plan.BatchNoByFiringOpRec[$record] = $batch
            $plan.QueueFiringOpRecs.Add($record)
        }
    }
    return $plan
}
function Export-Swk($assembly, $plan, [string]$destination) {
    $exporter = [Activator]::CreateInstance($assembly.GetType($prefix + 'swkOptimizer_vf'))
    $exporter.ExportSwkPlanToCsv($plan, $destination)
}
try {
    [Environment]::CurrentDirectory = $folder
    $rows = New-Object 'System.Collections.Generic.List[ExportFixture]'
    $rows.Add((New-Object ExportFixture))
    foreach ($culture in @('en-US','en-IN','fr-FR')) {
        [Threading.Thread]::CurrentThread.CurrentCulture = [Globalization.CultureInfo]::GetCultureInfo($culture)
        $signatures = @()
        foreach ($assembly in @($baseline,$optimized)) {
            [ExportFixture]::Reads.Clear()
            $path = Join-Path $folder ($assembly.Location.GetHashCode().ToString() + '.csv')
            Export-Csv $assembly $path $rows
            $signatures += [Convert]::ToBase64String([IO.File]::ReadAllBytes($path)) + '|' + ([ExportFixture]::Reads -join ',')
        }
        if ($signatures[0] -cne $signatures[1]) { throw "CSV bytes/getter order differ: $culture" }
    }
    [ExportFixture]::ThrowOnSecond = $true
    $signatures = @()
    foreach ($assembly in @($baseline,$optimized)) {
        [ExportFixture]::Reads.Clear()
        $path = Join-Path $folder ($assembly.Location.GetHashCode().ToString() + '.csv')
        try { Export-Csv $assembly $path $rows; throw 'Expected getter failure.' } catch {
            if ($_.Exception.ToString() -notmatch 'fixture') { throw }
        }
        $signatures += [Convert]::ToBase64String([IO.File]::ReadAllBytes($path)) + '|' + ([ExportFixture]::Reads -join ',')
    }
    if ($signatures[0] -cne $signatures[1]) { throw 'Throwing getter partial output/evaluation differs.' }
    [ExportFixture]::ThrowOnSecond = $false
    $legacySignatures = @()
    foreach ($assembly in @($baseline,$optimized)) {
        $legacyType = $assembly.GetType($prefix + 'AlgoSeq3')
        $legacy = [Activator]::CreateInstance($legacyType)
        $privateFlags = [Reflection.BindingFlags]'Instance,NonPublic'
        $legacyType.GetMethod('InitLogging', $privateFlags).Invoke($legacy, @()) | Out-Null
        $legacyType.GetField('_runId', $privateFlags).SetValue($legacy, 'fixed-run')
        foreach ($name in @('LogDecision','ExportPlannedRow')) {
            $method = $legacyType.GetMethod($name, $privateFlags)
            $parameters = $method.GetParameters()
            $arguments = New-Object object[] $parameters.Length
            for ($i = 0; $i -lt $parameters.Length; $i++) {
                $valueType = $parameters[$i].ParameterType
                if ($valueType -eq [string]) { $arguments[$i] = 'quoted,"value"' }
                elseif ($valueType -eq [int]) { $arguments[$i] = [int]7 }
                elseif ($valueType -eq [double]) { $arguments[$i] = [double]12.5 }
                else { $arguments[$i] = [DateTime]'2026-11-03T08:09:10' }
            }
            $method.Invoke($legacy, $arguments) | Out-Null
        }
        $signature = ''
        foreach ($field in @('_decisionLogPath','_scheduleExportPath')) {
            $path = [string]$legacyType.GetField($field, $privateFlags).GetValue($legacy)
            $signature += [Convert]::ToBase64String([IO.File]::ReadAllBytes($path))
        }
        $legacySignatures += $signature
    }
    if ($legacySignatures[0] -cne $legacySignatures[1]) { throw 'Enabled legacy decision/schedule log bytes differ.' }
    foreach ($custom in @($false,$true)) {
        $outputs = @()
        foreach ($assembly in @($baseline,$optimized)) {
            $table = New-Object Data.DataTable
            if ($custom) { $table.Columns.Add('A',[object]) | Out-Null } else { $table.Columns.Add('A',[string]) | Out-Null }
            $table.Columns.Add('B',[string]) | Out-Null
            $row = $table.NewRow(); $row[0] = 'a,b'; $row[1] = 'original'; $table.Rows.Add($row)
            if ($custom) { $cell = New-Object MutatingCell; $cell.Row = $row; $row[0] = $cell }
            $action = [Activator]::CreateInstance($assembly.GetType($prefix + 'AlgoSeq2'))
            $action.ExportDataTableToCsv($table, 'legacy')
            $outputs += [Convert]::ToBase64String([IO.File]::ReadAllBytes((Join-Path $folder 'legacy.csv')))
        }
        if ($outputs[0] -cne $outputs[1]) { throw "Legacy export differs (custom=$custom)." }
    }
    $oldPlan = New-SwkFixture $baseline 500
    $newPlan = New-SwkFixture $optimized 500
    Export-Swk $baseline $oldPlan (Join-Path $folder 'old-swk')
    Export-Swk $optimized $newPlan (Join-Path $folder 'new-swk')
    foreach ($file in @('SWK_BatchSummary.csv','SWK_FiringQueue_Return.csv')) {
        $oldBytes = [Convert]::ToBase64String([IO.File]::ReadAllBytes((Join-Path $folder ('old-swk/' + $file))))
        $newBytes = [Convert]::ToBase64String([IO.File]::ReadAllBytes((Join-Path $folder ('new-swk/' + $file))))
        if ($oldBytes -cne $newBytes) { throw "SWK output differs: $file" }
    }
    foreach ($assembly in @($baseline,$optimized)) {
        $emptyPlan = New-SwkFixture $assembly 0
        $emptyPlan.BatchNoByFiringOpRec = $null
        Export-Swk $assembly $emptyPlan (Join-Path $folder 'empty-swk')
    }
    $mappingPath = Join-Path $folder 'mapping.csv'
    [IO.File]::WriteAllText($mappingPath, "Id,Operation_name,Operation_number,150 VT`r`n1,Pressing,200,1.25`r`n2,Firing,300,2.5`r`n")
    $mappingResults = @()
    foreach ($assembly in @($baseline,$optimized)) {
        $helper = $assembly.GetType($prefix + 'Helper')
        $loader = $helper.GetMethod('LoadFromCsv')
        $parameters = $loader.GetParameters()
        $arguments = New-Object object[] $parameters.Length
        $arguments[0] = [string]$mappingPath
        for ($i = 1; $i -lt $parameters.Length; $i++) { $arguments[$i] = $parameters[$i].DefaultValue }
        $matrix = $loader.Invoke($null, $arguments)
        $signature = @()
        foreach ($methodName in @('TryGetDays','TryGetDaysByName')) {
            foreach ($key in @('200','Pressing','missing','')) {
                $arguments = [object[]]@('150 VT', $key, [decimal]99)
                $found = $helper.GetMethod($methodName).Invoke($matrix, $arguments)
                $signature += "$methodName|$key|$found|$($arguments[2])"
            }
        }
        $mappingResults += $signature -join ';'
    }
    if ($mappingResults[0] -cne $mappingResults[1]) { throw 'Process mapping lookup behavior differs.' }
    $oldRows = New-DiagnosticFixture $baseline 6000
    $newRows = New-DiagnosticFixture $optimized 6000
    $old = Diagnose $baseline $oldRows
    $new = Diagnose $optimized $newRows
    $oldSignature = $old.StageEligibilityRows | Select-Object * -ExcludeProperty RunId | ConvertTo-Json -Depth 4 -Compress
    $newSignature = $new.StageEligibilityRows | Select-Object * -ExcludeProperty RunId | ConvertTo-Json -Depth 4 -Compress
    if ($oldSignature -cne $newSignature) { throw 'Stage eligibility differs, including missing/zero/negative-parent groups.' }
    $results = @()
    # Warm both implementations before five alternating paired measurements.
    Diagnose $baseline $oldRows | Out-Null
    Diagnose $optimized $newRows | Out-Null
    for ($iteration = 0; $iteration -lt 5; $iteration++) {
        $order = if ($iteration % 2) { @('optimized','baseline') } else { @('baseline','optimized') }
        foreach ($name in $order) {
            $assembly = if ($name -eq 'baseline') { $baseline } else { $optimized }
            $inputRows = if ($name -eq 'baseline') { ,$oldRows } else { ,$newRows }
            $allocated = [AllocationMeter]::Read()
            $gen0 = [GC]::CollectionCount(0)
            $watch = [Diagnostics.Stopwatch]::StartNew()
            Diagnose $assembly $inputRows | Out-Null
            $watch.Stop()
            $results += [pscustomobject]@{ Implementation=$name; Iteration=$iteration; Milliseconds=$watch.Elapsed.TotalMilliseconds; AllocatedBytes=([AllocationMeter]::Read()-$allocated); Gen0Collections=([GC]::CollectionCount(0)-$gen0) }
        }
    }
    $evidenceRoot = Join-Path $PSScriptRoot 'obj/OptimizationEvidence'
    [IO.Directory]::CreateDirectory($evidenceRoot) | Out-Null
    $results | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $evidenceRoot 'diagnostic-benchmark.json')
    foreach ($name in @('baseline','optimized')) {
        $samples = @($results | Where-Object Implementation -EQ $name)
        $medianTime = @($samples.Milliseconds | Sort-Object)[2]
        $medianBytes = @($samples.AllocatedBytes | Sort-Object)[2]
        Write-Output "$name diagnostic fixture: median $medianTime ms; $medianBytes allocated bytes."
    }
    $swkResults = @()
    for ($iteration = 0; $iteration -lt 5; $iteration++) {
        $order = if ($iteration % 2) { @('optimized','baseline') } else { @('baseline','optimized') }
        foreach ($name in $order) {
            $assembly = if ($name -eq 'baseline') { $baseline } else { $optimized }
            $plan = if ($name -eq 'baseline') { $oldPlan } else { $newPlan }
            $allocated = [AllocationMeter]::Read()
            $watch = [Diagnostics.Stopwatch]::StartNew()
            Export-Swk $assembly $plan (Join-Path $folder ($name + '-swk-benchmark'))
            $watch.Stop()
            $swkResults += [pscustomobject]@{ Implementation=$name; Iteration=$iteration; Milliseconds=$watch.Elapsed.TotalMilliseconds; AllocatedBytes=([AllocationMeter]::Read()-$allocated) }
        }
    }
    $swkResults | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $evidenceRoot 'swk-benchmark.json')
    foreach ($name in @('baseline','optimized')) {
        $samples = @($swkResults | Where-Object Implementation -EQ $name)
        Write-Output "$name SWK export: median $(@($samples.Milliseconds | Sort-Object)[2]) ms; $(@($samples.AllocatedBytes | Sort-Object)[2]) allocated bytes."
    }
    Write-Output 'PASS: public API, CSV/getter failures, custom-object snapshots, SWK bytes/empty plans, mapping lookups, and 6000-row stage parity.'
} finally {
    [Environment]::CurrentDirectory = $previous
    $resolved = [IO.Path]::GetFullPath($folder)
    if ($resolved.StartsWith([IO.Path]::GetFullPath([IO.Path]::GetTempPath()),[StringComparison]::OrdinalIgnoreCase)) {
        Remove-Item -LiteralPath $resolved -Recurse -Force
    }
}
