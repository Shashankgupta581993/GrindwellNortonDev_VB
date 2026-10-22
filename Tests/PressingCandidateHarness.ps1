param([Parameter(Mandatory = $true)][string]$AssemblyPath)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Exercise the real compiled candidate builder, including the shared completion
# helpers. This catches integration mistakes that a pure calendar test cannot.
$assembly = [Reflection.Assembly]::LoadFrom((Resolve-Path -LiteralPath $AssemblyPath))
$helpers = $assembly.GetType('GrindwellNortonDev_VB.SharedHelpers', $true)
$optimizerType = $assembly.GetType('GrindwellNortonDev_VB.pressingOptimizer_vf', $true)
$optimizer = [Activator]::CreateInstance($optimizerType)
$rank = $optimizerType.GetMethod('BuildRankedCandidates', [Reflection.BindingFlags]'Instance,NonPublic')
$table = $helpers.GetMethod('BuildRoutingSchema').Invoke($null, @())
$today = [datetime]'2026-10-07T08:00:00'

function Add-Order([int]$Id, [int]$Score, [datetime]$Due, [string]$Ring) {
    $r = $table.NewRow()
    $r['OrdersID'] = $Id
    $r['Order No'] = "Order-$Id"
    $r['Operation Number'] = 200
    $r['parent_record'] = $Id
    $r['is_scheduled'] = $false
    $r['Resource Group'] = 'PRESS'
    $r['Pressing earliest start'] = $today.Date
    $r['Pressing Due date'] = $Due
    $r['Wheel Dia'] = $Ring
    $r['Wheel thickness'] = 'PIN'
    $r['Cycle Type'] = '150VT'
    $r['wip_status'] = 'Candidate'
    $r['wip_score'] = $Score
    $r['wip_prev_op_rec'] = 1000 + $Id
    $r['wip_prev_op_scheduled'] = $true
    $r['wip_prev_op_released'] = $true
    $r['wip_ready_time'] = $today
    [void]$table.Rows.Add($r)
    return $r
}
function Check([bool]$Value, [string]$Message) {
    if (-not $Value) { throw "FAIL: $Message" }
}

[void](Add-Order 1 100 ($today.AddDays(2)) 'A')
[void](Add-Order 2 200 ($today.AddDays(2)) 'B')
[void](Add-Order 3 100 ($today.AddDays(3)) 'A')
$r = Add-Order 4 900 ($today.AddDays(2)) 'A'
$r['source_is_completed'] = $true
$r = Add-Order 5 900 ($today.AddDays(2)) 'A'
$r['is_scheduled'] = $true
$r = Add-Order 6 900 ($today.AddDays(2)) 'A'
$r['wip_prev_op_released'] = $false
$r = Add-Order 7 900 ($today.AddDays(2)) 'A'
$r['wip_ready_time'] = [DBNull]::Value
$r = Add-Order 8 50 ($today.AddDays(2)) 'A'
$r['wip_prev_op_rec'] = 0
$r['wip_ready_time'] = [DBNull]::Value
$r = Add-Order 9 900 ($today.AddDays(2)) 'A'
$r['opcenter_use_actual'] = $true
$r['actual_end_time'] = $today.AddHours(-1)
$r = Add-Order 10 900 ($today.AddDays(2)) 'A'
$r['Operation Number'] = 260
$r = Add-Order 11 900 ($today.AddDays(2)) 'A'
$r['wip_status'] = 'Rejected'

$result = $rank.Invoke($optimizer, @($table, $today, 2, $true, $null))
$ids = @($result | ForEach-Object { $_.RecordNo })
Check (($ids -join ',') -eq '2,1,3,8') 'WIP/date order, completion filtering, predecessor gate, and first-operation eligibility'
Check ($result[2].WheelPin -eq 'PIN') 'Legacy thickness column must carry bore/pin'
$publicQueue = $optimizerType.GetMethod('BuildPressing200Queue').Invoke($optimizer, @($table, $today, 2, $true, $null))
Check (($publicQueue -join ',') -eq '2,1,3,8') 'Public queue signature still returns parent records'
Write-Output 'PASS: candidate ranking, WIP/actualized/scheduled exclusions, readiness, field mapping, and public queue compatibility.'
