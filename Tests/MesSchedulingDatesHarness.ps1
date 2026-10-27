param([Parameter(Mandatory = $true)][string]$AssemblyPath)

# Run with Windows PowerShell (the application and SDK target .NET Framework).
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$assembly = [Reflection.Assembly]::LoadFrom((Resolve-Path -LiteralPath $AssemblyPath))
$sdk = [Reflection.Assembly]::LoadFrom((Join-Path (Split-Path $AssemblyPath) 'Preactor.Extensibility.dll'))
$flags = [Reflection.BindingFlags]'Instance,NonPublic,Public'
$static = [Reflection.BindingFlags]'Static,NonPublic,Public'
$cycleType = $assembly.GetType('GrindwellNortonDev_VB.MesSchedulingCycle', $true)
$moduleType = $assembly.GetType('GrindwellNortonDev_VB.MesSchedulingDates', $true)

Add-Type -ReferencedAssemblies $sdk.Location -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Runtime.Remoting.Messaging;
using System.Runtime.Remoting.Proxies;
using Preactor;

public sealed class MesDateWriter {
    public readonly Dictionary<int, DateTime> Dates = new Dictionary<int, DateTime>();
    public readonly Dictionary<int, int> Counts = new Dictionary<int, int>();
    public int FailRecord;
    public void Write(int record, DateTime date) {
        if (record == FailRecord) throw new InvalidOperationException("Simulated date write failure");
        Dates[record] = date;
        Counts[record] = Counts.ContainsKey(record) ? Counts[record] + 1 : 1;
    }
    public Action<int, DateTime> Callback { get { return Write; } }
}

public sealed class MesPreactorProxy : RealProxy {
    public MesPreactorProxy() : base(typeof(IPreactor)) { }
    public override IMessage Invoke(IMessage message) {
        var call = (IMethodCallMessage)message;
        if (call.MethodName == "GetType") return new ReturnMessage(typeof(IPreactor), null, 0, null, call);
        if (call.MethodName == "GetHashCode") return new ReturnMessage(1, null, 0, null, call);
        if (call.MethodName == "Equals") return new ReturnMessage(Object.ReferenceEquals(GetTransparentProxy(), call.Args[0]), null, 0, null, call);
        return new ReturnMessage(new InvalidOperationException("Unexpected Preactor call: " + call.MethodName), call);
    }
    public IPreactor Client { get { return (IPreactor)GetTransparentProxy(); } }
}

public sealed class MesBoardProxy : RealProxy {
    public bool Success = true;
    public bool ThrowPlacement;
    public int Placements;
    public int Record;
    public int Resource;
    public DateTime Start;
    public MesBoardProxy() : base(typeof(IPlanningBoard)) { }
    public override IMessage Invoke(IMessage message) {
        var call = (IMethodCallMessage)message;
        if (call.MethodName == "GetType") return new ReturnMessage(typeof(IPlanningBoard), null, 0, null, call);
        if (call.MethodName != "PutOperationOnResource")
            return new ReturnMessage(new InvalidOperationException("Unexpected board call: " + call.MethodName), call);
        Placements++;
        if (ThrowPlacement) return new ReturnMessage(new InvalidOperationException("Simulated placement failure"), call);
        Record = (int)call.Args[0];
        Resource = (int)call.Args[1];
        Start = (DateTime)call.Args[2];
        return new ReturnMessage(Success, null, 0, null, call);
    }
    public IPlanningBoard Client { get { return (IPlanningBoard)GetTransparentProxy(); } }
}
'@

function Assert([bool]$Condition, [string]$Message) {
    if (-not $Condition) { throw "FAIL: $Message" }
}

function New-Cycle([datetime]$Date) {
    return [Activator]::CreateInstance($cycleType, $flags, $null, [object[]]@($Date), $null)
}

function Stamp($Cycle, [int]$Record, $Writer) {
    $cycleType.GetMethod('Stamp', $flags).Invoke($Cycle, [object[]]@($Record, $Writer.Callback))
}

function Expect-MesFailure([scriptblock]$Action) {
    try { & $Action } catch {
        $errorObject = $_.Exception
        while ($null -ne $errorObject.InnerException) {
            if ($errorObject.GetType().Name -eq 'MesSchedulingException') { return }
            $errorObject = $errorObject.InnerException
        }
        if ($errorObject.GetType().Name -eq 'MesSchedulingException') { return }
        throw
    }
    throw 'Expected an MES scheduling failure.'
}

$date = [datetime]'2026-10-27T14:30:00'
$writer = New-Object MesDateWriter
$cycle = New-Cycle $date
$writer.Dates[900] = [datetime]'2026-10-01' # An untouched operation.
Stamp $cycle 10 $writer
Stamp $cycle 10 $writer # Identical replacement / move back / another rule button.
Stamp $cycle 11 $writer
Assert ($writer.Dates[10] -eq $date.Date) 'The selected business date must be stored without a time.'
Assert ($writer.Counts[10] -eq 1) 'Repeated placements in one cycle must not repeat the date write.'
Assert ($writer.Dates[11] -eq $date.Date) 'Different operations must share the cycle date.'
Assert ($writer.Dates[900] -eq [datetime]'2026-10-01') 'Untouched operations must retain their dates.'

$writer.FailRecord = 12
Expect-MesFailure { Stamp $cycle 12 $writer }
Assert (-not $writer.Dates.ContainsKey(12)) 'A failed write must not be marked as stamped.'
$writer.FailRecord = 0
$cycleType.GetMethod('RetryPending', $flags).Invoke($cycle, [object[]]@($writer.Callback))
Assert ($writer.Dates[12] -eq $date.Date) 'Pending writes must be repaired using the original date.'
Stamp $cycle 12 $writer
Assert ($writer.Counts[12] -eq 1) 'A repaired write must enter the deduplication set.'
$newCycle = New-Cycle ([datetime]'2026-10-28')
Stamp $newCycle 10 $writer
Assert ($writer.Dates[10] -eq [datetime]'2026-10-28') 'A new cycle must allow overwriting the previous date.'
Assert ($writer.Counts[10] -eq 2) 'The new cycle must not reuse the old deduplication set.'

# Exercise the real shared placement helper, using SDK interface proxies. These
# proxies reject every unexpected API read, including extra schedule snapshots.
$fakePreactor = New-Object MesPreactorProxy
$fakeBoard = New-Object MesBoardProxy
$preparedType = $moduleType.GetNestedType('PreparedAction', [Reflection.BindingFlags]::NonPublic)
$prepared = [Activator]::CreateInstance($preparedType, $true)
$preparedType.GetField('Cycle', $flags).SetValue($prepared, $cycle)
$preparedType.GetField('WriteDate', $flags).SetValue($prepared, $writer.Callback)
$actions = $moduleType.GetField('actions', $static).GetValue($null)
$actions.GetType().GetMethod('Add').Invoke($actions, [object[]]@($fakePreactor.Client, $prepared))
$algoType = $assembly.GetType('GrindwellNortonDev_VB.AlgoSeq4', $true)
$algo = [Activator]::CreateInstance($algoType)
$place = $algoType.GetMethod('PutOperationWithTrace', $flags)
function Place([int]$Record, [int]$Resource, [datetime]$Start) {
    $place.Invoke($algo, [object[]]@($fakePreactor.Client, $fakeBoard.Client,
        'Test', $Record, $Resource, $Start, 'Forward', $null, $null))
}

$start = [datetime]'2026-11-02T08:00:00'
Place 20 7 $start
Assert ($writer.Dates[20] -eq $date.Date) 'A successful placement must receive the cycle date.'
Assert ($fakeBoard.Record -eq 20 -and $fakeBoard.Resource -eq 7 -and $fakeBoard.Start -eq $start) 'Placement arguments must remain unchanged.'
Place 20 8 ($start.AddHours(1))
Place 20 7 $start
Assert ($writer.Dates[20] -eq $date.Date -and $writer.Counts[20] -eq 1) 'Moving back must retain the new date without another write.'
$fakeBoard.Success = $false
Place 21 7 $start
Assert (-not $writer.Dates.ContainsKey(21)) 'A false placement result must not be stamped.'
$fakeBoard.ThrowPlacement = $true
try { Place 22 7 $start; throw 'Expected placement failure.' } catch {
    Assert ($_.Exception.ToString().Contains('Simulated placement failure')) 'Placement exceptions must propagate.'
}
Assert (-not $writer.Dates.ContainsKey(22)) 'A throwing placement must not be stamped.'
$fakeBoard.ThrowPlacement = $false
$fakeBoard.Success = $true
$writer.FailRecord = 23
Expect-MesFailure { Place 23 7 $start }
Assert (-not $writer.Dates.ContainsKey(23)) 'A stamp failure must propagate after placement.'

# Every compiled placement site must branch on the SDK Boolean before stamping.
foreach ($name in @('AlgoSeq.vb', 'AlgoSeq3.vb', 'AlgoSeq4.vb', 'EventSeq.vb', 'PostFiringScheduler.vb')) {
    $source = Get-Content -Raw -LiteralPath (Join-Path $root $name)
    $calls = [regex]::Matches($source, '(?m)^\s*Dim mesPlacementSucceeded As Boolean = \w+\.PutOperationOnResource\(').Count
    $stamps = [regex]::Matches($source, 'If mesPlacementSucceeded Then MesSchedulingDates\.StampSuccessfulPlacement').Count
    Assert ($calls -gt 0 -and $calls -eq $stamps) "Every placement in $name must conditionally stamp its operation."
    Assert (-not ($source -match '(?m)^\s*\w+\.PutOperationOnResource\(')) "No uninstrumented placement may remain in $name."
}
$source = Get-Content -Raw -LiteralPath (Join-Path $root 'AlgoSeq4.vb')
foreach ($name in @('ActualSchedule', 'Schedule400', 'runFiring', 'runSWKFiring', 'runFiring2', 'runFiringFollowOn', 'runFix', 'afterFiring', 'untilPress', 'runPress', 'runPressToFiring')) {
    $body = [regex]::Match($source, "(?ms)^    Public Function $name\(.*?^    End Function").Value
    Assert ($body.Contains('If Not MesSchedulingDates.TryBeginAction(preactor, preactorComObject) Then Return -1')) "Rule $name must prepare its date before scheduling."
}
Write-Output 'MES scheduling date harness passed (cycle semantics, retries, real placement adapter, and placement coverage).'
