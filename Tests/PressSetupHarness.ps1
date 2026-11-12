param([Parameter(Mandatory = $true)][string]$AssemblyPath)
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$assembly = [Reflection.Assembly]::LoadFrom((Resolve-Path -LiteralPath $AssemblyPath))
$sdk = Join-Path (Split-Path $assembly.Location) 'Preactor.Extensibility.dll'
[void][Reflection.Assembly]::LoadFrom($sdk)
Add-Type -ReferencedAssemblies $sdk -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Runtime.Remoting.Messaging;
using System.Runtime.Remoting.Proxies;
using Preactor;

public sealed class PressSetupProxy : RealProxy
{
    private readonly Func<string, object[], object> call;
    public PressSetupProxy(Type type, Func<string, object[], object> call) : base(type) { this.call = call; }
    public override IMessage Invoke(IMessage message)
    {
        var method = (IMethodCallMessage)message;
        try
        {
            object result = method.MethodName == "GetType" ? GetProxiedType() : call(method.MethodName, method.Args);
            return new ReturnMessage(result, null, 0, method.LogicalCallContext, method);
        }
        catch (Exception error) { return new ReturnMessage(error, method); }
    }
}

public sealed class PressSetupFakeBoard
{
    public readonly DateTime Start = new DateTime(2026, 10, 7, 8, 0, 0);
    public readonly Dictionary<int, OperationResourceTimes> Slots = new Dictionary<int, OperationResourceTimes>();
    public readonly Dictionary<int, double> SetupTimes = new Dictionary<int, double>();
    public readonly Dictionary<int, DateTime> ActualSetupStarts = new Dictionary<int, DateTime>();
    public readonly List<double> QueryValues = new List<double>();
    public readonly List<int> RecalculatedResources = new List<int>();
    public int AllocationCalls;
    private Dictionary<int, OperationResourceTimes> snapshot;
    private int snapshotId;
    public DateTime AvailableFrom;
    public int DropDuringRecalculation;
    public int IgnoreSetupDuringRecalculation;
    public int FailWriteForRecord;
    public int MoveDuringRecalculation;
    public int FailReadinessAfterRecalculation;
    public bool RejectSnapshot;
    public string ConfigurationFolder;
    public double Value(int record) { return SetupTimes.ContainsKey(record) ? SetupTimes[record] : 0; }
    public void Place(int record, int resource, DateTime start)
    {
        BoardCall("PutOperationOnResource", new object[] { record, resource, start });
    }
    public IPreactor GetPreactor()
    {
        var board = (IPlanningBoard)new PressSetupProxy(typeof(IPlanningBoard), BoardCall).GetTransparentProxy();
        return (IPreactor)new PressSetupProxy(typeof(IPreactor), (method, args) =>
        {
            switch (method)
            {
                case "get_PlanningBoard": return board;
                case "GetFormatNumber": return 1;
                case "GetFieldNumber":
                    return Array.IndexOf(new[] { "", "Setup Time", "String Attribute 5",
                        "String Attribute 4", "Finite or Infinite", "Lock Operation", "Actual Setup Start",
                        "Total Setup Time" }, (string)args[1]);
                case "GetFieldType": return PreactorFieldType.Duration;
                case "ReadFieldString": return "Finite";
                case "ReadFieldBool": return false;
                case "ReadFieldDateTime":
                    DateTime actualSetup;
                    return ActualSetupStarts.TryGetValue((int)args[2], out actualSetup)
                        ? actualSetup : DateTime.FromOADate(0);
                case "ReadFieldDouble":
                    int readRecord = (int)args[2];
                    if ((int)args[1] == 7)
                    {
                        var operation = Slots[readRecord].OperationTimes;
                        return (operation.ProcessStart - operation.ChangeStart).TotalDays;
                    }
                    return Value(readRecord);
                case "WriteField":
                    int record = (int)args[2];
                    double value = (double)args[3];
                    SetupTimes[record] = value;
                    if (record == FailWriteForRecord)
                    {
                        FailWriteForRecord = 0;
                        throw new InvalidOperationException("Write failed after changing the field");
                    }
                    return null;
                case "ParseShellString": return ConfigurationFolder;
                default: throw new NotSupportedException(method);
            }
        }).GetTransparentProxy();
    }
    private OperationTimes Times(int record, DateTime start)
    {
        if (AvailableFrom > start) start = AvailableFrom;
        double duration = Value(record);
        DateTime process = start.AddDays(duration);
        return new OperationTimes(start, process, process.AddMinutes(30));
    }
    private void Recalculate(int resource)
    {
        RecalculatedResources.Add(resource);
        var records = new List<int>();
        foreach (var slot in Slots)
            if (slot.Value.Resource == resource) records.Add(slot.Key);
        records.Sort((left, right) => Slots[left].OperationTimes.ChangeStart.CompareTo(Slots[right].OperationTimes.ChangeStart));
        DateTime end = DateTime.MinValue;
        foreach (int record in records)
        {
            if (record == DropDuringRecalculation) { Slots.Remove(record); continue; }
            var before = Slots[record].OperationTimes;
            var start = before.ChangeStart < end ? end : before.ChangeStart;
            if (record == MoveDuringRecalculation) start = start.AddMinutes(1);
            var setup = record == IgnoreSetupDuringRecalculation
                ? before.ProcessStart - before.ChangeStart : TimeSpan.FromDays(Value(record));
            var process = start + setup;
            end = process + (before.ProcessEnd - before.ProcessStart);
            Slots[record] = new OperationResourceTimes(resource, new OperationTimes(start, process, end));
        }
    }
    private object BoardCall(string method, object[] args)
    {
        switch (method)
        {
            case "get_TerminatorTime": return Start;
            case "get_ScheduleHorizon": return new DateTimeSpan(Start.AddDays(-1), Start.AddDays(10));
            case "GetResourceNumber":
                if ((string)args[0] == "PL06001A") return 1;
                if ((string)args[0] == "PS00761A") return 2;
                throw new ArgumentException("Unknown resource");
            case "GetResourceName": return "Press " + args[0];
            case "TestOperationOnResource":
                AllocationCalls++;
                throw new InvalidOperationException("Repair must not test or replace allocations");
            case "QueryOperationOnResource":
                QueryValues.Add(Value((int)args[0]));
                return Times((int)args[0], (DateTime)args[2]);
            case "PutOperationOnResource":
                AllocationCalls++;
                Slots[(int)args[0]] = new OperationResourceTimes((int)args[1], Times((int)args[0], (DateTime)args[2]));
                return true;
            case "IsOperationScheduled": return Slots.ContainsKey((int)args[0]);
            case "GetOperationTimes": return Slots.ContainsKey((int)args[0]) ? (object)Slots[(int)args[0]] : null;
            case "GetNextOperation": return 0;
            case "UnallocateOperation": AllocationCalls++; Slots.Remove((int)args[0]); return null;
            case "CanOperationStart": return RecalculatedResources.Count == 0 || (int)args[0] != FailReadinessAfterRecalculation;
            case "GetCurrentCalendarState": return null;
            case "RecalculateOperationTimes": Recalculate((int)args[0]); return null;
            case "SaveScheduleState":
                if (RejectSnapshot) throw new InvalidOperationException("Snapshot failed");
                snapshotId = (int)args[0];
                snapshot = new Dictionary<int, OperationResourceTimes>(Slots);
                return 0;
            case "RestoreScheduleState":
                if ((int)args[0] != snapshotId) throw new InvalidOperationException("Wrong snapshot ID");
                Slots.Clear();
                foreach (var slot in snapshot) Slots.Add(slot.Key, slot.Value);
                return Start;
            default: throw new NotSupportedException(method);
        }
    }
}
'@

$flags = [Reflection.BindingFlags]'Instance,NonPublic'
$boardType = $assembly.GetType('GrindwellNortonDev_VB.PressSetupBoard', $true)
$slotType = $assembly.GetType('GrindwellNortonDev_VB.PressCampaignSlot', $true)
$orderType = $assembly.GetType('GrindwellNortonDev_VB.PressCampaignOrder', $true)
$helpers = $assembly.GetType('GrindwellNortonDev_VB.SharedHelpers', $true)
$constructor = $boardType.GetConstructors($flags)[0]
$query = $boardType.GetMethod('Query', $flags)
$commit = $boardType.GetMethod('Commit', $flags)
$repair = $boardType.GetMethod('Repair', $flags)
$loadTimeline = $boardType.GetMethod('LoadTimeline', $flags)
$readRates = $boardType.GetMethod('ReadRates', [Reflection.BindingFlags]'Static,NonPublic')
$load = $boardType.GetMethod('Load', [Reflection.BindingFlags]'Static,NonPublic')
$readSetting = $boardType.GetMethod('ReadSetting', [Reflection.BindingFlags]'Static,NonPublic')

function Check([bool]$Result, [string]$Message) {
    if (-not $Result) { throw "FAIL: $Message" }
}
function Slot([int]$Resource, [datetime]$Start, [datetime]$Finish) {
    $slot = [Activator]::CreateInstance($slotType)
    $slot.RecordNo = 10
    $slot.Resource = $Resource
    $slot.Ring = '305-D'
    $slot.Pin = '10'
    $slot.IsPress = $true
    $slot.Start = $Start
    $slot.Finish = $Finish
    return $slot
}
function Minutes($Slot) { return $Slot.SetupDays * 1440 }

$folder = Join-Path ([IO.Path]::GetTempPath()) ('press-setup-harness-' + [guid]::NewGuid())
[void][IO.Directory]::CreateDirectory($folder)
$file = [IO.Path]::Combine($folder, 'PressSetupRates.csv')
try {
    $ratesCsv = "Workstation,RingSetupMinutes,BorePinSetupMinutes`nPL06001A,6,9`nPS00761A,5,`n"
    [IO.File]::WriteAllText($file, $ratesCsv)
    $fake = New-Object PressSetupFakeBoard
    $fake.ConfigurationFolder = $folder
    $adapter = $load.Invoke($null, @($fake.GetPreactor()))
    Check ($null -ne $adapter) 'DLL configuration enables the setup action without a host setting'
    $hostSettings = [System.Collections.Specialized.NameValueCollection]::new()
    $hostSettings['PressSetupCsvPath'] = ''
    $librarySettings = [System.Configuration.KeyValueConfigurationCollection]::new()
    $librarySettings.Add('PressSetupCsvPath', 'PressSetupRates.csv')
    Check (($readSetting.Invoke($null, @($hostSettings, $librarySettings, 'PressSetupCsvPath', 'default'))) -eq '') 'host setting can disable the DLL configuration'
    $start = $fake.Start
    $previous = Slot 1 ($start.AddHours(-1)) $start
    $partial = Slot 2 ($start.AddHours(-1)) $start
    $history = [Array]::CreateInstance($slotType, 2)
    $history.SetValue($previous, 0)
    $history.SetValue($partial, 1)
    [void]$loadTimeline.Invoke($adapter, @($history, [int[]]@(1, 2, 3)))
    $order = [Activator]::CreateInstance($orderType)
    $order.RecordNo = 11
    $order.Ring = ' 305-d '
    $order.Pin = '10'
    Check ((Minutes ($query.Invoke($adapter, @( $order, 1, $start, $null )))) -eq 0) 'unchanged contiguous tooling'
    $order.Ring = '310-D'
    Check ((Minutes ($query.Invoke($adapter, @( $order, 1, $start, $null )))) -eq 6) 'ring only'
    $order.Ring = '305-D'
    $order.Pin = '11'
    Check ((Minutes ($query.Invoke($adapter, @( $order, 1, $start, $null )))) -eq 9) 'pin only'
    Check ((Minutes ($query.Invoke($adapter, @( $order, 2, $start, $null )))) -eq 0) 'missing pin uses zero baseline'
    $order.Ring = '310-D'
    Check ((Minutes ($query.Invoke($adapter, @( $order, 1, $start, $null )))) -eq 15) 'both changed'
    Check ((Minutes ($query.Invoke($adapter, @( $order, 2, $start, $null )))) -eq 5) 'available ring applied without pin rate'
    Check ((Minutes ($query.Invoke($adapter, @( $order, 3, $start, $null )))) -eq 0) 'unrated press uses zero baseline'
    $order.Pin = ''
    Check ((Minutes ($query.Invoke($adapter, @( $order, 1, $start, $null )))) -eq 6) 'unknown pin does not prevent ring setup'
    $repairCases = @(
        @{ Ring = '305-D'; Pin = '10'; PriorRing = '305-D'; PriorPin = '10'; Expected = 0 }
        @{ Ring = '305-D'; Pin = ''; PriorRing = '305-D'; PriorPin = '10'; Expected = 0 }
        @{ Ring = '305-D'; Pin = '11'; PriorRing = '305-D'; PriorPin = ''; Expected = 0 }
        @{ Ring = ''; Pin = '11'; PriorRing = ''; PriorPin = '10'; Expected = 9 }
        @{ Ring = '310-D'; Pin = ''; PriorRing = '305-D'; PriorPin = '10'; Expected = 6 }
        @{ Ring = '310-D'; Pin = '11'; PriorRing = '305-D'; PriorPin = '10'; Expected = 15 }
        @{ Ring = ''; Pin = '11'; PriorRing = '305-D'; PriorPin = ''; Expected = 0 }
    )
    $previous.Ring = '305-D'
    $previous.Pin = '10'
    $order.Ring = '305-D'
    $order.Pin = '10'
    $fake.SetupTimes[11] = 23.0 / 1440
    $adapter = $load.Invoke($null, @($fake.GetPreactor()))
    [void]$loadTimeline.Invoke($adapter, @($history, [int[]]@(1, 2, 3)))
    $fake.QueryValues.Clear()
    $fake.AvailableFrom = $start.AddMilliseconds(1)
    $rounded = $query.Invoke($adapter, @( $order, 1, $start, $null ))
    Check ((Minutes $rounded) -eq 0 -and $fake.QueryValues.Count -eq 1 -and
        $rounded.Start -eq $fake.AvailableFrom) 'native timestamp rounding preserves contiguous setup'
    $fake.QueryValues.Clear()
    $fake.AvailableFrom = $start.AddMinutes(1)
    $result = $query.Invoke($adapter, @( $order, 1, $start, $null ))
    Check ((Minutes $result) -eq 15 -and $fake.QueryValues.Count -eq 2 -and $fake.QueryValues[0] -eq 0) 'shift into idle gap re-queries with full setup'
    Check (($result.ProcessStart - $result.Start).TotalMinutes -eq 15 -and $fake.Value(11) * 1440 -eq 23) 'query reserves setup and restores the prior Setup Time'
    $put = [Action[Action]]{
        param($Verify)
        $fake.Place(11, 1, $result.Start)
        $Verify.Invoke()
    }
    $placed = $commit.Invoke($adapter, @($result, $put))
    Check ($placed.ProcessStart -eq $fake.Slots[11].OperationTimes.ProcessStart -and
        $fake.Value(11) * 1440 -eq 15) 'commit persists the verified Setup Time once'

    $fixedFake = New-Object PressSetupFakeBoard
    $fixedAdapter = $constructor.Invoke(@($fixedFake.GetPreactor(), $file, 'String Attribute 5', 'String Attribute 4'))
    $successor = Slot 1 ($start.AddMinutes(30)) ($start.AddHours(1))
    $successor.RecordNo = 12
    $fixedHistory = [Array]::CreateInstance($slotType, 2)
    $fixedHistory.SetValue($previous, 0)
    $fixedHistory.SetValue($successor, 1)
    [void]$loadTimeline.Invoke($fixedAdapter, @($fixedHistory, [int[]]@(1)))
    $afterFixed = $query.Invoke($fixedAdapter, @($order, 1, $start, $null))
    Check ($afterFixed.Start -eq $successor.Finish) 'reject insertion that would change fixed successor setup'

    function Add-RepairRow($table, $board, [int]$record, [int]$resource, $begin, [double]$minutes,
                           [string]$ring, [string]$pin, [int]$operation = 200) {
        $row = $table.NewRow()
        $row['OrdersID'] = $record
        $row['Order No'] = "ORDER-$record"
        $row['Operation Number'] = $operation
        $row['is_scheduled'] = $null -ne $begin
        $row['Wheel Dia'] = $ring
        $row['Wheel thickness'] = $pin
        [void]$table.Rows.Add($row)
        $board.SetupTimes[$record] = $minutes / 1440
        if ($null -ne $begin) {
            $process = $begin.AddMinutes($minutes)
            $times = [Preactor.OperationTimes]::new($begin, $process, $process.AddMinutes(30))
            $board.Slots[$record] = [Preactor.OperationResourceTimes]::new($resource, $times)
        }
    }

    function Invoke-Repair($board, $table) {
        $adapter = $constructor.Invoke(@($board.GetPreactor(), $file, 'String Attribute 5', 'String Attribute 4'))
        $repair.Invoke($adapter, @(,$table))
    }

    $routing = $helpers.GetMethod('BuildRoutingSchema').Invoke($null, @())
    $repairCases += @(
        @{ Ring = '305-D'; Pin = '11'; PriorRing = '305-D'; PriorPin = '10'; Expected = 9 }
        @{ Ring = '310-D'; Pin = '11'; PriorRing = ''; PriorPin = '11'; Expected = 0 }
        @{ Ring = ''; Pin = '11'; PriorRing = '305-D'; PriorPin = '10'; Expected = 9 }
        @{ Ring = '310-D'; Pin = '11'; PriorRing = '305-D'; PriorPin = '10'; Expected = 5; Resource = 2 }
    )
    foreach ($case in $repairCases) {
        $testBoard = New-Object PressSetupFakeBoard
        $table = $routing.Clone()
        $resource = if ($case.ContainsKey('Resource')) { $case.Resource } else { 1 }
        Add-RepairRow $table $testBoard 10 $resource ($start.AddHours(-2)) 15 $case.PriorRing $case.PriorPin
        Add-RepairRow $table $testBoard 11 $resource $start 23 $case.Ring $case.Pin
        Check ((Invoke-Repair $testBoard $table) -eq 1 -and
            [math]::Abs($testBoard.Value(11) * 1440 - $case.Expected) -lt 0.0001 -and
            $testBoard.Slots[11].Resource -eq $resource -and $testBoard.AllocationCalls -eq 0) "repair compares known tooling across gaps, expected $($case.Expected) minutes"
    }

    $repairFake = New-Object PressSetupFakeBoard
    Add-RepairRow $routing $repairFake 11 1 $start 0 '305-D' '10'
    Add-RepairRow $routing $repairFake 12 1 ($start.AddMinutes(30)) 0 '305-D' '10'
    Add-RepairRow $routing $repairFake 13 1 $null 23 '310-D' '11'
    Check ((Invoke-Repair $repairFake $routing) -eq 1 -and
        ($repairFake.Slots[11].OperationTimes.ProcessStart - $start).TotalMinutes -eq 15 -and
        $repairFake.Slots[12].OperationTimes.ChangeStart -eq $start.AddMinutes(45) -and
        $repairFake.Value(12) -eq 0) 'blank native actual setup allows full setup on the first operation and native movement of the next'
    Check ((Invoke-Repair $repairFake $routing) -eq 0 -and
        $repairFake.RecalculatedResources.Count -eq 1 -and $repairFake.AllocationCalls -eq 0 -and
        -not $repairFake.Slots.ContainsKey(13) -and $repairFake.Value(13) * 1440 -eq 23) 'second pass performs no writes or recalculation and unallocated order stays untouched'

    foreach ($mode in 'success','allocation loss','fixed move','partial write','setup ignored','readiness','snapshot') {
        $chain = $routing.Clone()
        $chainFake = New-Object PressSetupFakeBoard
        Add-RepairRow $chain $chainFake 11 1 $start 15 '305-D' '10'
        Add-RepairRow $chain $chainFake 12 1 ($start.AddMinutes(45)) 15 '305-D' '10'
        Add-RepairRow $chain $chainFake 13 1 ($start.AddMinutes(90)) 15 '305-D' '10'
        Add-RepairRow $chain $chainFake 14 1 ($start.AddMinutes(150)) 0 '' '' 210
        $chainFake.ActualSetupStarts[11] = $start
        $expected = switch ($mode) {
            'allocation loss' { $chainFake.DropDuringRecalculation = 12; 'verify allocations.*ORDER-12.*restored and verified' }
            'fixed move' { $chainFake.MoveDuringRecalculation = 14; 'verify allocations.*ORDER-14.*fixed or unrelated.*restored and verified' }
            'partial write' { $chainFake.FailWriteForRecord = 13; 'write Setup Time.*ORDER-13.*restored and verified' }
            'setup ignored' { $chainFake.IgnoreSetupDuringRecalculation = 12; 'verify press timing and setup.*ORDER-12.*restored and verified' }
            'readiness' { $chainFake.FailReadinessAfterRecalculation = 12; 'verify press timing and setup.*ORDER-12.*readiness.*restored and verified' }
            'snapshot' { $chainFake.RejectSnapshot = $true; 'could not save the schedule state.*No setup writes' }
        }
        if ($mode -eq 'success') {
            Check ((Invoke-Repair $chainFake $chain) -eq 2 -and $chainFake.Value(11) * 1440 -eq 15 -and
                $chainFake.Value(12) -eq 0 -and $chainFake.Value(13) -eq 0 -and
                $chainFake.RecalculatedResources.Count -eq 1 -and $chainFake.Slots.Count -eq 4 -and
                $chainFake.Slots[14].OperationTimes.ChangeStart -eq $start.AddMinutes(150)) 'recalculate once for all updates on the press and preserve protected and unrelated allocations'
        } else {
            $failure = ''
            try { [void](Invoke-Repair $chainFake $chain) }
            catch { $failure = $_.Exception.InnerException.Message }
            Check ($failure -match $expected) "repair reports $mode at the correct stage. Got $failure"
            Check ($chainFake.Slots.Count -eq 4 -and $chainFake.Value(12) * 1440 -eq 15 -and
                $chainFake.Value(13) * 1440 -eq 15 -and
                $chainFake.Slots[12].OperationTimes.ChangeStart -eq $start.AddMinutes(45) -and
                $chainFake.Slots[14].OperationTimes.ChangeStart -eq $start.AddMinutes(150)) "repair preserves original allocation and setup on $mode"
        }
        Check ($chainFake.AllocationCalls -eq 0) "repair never explicitly changes allocations on $mode"
    }

    [IO.File]::WriteAllText($file, "Workstation,RingSetupMinutes,BorePinSetupMinutes`nPS00761A,5,`n")
    Check ((Invoke-Repair $repairFake $routing) -eq 1 -and $repairFake.Value(11) -eq 0 -and
        $repairFake.Slots[11].OperationTimes.ProcessStart -eq $start) 'removed rates reset prior calculated setup'

    [IO.File]::WriteAllText($file, $ratesCsv + "PL06001A,2,1`nBAD,-3,0`n")
    $resolver = [Func[string,int]]{ param($Name) if ($Name -eq 'PL06001A') { return 1 }; return 2 }
    $invalid = $false
    try { [void]$readRates.Invoke($null, @($file, $resolver)) }
    catch { $invalid = $_.Exception.InnerException.Message -match '(?s)Duplicate workstation.*nonnegative' }
    Check $invalid 'CSV groups duplicate and invalid-duration errors'
    Write-Host 'PASS: press setup rates, native adapter re-query, fallback, repair and rollback.'
}
finally {
    if (Test-Path -LiteralPath $file) { Remove-Item -LiteralPath $file }
    Remove-Item -LiteralPath $folder
}
