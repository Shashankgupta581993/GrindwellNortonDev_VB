param([Parameter(Mandatory = $true)][string]$AssemblyPath)

# Windows PowerShell: production and SDK target .NET Framework 4.8.
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$assembly = [Reflection.Assembly]::LoadFrom((Resolve-Path -LiteralPath $AssemblyPath))
$sdk = [Reflection.Assembly]::LoadFrom((Join-Path (Split-Path $assembly.Location) 'Preactor.Extensibility.dll'))
Add-Type -ReferencedAssemblies @($sdk.Location, 'System.Data', 'System.Xml') -TypeDefinition @'
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Reflection;
using System.Runtime.Remoting.Messaging;
using System.Runtime.Remoting.Proxies;
using Preactor;

public sealed class LoadingBoard : RealProxy {
    public readonly Dictionary<int, OperationResourceTimes> Times = new Dictionary<int, OperationResourceTimes>();
    public TimeSpan Process = TimeSpan.FromMinutes(30), Setup = TimeSpan.FromMinutes(10);
    public bool Accept = true, Shift, WrongResource, ThrowPlacement, MissingFiringTimes, HideLoading;
    public int Placements, Previous = 1;
    public DateTime Finish = new DateTime(2026, 11, 3, 8, 0, 0);
    public LoadingBoard() : base(typeof(IPlanningBoard)) {
        Times[2] = new OperationResourceTimes(90, new OperationTimes(Finish, Finish.AddMinutes(20), Finish.AddHours(8)));
    }
    public override IMessage Invoke(IMessage message) {
        var c = (IMethodCallMessage)message;
        object result;
        try {
            switch (c.MethodName) {
                case "GetType": result = typeof(IPlanningBoard); break;
                case "IsOperationScheduled": result = Times.ContainsKey((int)c.Args[0]); break;
                case "GetPreviousOperation": result = Previous; break;
                case "GetOperationTimes":
                    int record = (int)c.Args[0];
                    result = (record == 2 && MissingFiringTimes) || (record == 1 && HideLoading) || !Times.ContainsKey(record)
                        ? (object)null : Times[record]; break;
                case "GetProcessTime":
                    if ((DateTime)c.Args[2] != Finish) throw new Exception("Process time must be evaluated at firing setup start");
                    result = Process; break;
                case "GetSetupTime":
                    if ((DateTime)c.Args[2] != Finish.Subtract(Process)) throw new Exception("Setup time must be evaluated at loading process start");
                    result = Setup; break;
                case "QuickPutOperationOnResource":
                    Placements++;
                    if (ThrowPlacement) throw new InvalidOperationException("Simulated quick placement exception");
                    var t = (OperationTimes)c.Args[2];
                    if (Shift) t = new OperationTimes(t.ChangeStart, t.ProcessStart, t.ProcessEnd.AddMinutes(1));
                    if (Accept) Times[(int)c.Args[0]] = new OperationResourceTimes(WrongResource ? 999 : (int)c.Args[1], t);
                    result = Accept; break;
                default: throw new Exception("Unexpected board call: " + c.MethodName);
            }
            return new ReturnMessage(result, null, 0, null, c);
        } catch (Exception e) { return new ReturnMessage(e, c); }
    }
    public IPlanningBoard Client { get { return (IPlanningBoard)GetTransparentProxy(); } }
}

public sealed class LoadingPreactor : RealProxy {
    public LoadingPreactor() : base(typeof(IPreactor)) { }
    public override IMessage Invoke(IMessage message) {
        var c = (IMethodCallMessage)message;
        if (c.MethodName == "GetType") return new ReturnMessage(typeof(IPreactor), null, 0, null, c);
        if (c.MethodName == "GetHashCode") return new ReturnMessage(1, null, 0, null, c);
        if (c.MethodName == "Equals") return new ReturnMessage(Object.ReferenceEquals(GetTransparentProxy(), c.Args[0]), null, 0, null, c);
        return new ReturnMessage(new Exception("Unexpected Preactor call: " + c.MethodName), c);
    }
    public IPreactor Client { get { return (IPreactor)GetTransparentProxy(); } }
}

public sealed class LoadingCase {
    const BindingFlags Instance = BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic;
    const BindingFlags Static = BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic;
    public readonly LoadingBoard Board = new LoadingBoard();
    public readonly LoadingPreactor Preactor = new LoadingPreactor();
    public readonly Dictionary<int, DataRow> Rows = new Dictionary<int, DataRow>();
    public int Stamps;
    public bool FailStamp;
    readonly object algo;
    readonly object collector;
    readonly MethodInfo load, repair;
    public LoadingCase(Assembly assembly, int kilnType, int loadingNo) {
        var table = new DataTable();
        table.Columns.Add("OrdersID", typeof(int));
        table.Columns.Add("Operation Number", typeof(int));
        table.Columns.Add("Kiln Type", typeof(int));
        table.Columns.Add("operation_effective_completed", typeof(bool));
        table.Columns.Add("source_is_completed", typeof(bool));
        table.Columns.Add("opcenter_use_actual", typeof(bool));
        table.Columns.Add("actual_end_time", typeof(DateTime));
        table.Columns.Add("wip_status", typeof(string));
        // Loading is rejected by normal WIP gating because firing is scheduled.
        Rows[1] = table.Rows.Add(1, loadingNo, kilnType, false, false, false, DBNull.Value, "Rejected");
        Rows[2] = table.Rows.Add(2, 300, kilnType, false, false, false, DBNull.Value, "Rejected");
        var type = assembly.GetType("GrindwellNortonDev_VB.AlgoSeq4", true);
        algo = Activator.CreateInstance(type);
        collector = Activator.CreateInstance(assembly.GetType("GrindwellNortonDev_VB.SchedulerDebugCollector", true));
        type.GetField("_schedulerDebug", Instance).SetValue(algo, collector);
        load = type.GetMethod("ScheduleLoadingBeforeFiring", Instance);
        repair = type.GetMethod("RepairScheduledFiringLoading", Instance);
        var module = assembly.GetType("GrindwellNortonDev_VB.MesSchedulingDates", true);
        var preparedType = module.GetNestedType("PreparedAction", BindingFlags.NonPublic);
        var prepared = Activator.CreateInstance(preparedType, true);
        var cycleType = assembly.GetType("GrindwellNortonDev_VB.MesSchedulingCycle", true);
        var cycle = Activator.CreateInstance(cycleType, Instance, null, new object[] {new DateTime(2026, 11, 3)}, null);
        preparedType.GetField("Cycle", Instance).SetValue(prepared, cycle);
        preparedType.GetField("WriteDate", Instance).SetValue(prepared, new Action<int, DateTime>((rec, date) => {
            if (FailStamp) throw new Exception("Simulated stamp failure");
            if (rec != 1 || date != new DateTime(2026, 11, 3)) throw new Exception("Wrong MES stamp");
            Stamps++;
        }));
        var actions = module.GetField("actions", Static).GetValue(null);
        actions.GetType().GetMethod("Add").Invoke(actions, new object[] {Preactor.Client, prepared});
    }
    public void Load(int resource) { load.Invoke(algo, new object[] {Preactor.Client, Board.Client, Rows, 2, resource, "LoadingTest"}); }
    public void Repair(int kilnType, int resource) { repair.Invoke(algo, new object[] {Preactor.Client, Board.Client, Rows, kilnType, resource, "LoadingRepairTest"}); }
    public IList Traces { get { return (IList)collector.GetType().GetField("ScheduleAttemptTraceRows").GetValue(collector); } }
}

public static class LoadingTests {
    static int count;
    static void Check(bool value, string detail) { count++; if (!value) throw new Exception("FAIL: " + detail); }
    static void Failure(Action run, string detail) {
        try { run(); } catch (TargetInvocationException e) {
            Check(e.ToString().Contains(detail), "Expected failure containing " + detail + ": " + e);
            return;
        }
        throw new Exception("Expected failure: " + detail);
    }
    public static int Run(Assembly assembly) {
        for (int kind = 1; kind <= 3; kind++) {
            foreach (int loadingNo in new[] {290, 291}) {
                var c = new LoadingCase(assembly, kind, loadingNo);
                var firingBefore = c.Board.Times[2];
                c.Load(100 + kind);
                var actual = c.Board.Times[1];
                Check(actual.Resource == 100 + kind, "Correct loading resource");
                Check(actual.OperationTimes.ProcessEnd == c.Board.Finish, "Loading ends at actual firing ChangeStart");
                Check(actual.OperationTimes.ProcessStart == c.Board.Finish.AddMinutes(-30), "Process duration retained");
                Check(actual.OperationTimes.ChangeStart == c.Board.Finish.AddMinutes(-40), "Setup duration retained, including historical time");
                Check(c.Stamps == 1, "Successful placement stamped");
                c.Load(100 + kind);
                c.Repair(kind, 100 + kind);
                Check(c.Board.Placements == 1 && c.Stamps == 1, "Reruns preserve scheduled loading");
                Check(c.Board.Times[2].Equals(firingBefore), "Firing remains unchanged");
                var repair = new LoadingCase(assembly, kind, loadingNo);
                repair.Repair(kind == 3 ? 1 : kind + 1, 100 + kind);
                Check(repair.Board.Placements == 0, "Repair filters other kiln types");
                repair.Repair(kind, 100 + kind);
                Check(repair.Board.Placements == 1 && repair.Stamps == 1, "Repair bypasses rejected WIP status for scheduled firing");
            }
        }
        var zero = new LoadingCase(assembly, 1, 290);
        zero.Board.Process = zero.Board.Setup = TimeSpan.Zero;
        zero.Load(101);
        Check(zero.Board.Times[1].OperationTimes.ChangeStart == zero.Board.Finish, "Zero duration accepted");
        bool tracing = (bool)assembly.GetType("GrindwellNortonDev_VB.SchedulerDebugSettings", true)
            .GetField("MasterEnabled", BindingFlags.Static | BindingFlags.NonPublic).GetValue(null);
        Check(zero.Traces.Count == (tracing ? 1 : 0), "Forced loading obeys attempt trace switch");
        if (tracing) {
            var trace = zero.Traces[0];
            Check((bool)trace.GetType().GetProperty("ScheduledAfterAttempt").GetValue(trace, null), "Successful loading attempt traced");
        }
        foreach (string flag in new[] {"operation_effective_completed", "source_is_completed", "opcenter_use_actual"}) {
            var c = new LoadingCase(assembly, 1, 290);
            c.Rows[1][flag] = true;
            if (flag == "opcenter_use_actual") c.Rows[1]["actual_end_time"] = c.Board.Finish;
            c.Load(101); c.Repair(1, 101);
            Check(c.Board.Placements == 0 && c.Stamps == 0, "Preserve " + flag);
        }
        var other = new LoadingCase(assembly, 1, 280); other.Load(101);
        Check(other.Board.Placements == 0, "Do not force a drying predecessor");
        var none = new LoadingCase(assembly, 1, 290); none.Board.Previous = 0; none.Load(101);
        Check(none.Board.Placements == 0, "Absent predecessor skipped");
        var unfired = new LoadingCase(assembly, 1, 290); unfired.Board.Times.Clear(); unfired.Load(101); unfired.Repair(1, 101);
        Check(unfired.Board.Placements == 0, "Unscheduled firing never forces loading");
        var negative = new LoadingCase(assembly, 1, 290); negative.Board.Process = TimeSpan.FromMinutes(-1);
        Failure(() => negative.Load(101), "Invalid loading process duration");
        Check(negative.Board.Placements == 0 && negative.Stamps == 0, "Negative process duration rejected before placement");
        var setup = new LoadingCase(assembly, 1, 290); setup.Board.Setup = TimeSpan.FromMinutes(-1);
        Failure(() => setup.Load(101), "Invalid loading setup duration");
        Check(setup.Board.Placements == 0 && setup.Stamps == 0, "Negative setup rejected before placement");
        var missing = new LoadingCase(assembly, 1, 290); missing.Board.MissingFiringTimes = true;
        Failure(() => missing.Load(101), "no valid start time");
        var resource = new LoadingCase(assembly, 1, 290);
        Failure(() => resource.Load(0), "Loading resource not found");
        var row = new LoadingCase(assembly, 1, 290); row.Rows.Remove(1);
        Failure(() => row.Load(101), "Loading routing row missing");
        var rejected = new LoadingCase(assembly, 1, 290); rejected.Board.Accept = false;
        Failure(() => rejected.Load(101), "placement rejected");
        Check(rejected.Stamps == 0, "Rejected placement is not stamped");
        if (tracing) {
            Check(rejected.Traces.Count == 1, "Rejected loading attempt traced");
            var trace = rejected.Traces[0];
            Check(trace.GetType().GetProperty("ExceptionMessage").GetValue(trace, null).ToString().Contains("placement rejected"), "Loading rejection detail retained");
        }
        rejected.Board.Accept = true; rejected.Repair(1, 101);
        Check(rejected.Stamps == 1 && rejected.Board.Placements == 2, "Rejected placement retry repairs loading");
        var shifted = new LoadingCase(assembly, 1, 290); shifted.Board.Shift = true;
        Failure(() => shifted.Load(101), "timings were not retained");
        Check(shifted.Stamps == 0, "Incorrect timing not stamped");
        var wrong = new LoadingCase(assembly, 1, 290); wrong.Board.WrongResource = true;
        Failure(() => wrong.Load(101), "timings were not retained");
        var hidden = new LoadingCase(assembly, 1, 290); hidden.Board.HideLoading = true;
        Failure(() => hidden.Load(101), "timings were not retained");
        var throwing = new LoadingCase(assembly, 1, 290); throwing.Board.ThrowPlacement = true;
        Failure(() => throwing.Load(101), "Simulated quick placement exception");
        Check(throwing.Stamps == 0, "Throwing placement not stamped");
        var stamp = new LoadingCase(assembly, 1, 290); stamp.FailStamp = true;
        Failure(() => stamp.Load(101), "MesSchedulingException");
        Check(stamp.Board.Placements == 1 && stamp.Stamps == 0, "Stamp errors propagate after placement");
        return count;
    }
}
'@
$count = [LoadingTests]::Run($assembly)

# Check the production action wiring in addition to the behavioral proxy tests.
$source = Get-Content -Raw -LiteralPath (Join-Path (Split-Path -Parent $PSScriptRoot) 'AlgoSeq4.vb')
foreach ($name in @('runFiring', 'runSWKFiring', 'runFiring2')) {
    $body = [regex]::Match($source, "(?ms)^    Public Function $name\(.*?^    End Function").Value
    if (-not $body.Contains('ScheduleLoadingBeforeFiring(') -or -not $body.Contains('RepairScheduledFiringLoading(') -or $body.Contains('BackTestOpOnResource(')) {
        throw "Forced loading/repair wiring missing in $name."
    }
}
Write-Output "Loading 290 harness passed ($count behavioral assertions; all three firing actions wired)."
