Option Strict On
Option Explicit On

Imports System
Imports System.Configuration
Imports System.Data
Imports System.Globalization
Imports System.IO
Imports System.Text
Imports System.Diagnostics
Imports System.Collections.Generic
Imports System.Linq
Imports System.Runtime.InteropServices
Imports Preactor
Imports Preactor.Interop.PreactorObject

<ComVisible(True)> _
<Microsoft.VisualBasic.ComClass("e8de5ac3-e957-43bc-b37f-1c5c110cc044", "f72af322-4acd-4ba9-99f0-016e15e7d939")> _
Public Class AlgoSeq3
    Public Function Run2(ByRef preactorComObject As PreactorObj, ByRef pespComObject As Object) As Integer

        Dim preactor As IPreactor = PreactorFactory.CreatePreactorObject(preactorComObject)
        Return 0
    End Function



    '=========================================================
    ' Algo Sequencing Rule - CSV-driven planner, Opcenter applier
    '
    ' Key premise from you:
    '   Routing.csv "OrdersID" == Opcenter operation record number (opRec)
    '
    ' So each row in the CSV represents one operation record in Opcenter,
    ' and OrdersID can be used directly as opRec for scheduling.
    '
    ' This class follows the same PlanningBoard usage pattern in your template:
    '   - CreateQueue / RankQueue
    '   - FindResources(opRec)
    '   - TestOperationOnResource(opRec, resRec, terminator)
    '   - PutOperationOnResource(opRec, resRec, startTime)
    '
    ' Reference template: algo.txt :contentReference[oaicite:1]{index=1}
    '=========================================================

    '===========================
    ' CONFIG (parameterized)
    '===========================
    Private Const MinOcc As Double = 0.8
        Private Const MaxOcc As Double = 1.0
        Private Const TunnelCarsPerDayDefault As Integer = 4

        ' Wheel dia cooldown: if used on Day D on a resource => block same dia on D+1 and D+2
        Private Const DiaCooldownDays As Integer = 2

        ' Morning shift for Op 290 (your final answer): 08:00 to 14:00
        Private ReadOnly MorningShiftStart As TimeSpan = New TimeSpan(8, 0, 0)
        Private ReadOnly MorningShiftEnd As TimeSpan = New TimeSpan(14, 0, 0)

        '===========================
        ' DEBUG/EXPORT FILES
        '===========================
        Private _runId As String
        Private _logDir As String
        Private _decisionLogPath As String
        Private _scheduleExportPath As String

        Private Sub InitLogging()
            _runId = DateTime.Now.ToString("yyyyMMdd_HHmmss", CultureInfo.InvariantCulture)
            _logDir = Path.Combine(Directory.GetCurrentDirectory(), "AlgoSeqLogs")
            Directory.CreateDirectory(_logDir)


            _decisionLogPath = Path.Combine(_logDir, $"DecisionLog_{_runId}.csv")
            _scheduleExportPath = Path.Combine(_logDir, $"PlannedSchedule_{_runId}.csv")

        File.WriteAllText(_decisionLogPath,
                              "RunId,Phase,OpRec,OrderNo,OpNo,ResGroup,ChosenRes,PlanDay,Start,Finish,BatchOrCartId,kilnType,CycleType,Occ,Score,Reason" & Environment.NewLine)

        File.WriteAllText(_scheduleExportPath,
                              "RunId,OpRec,OrderNo,OpNo,OpName,ResGroup,ChosenRes,PlanDay,Start,Finish,BatchOrCartId,kilnType,CycleType,Occ" & Environment.NewLine)
    End Sub

    Private Sub LogDecision(phase As String,
                                opRec As Integer,
                                orderNo As String,
                                opNo As Integer,
                                resGroup As String,
                                chosenRes As String,
                                planDay As DateTime,
                                startTime As DateTime?,
                                finishTime As DateTime?,
                                batchOrCartId As String,
                                kilnType As String,
                                cycleType As String,
                                occ As Double,
                                score As Double,
                                reason As String)

        Dim sStart As String = If(startTime.HasValue, startTime.Value.ToString("yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture), "")
        Dim sFinish As String = If(finishTime.HasValue, finishTime.Value.ToString("yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture), "")

        Dim line = String.Join(",",
                Csv(_runId),
                Csv(phase),
                opRec.ToString(CultureInfo.InvariantCulture),
                Csv(orderNo),
                opNo.ToString(CultureInfo.InvariantCulture),
                Csv(resGroup),
                Csv(chosenRes),
                Csv(planDay.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)),
                Csv(sStart),
                Csv(sFinish),
                Csv(batchOrCartId),
                Csv(kilnType),
                Csv(cycleType),
                occ.ToString(CultureInfo.InvariantCulture),
                score.ToString(CultureInfo.InvariantCulture),
                Csv(reason)
            ) & Environment.NewLine

        File.AppendAllText(_decisionLogPath, line)

        ' Also write to VS Output window (Debug)
        Debug.WriteLine($"[{phase}] opRec={opRec} order={orderNo} opNo={opNo} resGrp={resGroup} res={chosenRes} day={planDay:yyyy-MM-dd} reason={reason}")
    End Sub

    Private Sub ExportPlannedRow(opRec As Integer,
                                    orderNo As String,
                                    opNo As Integer,
                                    opName As String,
                                    resGroup As String,
                                    chosenRes As String,
                                    planDay As DateTime,
                                    startTime As DateTime,
                                    finishTime As DateTime,
                                    batchOrCartId As String,
                                    kilnType As String,
                                    cycleType As String,
                                    occ As Double)

        Dim line = String.Join(",",
                Csv(_runId),
                opRec.ToString(CultureInfo.InvariantCulture),
                Csv(orderNo),
                opNo.ToString(CultureInfo.InvariantCulture),
                Csv(opName),
                Csv(resGroup),
                Csv(chosenRes),
                Csv(planDay.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)),
                Csv(startTime.ToString("yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture)),
                Csv(finishTime.ToString("yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture)),
                Csv(batchOrCartId),
                Csv(kilnType),
                Csv(cycleType),
                occ.ToString(CultureInfo.InvariantCulture)
            ) & Environment.NewLine

        File.AppendAllText(_scheduleExportPath, line)
    End Sub

    Private Function Csv(s As String) As String
        If s Is Nothing Then Return """"""
        Dim t = s.Replace("""", """""")
        Return $"""{t}"""
    End Function


    '=========================================================
    ' CSV READER (DataTable)
    '=========================================================
    Public Function ReadRoutingCsv(csvPath As String) As DataTable
        Dim dt As New DataTable("Routing")

        Using sr As New StreamReader(csvPath)
            Dim headerLine As String = sr.ReadLine()
            If String.IsNullOrWhiteSpace(headerLine) Then Throw New Exception("CSV header missing.")

            Dim headers = SplitCsvLine(headerLine)
            For Each h In headers
                dt.Columns.Add(h.Trim())
            Next

            While Not sr.EndOfStream
                Dim line As String = sr.ReadLine()
                If String.IsNullOrWhiteSpace(line) Then Continue While

                Dim values = SplitCsvLine(line)
                Dim row = dt.NewRow()
                For i = 0 To dt.Columns.Count - 1
                    row(i) = If(i < values.Count, values(i), "")
                Next
                dt.Rows.Add(row)
            End While
        End Using

        Return dt
    End Function

    ' Minimal CSV splitting (handles quoted commas and "" escaping)
    Private Function SplitCsvLine(line As String) As List(Of String)
        Dim result As New List(Of String)()
        Dim sb As New StringBuilder()
        Dim inQuotes As Boolean = False

        For i = 0 To line.Length - 1
            Dim c = line(i)

            If c = """"c Then
                If inQuotes AndAlso i + 1 < line.Length AndAlso line(i + 1) = """"c Then
                    sb.Append(""""c)
                    i += 1
                Else
                    inQuotes = Not inQuotes
                End If
            ElseIf c = ","c AndAlso Not inQuotes Then
                result.Add(sb.ToString())
                sb.Clear()
            Else
                sb.Append(c)
            End If
        Next

        result.Add(sb.ToString())
        Return result
    End Function


    '=========================================================
    ' INTERNAL MODEL
    '=========================================================
    Private Class OpInfo
        Public Property OpRec As Integer
        Public Property OrderNo As String
        Public Property OpNo As Integer
        Public Property OpName As String
        Public Property ResGroup As String

        ' Pressing attributes (primarily Op 200)
        Public Property WheelDia As String
        Public Property WheelPin As String
        Public Property PressEarliestStart As DateTime
        Public Property PressDueDate As DateTime

        ' Firing attributes (primarily Op 300)
        Public Property kilnType As String    ' "Batch" or "Tunnel"
        Public Property CycleType As String   ' "150VT","102VT","65VT" for batch; tunnel may be blank
        Public Property VolumeOcc As Double
        Public Property WeekStart As DateTime ' proxy firing day bucket

        ' Convenience: plan-day bucket used in planners
        Public Property PlanDay As DateTime
    End Class


    '=========================================================
    ' ENTRY POINT (template-based)
    '=========================================================
    Public Function Run(ByRef preactorComObject As PreactorObj, ByRef pespComObject As Object) As Integer

        InitLogging()

        Dim preactor As IPreactor = PreactorFactory.CreatePreactorObject(preactorComObject)
        Dim planningboard As IPlanningBoard = preactor.PlanningBoard

        ' ------------------------------------------------------
        ' 1) Load Routing.csv (for logic) from the Opcenter/configured path.
        ' ------------------------------------------------------
        Dim configuredRoutingPath As String =
            ConfigurationManager.AppSettings("RoutingCsvPath")
        If String.IsNullOrWhiteSpace(configuredRoutingPath) Then
            configuredRoutingPath = "Routing.csv"
        End If

        Dim csvPath As String = configuredRoutingPath
        If Not Path.IsPathRooted(csvPath) Then
            csvPath = Path.Combine(preactor.ParseShellString("{PATH}"),
                                   csvPath)
        End If

        Dim routing As DataTable = ReadRoutingCsv(csvPath)
        LogDecision("INIT", 0, "", 0, "", "", DateTime.Today, Nothing, Nothing, "", "", "", 0, 0,
                        $"Loaded CSV rows={routing.Rows.Count} from {csvPath}. Exports: {_decisionLogPath} ; {_scheduleExportPath}")

        ' ------------------------------------------------------
        ' 2) Build OpInfo map from CSV.
        '    Assumption for dry run (Option B you chose):
        '       - treat all ops as unscheduled
        '       - mixing considered scheduled (so pressing gate always passes)
        ' ------------------------------------------------------
        Dim opsByRec As Dictionary(Of Integer, OpInfo) = BuildOpInfoMap(routing)

        ' Also build a quick lookup: (OrderNo, OpNo) -> opRec
        Dim opRecByOrderAndOpNo As Dictionary(Of String, Integer) = BuildOrderOpLookup(opsByRec.Values)

        ' ------------------------------------------------------
        ' 3) PRESSING PLAN (Op 200)
        '    - Choose sequence per press resource based on changeover (dia+pin),
        '      dia cooldown per resource, and due date feasibility.
        '    - Apply schedule to Opcenter via PlanningBoard.
        ' ------------------------------------------------------
        ApplyPressingPlan(preactor, planningboard, opsByRec, opRecByOrderAndOpNo)

        ' ------------------------------------------------------
        ' 4) FIRING PLAN (Op 290 + Op 300)
        '    - Build daily batches/carts using WeekStart as proxy firing day.
        '    - Schedule Op290 in morning shift window (infinite res; time gate only).
        '    - Schedule Op300 on kiln/tunnel resources (resource groups via FindResources).
        ' ------------------------------------------------------
        ApplyFiringPlan(preactor, planningboard, opsByRec, opRecByOrderAndOpNo)

        LogDecision("DONE", 0, "", 0, "", "", DateTime.Today, Nothing, Nothing, "", "", "", 0, 0,
                        $"Completed. DecisionLog={_decisionLogPath} PlannedSchedule={_scheduleExportPath}")

        Return 0
    End Function


    '=========================================================
    ' Build OpInfo from CSV rows
    '=========================================================
    Private Function BuildOpInfoMap(routing As DataTable) As Dictionary(Of Integer, OpInfo)

        Dim map As New Dictionary(Of Integer, OpInfo)()

        For Each r As DataRow In routing.Rows

            ' The critical join key:
            ' OrdersID in CSV == Opcenter opRec (per your statement)
            Dim opRec As Integer = ParseIntSafe(CStr(r("OrdersID")), 0)
            If opRec <= 0 Then Continue For

            Dim opNo As Integer = ParseIntSafe(CStr(r("Operation Number")), 0)
            Dim orderNo As String = CStr(r("Order No"))

            Dim info As New OpInfo() With {
                    .OpRec = opRec,
                    .OrderNo = orderNo,
                    .OpNo = opNo,
                    .OpName = SafeGet(r, "Operation Name"),
                    .ResGroup = SafeGet(r, "Resource Group"),
                    .WheelDia = SafeGet(r, "Wheel Dia"),
                    .WheelPin = SafeGet(r, "Wheel thickness"),
                    .kilnType = SafeGet(r, "Kiln Type"),
                    .CycleType = SafeGet(r, "Cycle Type"),
                    .VolumeOcc = ParseDoubleSafe(SafeGet(r, "Volume Occupancy"), 0.0),
                    .WeekStart = ParseDateSafe(SafeGet(r, "Week start"), DateTime.Today),
                    .PressEarliestStart = ParseDateSafe(SafeGet(r, "pressing batch+date"), DateTime.Today),
                    .PressDueDate = ParseDateSafe(SafeGet(r, "pressing due date"), DateTime.MaxValue)
                }

            ' PlanDay defaults:
            ' - For pressing (op200): day from pressing earliest start date
            ' - For firing: day from week start (proxy)
            If info.OpNo = 200 Then
                info.PlanDay = info.PressEarliestStart.Date
            ElseIf info.OpNo = 300 Then
                info.PlanDay = info.WeekStart.Date
            Else
                info.PlanDay = DateTime.Today
            End If

            map(opRec) = info
        Next

        Return map
    End Function

    Private Function BuildOrderOpLookup(allOps As IEnumerable(Of OpInfo)) As Dictionary(Of String, Integer)
        ' Key = $"{OrderNo}||{OpNo}"
        Dim d As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
        For Each o In allOps
            If String.IsNullOrWhiteSpace(o.OrderNo) OrElse o.OpNo <= 0 Then Continue For
            Dim k = $"{o.OrderNo}||{o.OpNo}"
            If Not d.ContainsKey(k) Then
                d(k) = o.OpRec
            End If
        Next
        Return d
    End Function

    Private Function SafeGet(r As DataRow, colName As String) As String
        If r.Table.Columns.Contains(colName) Then
            Return CStr(r(colName))
        End If
        Return ""
    End Function


    '=========================================================
    ' PRESSING PLAN (Op 200)
    '=========================================================
    Private Sub ApplyPressingPlan(preactor As IPreactor,
                                      planningboard As IPlanningBoard,
                                      opsByRec As Dictionary(Of Integer, OpInfo),
                                      opRecByOrderAndOpNo As Dictionary(Of String, Integer))

        ' Gather all Op 200 candidates from CSV
        Dim pressOps = opsByRec.Values.
                Where(Function(o) o.OpNo = 200).
                OrderBy(Function(o) o.PressEarliestStart).
                ThenBy(Function(o) o.PressDueDate).
                ToList()

        If pressOps.Count = 0 Then
            LogDecision("PRESSING", 0, "", 200, "", "", DateTime.Today, Nothing, Nothing, "", "", "", 0, 0, "No Op200 rows found in CSV.")
            Exit Sub
        End If

        ' We must choose a specific press machine per operation using PlanningBoard.FindResources(opRec).
        ' Because machines are identical but changeover/cooldown is per resource, we need per-resource memory:
        '
        ' lastKeyByResDay: for changeover minimization on each resource/day
        ' lastDiaUsedDayByRes: to enforce dia cooldown per resource
        Dim lastKeyByResDay As New Dictionary(Of String, Dictionary(Of DateTime, String))(StringComparer.OrdinalIgnoreCase)
        Dim lastDiaUsedDayByRes As New Dictionary(Of String, Dictionary(Of String, DateTime))(StringComparer.OrdinalIgnoreCase)

        For Each op In pressOps
            Dim planDay As DateTime = op.PressEarliestStart.Date
            VerifyOpRecMapping(preactor, op.OpRec, 200, op.OrderNo, "PRESSING")
            ' 1) Find candidate resources for this opRec from Opcenter
            Dim resRecs As IEnumerable(Of Integer) = planningboard.FindResources(op.OpRec)

            ' We score each resource and pick best feasible:
            ' - Must be schedulable (TestOperationOnResource has value)
            ' - Must not violate due date (finish <= pressing due)
            ' - Must satisfy dia cooldown on that resource for planDay
            ' - Minimize changeover: prefer same dia+pin as previous job on that resource on that planDay
            Dim bestResRec As Integer = 0
            Dim bestScore As Double = Double.MaxValue
            Dim bestTimes As Nullable(Of Preactor.OperationTimes) = Nothing
            Dim bestReason As String = ""

            For Each resRec In resRecs

                Dim resName As String = TryGetResourceNameSafe(preactor, resRec)

                ' Dia cooldown check (per resource)
                If Not CooldownAllows(resName, op.WheelDia, planDay, lastDiaUsedDayByRes) Then
                    Continue For
                End If

                ' Enforce earliest start:
                ' We test from TerminatorTime but then we will place at max(testStart, earliestStart)
                Dim test = planningboard.TestOperationOnResource(op.OpRec, resRec, planningboard.TerminatorTime)
                If Not test.HasValue Then
                    Continue For
                End If

                Dim candidateStart As DateTime = test.Value.ChangeStart
                If candidateStart < op.PressEarliestStart Then
                    candidateStart = op.PressEarliestStart
                End If

                ' Rough finish estimate: use OperationTimes if available.
                ' (Different installations expose different fields; safest is to place at ChangeStart and trust board timing.
                ' For due-date feasibility, we approximate using ChangeEnd if present; otherwise skip due-date hard check.)
                Dim approxFinish As DateTime? = TryGetApproxFinish(test.Value)
                If approxFinish.HasValue Then
                    ' Shift finish if we shifted start after ChangeStart
                    Dim delta = candidateStart - test.Value.ChangeStart
                    approxFinish = approxFinish.Value.Add(delta)

                    If op.PressDueDate < DateTime.MaxValue AndAlso approxFinish.Value > op.PressDueDate Then
                        Continue For
                    End If
                End If

                ' Changeover score: 0 if same (dia+pin) as last key on that resource/day, else 1
                Dim key As String = $"{op.WheelDia}||{op.WheelPin}"
                Dim lastKey As String = GetLastKey(lastKeyByResDay, resName, planDay)
                Dim changePenalty As Double = If(String.Equals(lastKey, key, StringComparison.OrdinalIgnoreCase), 0.0, 1.0)

                ' Due-date urgency as a small tie-breaker (lower due date => slightly better)
                Dim dueTie As Double = If(op.PressDueDate < DateTime.MaxValue, op.PressDueDate.Subtract(DateTime.Today).TotalDays, 99999)

                ' Composite score: prioritize changeover, then due date
                Dim score As Double = changePenalty * 1000.0 + dueTie

                If score < bestScore Then
                    bestScore = score
                    bestResRec = resRec
                    bestTimes = test
                    bestReason = $"Selected by score. changePenalty={changePenalty}, lastKey={lastKey}"
                End If
            Next

            If bestResRec <= 0 OrElse Not bestTimes.HasValue Then
                LogDecision("PRESSING_SKIP", op.OpRec, op.OrderNo, op.OpNo, op.ResGroup, "", planDay, Nothing, Nothing, "", "", "", op.VolumeOcc, 0,
                                "No feasible press resource (cooldown/due/availability).")
                Continue For
            End If

            ' 2) Place the operation on the chosen resource
            Dim chosenResName As String = TryGetResourceNameSafe(preactor, bestResRec)

            Dim startTime As DateTime = bestTimes.Value.ChangeStart
            If startTime < op.PressEarliestStart Then startTime = op.PressEarliestStart

            planningboard.PutOperationOnResource(op.OpRec, bestResRec, startTime)

            ' 3) Update per-resource memories (cooldown + last key for changeover)
            Dim keyNow As String = $"{op.WheelDia}||{op.WheelPin}"
            SetLastKey(lastKeyByResDay, chosenResName, planDay, keyNow)
            MarkDiaUsed(lastDiaUsedDayByRes, chosenResName, op.WheelDia, planDay)

            ' 4) Log + export
            Dim finishApprox As DateTime? = TryGetApproxFinish(bestTimes.Value)
            If finishApprox.HasValue Then
                Dim delta = startTime - bestTimes.Value.ChangeStart
                finishApprox = finishApprox.Value.Add(delta)
            End If

            LogDecision("PRESSING_PUT", op.OpRec, op.OrderNo, op.OpNo, op.ResGroup, chosenResName, planDay,
                            startTime, finishApprox, "", "", "", op.VolumeOcc, bestScore, bestReason)

            ExportPlannedRow(op.OpRec, op.OrderNo, op.OpNo, op.OpName, op.ResGroup, chosenResName, planDay,
                                 startTime, If(finishApprox.HasValue, finishApprox.Value, startTime), "", "", "", op.VolumeOcc)

            ' 5) OPTIONAL: forward schedule associated drying ops (200->280) until just before loading (290)
            '     We use planningboard.GetNextOperation(opRec, 1) exactly like your template,
            '     but stop when opNo reaches 290 to avoid bypassing firing batching logic.
            ForwardScheduleAssociatedOps(preactor, planningboard, opsByRec, op.OpRec, stopAtOpNo:=290)
        Next
    End Sub


    '=========================================================
    ' FIRING PLAN (Op 290 + Op 300)
    '=========================================================
    Private Sub ApplyFiringPlan(preactor As IPreactor,
                                    planningboard As IPlanningBoard,
                                    opsByRec As Dictionary(Of Integer, OpInfo),
                                    opRecByOrderAndOpNo As Dictionary(Of String, Integer))

        ' Separate Op300 candidates by kiln type, grouped by PlanDay (WeekStart proxy)
        Dim op300 = opsByRec.Values.Where(Function(o) o.OpNo = 300).ToList()
        If op300.Count = 0 Then
            LogDecision("FIRING", 0, "", 300, "", "", DateTime.Today, Nothing, Nothing, "", "", "", 0, 0, "No Op300 rows found in CSV.")
            Exit Sub
        End If

        Dim batchOps = op300.Where(Function(o) o.kilnType.Equals("Batch", StringComparison.OrdinalIgnoreCase)).ToList()
        Dim tunnelOps = op300.Where(Function(o) o.kilnType.Equals("Tunnel", StringComparison.OrdinalIgnoreCase)).ToList()

        ' --- BATCH KILN (pure first: 150VT > 102VT > 65VT; then mixed adjacency)
        Dim batchByDay = batchOps.GroupBy(Function(o) o.PlanDay).OrderBy(Function(g) g.Key).ToList()
        For Each dayGroup In batchByDay
            Dim day = dayGroup.Key
            Dim pool = dayGroup.ToList()

            ' Sort small occupancy first (your packing preference), but we will still respect due dates later.
            pool = pool.OrderBy(Function(o) o.VolumeOcc).ToList()

            ' 1) Pure batches
            CreatePureBatches(day, pool, "150VT", preactor, planningboard, opsByRec, opRecByOrderAndOpNo)
            CreatePureBatches(day, pool, "102VT", preactor, planningboard, opsByRec, opRecByOrderAndOpNo)
            CreatePureBatches(day, pool, "65VT", preactor, planningboard, opsByRec, opRecByOrderAndOpNo)

            ' 2) Mixed batches adjacency
            ' 150 + 102 (time=150)
            CreateMixedBatches(day, pool, highCycle:="150VT", lowCycle:="102VT",
                                   preactor:=preactor, planningboard:=planningboard,
                                   opsByRec:=opsByRec, opRecByOrderAndOpNo:=opRecByOrderAndOpNo)

            ' 102 + 65 (time=102)
            CreateMixedBatches(day, pool, highCycle:="102VT", lowCycle:="65VT",
                                   preactor:=preactor, planningboard:=planningboard,
                                   opsByRec:=opsByRec, opRecByOrderAndOpNo:=opRecByOrderAndOpNo)

            ' 150 + 65 forbidden by your rule -> intentionally not attempted
        Next

        ' --- TUNNEL (pack into carts/day, default 4 cars/day)
        Dim tunnelByDay = tunnelOps.GroupBy(Function(o) o.PlanDay).OrderBy(Function(g) g.Key).ToList()
        For Each dayGroup In tunnelByDay
            Dim day = dayGroup.Key
            Dim pool = dayGroup.OrderBy(Function(o) o.VolumeOcc).ToList()

            Dim car = 1
            Dim idx = 0

            While car <= TunnelCarsPerDayDefault AndAlso idx < pool.Count
                Dim cartId = $"CART_{day:yyyyMMdd}_{car}"
                Dim sumOcc As Double = 0
                Dim chosen As New List(Of OpInfo)()

                ' Greedy pack: small-first until max
                While idx < pool.Count AndAlso sumOcc + pool(idx).VolumeOcc <= MaxOcc
                    chosen.Add(pool(idx))
                    sumOcc += pool(idx).VolumeOcc
                    idx += 1
                End While

                If sumOcc < MinOcc Then Exit While

                ' Apply: schedule Op290 (morning) then Op300 (tunnel)
                For Each o In chosen
                    ApplyFiringForOneOrder(day, o, cartId, cycleType:="TUNNEL",
                                               preactor:=preactor, planningboard:=planningboard,
                                               opsByRec:=opsByRec, opRecByOrderAndOpNo:=opRecByOrderAndOpNo)
                Next

                car += 1
            End While
        Next

    End Sub


    '=========================================================
    ' BATCH BUILDERS
    '=========================================================
    Private Sub CreatePureBatches(day As DateTime,
                                      pool As List(Of OpInfo),
                                      cycle As String,
                                      preactor As IPreactor,
                                      planningboard As IPlanningBoard,
                                      opsByRec As Dictionary(Of Integer, OpInfo),
                                      opRecByOrderAndOpNo As Dictionary(Of String, Integer))

        Dim candidates = pool.Where(Function(o) o.CycleType.Equals(cycle, StringComparison.OrdinalIgnoreCase)).OrderBy(Function(o) o.VolumeOcc).ToList()
        If candidates.Count = 0 Then Exit Sub

        Dim idx As Integer = 1
        Dim cursor As Integer = 0

        While cursor < candidates.Count
            Dim batchId = $"PURE_{cycle}_{day:yyyyMMdd}_{idx}"
            idx += 1

            Dim sumOcc As Double = 0
            Dim chosen As New List(Of OpInfo)()

            While cursor < candidates.Count AndAlso sumOcc + candidates(cursor).VolumeOcc <= MaxOcc
                chosen.Add(candidates(cursor))
                sumOcc += candidates(cursor).VolumeOcc
                cursor += 1
            End While

            If sumOcc < MinOcc Then Exit While

            For Each o In chosen
                ApplyFiringForOneOrder(day, o, batchId, cycleType:=cycle,
                                           preactor:=preactor, planningboard:=planningboard,
                                           opsByRec:=opsByRec, opRecByOrderAndOpNo:=opRecByOrderAndOpNo)
            Next
        End While
    End Sub


    Private Sub CreateMixedBatches(day As DateTime,
                                       pool As List(Of OpInfo),
                                       highCycle As String,
                                       lowCycle As String,
                                       preactor As IPreactor,
                                       planningboard As IPlanningBoard,
                                       opsByRec As Dictionary(Of Integer, OpInfo),
                                       opRecByOrderAndOpNo As Dictionary(Of String, Integer))

        Dim highs = pool.Where(Function(o) o.CycleType.Equals(highCycle, StringComparison.OrdinalIgnoreCase)).OrderBy(Function(o) o.VolumeOcc).ToList()
        Dim lows = pool.Where(Function(o) o.CycleType.Equals(lowCycle, StringComparison.OrdinalIgnoreCase)).OrderBy(Function(o) o.VolumeOcc).ToList()
        If highs.Count = 0 OrElse lows.Count = 0 Then Exit Sub

        Dim i As Integer = 0
        Dim j As Integer = 0
        Dim batchIdx As Integer = 1

        While i < highs.Count OrElse j < lows.Count

            Dim batchId = $"MIX_{highCycle}_{lowCycle}_{day:yyyyMMdd}_{batchIdx}"
            batchIdx += 1

            Dim sumOcc As Double = 0
            Dim chosen As New List(Of OpInfo)()

            ' Always include at least one "high" if possible (to ensure the batch time is high cycle)
            If i < highs.Count AndAlso sumOcc + highs(i).VolumeOcc <= MaxOcc Then
                chosen.Add(highs(i))
                sumOcc += highs(i).VolumeOcc
                i += 1
            End If

            ' Fill with lows as possible
            While j < lows.Count AndAlso sumOcc + lows(j).VolumeOcc <= MaxOcc
                chosen.Add(lows(j))
                sumOcc += lows(j).VolumeOcc
                j += 1
            End While

            If sumOcc < MinOcc Then Exit While

            ' Mixed batch runs at high cycle time (your rule)
            For Each o In chosen
                ApplyFiringForOneOrder(day, o, batchId, cycleType:=highCycle,
                                           preactor:=preactor, planningboard:=planningboard,
                                           opsByRec:=opsByRec, opRecByOrderAndOpNo:=opRecByOrderAndOpNo)
            Next
        End While
    End Sub


    '=========================================================
    ' Apply firing for one Op300 row:
    '   - schedule Op290 in morning window
    '   - schedule Op300 on kiln/tunnel resource
    '=========================================================
    Private Sub ApplyFiringForOneOrder(day As DateTime,
                                           op300Info As OpInfo,
                                           batchOrCartId As String,
                                           cycleType As String,
                                           preactor As IPreactor,
                                           planningboard As IPlanningBoard,
                                           opsByRec As Dictionary(Of Integer, OpInfo),
                                           opRecByOrderAndOpNo As Dictionary(Of String, Integer))

        ' 1) Schedule Op290 (loading) in morning shift (08:00–14:00), infinite resource.
        '    We still place it on the PlanningBoard so it respects time window.
        Dim op290Rec As Integer = FindOpRec(opRecByOrderAndOpNo, op300Info.OrderNo, 290)
        If op290Rec > 0 Then
            Dim op290Start As DateTime = day.Add(MorningShiftStart)

            VerifyOpRecMapping(preactor, op290Rec, 290, op300Info.OrderNo, "FIRING_290")
            ' For infinite resource, FindResources should return something; we just take first feasible.
            Dim resRecs = planningboard.FindResources(op290Rec)
            Dim placed290 As Boolean = False

            For Each resRec In resRecs
                Dim test = planningboard.TestOperationOnResource(op290Rec, resRec, planningboard.TerminatorTime)
                If Not test.HasValue Then Continue For

                Dim startTime = test.Value.ChangeStart
                If startTime < op290Start Then startTime = op290Start

                ' Ensure start in window (if not, skip)
                If startTime.TimeOfDay > MorningShiftEnd Then Continue For

                planningboard.PutOperationOnResource(op290Rec, resRec, startTime)

                Dim resName = TryGetResourceNameSafe(preactor, resRec)
                LogDecision("FIRING_PUT290", op290Rec, op300Info.OrderNo, 290, "(LOAD)", resName, day,
                                startTime, TryGetApproxFinish(test.Value), batchOrCartId, op300Info.kilnType, cycleType, op300Info.VolumeOcc, 0,
                                "Placed Op290 in morning window.")
                placed290 = True
                Exit For
            Next

            If Not placed290 Then
                LogDecision("FIRING_SKIP290", op290Rec, op300Info.OrderNo, 290, "(LOAD)", "", day,
                                Nothing, Nothing, batchOrCartId, op300Info.kilnType, cycleType, op300Info.VolumeOcc, 0,
                                "Could not place Op290 within morning window.")
            End If
        Else
            LogDecision("FIRING_NO290", 0, op300Info.OrderNo, 290, "(LOAD)", "", day,
                            Nothing, Nothing, batchOrCartId, op300Info.kilnType, cycleType, op300Info.VolumeOcc, 0,
                            "No Op290 found for this order in CSV lookup.")
        End If
        VerifyOpRecMapping(preactor, op300Info.OpRec, 300, op300Info.OrderNo, "FIRING_300")
        ' 2) Schedule Op300 on kiln/tunnel resources
        Dim op300Rec As Integer = op300Info.OpRec
        Dim resRecs300 = planningboard.FindResources(op300Rec)

        Dim placed300 As Boolean = False
        For Each resRec In resRecs300
            Dim test = planningboard.TestOperationOnResource(op300Rec, resRec, planningboard.TerminatorTime)
            If Not test.HasValue Then Continue For

            ' Place at earliest feasible start (Opcenter will account for precedence if configured)
            Dim startTime = test.Value.ChangeStart
            planningboard.PutOperationOnResource(op300Rec, resRec, startTime)

            Dim resName = TryGetResourceNameSafe(preactor, resRec)

            LogDecision("FIRING_PUT300", op300Rec, op300Info.OrderNo, 300, op300Info.ResGroup, resName, day,
                            startTime, TryGetApproxFinish(test.Value), batchOrCartId, op300Info.kilnType, cycleType, op300Info.VolumeOcc, 0,
                            "Placed Op300 based on batch/cart assignment.")

            ' Export planned row (use test-based finish approximation when available)
            Dim finishApprox = TryGetApproxFinish(test.Value)
            If Not finishApprox.HasValue Then finishApprox = startTime

            ExportPlannedRow(op300Rec, op300Info.OrderNo, 300, op300Info.OpName, op300Info.ResGroup, resName, day,
                                 startTime, finishApprox.Value, batchOrCartId, op300Info.kilnType, cycleType, op300Info.VolumeOcc)

            placed300 = True
            Exit For ' V1: take first feasible resource
        Next

        If Not placed300 Then
            LogDecision("FIRING_SKIP300", op300Rec, op300Info.OrderNo, 300, op300Info.ResGroup, "", day,
                            Nothing, Nothing, batchOrCartId, op300Info.kilnType, cycleType, op300Info.VolumeOcc, 0,
                            "Could not place Op300 on any resource.")
        End If

        ' 3) Optional: forward schedule unload and other ops up to 320 (as you requested)
        '    We schedule forward from op300Rec until OpNo > 320 to keep V1 bounded.
        ForwardScheduleAssociatedOps(preactor, planningboard, opsByRec, op300Rec, stopAtOpNo:=321)
    End Sub


    '=========================================================
    ' Forward schedule associated ops (template-like)
    '=========================================================
    Private Sub ForwardScheduleAssociatedOps(preactor As IPreactor,
                                                 planningboard As IPlanningBoard,
                                                 opsByRec As Dictionary(Of Integer, OpInfo),
                                                 startOpRec As Integer,
                                                 stopAtOpNo As Integer)

        ' This method matches your template pattern:
        '   opRec = planningboard.GetNextOperation(opRec, 1)
        ' and schedules each next op on the first feasible resource.
        ' We stop once we reach stopAtOpNo or cannot find op info.

        Dim opRec As Integer = startOpRec

        While opRec > 0
            Dim nextRec As Integer = planningboard.GetNextOperation(opRec, 1)
            If nextRec <= 0 Then Exit While

            If Not opsByRec.ContainsKey(nextRec) Then
                ' If CSV didn't contain this op record, we can't classify it;
                ' still, we could attempt to schedule it generically, but for V1 we stop.
                Exit While
            End If

            Dim info = opsByRec(nextRec)
            If info.OpNo >= stopAtOpNo Then Exit While

            Dim resRecs = planningboard.FindResources(nextRec)
            For Each resRec In resRecs
                Dim test = planningboard.TestOperationOnResource(nextRec, resRec, planningboard.TerminatorTime)
                If Not test.HasValue Then Continue For

                planningboard.PutOperationOnResource(nextRec, resRec, test.Value.ChangeStart)

                Dim resName = TryGetResourceNameSafe(preactor, resRec)
                LogDecision("FWD_PUT", nextRec, info.OrderNo, info.OpNo, info.ResGroup, resName, DateTime.Today,
                                test.Value.ChangeStart, TryGetApproxFinish(test.Value), "", info.kilnType, info.CycleType, info.VolumeOcc, 0,
                                "Forward-scheduled associated operation.")
                Exit For
                Next

                opRec = nextRec
            End While

        End Sub


        '=========================================================
        ' Cooldown helpers (per resource)
        '=========================================================
        Private Function CooldownAllows(resName As String,
                                        wheelDia As String,
                                        day As DateTime,
                                        lastDiaUsedDayByRes As Dictionary(Of String, Dictionary(Of String, DateTime))) As Boolean

            If String.IsNullOrWhiteSpace(resName) OrElse String.IsNullOrWhiteSpace(wheelDia) Then Return True

            If Not lastDiaUsedDayByRes.ContainsKey(resName) Then Return True
            Dim d = lastDiaUsedDayByRes(resName)
            If Not d.ContainsKey(wheelDia) Then Return True

            Dim lastUsed = d(wheelDia).Date
            Dim diffDays = (day.Date - lastUsed).TotalDays

            ' If used today => diff=0 (block), used yesterday => diff=1 (block), used 2 days ago => diff=2 (block)
            ' Allowed again when diff >= 3 (day+3), because cooldown is 2 days after use.
            Return diffDays >= (DiaCooldownDays + 1)
        End Function

        Private Sub MarkDiaUsed(lastDiaUsedDayByRes As Dictionary(Of String, Dictionary(Of String, DateTime)),
                                resName As String,
                                wheelDia As String,
                                day As DateTime)

            If String.IsNullOrWhiteSpace(resName) OrElse String.IsNullOrWhiteSpace(wheelDia) Then Exit Sub

            If Not lastDiaUsedDayByRes.ContainsKey(resName) Then
                lastDiaUsedDayByRes(resName) = New Dictionary(Of String, DateTime)(StringComparer.OrdinalIgnoreCase)
            End If

            lastDiaUsedDayByRes(resName)(wheelDia) = day.Date
        End Sub

        Private Function GetLastKey(lastKeyByResDay As Dictionary(Of String, Dictionary(Of DateTime, String)),
                                    resName As String,
                                    day As DateTime) As String

            If Not lastKeyByResDay.ContainsKey(resName) Then Return ""
            Dim d = lastKeyByResDay(resName)
            If Not d.ContainsKey(day.Date) Then Return ""
            Return d(day.Date)
        End Function

        Private Sub SetLastKey(lastKeyByResDay As Dictionary(Of String, Dictionary(Of DateTime, String)),
                               resName As String,
                               day As DateTime,
                               key As String)

            If Not lastKeyByResDay.ContainsKey(resName) Then
                lastKeyByResDay(resName) = New Dictionary(Of DateTime, String)()
            End If
            lastKeyByResDay(resName)(day.Date) = key
        End Sub


        '=========================================================
        ' Small utilities
        '=========================================================
        Private Function FindOpRec(lookup As Dictionary(Of String, Integer), orderNo As String, opNo As Integer) As Integer
            Dim k = $"{orderNo}||{opNo}"
            If lookup.ContainsKey(k) Then Return lookup(k)
            Return 0
        End Function

        Private Function TryGetResourceNameSafe(preactor As IPreactor, resRec As Integer) As String
            ' Depending on your model, you may have a field name for resource names.
            ' If not available, return the resRec as text.
            Try
                ' Many models use "Resource Name" or similar; but to avoid undocumented calls,
                ' we keep this safe and non-failing.
                Return $"ResRec_{resRec}"
            Catch
                Return $"ResRec_{resRec}"
            End Try
        End Function

        Private Function TryGetApproxFinish(opTimes As Preactor.OperationTimes) As DateTime?
            ' OperationTimes members can vary per environment; ChangeEnd exists in many installs.
            ' If missing, we return Nothing and caller will log without finish time.
            Try
            Return opTimes.ProcessStart
        Catch
                Return Nothing
            End Try
        End Function

        Private Function ParseDateSafe(text As String, fallback As DateTime) As DateTime
            If String.IsNullOrWhiteSpace(text) Then Return fallback
            Dim dt As DateTime
            Dim formats = New String() {"dd-MM-yyyy", "dd/MM/yyyy", "yyyy-MM-dd", "yyyy-MM-dd HH:mm", "dd-MM-yyyy HH:mm", "dd/MM/yyyy HH:mm"}
            If DateTime.TryParseExact(text.Trim(), formats, CultureInfo.InvariantCulture, DateTimeStyles.None, dt) Then Return dt
            If DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.None, dt) Then Return dt
            Return fallback
        End Function

        Private Function ParseDoubleSafe(text As String, fallback As Double) As Double
            If String.IsNullOrWhiteSpace(text) Then Return fallback
            Dim x As Double
            If Double.TryParse(text.Trim(), NumberStyles.Any, CultureInfo.InvariantCulture, x) Then Return x
            If Double.TryParse(text.Trim(), NumberStyles.Any, CultureInfo.CurrentCulture, x) Then Return x
            Return fallback
        End Function

        Private Function ParseIntSafe(text As String, fallback As Integer) As Integer
            If String.IsNullOrWhiteSpace(text) Then Return fallback
            Dim x As Integer
            If Integer.TryParse(text.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, x) Then Return x
            Return fallback
        End Function
    '---------------------------------------------
    ' Verifies that an opRec points to the expected operation number in Opcenter.
    '---------------------------------------------
    Private Sub VerifyOpRecMapping(preactor As IPreactor,
                               opRec As Integer,
                               expectedOpNo As Integer,
                               csvOrderNo As String,
                               phase As String)

        ' IMPORTANT: Replace "Operations" with your actual operations table name
        ' if it's different in your dataset (common names: "Operations", "Order Operations")
        Dim opsTable As String = "Orders"

        If opRec <= 0 Then
            Debug.WriteLine($"[VERIFY][{phase}] opRec invalid: {opRec}")
            Exit Sub
        End If

        ' RecordCount check avoids exceptions for invalid opRec
        Dim rc As Integer
        Try
            rc = preactor.RecordCount(opsTable)
        Catch ex As Exception
            Debug.WriteLine($"[VERIFY][{phase}] Could not access table '{opsTable}'. Exception: {ex.Message}")
            Exit Sub
        End Try

        If opRec > rc Then
            Debug.WriteLine($"[VERIFY][{phase}] opRec={opRec} > RecordCount({opsTable})={rc}. CSV OrderNo={csvOrderNo}")
            Exit Sub
        End If

        ' Read operation number from Opcenter
        ' Column name MUST be exactly your standard: "Operation Number"
        Dim actualOpNo As Integer = 0
        Try
            actualOpNo = preactor.ReadFieldInt(opsTable, "Op. No.", opRec)
        Catch ex As Exception
            Debug.WriteLine($"[VERIFY][{phase}] Failed ReadFieldInt({opsTable}, 'Operation Number', {opRec}). Exception: {ex.Message}")
            Exit Sub
        End Try

        ' Optional: read Order No from operations to cross-check (if the field exists)
        Dim actualOrderNo As String = ""
        Try
            actualOrderNo = preactor.ReadFieldString(opsTable, "Order No", opRec)
        Catch
            ' field may not exist; ignore
        End Try

        If actualOpNo <> expectedOpNo Then
            Debug.WriteLine($"[VERIFY][{phase}] MISMATCH: opRec={opRec} CSV OrderNo={csvOrderNo} OpNoExpected={expectedOpNo} OpNoActual={actualOpNo} OpOrderNo={actualOrderNo}")
        Else
            Debug.WriteLine($"[VERIFY][{phase}] OK: opRec={opRec} CSV OrderNo={csvOrderNo} OpNo={actualOpNo} OpOrderNo={actualOrderNo}")
        End If
    End Sub


    '---------------------------------------------
    ' (Optional) helper to dump a few key fields for the opRec to Output window.
    ' Use this once if you need more visibility.
    '---------------------------------------------
    Private Sub DumpOpRec(preactor As IPreactor, opRec As Integer, phase As String)
        Dim opsTable As String = "Orders"
        Try
            Dim opNo = preactor.ReadFieldInt(opsTable, "Op. No.", opRec)
            Dim ordNo As String = ""
            Try : ordNo = preactor.ReadFieldString(opsTable, "Order No", opRec) : Catch : End Try

            Debug.WriteLine($"[DUMP][{phase}] opRec={opRec} OrderNo={ordNo} OpNo={opNo}")
        Catch ex As Exception
            Debug.WriteLine($"[DUMP][{phase}] Failed to dump opRec={opRec}: {ex.Message}")
        End Try
    End Sub

End Class
