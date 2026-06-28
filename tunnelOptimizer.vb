Option Strict On
Option Explicit On

Imports System
Imports System.Collections.Generic
Imports System.Data
Imports System.Globalization
Imports System.Runtime.InteropServices

' NOTE:
' - This optimizer DOES NOT call PutOperationOnResource itself.
' - It returns a plan you can apply in your main Run() function.
' - Column meanings are exactly as you described:
'   OrdersId = operation record (opRec)
'   firing due date = due date for tunnel firing
'   scheduled_end_time = scheduled end time 
'   Volume Occupancy = cart occupancy share
'   is_scheduled = schedule flag

<ComVisible(True)>
<Microsoft.VisualBasic.ComClass("b9d2b6f9-6cf8-44d6-95e6-7a6c7a5b1d11", "9c3b3b67-2c02-4d13-a9d2-3d6b4b11d2aa")>
Public Class tunnelOptimizer_vf

    ' -----------------------------
    ' DataTable column names (your export)
    ' -----------------------------
    Private Const COL_ORDERNO As String = "Order No"
    Private Const COL_OPREC As String = "OrdersID"
    Private Const COL_OPNO As String = "Operation Number"
    Private Const COL_KILNTYPE As String = "Kiln Type"          ' Tunnel orders: "2" (per your rule)
    Private Const COL_IS_SCHEDULED As String = "is_scheduled"
    Private Const COL_SCHED_END As String = "scheduled_end_time"
    Private Const COL_FIRING_DUE As String = "firing due date"
    Private Const COL_OCC As String = "Volume Occupancy"

    ' -----------------------------
    ' Plan objects returned
    ' -----------------------------
    Public Class TunnelPlan
        ' Global pacing parameters derived from carts/day and total carts
        Public Property CartPitch As TimeSpan
        Public Property TunnelDuration As TimeSpan

        ' Timeline of carts (each cart is a packing container)
        Public Property Carts As New List(Of CartSlot)

        ' Per firing operation record schedule (op 300 opRec -> start/end)
        Public Property StartByFiringOpRec As New Dictionary(Of Integer, DateTime)
        Public Property EndByFiringOpRec As New Dictionary(Of Integer, DateTime)

        ' Which cart each firing opRec landed into
        Public Property CartNoByFiringOpRec As New Dictionary(Of Integer, Integer)

        ' Per cart number -> schedule window + resource
        Public Property CartStartByCartNo As New Dictionary(Of Integer, DateTime)
        Public Property CartEndByCartNo As New Dictionary(Of Integer, DateTime)
        Public Property ResourceByCartNo As New Dictionary(Of Integer, String)
        Public Property OccSumByCartNo As New Dictionary(Of Integer, Double)

    End Class

    Public Class CartSlot
        Public Property CartNo As Integer
        Public Property CartStart As DateTime
        Public Property CartEnd As DateTime
        Public Property OccSum As Double

        ' For traceability
        Public Property Orders As New List(Of String)
        Public Property FiringOpRecs As New List(Of Integer)
        Public Property DueTimes As New List(Of DateTime)
        Public Property ReadyTimes As New List(Of DateTime)
    End Class

    ' Internal candidate built from DataTable
    Private Class TunnelCandidate
        Public Property OrderNo As String
        Public Property FiringOpRec As Integer   ' op 300 record id (OrdersID)
        Public Property ReadyTime As DateTime
        Public Property DueTime As DateTime
        Public Property Occ As Double
        Public Property WipScore As Integer
        Public Property WipRejectReason As String
    End Class

    ' -----------------------------
    ' Public entry point
    ' -----------------------------
    ' dt: your schedule-export DataTable at the “ops until 290 scheduled” stage
    ' startTime: baseline time to start creating carts (often Now or earliest planning horizon)
    ' cartsPerDay: ex 4.5 -> pitch = 24/4.5 hours (exact)
    ' totalCarts: ex 18 -> tunnel duration = totalCarts * pitch
    ' minOccPreferred: soft target (preference)
    ' maxOcc: hard cap
    ' dryingToFiringBufferHours: default 6 hours
    Public Function BuildTunnelPlan(dt As DataTable,
                                    startTime As DateTime,
                                    cartsPerDay As Double,
                                    totalCarts As Integer,
                                    minOccPreferred As Double,
                                    maxOcc As Double,
                                    Optional dryingToFiringBufferHours As Double = 6.0,
                                    Optional debug As SchedulerDebugCollector = Nothing) As TunnelPlan

        ValidateInputs(dt, cartsPerDay, totalCarts, minOccPreferred, maxOcc)

        ' 1) Derive pitch and tunnel duration from carts/day and total carts
        Dim pitchMinutes As Double = 1440.0 / cartsPerDay          ' 24h * 60
        Dim pitch As TimeSpan = TimeSpan.FromMinutes(pitchMinutes)

        ' Tunnel duration = total carts * pitch (exact)
        Dim tunnelDurTicks As Long = pitch.Ticks * CLng(totalCarts)
        Dim tunnelDuration As TimeSpan = TimeSpan.FromTicks(tunnelDurTicks)

        ' 2) Build candidates (op300 unscheduled, kiln type 2) AND compute readiness from ops<290
        Dim candidates As List(Of TunnelCandidate) =
            BuildCandidates(dt, dryingToFiringBufferHours, debug)

        ' Nothing to do
        Dim plan As New TunnelPlan With {
            .CartPitch = pitch,
            .TunnelDuration = tunnelDuration
        }
        If candidates.Count = 0 Then Return plan

        ' Deterministic ordering of remaining items (we will still filter by readiness per cart start)
        ' Primary: earliest due date. Tie: earliest ready. Tie: bigger occupancy first (packs faster).
        candidates.Sort(AddressOf CompareByDueReadyOcc)

        ' 3) Timeline-driven cart creation
        Dim remaining As New List(Of TunnelCandidate)(candidates)

        Dim nextCartStart As DateTime = startTime
        Dim cartNo As Integer = 0

        While remaining.Count > 0

            ' Build the ready pool and, when necessary, the earliest future
            ' readiness group in one order-preserving pass.
            Dim earliestReady As DateTime
            Dim advancedToEarliest As Boolean
            Dim readyNow As List(Of TunnelCandidate) =
                GetReadyCandidates(remaining,
                                   nextCartStart,
                                   earliestReady,
                                   advancedToEarliest)

            If advancedToEarliest Then
                nextCartStart = earliestReady
            End If

            ' Safety: if still empty, stop (prevents infinite loop)
            If readyNow.Count = 0 Then Exit While

            ' 4) Pack a cart from ready orders by EDD (Rule 5), respecting maxOcc (Rule 2)
            Dim chosen As List(Of TunnelCandidate) = PackCart(readyNow, minOccPreferred, maxOcc)

            ' Must place at least one order on a cart to advance.
            ' If the "best" order alone exceeds maxOcc, we skip it (data issue) to avoid deadlock.
            If chosen.Count = 0 Then
                ' Remove one problematic candidate deterministically (earliest due among readyNow)
                Dim bad As TunnelCandidate = readyNow(0)
                remaining.Remove(bad)
                Continue While
            End If

            cartNo += 1
            Dim cartStart As DateTime = nextCartStart
            Dim cartEnd As DateTime = cartStart.Add(tunnelDuration)
            Dim chosenSet As New HashSet(Of TunnelCandidate)(chosen)

            ' Commit cart slot
            Dim slot As New CartSlot With {
                .CartNo = cartNo,
                .CartStart = cartStart,
                .CartEnd = cartEnd,
                .OccSum = 0.0
            }

            For Each o In chosen
                slot.Orders.Add(o.OrderNo)
                slot.FiringOpRecs.Add(o.FiringOpRec)
                slot.DueTimes.Add(o.DueTime)
                slot.ReadyTimes.Add(o.ReadyTime)
                slot.OccSum += o.Occ

                plan.StartByFiringOpRec(o.FiringOpRec) = cartStart
                plan.EndByFiringOpRec(o.FiringOpRec) = cartEnd
                plan.CartNoByFiringOpRec(o.FiringOpRec) = cartNo
            Next

            remaining.RemoveAll(Function(candidate) chosenSet.Contains(candidate))

            plan.Carts.Add(slot)

            ' 5) Next cart start follows exact pacing rule:
            ' "NextCartStart = ActualCartStart + Pitch"
            nextCartStart = cartStart.Add(pitch)

        End While

        Return plan
    End Function

    Private Function GetReadyCandidates(remaining As List(Of TunnelCandidate),
                                        currentStart As DateTime,
                                        ByRef earliestReady As DateTime,
                                        ByRef advancedToEarliest As Boolean) As List(Of TunnelCandidate)

        Dim readyNow As New List(Of TunnelCandidate)()
        Dim earliestCandidates As New List(Of TunnelCandidate)()
        earliestReady = DateTime.MaxValue
        advancedToEarliest = False

        For Each candidate As TunnelCandidate In remaining
            If candidate.ReadyTime <= currentStart Then
                readyNow.Add(candidate)
            End If

            If candidate.ReadyTime < earliestReady Then
                earliestReady = candidate.ReadyTime
                earliestCandidates.Clear()
                earliestCandidates.Add(candidate)
            ElseIf candidate.ReadyTime = earliestReady Then
                earliestCandidates.Add(candidate)
            End If
        Next

        If readyNow.Count = 0 AndAlso earliestReady <> DateTime.MaxValue Then
            readyNow.AddRange(earliestCandidates)
            advancedToEarliest = True
        End If

        Return readyNow
    End Function

    ' -----------------------------
    ' Candidate building
    ' -----------------------------
    Private Function BuildCandidates(dt As DataTable,
                                     dryingToFiringBufferHours As Double,
                                     Optional debug As SchedulerDebugCollector = Nothing) As List(Of TunnelCandidate)

        ' readiness per order = max scheduled_end_time among scheduled ops with opNo < 290
        'Dim readyByOrder As New Dictionary(Of String, DateTime)(StringComparer.OrdinalIgnoreCase)

        'For Each r As DataRow In dt.Rows
        '    If Not SafeBool(r(COL_IS_SCHEDULED)) Then Continue For

        '    ' Area that I did not add the code for WIP. 

        '    Dim opNo As Integer = SafeInt(r(COL_OPNO))
        '    If opNo >= 290 Then Continue For

        '    Dim orderNo As String = SafeStr(r(COL_ORDERNO)).Trim()
        '    If orderNo = "" Then Continue For

        '    Dim endT As DateTime = SafeDate(r(COL_SCHED_END))
        '    If endT = DateTime.MinValue Then Continue For

        '    If (Not readyByOrder.ContainsKey(orderNo)) OrElse endT > readyByOrder(orderNo) Then
        '        readyByOrder(orderNo) = endT
        '    End If
        'Next
        Dim readinessByOrder As Dictionary(Of String, SharedHelpers.FiringReadinessInfo) =
    SharedHelpers.BuildFiringReadinessByOrder(dt)


        ' Build candidates from op 300 rows (unscheduled) for tunnel kiln type = 2
        Dim list As New List(Of TunnelCandidate)()
        Dim beforeCount As Integer = 0
        If debug IsNot Nothing AndAlso debug.Enabled Then
            beforeCount = dt.AsEnumerable().Count(
                Function(row) SafeInt(row(COL_OPNO)) = 300)
        End If

        For Each r As DataRow In dt.Rows

            Dim opNo As Integer = SafeInt(r(COL_OPNO))
            If opNo <> 300 Then Continue For

            Dim kilnType As String = SafeStr(r(COL_KILNTYPE)).Trim()
            If Not kilnType.Equals("2", StringComparison.OrdinalIgnoreCase) Then
                TraceCandidateDecision(debug, r, beforeCount, list.Count, False,
                                       SchedulerDebugReasonCodes.FIRING_KILN_TYPE_UNSUPPORTED,
                                       "Expected tunnel kiln type 2; actual value='" &
                                       kilnType & "'.")
                Continue For
            End If

            If SafeBool(r(COL_IS_SCHEDULED)) Then
                TraceCandidateDecision(debug, r, beforeCount, list.Count, False,
                                       SchedulerDebugReasonCodes.WIP_CURRENT_OPERATION_ALREADY_SCHEDULED,
                                       "Firing operation 300 is already scheduled.")
                Continue For
            End If

            Dim wipStatus As String = SharedHelpers.SafeStr(r("wip_status")).Trim()
            If Not wipStatus.Equals("Candidate", StringComparison.OrdinalIgnoreCase) Then
                TraceCandidateDecision(debug, r, beforeCount, list.Count, False,
                                       SchedulerDebugReasonCodes.WIP_STATUS_NOT_CANDIDATE,
                                       "Expected wip_status='Candidate'; actual value='" &
                                       wipStatus & "'.")
                Continue For
            End If

            Dim wipScore As Integer = SharedHelpers.SafeInt(r("wip_score"))
            Dim wipRejectReason As String = SharedHelpers.SafeStr(r("wip_reject_reason"))


            Dim orderNo As String = SafeStr(r(COL_ORDERNO)).Trim()
            If orderNo = "" Then
                TraceCandidateDecision(debug, r, beforeCount, list.Count, False,
                                       SchedulerDebugReasonCodes.DATA_MISSING_ORDER_NO,
                                       "Order No is blank.")
                Continue For
            End If

            ' Rule 3: ensure prior ops (<290) are scheduled -> must have readiness
            'If Not readyByOrder.ContainsKey(orderNo) Then Continue For
            Dim readiness As SharedHelpers.FiringReadinessInfo = Nothing

            If Not readinessByOrder.TryGetValue(orderNo, readiness) Then
                TraceCandidateDecision(debug, r, beforeCount, list.Count, False,
                                       SchedulerDebugReasonCodes.FIRING_PREV_OP_NOT_READY,
                                       "No firing readiness was produced for this order.")
                Continue For
            End If

            Dim firingOpRec As Integer = SafeInt(r(COL_OPREC))
            If firingOpRec <= 0 Then
                TraceCandidateDecision(debug, r, beforeCount, list.Count, False,
                                       SchedulerDebugReasonCodes.DATA_MISSING_OPERATION,
                                       "OrdersID for firing operation 300 is invalid.")
                Continue For
            End If

            Dim occ As Double = SafeDbl(r(COL_OCC))
            If occ <= 0 Then
                TraceCandidateDecision(debug, r, beforeCount, list.Count, False,
                                       SchedulerDebugReasonCodes.DATA_INVALID_OCCUPANCY,
                                       "Volume Occupancy must be positive; actual value='" &
                                       SafeStr(r(COL_OCC)).Trim() & "'.")
                Continue For
            End If

            Dim due As DateTime = SharedHelpers.ParseDueAsEndOfDay(r(COL_FIRING_DUE))
            If due = DateTime.MinValue Then
                TraceCandidateDecision(debug, r, beforeCount, list.Count, False,
                                       SchedulerDebugReasonCodes.DATA_INVALID_DATE,
                                       "Invalid firing due date raw value='" &
                                       SafeStr(r(COL_FIRING_DUE)).Trim() & "'.")
                Continue For
            End If

            'Dim ready As DateTime = readyByOrder(orderNo).AddHours(dryingToFiringBufferHours)
            Dim ready As DateTime = readiness.ReadyTime

            If Not readiness.LoadingAlreadyReleased Then
                ready = ready.AddHours(dryingToFiringBufferHours)
            End If

            wipScore = Math.Max(wipScore, readiness.WipScore)
            list.Add(New TunnelCandidate With {
                .OrderNo = orderNo,
                .FiringOpRec = firingOpRec,
                .ReadyTime = ready,
                .DueTime = due,
                .Occ = occ,
                .WipScore = wipScore,
                .WipRejectReason = wipRejectReason
            })

            TraceCandidateDecision(debug, r, beforeCount, list.Count, True,
                                   SchedulerDebugReasonCodes.OK_INCLUDED,
                                   "Included. Due=" &
                                   due.ToString("yyyy-MM-dd HH:mm:ss.fffffff",
                                                CultureInfo.InvariantCulture) &
                                   "; Ready=" &
                                   ready.ToString("yyyy-MM-dd HH:mm:ss",
                                                  CultureInfo.InvariantCulture) & ".")
        Next

        Return list
    End Function

    Private Sub TraceCandidateDecision(debug As SchedulerDebugCollector,
                                       row As DataRow,
                                       beforeCount As Integer,
                                       afterCount As Integer,
                                       included As Boolean,
                                       reasonCode As String,
                                       reasonDetail As String)
        If debug Is Nothing OrElse Not debug.Enabled Then Return

        Dim parentRecord As Integer = 0
        If row.Table.Columns.Contains("parent_record") Then
            parentRecord = SafeInt(row("parent_record"))
        End If

        debug.TraceCandidateStep(New OptimizerCandidateTraceRow With {
            .OptimizerName = "tunnelOptimizer_vf",
            .Stage = "TunnelFiring",
            .StepName = "CandidateFilter",
            .OrderNo = SafeStr(row(COL_ORDERNO)).Trim(),
            .ParentRecordNo = parentRecord,
            .RecordNo = SafeInt(row(COL_OPREC)),
            .OperationNumber = SafeInt(row(COL_OPNO)),
            .BeforeCount = beforeCount,
            .AfterCount = afterCount,
            .Included = included,
            .ReasonCode = reasonCode,
            .ReasonDetail = reasonDetail
        })
    End Sub

    ' -----------------------------
    ' Cart packing (Rule 2 + Rule 5)
    ' -----------------------------
    'Private Function PackCart(readyNow As List(Of TunnelCandidate),
    '                          minOccPreferred As Double,
    '                          maxOcc As Double) As List(Of TunnelCandidate)

    '    ' Deterministic selection:
    '    ' - iterate by earliest due date (Rule 5)
    '    ' - tie by earliest ready time
    '    ' - tie by larger occupancy first (fills cart quicker)
    '    Dim sorted As New List(Of TunnelCandidate)(readyNow)
    '    sorted.Sort(AddressOf CompareByDueReadyOcc)

    '    Dim chosen As New List(Of TunnelCandidate)()
    '    Dim occSum As Double = 0.0

    '    For Each c In sorted
    '        If occSum + c.Occ <= maxOcc + 0.0000001 Then
    '            chosen.Add(c)
    '            occSum += c.Occ

    '            ' MinOccPreferred is only a preference:
    '            ' once we reach it, we *may still add* if there is space, but we stop early
    '            ' to keep carts moving (and reduce late risk). This is conservative and deterministic.
    '            'If occSum + 0.0000001 >= minOccPreferred Then
    '            '    Exit For
    '            'End If
    '        End If
    '    Next

    '    ' If we didn't reach minOccPreferred, we still return what we have (preference, not hard).
    '    Return chosen
    'End Function

    Private Function PackCart(readyNow As List(Of TunnelCandidate),
                          minOccPreferred As Double,
                          maxOcc As Double) As List(Of TunnelCandidate)

        Dim sorted As New List(Of TunnelCandidate)(readyNow)
        sorted.Sort(AddressOf CompareByDueReadyOcc)

        Dim chosen As New List(Of TunnelCandidate)()
        Dim chosenSet As New HashSet(Of TunnelCandidate)()
        Dim occSum As Double = 0.0

        ' First preserve EDD: take the most urgent feasible order.
        For Each c In sorted
            If c.Occ <= maxOcc + 0.0000001 Then
                chosen.Add(c)
                chosenSet.Add(c)
                occSum += c.Occ
                Exit For
            End If
        Next

        If chosen.Count = 0 Then Return chosen

        ' Then fill remaining space with any ready order that fits.
        For Each c In sorted
            If chosenSet.Contains(c) Then Continue For

            If occSum + c.Occ <= maxOcc + 0.0000001 Then
                chosen.Add(c)
                chosenSet.Add(c)
                occSum += c.Occ
            End If
        Next

        Return chosen
    End Function

    ' -----------------------------
    ' Validation & parsing helpers
    ' -----------------------------
    Private Sub ValidateInputs(dt As DataTable,
                               cartsPerDay As Double,
                               totalCarts As Integer,
                               minOccPreferred As Double,
                               maxOcc As Double)

        If dt Is Nothing Then Throw New ArgumentNullException(NameOf(dt))
        If cartsPerDay <= 0 Then Throw New ArgumentException("cartsPerDay must be > 0.")
        If totalCarts <= 0 Then Throw New ArgumentException("totalCarts must be > 0.")
        If maxOcc <= 0 Then Throw New ArgumentException("maxOcc must be > 0.")
        If minOccPreferred < 0 Then Throw New ArgumentException("minOccPreferred must be >= 0.")
        If minOccPreferred > maxOcc Then Throw New ArgumentException("minOccPreferred cannot exceed maxOcc.")

        RequireColumn(dt, COL_ORDERNO)
        RequireColumn(dt, COL_OPREC)
        RequireColumn(dt, COL_OPNO)
        RequireColumn(dt, COL_KILNTYPE)
        RequireColumn(dt, COL_IS_SCHEDULED)
        RequireColumn(dt, COL_SCHED_END)
        RequireColumn(dt, COL_FIRING_DUE)
        RequireColumn(dt, COL_OCC)
        RequireColumn(dt, "wip_status")
        RequireColumn(dt, "wip_score")
        RequireColumn(dt, "wip_reject_reason")
    End Sub

    Private Sub RequireColumn(dt As DataTable, name As String)
        If Not dt.Columns.Contains(name) Then
            Throw New ArgumentException("Missing required column: " & name)
        End If
    End Sub

    'Private Shared Function CompareByDueReadyOcc(a As TunnelCandidate, b As TunnelCandidate) As Integer
    '    Dim c As Integer = a.DueTime.CompareTo(b.DueTime)
    '    If c <> 0 Then Return c
    '    c = a.ReadyTime.CompareTo(b.ReadyTime)
    '    If c <> 0 Then Return c
    '    ' larger occupancy first
    '    Return (-a.Occ.CompareTo(b.Occ))
    'End Function
    Private Shared Function CompareByDueReadyOcc(a As TunnelCandidate, b As TunnelCandidate) As Integer
        Dim w As Integer = b.WipScore.CompareTo(a.WipScore)
        If w <> 0 Then Return w

        Dim c As Integer = a.DueTime.CompareTo(b.DueTime)
        If c <> 0 Then Return c

        c = a.ReadyTime.CompareTo(b.ReadyTime)
        If c <> 0 Then Return c

        Return (-a.Occ.CompareTo(b.Occ))
    End Function

    ' Minimal “safe” converters (swap to your SharedHelpers if you prefer)
    Private Function SafeStr(o As Object) As String
        If o Is Nothing OrElse o Is DBNull.Value Then Return ""
        Return Convert.ToString(o, CultureInfo.InvariantCulture)
    End Function

    Private Function SafeInt(o As Object) As Integer
        If o Is Nothing OrElse o Is DBNull.Value Then Return 0
        Dim s As String = Convert.ToString(o, CultureInfo.InvariantCulture).Trim()
        Dim v As Integer
        If Integer.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, v) Then Return v
        Dim d As Double
        If Double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, d) Then Return CInt(Math.Truncate(d))
        Return 0
    End Function

    Private Function SafeDbl(o As Object) As Double
        If o Is Nothing OrElse o Is DBNull.Value Then Return 0.0
        Dim s As String = Convert.ToString(o, CultureInfo.InvariantCulture).Trim()
        Dim v As Double
        If Double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, v) Then Return v
        Return 0.0
    End Function

    Private Function SafeBool(o As Object) As Boolean
        If o Is Nothing OrElse o Is DBNull.Value Then Return False
        If TypeOf o Is Boolean Then Return CBool(o)
        Dim s As String = Convert.ToString(o, CultureInfo.InvariantCulture).Trim().ToLowerInvariant()
        Return (s = "1" OrElse s = "true" OrElse s = "yes" OrElse s = "y")
    End Function

End Class
