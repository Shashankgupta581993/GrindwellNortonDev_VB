Option Strict On
Option Explicit On

Imports System
Imports System.Collections.Generic
Imports System.Data
Imports System.Globalization
Imports System.IO
Imports System.Linq
Imports System.Runtime.InteropServices
Imports System.Text

<ComVisible(True)>
<Microsoft.VisualBasic.ComClass("7b8bde4c-1b57-4d35-934c-3a6f7f2b6f11", "f62e2b53-85d4-46ce-86e5-b1d9f4e9c8d9")>
Public Class swkOptimizer_vf

    Private Const SWK_KILN_TYPE As Integer = 3
    Private Const SWK_ACTIVE_CYCLE As String = "HT024"
    Private Const SWK_FUTURE_FILLER_CYCLE As String = "65VT"

    Private Const COL_ORDERNO As String = "Order No"
    Private Const COL_OPREC As String = "OrdersID"
    Private Const COL_OPNO As String = "Operation Number"
    Private Const COL_KILNTYPE As String = "Kiln Type"
    Private Const COL_CYCLE As String = "Cycle Type"
    Private Const COL_TONNAGE As String = "Tonnage"
    Private Const COL_BATCHTIME As String = "Batch Time"
    Private Const COL_IS_SCHEDULED As String = "is_scheduled"
    Private Const COL_SCHED_END As String = "scheduled_end_time"
    Private Const COL_FIRING_DUE As String = "firing due date"
    Private Const COL_PARENT As String = "parent_record"
    Private Const COL_PREVOP_IS_SCH As String = "prev_op_is_scheduled"

    Public Class SwkBatchPlan
        Public Property QueueFiringOpRecs As New List(Of Integer)

        Public Property BatchNoByFiringOpRec As New Dictionary(Of Integer, Integer)
        Public Property BatchStartByBatchNo As New Dictionary(Of Integer, DateTime)
        Public Property BatchEndByBatchNo As New Dictionary(Of Integer, DateTime)

        Public Property ResourceByBatchNo As New Dictionary(Of Integer, String)
        Public Property CycleByBatchNo As New Dictionary(Of Integer, String)
        Public Property TotalTonnageByBatchNo As New Dictionary(Of Integer, Double)
        Public Property LateCountByBatchNo As New Dictionary(Of Integer, Integer)
        Public Property UnderfilledByBatchNo As New Dictionary(Of Integer, Boolean)

        Public Property TotalBatches As Integer = 0
        Public Property TotalLateOrders As Integer = 0
    End Class

    Private Class SwkCandidate
        Public Property OrderNo As String
        Public Property FiringOpRec As Integer
        Public Property ParentRecord As Integer

        Public Property ReadyTime As DateTime
        Public Property DueTime As DateTime

        Public Property Tonnage As Double
        Public Property FireMins As Integer
        Public Property LoadMins As Integer

        Public Property PrevOpIsScheduled As Boolean
    End Class

    Private Class SwkBatchCandidate
        Public Property BatchStart As DateTime
        Public Property BatchEnd As DateTime
        Public Property Orders As List(Of SwkCandidate)
        Public Property TotalTonnage As Double
        Public Property LateCount As Integer
        Public Property Underfilled As Boolean
    End Class

    Public Function BuildSwkPlan(dt As DataTable,
                                 startTime As DateTime,
                                 minTonnage As Double,
                                 maxTonnage As Double,
                                 Optional dailyBatchLimit As Integer = 2,
                                 Optional batchStartDelayMins As Integer = 60,
                                 Optional allowUnderfilledTail As Boolean = True,
                                 Optional swkResourceName As String = "SWBKILN") As SwkBatchPlan

        ValidateInputs(dt, minTonnage, maxTonnage)

        Dim candidates As List(Of SwkCandidate) = BuildCandidates(dt, maxTonnage)

        Dim plan As New SwkBatchPlan()
        If candidates.Count = 0 Then Return plan

        Dim unassigned As New Dictionary(Of Integer, SwkCandidate)()
        For Each c In candidates
            If Not unassigned.ContainsKey(c.FiringOpRec) Then
                unassigned.Add(c.FiringOpRec, c)
            End If
        Next

        Dim swkAvail As DateTime = startTime
        Dim batchNo As Integer = 0

        Dim capPerDay As Integer = If(dailyBatchLimit <= 0, Integer.MaxValue, dailyBatchLimit)
        Dim countByDay As New Dictionary(Of Date, Integer)()

        While unassigned.Count > 0

            Dim readyPool As List(Of SwkCandidate) = GetReadyPool(unassigned, swkAvail)

            If readyPool.Count = 0 Then
                Dim nextReady As DateTime = GetNextReadyTime(unassigned)
                If nextReady = DateTime.MaxValue Then Exit While
                swkAvail = If(nextReady > swkAvail, nextReady, swkAvail)
                Continue While
            End If

            Dim batch As SwkBatchCandidate =
                BuildBestPureHt024Batch(readyPool,
                                        swkAvail,
                                        minTonnage,
                                        maxTonnage,
                                        batchStartDelayMins)

            If batch Is Nothing OrElse batch.Orders.Count = 0 Then
                Exit While
            End If

            If batch.TotalTonnage < minTonnage Then
                Dim futureReady As DateTime = GetNextReadyTimeAfter(unassigned, swkAvail)

                If futureReady <> DateTime.MinValue AndAlso Not allowUnderfilledTail Then
                    swkAvail = futureReady
                    Continue While
                End If

                If futureReady <> DateTime.MinValue AndAlso allowUnderfilledTail = False Then
                    swkAvail = futureReady
                    Continue While
                End If

                If futureReady <> DateTime.MinValue AndAlso batch.TotalTonnage < minTonnage Then
                    ' Wait for more HT024 only if we are not at tail.
                    swkAvail = futureReady
                    Continue While
                End If

                If futureReady = DateTime.MinValue AndAlso Not allowUnderfilledTail Then
                    Exit While
                End If

                batch.Underfilled = True
            End If

            Dim d As Date = batch.BatchStart.Date
            Dim usedToday As Integer = 0
            If countByDay.ContainsKey(d) Then usedToday = countByDay(d)

            If usedToday >= capPerDay Then
                swkAvail = d.AddDays(1).Add(startTime.TimeOfDay)
                Continue While
            End If

            batchNo += 1
            countByDay(d) = usedToday + 1

            CommitBatch(plan, batch, batchNo, swkResourceName)

            swkAvail = batch.BatchEnd

            For Each o In batch.Orders
                unassigned.Remove(o.FiringOpRec)
            Next

        End While

        plan.TotalBatches = batchNo
        Return plan

    End Function

    Private Function BuildCandidates(dt As DataTable, maxTonnage As Double) As List(Of SwkCandidate)

        Dim list As New List(Of SwkCandidate)()

        Dim lastPreOpNoByOrder As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)

        For Each r As DataRow In dt.Rows
            Dim orderNo As String = SharedHelpers.SafeStr(r(COL_ORDERNO)).Trim()
            If orderNo = "" Then Continue For

            Dim opNo As Integer = SharedHelpers.SafeInt(r(COL_OPNO))
            If opNo <= 0 OrElse opNo >= 290 Then Continue For

            Dim current As Integer = 0
            If Not lastPreOpNoByOrder.TryGetValue(orderNo, current) OrElse opNo > current Then
                lastPreOpNoByOrder(orderNo) = opNo
            End If
        Next

        Dim readyByOrder As New Dictionary(Of String, DateTime)(StringComparer.OrdinalIgnoreCase)

        For Each r As DataRow In dt.Rows
            Dim orderNo As String = SharedHelpers.SafeStr(r(COL_ORDERNO)).Trim()
            If orderNo = "" Then Continue For

            Dim lastOpNo As Integer = 0
            If Not lastPreOpNoByOrder.TryGetValue(orderNo, lastOpNo) Then Continue For

            Dim opNo As Integer = SharedHelpers.SafeInt(r(COL_OPNO))
            If opNo <> lastOpNo Then Continue For

            If Not SharedHelpers.SafeBool(r(COL_IS_SCHEDULED)) Then Continue For

            Dim endT As DateTime = SharedHelpers.SafeDate(r(COL_SCHED_END))
            If endT = DateTime.MinValue Then Continue For

            If Not readyByOrder.ContainsKey(orderNo) OrElse endT > readyByOrder(orderNo) Then
                readyByOrder(orderNo) = endT
            End If
        Next

        Dim loadingMinsByOrder As Dictionary(Of String, Integer) = BuildLoadingMinsByOrder(dt)
        Dim hasPrevCol As Boolean = dt.Columns.Contains(COL_PREVOP_IS_SCH)

        For Each r As DataRow In dt.Rows

            Dim kilnType As Integer = SharedHelpers.SafeInt(r(COL_KILNTYPE))
            If kilnType <> SWK_KILN_TYPE Then Continue For

            Dim opNo As Integer = SharedHelpers.SafeInt(r(COL_OPNO))
            If opNo <> 300 Then Continue For

            If SharedHelpers.SafeBool(r(COL_IS_SCHEDULED)) Then Continue For

            Dim cycle As String = NormalizeCycle(SharedHelpers.SafeStr(r(COL_CYCLE)))
            If cycle = SWK_FUTURE_FILLER_CYCLE Then Continue For
            If cycle <> SWK_ACTIVE_CYCLE Then Continue For

            Dim orderNo As String = SharedHelpers.SafeStr(r(COL_ORDERNO)).Trim()
            If orderNo = "" Then Continue For

            Dim ready As DateTime
            If Not readyByOrder.TryGetValue(orderNo, ready) Then Continue For

            Dim due As DateTime = SharedHelpers.ParseDueAsEndOfDay(r(COL_FIRING_DUE))
            If due = DateTime.MinValue Then Continue For

            Dim tonnage As Double = SharedHelpers.SafeDbl(r(COL_TONNAGE))
            If tonnage <= 0 Then Continue For
            If tonnage > maxTonnage + 0.0000001 Then Continue For

            Dim fireMins As Integer = CInt(Math.Truncate(SharedHelpers.SafeDbl(r(COL_BATCHTIME))))
            If fireMins <= 0 Then Continue For

            Dim loadMins As Integer = 0
            If Not loadingMinsByOrder.TryGetValue(orderNo, loadMins) Then
                loadMins = 0
            End If

            Dim firingOpRec As Integer = SharedHelpers.SafeInt(r(COL_OPREC))
            If firingOpRec <= 0 Then Continue For

            Dim prevScheduled As Boolean = False
            If hasPrevCol Then prevScheduled = SharedHelpers.SafeBool(r(COL_PREVOP_IS_SCH))

            list.Add(New SwkCandidate With {
                .OrderNo = orderNo,
                .FiringOpRec = firingOpRec,
                .ParentRecord = SharedHelpers.SafeInt(r(COL_PARENT)),
                .ReadyTime = ready,
                .DueTime = due,
                .Tonnage = tonnage,
                .FireMins = fireMins,
                .LoadMins = loadMins,
                .PrevOpIsScheduled = prevScheduled
            })

        Next

        Return list

    End Function

    Private Function BuildLoadingMinsByOrder(dt As DataTable) As Dictionary(Of String, Integer)

        Dim dict As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)

        For Each r As DataRow In dt.Rows

            Dim orderNo As String = SharedHelpers.SafeStr(r(COL_ORDERNO)).Trim()
            If orderNo = "" Then Continue For

            Dim opNo As Integer = SharedHelpers.SafeInt(r(COL_OPNO))

            ' Keep same batch convention: loading rows are around 290/291.
            If opNo <> 290 AndAlso opNo <> 291 Then Continue For

            Dim mins As Integer = CInt(Math.Truncate(SharedHelpers.SafeDbl(r(COL_BATCHTIME))))

            If Not dict.ContainsKey(orderNo) OrElse mins > dict(orderNo) Then
                dict(orderNo) = mins
            End If

        Next

        Return dict

    End Function

    Private Function BuildBestPureHt024Batch(readyPool As List(Of SwkCandidate),
                                             swkAvail As DateTime,
                                             minTonnage As Double,
                                             maxTonnage As Double,
                                             batchStartDelayMins As Integer) As SwkBatchCandidate

        Dim sorted As List(Of SwkCandidate) = SortCandidates(readyPool)

        Dim selected As New List(Of SwkCandidate)()
        Dim total As Double = 0.0

        For Each c In sorted
            If total + c.Tonnage <= maxTonnage + 0.0000001 Then
                selected.Add(c)
                total += c.Tonnage
            End If
        Next

        If selected.Count = 0 Then Return Nothing

        Dim startT As DateTime = swkAvail

        Dim maxReady As DateTime = DateTime.MinValue
        Dim maxLoadMins As Integer = 0
        Dim maxFireMins As Integer = 0

        For Each o In selected
            If o.ReadyTime > maxReady Then maxReady = o.ReadyTime
            If o.LoadMins > maxLoadMins Then maxLoadMins = o.LoadMins
            If o.FireMins > maxFireMins Then maxFireMins = o.FireMins
        Next

        Dim readyPlusLoad As DateTime = maxReady.AddMinutes(maxLoadMins)
        If readyPlusLoad > startT Then startT = readyPlusLoad

        If batchStartDelayMins > 0 Then
            startT = startT.AddMinutes(batchStartDelayMins)
        End If

        Dim endT As DateTime = startT.AddMinutes(maxFireMins)

        Dim lateCount As Integer = 0
        For Each o In selected
            If endT > o.DueTime Then lateCount += 1
        Next

        Return New SwkBatchCandidate With {
            .BatchStart = startT,
            .BatchEnd = endT,
            .Orders = selected,
            .TotalTonnage = total,
            .LateCount = lateCount,
            .Underfilled = total < minTonnage
        }

    End Function

    Private Sub CommitBatch(plan As SwkBatchPlan,
                            b As SwkBatchCandidate,
                            batchNo As Integer,
                            swkResourceName As String)

        plan.BatchStartByBatchNo(batchNo) = b.BatchStart
        plan.BatchEndByBatchNo(batchNo) = b.BatchEnd
        plan.ResourceByBatchNo(batchNo) = swkResourceName
        plan.CycleByBatchNo(batchNo) = SWK_ACTIVE_CYCLE
        plan.TotalTonnageByBatchNo(batchNo) = b.TotalTonnage
        plan.LateCountByBatchNo(batchNo) = b.LateCount
        plan.UnderfilledByBatchNo(batchNo) = b.Underfilled

        plan.TotalLateOrders += b.LateCount

        Dim ordered As List(Of SwkCandidate) = SortCandidates(b.Orders)

        For Each o In ordered
            plan.QueueFiringOpRecs.Add(o.FiringOpRec)
            plan.BatchNoByFiringOpRec(o.FiringOpRec) = batchNo
        Next

    End Sub

    Private Function SortCandidates(input As List(Of SwkCandidate)) As List(Of SwkCandidate)

        Dim list As New List(Of SwkCandidate)(input)

        list.Sort(Function(a, b)
                      If a.PrevOpIsScheduled <> b.PrevOpIsScheduled Then
                          If a.PrevOpIsScheduled Then Return -1 Else Return 1
                      End If

                      Dim c As Integer = a.DueTime.CompareTo(b.DueTime)
                      If c <> 0 Then Return c

                      c = a.ReadyTime.CompareTo(b.ReadyTime)
                      If c <> 0 Then Return c

                      c = a.Tonnage.CompareTo(b.Tonnage)
                      If c <> 0 Then Return c

                      c = a.ParentRecord.CompareTo(b.ParentRecord)
                      If c <> 0 Then Return c

                      Return a.FiringOpRec.CompareTo(b.FiringOpRec)
                  End Function)

        Return list

    End Function

    Private Function GetReadyPool(unassigned As Dictionary(Of Integer, SwkCandidate),
                                  t As DateTime) As List(Of SwkCandidate)

        Dim pool As New List(Of SwkCandidate)()

        For Each kvp In unassigned
            If kvp.Value.ReadyTime <= t Then pool.Add(kvp.Value)
        Next

        Return pool

    End Function

    Private Function GetNextReadyTime(unassigned As Dictionary(Of Integer, SwkCandidate)) As DateTime

        Dim best As DateTime = DateTime.MaxValue

        For Each kvp In unassigned
            If kvp.Value.ReadyTime < best Then best = kvp.Value.ReadyTime
        Next

        Return best

    End Function

    Private Function GetNextReadyTimeAfter(unassigned As Dictionary(Of Integer, SwkCandidate),
                                           t As DateTime) As DateTime

        Dim best As DateTime = DateTime.MaxValue
        Dim found As Boolean = False

        For Each kvp In unassigned
            Dim rt As DateTime = kvp.Value.ReadyTime
            If rt > t AndAlso rt < best Then
                best = rt
                found = True
            End If
        Next

        If Not found Then Return DateTime.MinValue
        Return best

    End Function

    Private Function NormalizeCycle(value As String) As String
        If value Is Nothing Then Return ""
        Return value.Trim().Replace(" ", "").ToUpperInvariant()
    End Function

    Private Sub ValidateInputs(dt As DataTable,
                               minTonnage As Double,
                               maxTonnage As Double)

        If dt Is Nothing Then Throw New ArgumentNullException(NameOf(dt))
        If minTonnage <= 0 Then Throw New ArgumentException("SWK min tonnage must be > 0.")
        If maxTonnage <= 0 Then Throw New ArgumentException("SWK max tonnage must be > 0.")
        If minTonnage > maxTonnage Then Throw New ArgumentException("SWK min tonnage cannot exceed max tonnage.")

        SharedHelpers.RequireColumn(dt, COL_ORDERNO)
        SharedHelpers.RequireColumn(dt, COL_OPREC)
        SharedHelpers.RequireColumn(dt, COL_OPNO)
        SharedHelpers.RequireColumn(dt, COL_KILNTYPE)
        SharedHelpers.RequireColumn(dt, COL_CYCLE)
        SharedHelpers.RequireColumn(dt, COL_TONNAGE)
        SharedHelpers.RequireColumn(dt, COL_BATCHTIME)
        SharedHelpers.RequireColumn(dt, COL_IS_SCHEDULED)
        SharedHelpers.RequireColumn(dt, COL_SCHED_END)
        SharedHelpers.RequireColumn(dt, COL_FIRING_DUE)
        SharedHelpers.RequireColumn(dt, COL_PARENT)

    End Sub

    Public Sub ExportSwkPlanToCsv(plan As SwkBatchPlan, folderPath As String)

        If plan Is Nothing Then Throw New ArgumentNullException(NameOf(plan))
        If String.IsNullOrWhiteSpace(folderPath) Then Throw New ArgumentException("folderPath is empty.")

        Directory.CreateDirectory(folderPath)

        Dim summaryPath As String = Path.Combine(folderPath, "SWK_BatchSummary.csv")
        Dim queuePath As String = Path.Combine(folderPath, "SWK_FiringQueue_Return.csv")

        Using w As New StreamWriter(summaryPath, False, New UTF8Encoding(False))

            w.WriteLine("BatchNo,Resource,CycleType,BatchStart,BatchEnd,TotalTonnage,OrderCount,LateOrderCount,UnderfilledFlag")

            Dim batchNos As New List(Of Integer)(plan.BatchStartByBatchNo.Keys)
            batchNos.Sort()

            For Each b In batchNos

                Dim orderCount As Integer = 0
                For Each kvp In plan.BatchNoByFiringOpRec
                    If kvp.Value = b Then orderCount += 1
                Next

                w.WriteLine(String.Join(",", New String() {
                    b.ToString(CultureInfo.InvariantCulture),
                    SharedHelpers.Csv(plan.ResourceByBatchNo(b)),
                    SharedHelpers.Csv(plan.CycleByBatchNo(b)),
                    SharedHelpers.Csv(SharedHelpers.FormatDateOrBlank(plan.BatchStartByBatchNo(b))),
                    SharedHelpers.Csv(SharedHelpers.FormatDateOrBlank(plan.BatchEndByBatchNo(b))),
                    plan.TotalTonnageByBatchNo(b).ToString(CultureInfo.InvariantCulture),
                    orderCount.ToString(CultureInfo.InvariantCulture),
                    plan.LateCountByBatchNo(b).ToString(CultureInfo.InvariantCulture),
                    plan.UnderfilledByBatchNo(b).ToString(CultureInfo.InvariantCulture)
                }))

            Next

        End Using

        Using w As New StreamWriter(queuePath, False, New UTF8Encoding(False))

            w.WriteLine("QueueIndex,FiringOpRec,BatchNo,BatchStart,BatchEnd,Resource,CycleType,TotalTonnage")

            For i As Integer = 0 To plan.QueueFiringOpRecs.Count - 1

                Dim firingOpRec As Integer = plan.QueueFiringOpRecs(i)
                Dim batchNo As Integer = plan.BatchNoByFiringOpRec(firingOpRec)

                w.WriteLine(String.Join(",", New String() {
                    (i + 1).ToString(CultureInfo.InvariantCulture),
                    firingOpRec.ToString(CultureInfo.InvariantCulture),
                    batchNo.ToString(CultureInfo.InvariantCulture),
                    SharedHelpers.Csv(SharedHelpers.FormatDateOrBlank(plan.BatchStartByBatchNo(batchNo))),
                    SharedHelpers.Csv(SharedHelpers.FormatDateOrBlank(plan.BatchEndByBatchNo(batchNo))),
                    SharedHelpers.Csv(plan.ResourceByBatchNo(batchNo)),
                    SharedHelpers.Csv(plan.CycleByBatchNo(batchNo)),
                    plan.TotalTonnageByBatchNo(batchNo).ToString(CultureInfo.InvariantCulture)
                }))

            Next

        End Using

    End Sub

End Class