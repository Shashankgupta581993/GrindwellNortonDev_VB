Option Strict On
Option Explicit On

Imports System
Imports System.Data
Imports System.Diagnostics
Imports System.Linq
Imports Preactor

Public Class PostFiringScheduler

    Public Class QueueItem
        Public Property ParentRecord As Integer
        Public Property OrderNo As String
        Public Property KilnAckOpRec As Integer
        Public Property KilnAckEndTime As DateTime
        Public Property NextOpRec As Integer
        Public Property NextOpNo As Integer
        Public Property DueDate As DateTime
        Public Property Priority As Integer
        Public Property WipScore As Integer
        Public Property WipStatus As String
        Public Property WipRejectReason As String
    End Class

    Public Function BuildQueue(preactor As IPreactor,
                               planningboard As IPlanningBoard,
                               routingDt As DataTable,
                               Optional kilnAckName As String = "KILNACK") As List(Of QueueItem)

        If routingDt Is Nothing Then Throw New ArgumentNullException(NameOf(routingDt))

        RequireColumn(routingDt, "OrdersID")
        RequireColumn(routingDt, "Order No")
        RequireColumn(routingDt, "is_scheduled")
        RequireColumn(routingDt, "scheduled_end_time")
        RequireColumn(routingDt, "parent_record")
        RequireColumn(routingDt, "wip_score")
        RequireColumn(routingDt, "wip_status")
        RequireColumn(routingDt, "wip_reject_reason")

        Dim ordersTable As Integer = preactor.FindFirstClassificationString("LAUNCH TIME").Value.FormatNumber
        Dim opNoField As Integer = preactor.GetFieldNumber(ordersTable, "Op. No.")
        Dim dueDateField As Integer = TryGetFieldNumber(preactor, ordersTable, "Due Date")
        Dim priorityField As Integer = TryGetFieldNumber(preactor, ordersTable, "Priority")

        Dim queue As New List(Of QueueItem)()
        Dim rowByOpRec As New Dictionary(Of Integer, DataRow)()

        For Each row As DataRow In routingDt.Rows
            Dim opRec As Integer = SafeInt(row("OrdersID"))
            If opRec > 0 AndAlso Not rowByOpRec.ContainsKey(opRec) Then
                rowByOpRec.Add(opRec, row)
            End If
        Next

        For Each r As DataRow In routingDt.Rows

            If Not IsKilnAckRow(routingDt, r, kilnAckName) Then Continue For
            If Not SafeBool(r("is_scheduled")) Then Continue For

            Dim kilnAckOpRec As Integer = SafeInt(r("OrdersID"))
            If kilnAckOpRec <= 0 Then Continue For

            Dim kilnAckEnd As DateTime = SafeDate(r("scheduled_end_time"))
            If kilnAckEnd = DateTime.MinValue Then Continue For

            ' First unscheduled operation after KILNACK
            Dim nextOpRec As Integer = planningboard.GetNextOperation(kilnAckOpRec, 1)

            While nextOpRec > 0 AndAlso planningboard.IsOperationScheduled(nextOpRec)
                nextOpRec = planningboard.GetNextOperation(nextOpRec, 1)
            End While

            If nextOpRec <= 0 Then Continue For

            Dim nextOpNo As Integer
            Try
                nextOpNo = preactor.ReadFieldInt(ordersTable, opNoField, nextOpRec)
            Catch
                Continue For
            End Try

            Dim nextRow As DataRow = Nothing
            If Not rowByOpRec.TryGetValue(nextOpRec, nextRow) Then Continue For

            Dim wipStatus As String = SafeStr(nextRow("wip_status"))
            Dim wipScore As Integer = SafeInt(nextRow("wip_score"))
            Dim wipRejectReason As String = SafeStr(nextRow("wip_reject_reason"))

            ' Do not use snapshot WIP status as a hard gate here. A scheduled
            ' KILNACK and the live routing chain determine post-firing eligibility.
            queue.Add(New QueueItem With {
                    .ParentRecord = SafeInt(r("parent_record")),
                    .OrderNo = SafeStr(r("Order No")),
                    .KilnAckOpRec = kilnAckOpRec,
                    .KilnAckEndTime = kilnAckEnd,
                    .NextOpRec = nextOpRec,
                    .NextOpNo = nextOpNo,
                    .DueDate = ReadDueDate(preactor, ordersTable, dueDateField, nextOpRec),
                    .Priority = ReadPriority(preactor, ordersTable, priorityField, nextOpRec),
                    .WipScore = wipScore,
                    .WipStatus = wipStatus,
                    .WipRejectReason = wipRejectReason
})


        Next

        ' FIFO: oven exit first. Due date breaks conflict.
        Return queue _
                .OrderByDescending(Function(x) x.WipScore) _
                .ThenBy(Function(x) x.KilnAckEndTime) _
                .ThenBy(Function(x) x.DueDate) _
                .ThenBy(Function(x) x.Priority) _
                .ThenBy(Function(x) x.ParentRecord) _
                .ThenBy(Function(x) x.NextOpNo) _
                .ToList()

    End Function

    Public Function ScheduleQueue(preactor As IPreactor,
                                  planningboard As IPlanningBoard,
                                  queue As List(Of QueueItem)) As Integer

        If queue Is Nothing OrElse queue.Count = 0 Then Return 0

        Dim ordersTable As Integer = preactor.FindFirstClassificationString("LAUNCH TIME").Value.FormatNumber
        Dim opNoField As Integer = preactor.GetFieldNumber(ordersTable, "Op. No.")

        Dim scheduledCount As Integer = 0

        For Each item As QueueItem In queue

            Try
                Dim opRec As Integer = item.NextOpRec
                Dim testFrom As DateTime = item.KilnAckEndTime

                While opRec > 0
                    If planningboard.IsOperationScheduled(opRec) Then
                        testFrom = GetScheduledEnd(planningboard, opRec, testFrom)
                        opRec = planningboard.GetNextOperation(opRec, 1)
                        Continue While
                    End If

                    Dim liveOpNo As Integer =
                        preactor.ReadFieldInt(ordersTable, opNoField, opRec)

                    Dim bestResRec As Integer = 0
                    Dim bestTimes As OperationTimes? = Nothing
                    Dim resources As IEnumerable(Of Integer) =
                        planningboard.FindResources(opRec)

                    If resources IsNot Nothing Then
                        For Each resRec As Integer In resources
                            Dim testTimes As OperationTimes? =
                                planningboard.TestOperationOnResource(opRec,
                                                                      resRec,
                                                                      testFrom)

                            If testTimes.HasValue AndAlso
                               (Not bestTimes.HasValue OrElse
                                testTimes.Value.ChangeStart < bestTimes.Value.ChangeStart) Then

                                bestTimes = testTimes
                                bestResRec = resRec
                            End If
                        Next
                    End If

                    If bestTimes.HasValue AndAlso bestResRec > 0 Then
                        ' Recheck immediately before changing the live board.
                        If Not planningboard.IsOperationScheduled(opRec) Then
                            planningboard.PutOperationOnResource(opRec,
                                                                 bestResRec,
                                                                 bestTimes.Value.ChangeStart)
                            scheduledCount += 1
                        End If

                        testFrom = GetScheduledEnd(planningboard,
                                                   opRec,
                                                   bestTimes.Value.ProcessEnd)
                    Else
                        Debug.WriteLine("PostFiring: no feasible resource. Order=" &
                                        item.OrderNo &
                                        ", OpRec=" & opRec &
                                        ", OpNo=" & liveOpNo)
                    End If

                    opRec = planningboard.GetNextOperation(opRec, 1)
                End While

            Catch ex As Exception
                Debug.WriteLine("PostFiring failed. Order=" &
                                item.OrderNo &
                                ", OpRec=" & item.NextOpRec &
                                ", Error=" & ex.Message)
            End Try

        Next

        Return scheduledCount

    End Function

    Private Function IsKilnAckRow(dt As DataTable,
                                  r As DataRow,
                                  kilnAckName As String) As Boolean

        If dt.Columns.Contains("Operation Name") Then
            If SafeStr(r("Operation Name")).Equals(kilnAckName, StringComparison.OrdinalIgnoreCase) Then
                Return True
            End If
        End If

        If dt.Columns.Contains("Required Resource") Then
            If SafeStr(r("Required Resource")).Equals(kilnAckName, StringComparison.OrdinalIgnoreCase) Then
                Return True
            End If
        End If

        If dt.Columns.Contains("Resource Group") Then
            If SafeStr(r("Resource Group")).Equals(kilnAckName, StringComparison.OrdinalIgnoreCase) Then
                Return True
            End If
        End If

        Return False

    End Function

    Private Function ReadDueDate(preactor As IPreactor,
                                 ordersTable As Integer,
                                 dueField As Integer,
                                 opRec As Integer) As DateTime
        If dueField <= 0 Then Return DateTime.MaxValue

        Try
            Return preactor.ReadFieldDateTime(ordersTable, dueField, opRec)
        Catch
            Return DateTime.MaxValue
        End Try
    End Function

    Private Function ReadPriority(preactor As IPreactor,
                                  ordersTable As Integer,
                                  priorityField As Integer,
                                  opRec As Integer) As Integer
        If priorityField <= 0 Then Return 999999

        Try
            Return preactor.ReadFieldInt(ordersTable, priorityField, opRec)
        Catch
            Return 999999
        End Try
    End Function

    Private Function TryGetFieldNumber(preactor As IPreactor,
                                       ordersTable As Integer,
                                       fieldName As String) As Integer
        Try
            Return preactor.GetFieldNumber(ordersTable, fieldName)
        Catch
            Return 0
        End Try
    End Function

    Private Function GetScheduledEnd(planningboard As IPlanningBoard,
                                     opRec As Integer,
                                     fallback As DateTime) As DateTime
        Dim times As Nullable(Of OperationResourceTimes) =
            planningboard.GetOperationTimes(opRec)

        If times.HasValue Then
            Return times.Value.OperationTimes.ProcessEnd
        End If

        Return fallback
    End Function

    Private Sub RequireColumn(dt As DataTable, colName As String)
        If Not dt.Columns.Contains(colName) Then
            Throw New Exception("Required routingDt column missing: " & colName)
        End If
    End Sub

    Private Function SafeStr(v As Object) As String
        If v Is Nothing OrElse v Is DBNull.Value Then Return ""
        Return v.ToString().Trim()
    End Function

    Private Function SafeInt(v As Object) As Integer
        If v Is Nothing OrElse v Is DBNull.Value Then Return 0

        Dim result As Integer
        If Integer.TryParse(v.ToString(), result) Then Return result

        Return 0
    End Function

    Private Function SafeBool(v As Object) As Boolean
        If v Is Nothing OrElse v Is DBNull.Value Then Return False

        Dim result As Boolean
        If Boolean.TryParse(v.ToString(), result) Then Return result

        Dim s As String = v.ToString().Trim()
        Return s = "1" OrElse
               s.Equals("Y", StringComparison.OrdinalIgnoreCase) OrElse
               s.Equals("YES", StringComparison.OrdinalIgnoreCase)
    End Function

    Private Function SafeDate(v As Object) As DateTime
        If v Is Nothing OrElse v Is DBNull.Value Then Return DateTime.MinValue

        Dim result As DateTime
        If DateTime.TryParse(v.ToString(), result) Then Return result

        Return DateTime.MinValue
    End Function

End Class
