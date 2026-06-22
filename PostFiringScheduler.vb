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

        Dim ordersTable As Integer = preactor.FindFirstClassificationString("LAUNCH TIME").Value.FormatNumber
        Dim opNoField As Integer = preactor.GetFieldNumber(ordersTable, "Op. No.")

        Dim queue As New List(Of QueueItem)()

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

            queue.Add(New QueueItem With {
                .ParentRecord = SafeInt(r("parent_record")),
                .OrderNo = SafeStr(r("Order No")),
                .KilnAckOpRec = kilnAckOpRec,
                .KilnAckEndTime = kilnAckEnd,
                .NextOpRec = nextOpRec,
                .NextOpNo = nextOpNo,
                .DueDate = ReadDueDate(preactor, ordersTable, nextOpRec),
                .Priority = ReadPriority(preactor, ordersTable, nextOpRec)
            })

        Next

        ' FIFO: oven exit first. Due date breaks conflict.
        Return queue _
            .OrderBy(Function(x) x.KilnAckEndTime) _
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
                ' Live validation because routingDt is only a snapshot
                If planningboard.IsOperationScheduled(item.NextOpRec) Then Continue For

                Dim liveOpNo As Integer = preactor.ReadFieldInt(ordersTable, opNoField, item.NextOpRec)
                If liveOpNo <> item.NextOpNo Then Continue For

                Dim resources As IEnumerable(Of Integer) = planningboard.FindResources(item.NextOpRec)

                Dim bestResRec As Integer = 0
                Dim bestTimes As OperationTimes? = Nothing

                For Each resRec As Integer In resources

                    Dim testTimes As OperationTimes? =
                        planningboard.TestOperationOnResource(item.NextOpRec, resRec, item.KilnAckEndTime)

                    If testTimes.HasValue Then
                        If Not bestTimes.HasValue OrElse
                           testTimes.Value.ChangeStart < bestTimes.Value.ChangeStart Then

                            bestTimes = testTimes
                            bestResRec = resRec

                        End If
                    End If

                Next

                If bestTimes.HasValue AndAlso bestResRec > 0 Then
                    planningboard.PutOperationOnResource(item.NextOpRec,
                                                         bestResRec,
                                                         bestTimes.Value.ChangeStart)
                    scheduledCount += 1
                Else
                    Debug.WriteLine("PostFiring: no feasible resource. Order=" &
                                    item.OrderNo &
                                    ", OpRec=" & item.NextOpRec &
                                    ", OpNo=" & item.NextOpNo)
                End If

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
                                 opRec As Integer) As DateTime
        Try
            Dim dueField As Integer = preactor.GetFieldNumber(ordersTable, "Due Date")
            Return preactor.ReadFieldDateTime(ordersTable, dueField, opRec)
        Catch
            Return DateTime.MaxValue
        End Try
    End Function

    Private Function ReadPriority(preactor As IPreactor,
                                  ordersTable As Integer,
                                  opRec As Integer) As Integer
        Try
            Dim priorityField As Integer = preactor.GetFieldNumber(ordersTable, "Priority")
            Return preactor.ReadFieldInt(ordersTable, priorityField, opRec)
        Catch
            Return 999999
        End Try
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