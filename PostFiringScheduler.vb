Option Strict On
Option Explicit On

Imports System
Imports System.Data
Imports System.Diagnostics
Imports System.Linq
Imports Preactor

Public Class PostFiringScheduler
    Private Const POSTFIRING_START_OP_NO As Integer = 400

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
        Public Property OperationRows As IDictionary(Of Integer, DataRow)
    End Class

    Public Function BuildQueue(preactor As IPreactor,
                               planningboard As IPlanningBoard,
                               routingDt As DataTable,
                               Optional kilnAckName As String = "KILNACK",
                               Optional debug As SchedulerDebugCollector = Nothing) As List(Of QueueItem)

        If routingDt Is Nothing Then Throw New ArgumentNullException(NameOf(routingDt))

        RequireColumn(routingDt, "OrdersID")
        RequireColumn(routingDt, "Order No")
        RequireColumn(routingDt, "is_scheduled")
        RequireColumn(routingDt, "scheduled_end_time")
        RequireColumn(routingDt, "parent_record")
        RequireColumn(routingDt, "wip_score")
        RequireColumn(routingDt, "wip_status")
        RequireColumn(routingDt, "wip_reject_reason")
        RequireColumn(routingDt, "operation_releases_next")
        RequireColumn(routingDt, "operation_release_time")
        RequireColumn(routingDt, "operation_effective_completed")
        RequireColumn(routingDt, "Operation Number")

        Dim ordersTable As Integer = preactor.FindFirstClassificationString("LAUNCH TIME").Value.FormatNumber
        Dim opNoField As Integer = preactor.GetFieldNumber(ordersTable, "Op. No.")
        Dim dueDateField As Integer = TryGetFieldNumber(preactor, ordersTable, "Due Date")
        Dim priorityField As Integer = TryGetFieldNumber(preactor, ordersTable, "Priority")

        Dim queue As New List(Of QueueItem)()
        Dim rowByOpRec As Dictionary(Of Integer, DataRow) =
            SharedHelpers.BuildOperationRowIndex(routingDt)
        Dim boundaries As List(Of SharedHelpers.ReleaseBoundaryInfo) =
            SharedHelpers.BuildLatestReleaseBoundaries(routingDt)

        If boundaries.Count = 0 AndAlso debug IsNot Nothing AndAlso debug.Enabled Then
            debug.TraceCandidateStep(New OptimizerCandidateTraceRow With {
                .OptimizerName = "PostFiringScheduler",
                .Stage = "PostFiring400Plus",
                .StepName = "LatestReleaseBoundary",
                .OrderNo = "",
                .ParentRecordNo = 0,
                .RecordNo = 0,
                .OperationNumber = 0,
                .BeforeCount = routingDt.Rows.Count,
                .AfterCount = 0,
                .Included = False,
                .ReasonCode = SchedulerDebugReasonCodes.POSTFIRING_NO_RELEASE_BOUNDARY,
                .ReasonDetail = "No operation_releases_next boundary with valid operation_release_time was found.",
                .RankScore = 0,
                .RankBreakdown = ""
            })
        End If

        For Each boundary As SharedHelpers.ReleaseBoundaryInfo In boundaries

            TraceBoundary(debug,
                          boundary,
                          True,
                          SchedulerDebugReasonCodes.OK_INCLUDED,
                          "Latest completed/released boundary selected for 400+ postfiring.")

            Dim nextOpRec As Integer =
                FindNextPostFiringCandidate(preactor,
                                            planningboard,
                                            ordersTable,
                                            opNoField,
                                            rowByOpRec,
                                            boundary,
                                            debug)
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
            ' release boundary and the live routing chain determine post-firing eligibility.
            queue.Add(New QueueItem With {
                    .ParentRecord = boundary.ParentRecord,
                    .OrderNo = boundary.OrderNo,
                    .KilnAckOpRec = boundary.OpRec,
                    .KilnAckEndTime = boundary.ReleaseTime,
                    .NextOpRec = nextOpRec,
                    .NextOpNo = nextOpNo,
                    .DueDate = ReadDueDate(preactor, ordersTable, dueDateField, nextOpRec),
                    .Priority = ReadPriority(preactor, ordersTable, priorityField, nextOpRec),
                    .WipScore = wipScore,
                    .WipStatus = wipStatus,
                    .WipRejectReason = wipRejectReason,
                    .OperationRows = rowByOpRec
})


        Next

        ' FIFO: oven exit first. Due date breaks conflict.
        Dim ranked As List(Of QueueItem) = queue _
                .OrderByDescending(Function(x) x.WipScore) _
                .ThenBy(Function(x) x.KilnAckEndTime) _
                .ThenBy(Function(x) x.DueDate) _
                .ThenBy(Function(x) x.Priority) _
                .ThenBy(Function(x) x.ParentRecord) _
                .ThenBy(Function(x) x.NextOpNo) _
                .ToList()
        If debug IsNot Nothing AndAlso debug.Enabled Then
            For i As Integer = 0 To ranked.Count - 1
                Dim item As QueueItem = ranked(i)
                debug.TraceCandidateStep(New OptimizerCandidateTraceRow With {
                    .OptimizerName = "PostFiringScheduler", .Stage = "PostFiring400Plus", .StepName = "FinalRankedQueue",
                    .OrderNo = item.OrderNo, .ParentRecordNo = item.ParentRecord, .RecordNo = item.NextOpRec,
                    .OperationNumber = item.NextOpNo, .BeforeCount = routingDt.Rows.Count, .AfterCount = ranked.Count,
                    .Included = True, .ReasonCode = SchedulerDebugReasonCodes.OK_INCLUDED,
                    .ReasonDetail = "Included in 400+ postfiring queue.", .RankScore = item.WipScore
                })
            Next
        End If
        Return ranked

    End Function

    Public Function ScheduleQueue(preactor As IPreactor,
                                  planningboard As IPlanningBoard,
                                  queue As List(Of QueueItem),
                                  Optional debug As SchedulerDebugCollector = Nothing) As Integer

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

                    If SharedHelpers.IsCompletedOrActualizedOp(item.OperationRows, opRec) Then
                        Dim releaseTime As DateTime =
                            SharedHelpers.GetOperationReleaseTime(item.OperationRows, opRec)

                        If releaseTime <> DateTime.MinValue Then
                            testFrom = releaseTime
                        End If

                        opRec = planningboard.GetNextOperation(opRec, 1)
                        Continue While
                    End If

                    Dim liveOpNo As Integer =
                        preactor.ReadFieldInt(ordersTable, opNoField, opRec)

                    If liveOpNo < POSTFIRING_START_OP_NO Then
                        Exit While
                    End If

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
                            Dim trace As ScheduleAttemptTraceRow = CreateAttempt(debug, item, opRec, liveOpNo, bestResRec, bestTimes.Value.ChangeStart)
                            Try
                                planningboard.PutOperationOnResource(opRec,
                                                                 bestResRec,
                                                                 bestTimes.Value.ChangeStart)
                                CompleteAttempt(planningboard, trace, opRec, True, Nothing)
                            Catch ex As Exception
                                CompleteAttempt(planningboard, trace, opRec, False, ex)
                                Throw
                            End Try
                            scheduledCount += 1
                        End If

                        testFrom = GetScheduledEnd(planningboard,
                                                   opRec,
                                                   bestTimes.Value.ProcessEnd)
                    Else
                        System.Diagnostics.Debug.WriteLine("PostFiring: no feasible resource. Order=" &
                                        item.OrderNo &
                                        ", OpRec=" & opRec &
                                        ", OpNo=" & liveOpNo)
                        TraceCandidate(debug,
                                       item,
                                       opRec,
                                       liveOpNo,
                                       False,
                                       SchedulerDebugReasonCodes.POSTFIRING_NO_FEASIBLE_RESOURCE,
                                       "No feasible resource was found for 400+ postfiring placement.")
                    End If

                    opRec = planningboard.GetNextOperation(opRec, 1)
                End While

            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("PostFiring failed. Order=" &
                                item.OrderNo &
                                ", OpRec=" & item.NextOpRec &
                                ", Error=" & ex.Message)
            End Try

        Next

        Return scheduledCount

    End Function

    Private Function FindNextPostFiringCandidate(preactor As IPreactor,
                                                 planningboard As IPlanningBoard,
                                                 ordersTable As Integer,
                                                 opNoField As Integer,
                                                 rowByOpRec As IDictionary(Of Integer, DataRow),
                                                 boundary As SharedHelpers.ReleaseBoundaryInfo,
                                                 debug As SchedulerDebugCollector) As Integer

        Dim nextOpRec As Integer = planningboard.GetNextOperation(boundary.OpRec, 1)
        If nextOpRec <= 0 Then
            TraceBoundary(debug,
                          boundary,
                          False,
                          SchedulerDebugReasonCodes.POSTFIRING_NO_NEXT_OPERATION,
                          "Released boundary has no next operation in the live planning-board chain.")
            Return 0
        End If

        While nextOpRec > 0

            Dim nextOpNo As Integer = 0
            Try
                nextOpNo = preactor.ReadFieldInt(ordersTable, opNoField, nextOpRec)
            Catch
                TraceBoundary(debug,
                              boundary,
                              False,
                              SchedulerDebugReasonCodes.DATA_MISSING_OPERATION_NUMBER,
                              "Unable to read next operation number. NextOpRec=" &
                              nextOpRec.ToString(Globalization.CultureInfo.InvariantCulture) & ".")
                Return 0
            End Try

            If planningboard.IsOperationScheduled(nextOpRec) Then
                TraceBoundary(debug,
                              boundary,
                              False,
                              SchedulerDebugReasonCodes.POSTFIRING_NEXT_ALREADY_SCHEDULED,
                              "Next operation is already scheduled; advancing. NextOpRec=" &
                              nextOpRec.ToString(Globalization.CultureInfo.InvariantCulture) &
                              "; NextOpNo=" &
                              nextOpNo.ToString(Globalization.CultureInfo.InvariantCulture) & ".")
                nextOpRec = planningboard.GetNextOperation(nextOpRec, 1)
                Continue While
            End If

            If SharedHelpers.IsCompletedOrActualizedOp(rowByOpRec, nextOpRec) Then
                TraceBoundary(debug,
                              boundary,
                              False,
                              SchedulerDebugReasonCodes.POSTFIRING_NEXT_COMPLETED,
                              "Next operation is completed/actualized; advancing. NextOpRec=" &
                              nextOpRec.ToString(Globalization.CultureInfo.InvariantCulture) &
                              "; NextOpNo=" &
                              nextOpNo.ToString(Globalization.CultureInfo.InvariantCulture) & ".")
                nextOpRec = planningboard.GetNextOperation(nextOpRec, 1)
                Continue While
            End If

            If nextOpNo < POSTFIRING_START_OP_NO Then
                TraceBoundary(debug,
                              boundary,
                              False,
                              SchedulerDebugReasonCodes.POSTFIRING_NEXT_BELOW_RANGE_BLOCKED,
                              "First pending operation is below 400, so 400+ is not released. NextOpRec=" &
                              nextOpRec.ToString(Globalization.CultureInfo.InvariantCulture) &
                              "; NextOpNo=" &
                              nextOpNo.ToString(Globalization.CultureInfo.InvariantCulture) & ".")
                Return 0
            End If

            TraceBoundary(debug,
                          boundary,
                          True,
                          SchedulerDebugReasonCodes.OK_INCLUDED,
                          "Selected next pending 400+ operation. NextOpRec=" &
                          nextOpRec.ToString(Globalization.CultureInfo.InvariantCulture) &
                          "; NextOpNo=" &
                          nextOpNo.ToString(Globalization.CultureInfo.InvariantCulture) & ".")
            Return nextOpRec

        End While

        TraceBoundary(debug,
                      boundary,
                      False,
                      SchedulerDebugReasonCodes.POSTFIRING_NO_NEXT_OPERATION,
                      "No pending 400+ operation found after released boundary.")
        Return 0

    End Function

    Private Sub TraceBoundary(debug As SchedulerDebugCollector,
                              boundary As SharedHelpers.ReleaseBoundaryInfo,
                              included As Boolean,
                              reasonCode As String,
                              reasonDetail As String)

        If debug Is Nothing OrElse Not debug.Enabled Then Return
        If boundary Is Nothing Then Return

        debug.TraceCandidateStep(New OptimizerCandidateTraceRow With {
            .OptimizerName = "PostFiringScheduler",
            .Stage = "PostFiring400Plus",
            .StepName = "LatestReleaseBoundary",
            .OrderNo = boundary.OrderNo,
            .ParentRecordNo = boundary.ParentRecord,
            .RecordNo = boundary.OpRec,
            .OperationNumber = boundary.OpNo,
            .BeforeCount = 0,
            .AfterCount = 0,
            .Included = included,
            .ReasonCode = reasonCode,
            .ReasonDetail = reasonDetail &
                            " GroupKey=" & boundary.GroupKey &
                            "; ReleasedOpRec=" & boundary.OpRec.ToString(Globalization.CultureInfo.InvariantCulture) &
                            "; ReleasedOpNo=" & boundary.OpNo.ToString(Globalization.CultureInfo.InvariantCulture) &
                            "; ReleaseTime=" & FormatDateForTrace(boundary.ReleaseTime) & ".",
            .RankScore = boundary.WipScore,
            .RankBreakdown = ""
        })

    End Sub

    Private Sub TraceCandidate(debug As SchedulerDebugCollector,
                               item As QueueItem,
                               opRec As Integer,
                               opNo As Integer,
                               included As Boolean,
                               reasonCode As String,
                               reasonDetail As String)

        If debug Is Nothing OrElse Not debug.Enabled Then Return
        If item Is Nothing Then Return

        debug.TraceCandidateStep(New OptimizerCandidateTraceRow With {
            .OptimizerName = "PostFiringScheduler",
            .Stage = "PostFiring400Plus",
            .StepName = "ScheduleTraversal",
            .OrderNo = item.OrderNo,
            .ParentRecordNo = item.ParentRecord,
            .RecordNo = opRec,
            .OperationNumber = opNo,
            .BeforeCount = 0,
            .AfterCount = 0,
            .Included = included,
            .ReasonCode = reasonCode,
            .ReasonDetail = reasonDetail,
            .RankScore = item.WipScore,
            .RankBreakdown = ""
        })

    End Sub

    Private Function CreateAttempt(debug As SchedulerDebugCollector, item As QueueItem, opRec As Integer,
                                   opNo As Integer, resourceRec As Integer, startTime As DateTime) As ScheduleAttemptTraceRow
        If debug Is Nothing OrElse Not debug.Enabled Then Return Nothing
        Dim row As New ScheduleAttemptTraceRow With {
            .Stage = "PostFiring400Plus", .OrderNo = item.OrderNo, .ParentRecordNo = item.ParentRecord,
            .RecordNo = opRec, .OperationNumber = opNo, .RequestedResource = resourceRec.ToString(),
            .RequestedStartTime = startTime, .SchedulingDirection = "Forward", .WasAttempted = True
        }
        debug.TraceScheduleAttempt(row)
        Return row
    End Function

    Private Sub CompleteAttempt(planningboard As IPlanningBoard, row As ScheduleAttemptTraceRow,
                                opRec As Integer, expectedSuccess As Boolean, ex As Exception)
        If row Is Nothing Then Return
        If ex IsNot Nothing Then
            row.ExceptionType = ex.GetType().FullName
            row.ExceptionMessage = ex.Message
            row.FailureReasonCode = SchedulerDebugReasonCodes.SCHEDULE_EXCEPTION_THROWN
            row.FailureReasonDetail = ex.Message
            Return
        End If
        row.ScheduledAfterAttempt = planningboard.IsOperationScheduled(opRec)
        row.PlanningBoardResultCode = If(row.ScheduledAfterAttempt, 0, -1)
        row.PlanningBoardResultMeaning = If(row.ScheduledAfterAttempt, "Scheduled", "Operation remained unscheduled")
        row.FailureReasonCode = If(row.ScheduledAfterAttempt, SchedulerDebugReasonCodes.OK_SCHEDULED,
                                   SchedulerDebugReasonCodes.SCHEDULE_RESULT_NOT_SCHEDULED)
        Dim times As Nullable(Of OperationResourceTimes) = planningboard.GetOperationTimes(opRec)
        If times.HasValue Then
            row.ActualStartTime = times.Value.OperationTimes.ProcessStart
            row.ActualEndTime = times.Value.OperationTimes.ProcessEnd
        End If
    End Sub

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

    Private Function FormatDateForTrace(value As DateTime) As String
        If value = DateTime.MinValue Then Return ""
        Return value.ToString("yyyy-MM-dd HH:mm:ss", Globalization.CultureInfo.InvariantCulture)
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

End Class
