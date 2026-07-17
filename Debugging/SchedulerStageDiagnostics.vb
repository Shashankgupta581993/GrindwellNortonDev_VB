Option Strict On
Option Explicit On

Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports Preactor

Public Class SchedulerStageDiagnostics
    Public Shared Function BuildOrderOperationSnapshot(preactor As IPreactor,
                                                       debug As SchedulerDebugCollector) As List(Of OperationSnapshot)
        Dim cache As New SchedulerDebugFieldCache(preactor)
        debug.FieldMapRows.Clear()
        For Each fieldRow In cache.FieldMapRows
            debug.AddFieldMap(fieldRow)
        Next

        Dim result As New List(Of OperationSnapshot)()
        Dim board As IPlanningBoard = preactor.PlanningBoard
        Dim count As Integer = preactor.RecordCount(cache.OrdersFormatNumber)
        For rec As Integer = 1 To count
            Dim startTime As DateTime? = FirstDate(cache.ReadDateNullable(rec, "scheduled_start_time"),
                                                   cache.ReadDateNullable(rec, "Start Time"))
            Dim endTime As DateTime? = FirstDate(cache.ReadDateNullable(rec, "scheduled_end_time"),
                                                 cache.ReadDateNullable(rec, "End Time"))
            Dim isScheduled As Boolean
            If cache.HasField("is_scheduled") Then
                isScheduled = cache.ReadBool(rec, "is_scheduled")
            ElseIf startTime.HasValue OrElse endTime.HasValue Then
                isScheduled = True
            Else
                Try
                    isScheduled = board.IsOperationScheduled(rec)
                Catch
                    isScheduled = False
                End Try
            End If

            Dim row As New OperationSnapshot With {
                .RunId = debug.RunId,
                .ExportedAt = debug.ExportedAt,
                .RecordNo = rec,
                .ParentRecordNo = cache.ReadInt(rec, "parent_record"),
                .OrderNo = cache.ReadString(rec, "Order No").Trim(),
                .OperationNumber = cache.ReadInt(rec, "Operation Number"),
                .OperationName = cache.ReadString(rec, "Operation Name").Trim(),
                .ResourceGroup = cache.ReadString(rec, "Resource Group").Trim(),
                .RequiredResource = cache.ReadString(rec, "Required Resource").Trim(),
                .KilnType = cache.ReadString(rec, "Kiln Type").Trim(),
                .CycleType = cache.ReadString(rec, "Cycle Type").Trim(),
                .VolumeOccupancy = cache.ReadDouble(rec, "Volume Occupancy"),
                .Quantity = cache.ReadDouble(rec, "Quantity"),
                .DueDate = cache.ReadDateNullable(rec, "Due Date"),
                .EarliestStart = cache.ReadDateNullable(rec, "Earliest Start Date"),
                .ScheduledStartTime = startTime,
                .ScheduledEndTime = endTime,
                .IsScheduled = isScheduled,
                .IsDisabled = cache.ReadBool(rec, "Disable Op"),
                .IsComplete = cache.ReadBool(rec, "Complete"),
                .Show = If(cache.HasField("Show"), cache.ReadBool(rec, "Show"), True),
                .TableAttribute1 = cache.ReadString(rec, "Table Attribute 1"),
                .TableAttribute2 = cache.ReadString(rec, "Table Attribute 2"),
                .TableAttribute3 = cache.ReadString(rec, "Table Attribute 3"),
                .WheelDia = cache.ReadString(rec, "Wheel Dia"),
                .WheelThickness = cache.ReadString(rec, "Wheel thickness")
            }
            result.Add(row)
        Next

        For Each group In result.GroupBy(Function(x) If(x.ParentRecordNo > 0, "P:" & x.ParentRecordNo.ToString(), "O:" & x.OrderNo),
                                         StringComparer.OrdinalIgnoreCase)
            Dim rows As List(Of OperationSnapshot) = group.OrderBy(Function(x) x.OperationNumber).ThenBy(Function(x) x.RecordNo).ToList()
            Dim derivedParent As Integer = rows(0).RecordNo
            For i As Integer = 0 To rows.Count - 1
                If rows(i).ParentRecordNo <= 0 AndAlso rows(i).OrderNo.Length > 0 Then rows(i).ParentRecordNo = derivedParent
                If i > 0 Then
                    rows(i).PrevOperationRecordNo = rows(i - 1).RecordNo
                    rows(i).PrevOperationNumber = rows(i - 1).OperationNumber
                    rows(i).PrevOperationIsScheduled = rows(i - 1).IsScheduled
                    rows(i).PrevOperationEndTime = rows(i - 1).ScheduledEndTime
                End If
                If i < rows.Count - 1 Then
                    rows(i).NextOperationRecordNo = rows(i + 1).RecordNo
                    rows(i).NextOperationNumber = rows(i + 1).OperationNumber
                End If
            Next
        Next

        debug.OperationSnapshots.Clear()
        For Each row In result
            debug.AddSnapshot(row)
        Next
        Return result
    End Function

    Public Shared Sub DiagnoseWip(snapshot As List(Of OperationSnapshot), debug As SchedulerDebugCollector)
        debug.WipDiagnostics.Clear()
        Dim byRecord As Dictionary(Of Integer, OperationSnapshot) = snapshot.ToDictionary(Function(x) x.RecordNo)
        For Each op In snapshot
            Dim status As String = "READY_FOR_THIS_OPERATION"
            Dim code As String = SchedulerDebugReasonCodes.OK_INCLUDED
            Dim detail As String = "Operation is ready for stage eligibility checks."

            If op.OperationNumber <= 0 Then
                status = "INVALID_DATA" : code = SchedulerDebugReasonCodes.DATA_MISSING_OPERATION_NUMBER : detail = "Operation Number is missing or zero."
            ElseIf op.ParentRecordNo <= 0 Then
                status = "INVALID_DATA" : code = SchedulerDebugReasonCodes.DATA_MISSING_PARENT : detail = "Parent record could not be resolved."
            ElseIf op.IsScheduled Then
                status = "ALREADY_SCHEDULED" : code = SchedulerDebugReasonCodes.OK_ALREADY_SCHEDULED : detail = "Operation is already scheduled."
            ElseIf op.PrevOperationRecordNo > 0 AndAlso Not op.PrevOperationIsScheduled Then
                status = "BLOCKED_BY_PREVIOUS_OPERATION" : code = SchedulerDebugReasonCodes.WIP_PREVIOUS_OPERATION_NOT_SCHEDULED : detail = "Previous operation is not scheduled."
            ElseIf op.PrevOperationRecordNo > 0 AndAlso op.PrevOperationIsScheduled AndAlso Not op.PrevOperationEndTime.HasValue Then
                status = "BLOCKED_BY_PREVIOUS_OPERATION" : code = SchedulerDebugReasonCodes.WIP_PREVIOUS_OPERATION_END_MISSING : detail = "Previous operation is scheduled but has no end time."
            ElseIf op.PrevOperationRecordNo > 0 AndAlso Not byRecord.ContainsKey(op.PrevOperationRecordNo) Then
                status = "BLOCKED_BY_MISSING_PREVIOUS_OPERATION" : code = SchedulerDebugReasonCodes.WIP_CHAIN_BROKEN : detail = "Previous operation record is absent from the snapshot."
            End If

            Dim nextOp As OperationSnapshot = Nothing
            If Not op.IsScheduled AndAlso op.NextOperationRecordNo > 0 AndAlso
               byRecord.TryGetValue(op.NextOperationRecordNo, nextOp) AndAlso nextOp.IsScheduled Then
                status = "CHAIN_INCONSISTENT" : code = SchedulerDebugReasonCodes.WIP_NEXT_OPERATION_ALREADY_SCHEDULED : detail = "A later operation is scheduled while this operation is unscheduled."
            End If

            debug.AddWipDiagnostic(New WipDiagnosticRow With {
                .RunId = debug.RunId, .OrderNo = op.OrderNo, .ParentRecordNo = op.ParentRecordNo,
                .RecordNo = op.RecordNo, .OperationNumber = op.OperationNumber, .OperationName = op.OperationName,
                .Status = status, .ReasonCode = code, .ReasonDetail = detail,
                .PreviousOperationRecordNo = op.PrevOperationRecordNo, .PreviousOperationNumber = op.PrevOperationNumber,
                .PreviousOperationScheduled = op.PrevOperationIsScheduled, .PreviousOperationEndTime = op.PrevOperationEndTime,
                .NextOperationRecordNo = op.NextOperationRecordNo, .NextOperationNumber = op.NextOperationNumber
            })
        Next
    End Sub

    Public Shared Sub DiagnosePressingEligibility(snapshot As List(Of OperationSnapshot), debug As SchedulerDebugCollector)
        DiagnoseSimpleStage(snapshot.Where(Function(x) x.OperationNumber >= 200 AndAlso x.OperationNumber <= 280),
                            debug, "Pressing", True, True)
    End Sub
    Public Shared Sub DiagnoseDryingEligibility(snapshot As List(Of OperationSnapshot), debug As SchedulerDebugCollector)
        DiagnoseSimpleStage(snapshot.Where(Function(x) x.OperationNumber > 200 AndAlso x.OperationNumber < 290 AndAlso
                                                       x.OperationNumber <> 290 AndAlso x.OperationNumber <> 300 AndAlso x.OperationNumber <> 310),
                            debug, "Drying", True, False)
    End Sub
    Public Shared Sub DiagnoseBatchFiringEligibility(snapshot As List(Of OperationSnapshot), debug As SchedulerDebugCollector)
        DiagnoseFiring(snapshot, debug, "BatchFiring", Function(k) IsKiln(k, "BATCH", "1"))
    End Sub
    Public Shared Sub DiagnoseTunnelFiringEligibility(snapshot As List(Of OperationSnapshot), debug As SchedulerDebugCollector)
        DiagnoseFiring(snapshot, debug, "TunnelFiring", Function(k) IsKiln(k, "TUNNEL", "2"))
    End Sub
    Public Shared Sub DiagnoseSwkFiringEligibility(snapshot As List(Of OperationSnapshot), debug As SchedulerDebugCollector)
        DiagnoseFiring(snapshot, debug, "SWK", Function(k) IsKiln(k, "SWK", "3"))
    End Sub
    Public Shared Sub DiagnosePostFiringEligibility(snapshot As List(Of OperationSnapshot), debug As SchedulerDebugCollector)
        DiagnoseSimpleStage(snapshot.Where(Function(x) x.OperationNumber >= 320 AndAlso x.OperationNumber < 400),
                            debug, "FiringFollowOn", True, False)
        DiagnoseSimpleStage(snapshot.Where(Function(x) x.OperationNumber >= 400),
                            debug, "PostFiring400Plus", True, False)
    End Sub

    Public Shared Sub DiagnoseAll(snapshot As List(Of OperationSnapshot), debug As SchedulerDebugCollector)
        debug.StageEligibilityRows.Clear()
        DiagnosePressingEligibility(snapshot, debug)
        DiagnoseDryingEligibility(snapshot, debug)
        DiagnoseBatchFiringEligibility(snapshot, debug)
        DiagnoseTunnelFiringEligibility(snapshot, debug)
        DiagnoseSwkFiringEligibility(snapshot, debug)
        DiagnosePostFiringEligibility(snapshot, debug)
    End Sub

    Private Shared Sub DiagnoseSimpleStage(ops As IEnumerable(Of OperationSnapshot), debug As SchedulerDebugCollector,
                                           stage As String, requirePrevious As Boolean, requirePressAttributes As Boolean)
        Dim rank As Integer
        For Each op In ops.OrderBy(Function(x) x.DueDate).ThenBy(Function(x) x.OperationNumber)
            rank += 1
            Dim code As String = SchedulerDebugReasonCodes.OK_INCLUDED
            Dim detail As String = "Eligible for optimizer consideration."
            Dim candidate As Boolean = True
            If op.IsScheduled Then
                candidate = False : code = SchedulerDebugReasonCodes.OK_ALREADY_SCHEDULED : detail = "Already scheduled."
            ElseIf requirePrevious AndAlso op.PrevOperationRecordNo > 0 AndAlso Not op.PrevOperationIsScheduled Then
                candidate = False : code = If(stage = "PostFiring400Plus", SchedulerDebugReasonCodes.POSTFIRING_PREV_OP_NOT_READY, SchedulerDebugReasonCodes.PRESSING_PREV_OP_NOT_READY) : detail = "Previous operation is not ready."
            ElseIf String.IsNullOrWhiteSpace(op.RequiredResource) AndAlso String.IsNullOrWhiteSpace(op.ResourceGroup) Then
                candidate = False : code = If(stage = "PostFiring400Plus", SchedulerDebugReasonCodes.POSTFIRING_RESOURCE_MISSING, SchedulerDebugReasonCodes.DATA_MISSING_REQUIRED_RESOURCE) : detail = "Required Resource and Resource Group are blank."
            ElseIf requirePressAttributes AndAlso String.IsNullOrWhiteSpace(op.WheelDia) Then
                candidate = False : code = SchedulerDebugReasonCodes.DATA_MISSING_OPERATION : detail = "Wheel Dia is blank."
            ElseIf requirePressAttributes AndAlso String.IsNullOrWhiteSpace(op.WheelThickness) Then
                candidate = False : code = SchedulerDebugReasonCodes.DATA_MISSING_OPERATION : detail = "Wheel thickness is blank."
            End If
            AddEligibility(debug, op, stage, rank, candidate, code, detail)
        Next
    End Sub

    Private Shared Sub DiagnoseFiring(snapshot As List(Of OperationSnapshot), debug As SchedulerDebugCollector,
                                      stage As String, kilnMatch As Func(Of String, Boolean))
        Dim rank As Integer
        For Each op In snapshot.Where(Function(x) x.OperationNumber = 300 AndAlso kilnMatch(x.KilnType)).
                                OrderBy(Function(x) x.DueDate)
            rank += 1
            Dim chain = snapshot.Where(Function(x) x.ParentRecordNo = op.ParentRecordNo).ToList()
            Dim code As String = SchedulerDebugReasonCodes.OK_INCLUDED
            Dim detail As String = "Eligible for firing optimizer consideration."
            Dim candidate As Boolean = True
            If op.IsScheduled Then
                candidate = False : code = SchedulerDebugReasonCodes.OK_ALREADY_SCHEDULED : detail = "Already scheduled."
            ElseIf Not chain.Any(Function(x) x.OperationNumber = 290) Then
                candidate = False : code = SchedulerDebugReasonCodes.FIRING_LOADING_OP_MISSING : detail = "Operation 290 is missing."
            ElseIf Not chain.Any(Function(x) x.OperationNumber = 310) Then
                candidate = False : code = SchedulerDebugReasonCodes.FIRING_UNLOADING_OP_MISSING : detail = "Operation 310 is missing."
            ElseIf op.PrevOperationRecordNo > 0 AndAlso Not op.PrevOperationIsScheduled Then
                candidate = False : code = SchedulerDebugReasonCodes.FIRING_PREV_OP_NOT_READY : detail = "Previous operation is not scheduled."
            ElseIf String.IsNullOrWhiteSpace(op.CycleType) Then
                candidate = False : code = SchedulerDebugReasonCodes.DATA_MISSING_CYCLE_TYPE : detail = "Cycle Type is blank."
            ElseIf op.VolumeOccupancy <= 0 Then
                candidate = False : code = SchedulerDebugReasonCodes.DATA_INVALID_OCCUPANCY : detail = "Volume Occupancy must be greater than zero."
            ElseIf String.IsNullOrWhiteSpace(op.RequiredResource) AndAlso String.IsNullOrWhiteSpace(op.ResourceGroup) Then
                candidate = False
                code = If(stage = "TunnelFiring", SchedulerDebugReasonCodes.TUNNEL_RESOURCE_MISSING,
                          If(stage = "SWK", SchedulerDebugReasonCodes.SWK_RESOURCE_MISSING, SchedulerDebugReasonCodes.FIRING_NO_KILN_AVAILABLE))
                detail = "Kiln resource cannot be resolved from operation data."
            End If
            AddEligibility(debug, op, stage, rank, candidate, code, detail)
        Next
    End Sub

    Private Shared Sub AddEligibility(debug As SchedulerDebugCollector, op As OperationSnapshot, stage As String,
                                      rank As Integer, candidate As Boolean, code As String, detail As String)
        debug.AddStageEligibility(New StageEligibilityRow With {
            .RunId = debug.RunId, .Stage = stage, .OrderNo = op.OrderNo, .ParentRecordNo = op.ParentRecordNo,
            .RecordNo = op.RecordNo, .OperationNumber = op.OperationNumber, .IsCandidate = candidate,
            .CandidateRank = rank, .IncludedInOptimizer = candidate, .ExcludedReasonCode = code,
            .ExcludedReasonDetail = detail, .RequiredResource = op.RequiredResource, .ResourceGroup = op.ResourceGroup,
            .KilnType = op.KilnType, .CycleType = op.CycleType, .VolumeOccupancy = op.VolumeOccupancy,
            .DueDate = op.DueDate, .EarliestAllowedStart = op.EarliestStart,
            .PreviousOperationEndTime = op.PrevOperationEndTime,
            .WipStatus = If(candidate, "READY_FOR_THIS_OPERATION", "BLOCKED")
        })
    End Sub

    Private Shared Function IsKiln(value As String, name As String, number As String) As Boolean
        Dim normalized As String = If(value, "").Trim()
        Return normalized.Equals(name, StringComparison.OrdinalIgnoreCase) OrElse normalized = number
    End Function

    Private Shared Function FirstDate(first As DateTime?, second As DateTime?) As DateTime?
        If first.HasValue Then Return first
        Return second
    End Function
End Class
