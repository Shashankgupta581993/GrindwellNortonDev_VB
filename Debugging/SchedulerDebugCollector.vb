Option Strict On
Option Explicit On

Imports System
Imports System.Collections.Generic
Imports System.IO
Imports Preactor

Public Class SchedulerDebugCollector
    Private _attemptNo As Integer
    Public ReadOnly Property RunId As String
    Public ReadOnly Property ExportedAt As DateTime
    Public Property Enabled As Boolean
    Public Property ExportFolder As String
    Public FieldMapRows As New List(Of DebugFieldMapRow)
    Public OperationSnapshots As New List(Of OperationSnapshot)
    Public WipDiagnostics As New List(Of WipDiagnosticRow)
    Public StageEligibilityRows As New List(Of StageEligibilityRow)
    Public OptimizerCandidateTraceRows As New List(Of OptimizerCandidateTraceRow)
    Public BatchTunnelSwkPlanTraceRows As New List(Of BatchTunnelSwkPlanTraceRow)
    Public ScheduleAttemptTraceRows As New List(Of ScheduleAttemptTraceRow)
    Public ResourceValidationRows As New List(Of ResourceValidationRow)

    Public Sub New()
        ExportedAt = DateTime.Now
        RunId = Guid.NewGuid().ToString("N")
        Enabled = True
        ExportFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                                    "GN_Opcenter_Debug",
                                    ExportedAt.ToString("yyyy-MM-dd_HHmmss"))
    End Sub

    Public Function NextAttemptNo() As Integer
        _attemptNo += 1
        Return _attemptNo
    End Function

    Public Sub AddFieldMap(row As DebugFieldMapRow)
        If row Is Nothing Then Return
        row.RunId = RunId
        FieldMapRows.Add(row)
    End Sub
    Public Sub AddSnapshot(row As OperationSnapshot)
        If row IsNot Nothing Then OperationSnapshots.Add(row)
    End Sub
    Public Sub AddWipDiagnostic(row As WipDiagnosticRow)
        If row IsNot Nothing Then WipDiagnostics.Add(row)
    End Sub
    Public Sub AddStageEligibility(row As StageEligibilityRow)
        If row IsNot Nothing Then StageEligibilityRows.Add(row)
    End Sub
    Public Sub TraceCandidateStep(row As OptimizerCandidateTraceRow)
        If row IsNot Nothing Then
            row.RunId = RunId
            OptimizerCandidateTraceRows.Add(row)
        End If
    End Sub
    Public Sub TracePlanStep(row As BatchTunnelSwkPlanTraceRow)
        If row IsNot Nothing Then
            row.RunId = RunId
            BatchTunnelSwkPlanTraceRows.Add(row)
        End If
    End Sub
    Public Sub TraceScheduleAttempt(row As ScheduleAttemptTraceRow)
        If row IsNot Nothing Then
            row.RunId = RunId
            If row.AttemptNo <= 0 Then row.AttemptNo = NextAttemptNo()
            ScheduleAttemptTraceRows.Add(row)
        End If
    End Sub
    Public Sub AddResourceValidation(row As ResourceValidationRow)
        If row IsNot Nothing Then
            row.RunId = RunId
            ResourceValidationRows.Add(row)
        End If
    End Sub
    Public Sub ExportAll(preactor As IPreactor)
        SchedulerDebugExporter.ExportAll(Me, preactor)
    End Sub
End Class
