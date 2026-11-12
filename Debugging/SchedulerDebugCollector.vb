Option Strict On
Option Explicit On

Imports System
Imports System.Collections.Generic
Imports System.IO
Imports Preactor


' Edit these Boolean values, rebuild, and reload the DLL in Opcenter.
' The master overrides every channel. No environment/config-file overrides apply.
Friend Module SchedulerDebugSettings
    ' BEGIN EDITABLE DEBUGGER SETTINGS
    Friend ReadOnly MasterEnabled As Boolean = False
    Friend ReadOnly ActionMetricsEnabled As Boolean = True
    Friend ReadOnly FieldMapEnabled As Boolean = True
    Friend ReadOnly OperationSnapshotsEnabled As Boolean = True
    Friend ReadOnly WipDiagnosticsEnabled As Boolean = True
    Friend ReadOnly StageEligibilityEnabled As Boolean = True
    Friend ReadOnly OptimizerCandidateTraceEnabled As Boolean = True
    Friend ReadOnly BatchPlanTraceEnabled As Boolean = True
    Friend ReadOnly ScheduleAttemptTraceEnabled As Boolean = True
    Friend ReadOnly ResourceValidationEnabled As Boolean = True
    Friend ReadOnly ReasonSummaryEnabled As Boolean = True
    Friend ReadOnly ProblemStatementExportEnabled As Boolean = True
    Friend ReadOnly BatchDebugCsvEnabled As Boolean = True
    Friend ReadOnly SwkDebugCsvEnabled As Boolean = True
    Friend ReadOnly AlgoSeq3DecisionLogEnabled As Boolean = True
    Friend ReadOnly AlgoSeq3ScheduleLogEnabled As Boolean = True
    Friend ReadOnly DebugOutputEnabled As Boolean = True
    Friend ReadOnly DiagnosticPopupsEnabled As Boolean = True

    ' END EDITABLE DEBUGGER SETTINGS

    Friend Function IsEnabled(channel As SchedulerDebugChannel) As Boolean
        If Not MasterEnabled Then Return False
        Select Case channel
            Case SchedulerDebugChannel.ActionMetrics : Return ActionMetricsEnabled
            Case SchedulerDebugChannel.FieldMap : Return FieldMapEnabled
            Case SchedulerDebugChannel.OperationSnapshots : Return OperationSnapshotsEnabled
            Case SchedulerDebugChannel.WipDiagnostics : Return WipDiagnosticsEnabled
            Case SchedulerDebugChannel.StageEligibility : Return StageEligibilityEnabled
            Case SchedulerDebugChannel.OptimizerCandidateTrace : Return OptimizerCandidateTraceEnabled
            Case SchedulerDebugChannel.BatchPlanTrace : Return BatchPlanTraceEnabled
            Case SchedulerDebugChannel.ScheduleAttemptTrace : Return ScheduleAttemptTraceEnabled
            Case SchedulerDebugChannel.ResourceValidation : Return ResourceValidationEnabled
            Case SchedulerDebugChannel.ReasonSummary : Return ReasonSummaryEnabled
            Case SchedulerDebugChannel.ProblemStatementExport : Return ProblemStatementExportEnabled
            Case SchedulerDebugChannel.BatchDebugCsv : Return BatchDebugCsvEnabled
            Case SchedulerDebugChannel.SwkDebugCsv : Return SwkDebugCsvEnabled
            Case SchedulerDebugChannel.AlgoSeq3DecisionLog : Return AlgoSeq3DecisionLogEnabled
            Case SchedulerDebugChannel.AlgoSeq3ScheduleLog : Return AlgoSeq3ScheduleLogEnabled
            Case SchedulerDebugChannel.DebugOutput : Return DebugOutputEnabled
            Case SchedulerDebugChannel.DiagnosticPopups : Return DiagnosticPopupsEnabled
            Case Else : Return False
        End Select
    End Function
End Module

Friend Enum SchedulerDebugChannel
    ActionMetrics
    FieldMap
    OperationSnapshots
    WipDiagnostics
    StageEligibility
    OptimizerCandidateTrace
    BatchPlanTrace
    ScheduleAttemptTrace
    ResourceValidation
    ReasonSummary
    ProblemStatementExport
    BatchDebugCsv
    SwkDebugCsv
    AlgoSeq3DecisionLog
    AlgoSeq3ScheduleLog
    DebugOutput
    DiagnosticPopups
End Enum

Public Class SchedulerDebugCollector
    Private _attemptNo As Integer
    Private _operationSnapshotByRecord As Dictionary(Of Integer, OperationSnapshot)
    Public ReadOnly Property RunId As String
    Public ReadOnly Property ExportedAt As DateTime
    Private _enabled As Boolean = True
    Private _metricsEnabled As Boolean = True

    Public Property Enabled As Boolean
        Get
            Return _enabled AndAlso IsEnabledByConfiguration()
        End Get
        Set(value As Boolean)
            _enabled = value
        End Set
    End Property

    Public Property MetricsEnabled As Boolean
        Get
            Return _metricsEnabled AndAlso IsMetricsEnabledByConfiguration()
        End Get
        Set(value As Boolean)
            _metricsEnabled = value
        End Set
    End Property

    Friend Function Collects(channel As SchedulerDebugChannel) As Boolean
        If channel = SchedulerDebugChannel.ActionMetrics Then Return MetricsEnabled
        Return Enabled AndAlso SchedulerDebugSettings.IsEnabled(channel)
    End Function

    Friend ReadOnly Property NeedsSnapshot As Boolean
        Get
            Return Collects(SchedulerDebugChannel.OperationSnapshots) OrElse
                   Collects(SchedulerDebugChannel.WipDiagnostics) OrElse
                   Collects(SchedulerDebugChannel.StageEligibility) OrElse
                   Collects(SchedulerDebugChannel.OptimizerCandidateTrace) OrElse
                   Collects(SchedulerDebugChannel.ScheduleAttemptTrace)
        End Get
    End Property
    Friend ReadOnly GeneratedFiles As New HashSet(Of String)(StringComparer.Ordinal)

    Public Property ExportFolder As String
    Public FieldMapRows As New List(Of DebugFieldMapRow)
    Public OperationSnapshots As New List(Of OperationSnapshot)
    Public WipDiagnostics As New List(Of WipDiagnosticRow)
    Public StageEligibilityRows As New List(Of StageEligibilityRow)
    Public OptimizerCandidateTraceRows As New List(Of OptimizerCandidateTraceRow)
    Public BatchTunnelSwkPlanTraceRows As New List(Of BatchTunnelSwkPlanTraceRow)
    Public ScheduleAttemptTraceRows As New List(Of ScheduleAttemptTraceRow)
    Public ResourceValidationRows As New List(Of ResourceValidationRow)
    Public ActionMetricsRows As New List(Of SchedulerActionMetricsRow)

    Public Sub New()
        ExportedAt = DateTime.Now
        RunId = Guid.NewGuid().ToString("N")
        ExportFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                                    "GN_Opcenter_Debug",
                                    ExportedAt.ToString("yyyy-MM-dd_HHmmss"))
    End Sub

    Public Shared Function IsEnabledByConfiguration() As Boolean
        Return SchedulerDebugSettings.IsEnabled(SchedulerDebugChannel.FieldMap) OrElse
               SchedulerDebugSettings.IsEnabled(SchedulerDebugChannel.OperationSnapshots) OrElse
               SchedulerDebugSettings.IsEnabled(SchedulerDebugChannel.WipDiagnostics) OrElse
               SchedulerDebugSettings.IsEnabled(SchedulerDebugChannel.StageEligibility) OrElse
               SchedulerDebugSettings.IsEnabled(SchedulerDebugChannel.OptimizerCandidateTrace) OrElse
               SchedulerDebugSettings.IsEnabled(SchedulerDebugChannel.BatchPlanTrace) OrElse
               SchedulerDebugSettings.IsEnabled(SchedulerDebugChannel.ScheduleAttemptTrace) OrElse
               SchedulerDebugSettings.IsEnabled(SchedulerDebugChannel.ResourceValidation) OrElse
               SchedulerDebugSettings.IsEnabled(SchedulerDebugChannel.ReasonSummary) OrElse
               SchedulerDebugSettings.IsEnabled(SchedulerDebugChannel.ProblemStatementExport)
    End Function

    Public Shared Function IsMetricsEnabledByConfiguration() As Boolean
        Return SchedulerDebugSettings.IsEnabled(SchedulerDebugChannel.ActionMetrics)
    End Function

    Public Function NextAttemptNo() As Integer
        _attemptNo += 1
        Return _attemptNo
    End Function

    Public Sub AddFieldMap(row As DebugFieldMapRow)
        If Not Collects(SchedulerDebugChannel.FieldMap) Then Return
        If row Is Nothing Then Return
        row.RunId = RunId
        FieldMapRows.Add(row)
    End Sub
    Public Sub AddSnapshot(row As OperationSnapshot)
        If Not NeedsSnapshot Then Return
        If row IsNot Nothing Then
            OperationSnapshots.Add(row)
            _operationSnapshotByRecord = Nothing
        End If
    End Sub

    Public Function FindOperationSnapshot(recordNo As Integer) As OperationSnapshot
        If Not NeedsSnapshot OrElse recordNo <= 0 Then Return Nothing

        If _operationSnapshotByRecord Is Nothing Then
            _operationSnapshotByRecord = New Dictionary(Of Integer, OperationSnapshot)()
            For Each snapshot As OperationSnapshot In OperationSnapshots
                If Not _operationSnapshotByRecord.ContainsKey(snapshot.RecordNo) Then
                    _operationSnapshotByRecord.Add(snapshot.RecordNo, snapshot)
                End If
            Next
        End If

        Dim result As OperationSnapshot = Nothing
        _operationSnapshotByRecord.TryGetValue(recordNo, result)
        Return result
    End Function
    Public Sub AddWipDiagnostic(row As WipDiagnosticRow)
        If Not Collects(SchedulerDebugChannel.WipDiagnostics) Then Return
        If row IsNot Nothing Then WipDiagnostics.Add(row)
    End Sub
    Public Sub AddStageEligibility(row As StageEligibilityRow)
        If Not Collects(SchedulerDebugChannel.StageEligibility) Then Return
        If row IsNot Nothing Then StageEligibilityRows.Add(row)
    End Sub
    Public Sub TraceCandidateStep(row As OptimizerCandidateTraceRow)
        If Not Collects(SchedulerDebugChannel.OptimizerCandidateTrace) Then Return
        If row IsNot Nothing Then
            row.RunId = RunId
            OptimizerCandidateTraceRows.Add(row)
        End If
    End Sub
    Public Sub TracePlanStep(row As BatchTunnelSwkPlanTraceRow)
        If Not Collects(SchedulerDebugChannel.BatchPlanTrace) Then Return
        If row IsNot Nothing Then
            row.RunId = RunId
            BatchTunnelSwkPlanTraceRows.Add(row)
        End If
    End Sub
    Public Sub TraceScheduleAttempt(row As ScheduleAttemptTraceRow)
        If Not Collects(SchedulerDebugChannel.ScheduleAttemptTrace) Then Return
        If row IsNot Nothing Then
            row.RunId = RunId
            If row.AttemptNo <= 0 Then row.AttemptNo = NextAttemptNo()
            ScheduleAttemptTraceRows.Add(row)
        End If
    End Sub
    Public Sub AddResourceValidation(row As ResourceValidationRow)
        If Not Collects(SchedulerDebugChannel.ResourceValidation) Then Return
        If row IsNot Nothing Then
            row.RunId = RunId
            ResourceValidationRows.Add(row)
        End If
    End Sub
    Public Sub AddActionMetrics(row As SchedulerActionMetricsRow)
        If Not Collects(SchedulerDebugChannel.ActionMetrics) Then Return
        If row Is Nothing Then Return
        row.RunId = RunId
        ActionMetricsRows.Add(row)
    End Sub
    Public Sub ExportAll(preactor As IPreactor)
        SchedulerDebugExporter.ExportAll(Me, preactor)
    End Sub
End Class
