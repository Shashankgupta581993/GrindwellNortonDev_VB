Option Strict On
Option Explicit On

Imports System
Imports System.Collections.Generic
Imports System.Globalization
Imports System.IO
Imports System.Linq
Imports System.Reflection
Imports System.Text
Imports Preactor

Public Class SchedulerDebugExporter
    Private Shared ReadOnly CsvSpecialCharacters As Char() = {","c, """"c, ControlChars.Cr, ControlChars.Lf}
    Public Shared Sub ExportAll(debug As SchedulerDebugCollector, preactor As IPreactor)
        If debug Is Nothing OrElse (Not debug.Enabled AndAlso Not debug.MetricsEnabled) Then Return
        Directory.CreateDirectory(debug.ExportFolder)
        debug.GeneratedFiles.Clear()

        Dim configRows As New List(Of DebugConfigSnapshotRow) From {
            New DebugConfigSnapshotRow With {.RunId = debug.RunId, .Name = "ExportedAt", .Value = debug.ExportedAt.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture)},
            New DebugConfigSnapshotRow With {.RunId = debug.RunId, .Name = "ExportFolder", .Value = debug.ExportFolder},
            New DebugConfigSnapshotRow With {.RunId = debug.RunId, .Name = "SnapshotRows", .Value = debug.OperationSnapshots.Count.ToString(CultureInfo.InvariantCulture)}
        }

        WriteRunSummary(Path.Combine(debug.ExportFolder, "00_RunSummary.md"), debug)
        debug.GeneratedFiles.Add("00_RunSummary.md")
        If debug.MetricsEnabled Then
            WriteCsv(Path.Combine(debug.ExportFolder, "12_ActionMetrics.csv"), debug.ActionMetricsRows)
            debug.GeneratedFiles.Add("12_ActionMetrics.csv")
        End If

        If debug.Enabled Then
            WriteCsv(Path.Combine(debug.ExportFolder, "01_ConfigSnapshot.csv"), configRows)
            debug.GeneratedFiles.Add("01_ConfigSnapshot.csv")
            If debug.Collects(SchedulerDebugChannel.FieldMap) Then
                WriteCsv(Path.Combine(debug.ExportFolder, "02_FieldMap.csv"), debug.FieldMapRows)
                debug.GeneratedFiles.Add("02_FieldMap.csv")
            End If
            If debug.Collects(SchedulerDebugChannel.OperationSnapshots) Then
                WriteCsv(Path.Combine(debug.ExportFolder, "03_OrderOperationSnapshot.csv"), debug.OperationSnapshots)
                debug.GeneratedFiles.Add("03_OrderOperationSnapshot.csv")
                WriteCsv(Path.Combine(debug.ExportFolder, "04_UnscheduledOperations.csv"),
                         debug.OperationSnapshots.Where(Function(x) Not x.IsScheduled))
                debug.GeneratedFiles.Add("04_UnscheduledOperations.csv")
            End If
            If debug.Collects(SchedulerDebugChannel.WipDiagnostics) Then
                WriteCsv(Path.Combine(debug.ExportFolder, "05_WipChainDiagnostics.csv"), debug.WipDiagnostics)
                debug.GeneratedFiles.Add("05_WipChainDiagnostics.csv")
            End If
            If debug.Collects(SchedulerDebugChannel.StageEligibility) Then
                WriteCsv(Path.Combine(debug.ExportFolder, "06_StageEligibilityDiagnostics.csv"), debug.StageEligibilityRows)
                debug.GeneratedFiles.Add("06_StageEligibilityDiagnostics.csv")
            End If
            If debug.Collects(SchedulerDebugChannel.OptimizerCandidateTrace) Then
                WriteCsv(Path.Combine(debug.ExportFolder, "07_OptimizerCandidateTrace.csv"), debug.OptimizerCandidateTraceRows)
                debug.GeneratedFiles.Add("07_OptimizerCandidateTrace.csv")
            End If
            If debug.Collects(SchedulerDebugChannel.BatchPlanTrace) Then
                WriteCsv(Path.Combine(debug.ExportFolder, "08_BatchTunnelSwkPlanTrace.csv"), debug.BatchTunnelSwkPlanTraceRows)
                debug.GeneratedFiles.Add("08_BatchTunnelSwkPlanTrace.csv")
            End If
            If debug.Collects(SchedulerDebugChannel.ScheduleAttemptTrace) Then
                WriteCsv(Path.Combine(debug.ExportFolder, "09_ScheduleAttemptTrace.csv"), debug.ScheduleAttemptTraceRows)
                debug.GeneratedFiles.Add("09_ScheduleAttemptTrace.csv")
            End If
            If debug.Collects(SchedulerDebugChannel.ResourceValidation) Then
                WriteCsv(Path.Combine(debug.ExportFolder, "10_ResourceValidation.csv"), debug.ResourceValidationRows)
                debug.GeneratedFiles.Add("10_ResourceValidation.csv")
            End If
            If debug.Collects(SchedulerDebugChannel.ReasonSummary) Then
                WriteReasonSummary(Path.Combine(debug.ExportFolder, "11_ReasonCodeSummary.csv"), debug)
                debug.GeneratedFiles.Add("11_ReasonCodeSummary.csv")
            End If
            If debug.Collects(SchedulerDebugChannel.ProblemStatementExport) Then
                SchedulerGptProblemStatementBuilder.WriteFiles(debug)
            End If
        End If
    End Sub

    Public Shared Sub WriteCsv(Of T)(filePath As String, rows As IEnumerable(Of T))
        If Not CanWriteType(GetType(T)) Then Return
        Dim properties As PropertyInfo() = GetType(T).GetProperties(BindingFlags.Instance Or BindingFlags.Public)
        Using writer As New StreamWriter(filePath, False, New UTF8Encoding(False))
            writer.WriteLine(String.Join(",", properties.Select(Function(p) Escape(p.Name))))
            If rows Is Nothing Then Return
            Dim values(properties.Length - 1) As String
            For Each row As T In rows
                For propertyIndex As Integer = 0 To properties.Length - 1
                    values(propertyIndex) = Escape(FormatValue(properties(propertyIndex).GetValue(row, Nothing)))
                Next
                writer.WriteLine(String.Join(",", values))
            Next
        End Using
    End Sub

    Public Shared Sub WriteRunSummary(filePath As String, debug As SchedulerDebugCollector)
        If debug Is Nothing OrElse (Not debug.Enabled AndAlso Not debug.MetricsEnabled) Then Return
        Dim unscheduled As Integer = debug.OperationSnapshots.Where(Function(x) Not x.IsScheduled).Count()
        Dim text As String =
            "# GN Opcenter Scheduling Debug Run" & Environment.NewLine & Environment.NewLine &
            "- Run ID: " & debug.RunId & Environment.NewLine &
            "- Exported at: " & debug.ExportedAt.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture) & Environment.NewLine &
            "- Operations: " & debug.OperationSnapshots.Count.ToString(CultureInfo.InvariantCulture) & Environment.NewLine &
            "- Unscheduled operations: " & unscheduled.ToString(CultureInfo.InvariantCulture) & Environment.NewLine &
            "- WIP diagnostics: " & debug.WipDiagnostics.Count.ToString(CultureInfo.InvariantCulture) & Environment.NewLine &
            "- Stage eligibility rows: " & debug.StageEligibilityRows.Count.ToString(CultureInfo.InvariantCulture) & Environment.NewLine &
            "- Optimizer trace rows: " & debug.OptimizerCandidateTraceRows.Count.ToString(CultureInfo.InvariantCulture) & Environment.NewLine &
            "- Schedule attempts: " & debug.ScheduleAttemptTraceRows.Count.ToString(CultureInfo.InvariantCulture) & Environment.NewLine &
            "- Action metric rows: " & debug.ActionMetricsRows.Count.ToString(CultureInfo.InvariantCulture) & Environment.NewLine
        File.WriteAllText(filePath, text, New UTF8Encoding(False))
    End Sub

    Public Shared Sub WriteReasonSummary(filePath As String, debug As SchedulerDebugCollector)
        If debug Is Nothing OrElse Not debug.Collects(SchedulerDebugChannel.ReasonSummary) Then Return
        Dim items As New List(Of KeyValuePair(Of String, String))()
        If debug.Collects(SchedulerDebugChannel.WipDiagnostics) Then
            For Each row In debug.WipDiagnostics
                AddReason(items, row.ReasonCode, "WIP")
            Next
        End If
        If debug.Collects(SchedulerDebugChannel.StageEligibility) Then
            For Each row In debug.StageEligibilityRows
                AddReason(items, row.ExcludedReasonCode, "Eligibility")
            Next
        End If
        If debug.Collects(SchedulerDebugChannel.OptimizerCandidateTrace) Then
            For Each row In debug.OptimizerCandidateTraceRows
                AddReason(items, row.ReasonCode, "Optimizer")
            Next
        End If
        If debug.Collects(SchedulerDebugChannel.BatchPlanTrace) Then
            For Each row In debug.BatchTunnelSwkPlanTraceRows
                AddReason(items, row.ReasonCode, "Plan")
            Next
        End If
        If debug.Collects(SchedulerDebugChannel.ScheduleAttemptTrace) Then
            For Each row In debug.ScheduleAttemptTraceRows
                AddReason(items, row.FailureReasonCode, "ScheduleAttempt")
            Next
        End If

        Dim rows As List(Of ReasonCodeSummaryRow) =
            items.GroupBy(Function(x) x.Key, StringComparer.OrdinalIgnoreCase).
            Select(Function(g) New ReasonCodeSummaryRow With {
                .ReasonCode = g.Key,
                .Count = g.Count(),
                .Sources = String.Join("|", g.Select(Function(x) x.Value).Distinct().OrderBy(Function(x) x))
            }).
            OrderByDescending(Function(x) x.Count).
            ThenBy(Function(x) x.ReasonCode).
            ToList()
        WriteCsv(filePath, rows)
    End Sub

    Private Shared Function CanWriteType(rowType As Type) As Boolean
        If Not SchedulerDebugSettings.MasterEnabled Then Return False
        If rowType Is GetType(SchedulerActionMetricsRow) Then Return SchedulerDebugSettings.IsEnabled(SchedulerDebugChannel.ActionMetrics)
        If rowType Is GetType(DebugFieldMapRow) Then Return SchedulerDebugSettings.IsEnabled(SchedulerDebugChannel.FieldMap)
        If rowType Is GetType(OperationSnapshot) Then Return SchedulerDebugSettings.IsEnabled(SchedulerDebugChannel.OperationSnapshots)
        If rowType Is GetType(WipDiagnosticRow) Then Return SchedulerDebugSettings.IsEnabled(SchedulerDebugChannel.WipDiagnostics)
        If rowType Is GetType(StageEligibilityRow) Then Return SchedulerDebugSettings.IsEnabled(SchedulerDebugChannel.StageEligibility)
        If rowType Is GetType(OptimizerCandidateTraceRow) Then Return SchedulerDebugSettings.IsEnabled(SchedulerDebugChannel.OptimizerCandidateTrace)
        If rowType Is GetType(BatchTunnelSwkPlanTraceRow) Then Return SchedulerDebugSettings.IsEnabled(SchedulerDebugChannel.BatchPlanTrace)
        If rowType Is GetType(ScheduleAttemptTraceRow) Then Return SchedulerDebugSettings.IsEnabled(SchedulerDebugChannel.ScheduleAttemptTrace)
        If rowType Is GetType(ResourceValidationRow) Then Return SchedulerDebugSettings.IsEnabled(SchedulerDebugChannel.ResourceValidation)
        If rowType Is GetType(ReasonCodeSummaryRow) Then Return SchedulerDebugSettings.IsEnabled(SchedulerDebugChannel.ReasonSummary)
        Return SchedulerDebugCollector.IsEnabledByConfiguration() OrElse SchedulerDebugCollector.IsMetricsEnabledByConfiguration()
    End Function

    Private Shared Sub AddReason(items As List(Of KeyValuePair(Of String, String)), code As String, source As String)
        If Not String.IsNullOrWhiteSpace(code) Then items.Add(New KeyValuePair(Of String, String)(code, source))
    End Sub

    Private Shared Function FormatValue(value As Object) As String
        If value Is Nothing Then Return ""
        If TypeOf value Is DateTime Then Return DirectCast(value, DateTime).ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture)
        If TypeOf value Is Boolean Then Return If(DirectCast(value, Boolean), "TRUE", "FALSE")
        If TypeOf value Is IFormattable Then Return DirectCast(value, IFormattable).ToString(Nothing, CultureInfo.InvariantCulture)
        Return value.ToString()
    End Function

    Private Shared Function Escape(value As String) As String
        If value Is Nothing Then Return ""
        If value.IndexOfAny(CsvSpecialCharacters) >= 0 Then
            Return """" & value.Replace("""", """""") & """"
        End If
        Return value
    End Function
End Class
