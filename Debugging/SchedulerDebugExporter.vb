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
    Public Shared Sub ExportAll(debug As SchedulerDebugCollector, preactor As IPreactor)
        If debug Is Nothing OrElse Not debug.Enabled Then Return
        Directory.CreateDirectory(debug.ExportFolder)

        Dim configRows As New List(Of DebugConfigSnapshotRow) From {
            New DebugConfigSnapshotRow With {.RunId = debug.RunId, .Name = "ExportedAt", .Value = debug.ExportedAt.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture)},
            New DebugConfigSnapshotRow With {.RunId = debug.RunId, .Name = "ExportFolder", .Value = debug.ExportFolder},
            New DebugConfigSnapshotRow With {.RunId = debug.RunId, .Name = "SnapshotRows", .Value = debug.OperationSnapshots.Count.ToString(CultureInfo.InvariantCulture)}
        }

        WriteRunSummary(Path.Combine(debug.ExportFolder, "00_RunSummary.md"), debug)
        WriteCsv(Path.Combine(debug.ExportFolder, "01_ConfigSnapshot.csv"), configRows)
        WriteCsv(Path.Combine(debug.ExportFolder, "02_FieldMap.csv"), debug.FieldMapRows)
        WriteCsv(Path.Combine(debug.ExportFolder, "03_OrderOperationSnapshot.csv"), debug.OperationSnapshots)
        WriteCsv(Path.Combine(debug.ExportFolder, "04_UnscheduledOperations.csv"),
                 debug.OperationSnapshots.Where(Function(x) Not x.IsScheduled))
        WriteCsv(Path.Combine(debug.ExportFolder, "05_WipChainDiagnostics.csv"), debug.WipDiagnostics)
        WriteCsv(Path.Combine(debug.ExportFolder, "06_StageEligibilityDiagnostics.csv"), debug.StageEligibilityRows)
        WriteCsv(Path.Combine(debug.ExportFolder, "07_OptimizerCandidateTrace.csv"), debug.OptimizerCandidateTraceRows)
        WriteCsv(Path.Combine(debug.ExportFolder, "08_BatchTunnelSwkPlanTrace.csv"), debug.BatchTunnelSwkPlanTraceRows)
        WriteCsv(Path.Combine(debug.ExportFolder, "09_ScheduleAttemptTrace.csv"), debug.ScheduleAttemptTraceRows)
        WriteCsv(Path.Combine(debug.ExportFolder, "10_ResourceValidation.csv"), debug.ResourceValidationRows)
        WriteReasonSummary(Path.Combine(debug.ExportFolder, "11_ReasonCodeSummary.csv"), debug)
        SchedulerGptProblemStatementBuilder.WriteFiles(debug)
    End Sub

    Public Shared Sub WriteCsv(Of T)(filePath As String, rows As IEnumerable(Of T))
        Dim properties As PropertyInfo() = GetType(T).GetProperties(BindingFlags.Instance Or BindingFlags.Public)
        Using writer As New StreamWriter(filePath, False, New UTF8Encoding(False))
            writer.WriteLine(String.Join(",", properties.Select(Function(p) Escape(p.Name))))
            If rows Is Nothing Then Return
            For Each row As T In rows
                Dim values As New List(Of String)(properties.Length)
                For Each prop As PropertyInfo In properties
                    values.Add(Escape(FormatValue(prop.GetValue(row, Nothing))))
                Next
                writer.WriteLine(String.Join(",", values))
            Next
        End Using
    End Sub

    Public Shared Sub WriteRunSummary(filePath As String, debug As SchedulerDebugCollector)
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
            "- Schedule attempts: " & debug.ScheduleAttemptTraceRows.Count.ToString(CultureInfo.InvariantCulture) & Environment.NewLine
        File.WriteAllText(filePath, text, New UTF8Encoding(False))
    End Sub

    Public Shared Sub WriteReasonSummary(filePath As String, debug As SchedulerDebugCollector)
        Dim items As New List(Of KeyValuePair(Of String, String))()
        For Each row In debug.WipDiagnostics
            AddReason(items, row.ReasonCode, "WIP")
        Next
        For Each row In debug.StageEligibilityRows
            AddReason(items, row.ExcludedReasonCode, "Eligibility")
        Next
        For Each row In debug.OptimizerCandidateTraceRows
            AddReason(items, row.ReasonCode, "Optimizer")
        Next
        For Each row In debug.BatchTunnelSwkPlanTraceRows
            AddReason(items, row.ReasonCode, "Plan")
        Next
        For Each row In debug.ScheduleAttemptTraceRows
            AddReason(items, row.FailureReasonCode, "ScheduleAttempt")
        Next

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
        If value.IndexOfAny(New Char() {","c, """"c, ControlChars.Cr, ControlChars.Lf}) >= 0 Then
            Return """" & value.Replace("""", """""") & """"
        End If
        Return value
    End Function
End Class
