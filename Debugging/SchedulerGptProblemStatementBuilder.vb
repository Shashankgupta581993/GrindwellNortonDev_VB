Option Strict On
Option Explicit On

Imports System
Imports System.Collections.Generic
Imports System.Globalization
Imports System.IO
Imports System.Text
Imports System.Web.Script.Serialization

Public Class SchedulerGptProblemStatementBuilder
    Private Shared ReadOnly Files As String() = {
        "00_RunSummary.md", "01_ConfigSnapshot.csv", "02_FieldMap.csv",
        "03_OrderOperationSnapshot.csv", "04_UnscheduledOperations.csv",
        "05_WipChainDiagnostics.csv", "06_StageEligibilityDiagnostics.csv",
        "07_OptimizerCandidateTrace.csv", "08_BatchTunnelSwkPlanTrace.csv",
        "09_ScheduleAttemptTrace.csv", "10_ResourceValidation.csv",
        "11_ReasonCodeSummary.csv", "12_GptProblemStatement.md",
        "13_GptProblemStatement.json"
    }

    Private Shared ReadOnly Questions As String() = {
        "Which operations are unscheduled?",
        "Which operation numbers are unscheduled most often?",
        "Were unscheduled operations missing from optimizer candidate selection?",
        "Were they selected by optimizer but not attempted in Run?",
        "Were they attempted but rejected by planning-board scheduling?",
        "Were they blocked by previous operations?",
        "Were they blocked by WIP?",
        "Were they blocked by missing resources?",
        "Were they blocked by kiln type, cycle type, occupancy, or batch/cart rules?",
        "Were SWK orders silently skipped?",
        "Were parent records returned correctly?",
        "Are any operations marked scheduled in custom fields but not scheduled on the board?",
        "What exact code area is responsible?",
        "What minimal code change should be made?"
    }

    Public Shared Sub WriteFiles(debug As SchedulerDebugCollector)
        WriteMarkdown(Path.Combine(debug.ExportFolder, "12_GptProblemStatement.md"), debug)
        WriteJson(Path.Combine(debug.ExportFolder, "13_GptProblemStatement.json"), debug)
    End Sub

    Private Shared Sub WriteMarkdown(path As String, debug As SchedulerDebugCollector)
        Dim sb As New StringBuilder()
        sb.AppendLine("# GN Opcenter Scheduling Debug Package").AppendLine()
        sb.AppendLine("## Problem").AppendLine("Many orders/operations remain unscheduled after running custom GN Opcenter scheduling logic.").AppendLine()
        sb.AppendLine("## Scheduling stages")
        sb.AppendLine("- Pressing: 200–280")
        sb.AppendLine("- Firing loading: 290")
        sb.AppendLine("- Firing: 300")
        sb.AppendLine("- Firing unloading: 310")
        sb.AppendLine("- PostFiring: 380+").AppendLine()
        sb.AppendLine("## Firing types").AppendLine("- Batch").AppendLine("- Tunnel").AppendLine("- SWK").AppendLine()
        sb.AppendLine("## Recent logic context")
        sb.AppendLine("- Optimizers return parent records.")
        sb.AppendLine("- Batch firing uses occupancy, cycle type, kiln matrix, loading buffer, daily batch limit, and predecessor completion.")
        sb.AppendLine("- Tunnel firing uses cart pitch, carts per day, total carts, occupancy, and predecessor completion.")
        sb.AppendLine("- Pressing uses previous-operation readiness, resource group, wheel dia/pin, and cooldown/changeover rules.")
        sb.AppendLine("- PostFiring schedules forward after previous operation is scheduled.")
        sb.AppendLine("- Only unscheduled operations should be selected by optimizers.")
        sb.AppendLine("- WIP must be considered across all operations.").AppendLine()
        sb.AppendLine("## Files in this debug package")
        For Each fileName In Files
            sb.AppendLine("- " & fileName)
        Next
        sb.AppendLine().AppendLine("## Questions for GPT")
        For i As Integer = 0 To Questions.Length - 1
            sb.AppendLine((i + 1).ToString(CultureInfo.InvariantCulture) & ". " & Questions(i))
        Next
        File.WriteAllText(path, sb.ToString(), New UTF8Encoding(False))
    End Sub

    Private Shared Sub WriteJson(path As String, debug As SchedulerDebugCollector)
        Dim payload As New Dictionary(Of String, Object) From {
            {"project", "GN_Opcenter VB.NET"},
            {"debugPackageVersion", "1.0"},
            {"runId", debug.RunId},
            {"exportedAt", debug.ExportedAt.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture)},
            {"files", Files},
            {"stages", New String() {"Pressing", "Drying", "BatchFiring", "TunnelFiring", "SWK", "PostFiring"}},
            {"questionsForGpt", Questions}
        }
        Dim serializer As New JavaScriptSerializer()
        File.WriteAllText(path, serializer.Serialize(payload), New UTF8Encoding(False))
    End Sub
End Class
