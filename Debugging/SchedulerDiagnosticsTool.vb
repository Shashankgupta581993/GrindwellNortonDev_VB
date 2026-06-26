Option Strict On
Option Explicit On

Imports System
Imports System.Collections.Generic
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports Preactor
Imports Preactor.Interop.PreactorObject

<ComVisible(True)>
<Microsoft.VisualBasic.ComClass("CB33CBBB-86A2-4E5B-AB7F-CED625FC91F0", "74CB5EE0-B64D-47E8-BA3C-AF77977B5DA2")>
Public Class SchedulerDiagnosticsTool
    Public Function ExportSchedulingDiagnostics(ByRef preactorComObject As PreactorObj) As Integer
        Dim preactor As IPreactor = PreactorFactory.CreatePreactorObject(preactorComObject)
        Dim debug As New SchedulerDebugCollector()
        Try
            Dim snapshot As List(Of OperationSnapshot) =
                SchedulerStageDiagnostics.BuildOrderOperationSnapshot(preactor, debug)
            SchedulerStageDiagnostics.DiagnoseWip(snapshot, debug)
            SchedulerStageDiagnostics.DiagnoseAll(snapshot, debug)
            debug.ExportAll(preactor)
            MessageBox.Show("Scheduling diagnostics exported to:" & Environment.NewLine & debug.ExportFolder,
                            "GN Opcenter Scheduling Diagnostics",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information)
            Return 0
        Catch ex As Exception
            MessageBox.Show("Scheduling diagnostics export failed:" & Environment.NewLine & ex.Message,
                            "GN Opcenter Scheduling Diagnostics",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error)
            Return -1
        End Try
    End Function
End Class
