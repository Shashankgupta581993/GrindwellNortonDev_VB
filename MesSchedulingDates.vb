Option Strict On
Option Explicit On

Imports System
Imports System.Collections.Generic
Imports System.Configuration
Imports System.Drawing
Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports Preactor
Imports Preactor.Interop.PreactorObject

' Distinct from placement errors: existing per-order catch/continue handlers must
' never swallow a failed MES stamp after the live board has already been changed.
<ComVisible(False)>
Friend NotInheritable Class MesSchedulingException
    Inherits InvalidOperationException

    Friend Sub New(message As String, Optional inner As Exception = Nothing)
        MyBase.New(message, inner)
    End Sub
End Class

' No Opcenter references: allows retry/deduplication semantics to be tested without
' a running planning board. Record numbers are valid only until dataset reload.
<ComVisible(False)>
Friend NotInheritable Class MesSchedulingCycle
    Friend ReadOnly EffectiveDate As DateTime
    Private ReadOnly stamped As New HashSet(Of Integer)
    Private ReadOnly pending As New HashSet(Of Integer)

    Friend Sub New(value As DateTime)
        EffectiveDate = value.Date
    End Sub

    Friend Sub Stamp(record As Integer, writeDate As Action(Of Integer, DateTime))
        If record <= 0 Then Throw New ArgumentOutOfRangeException(NameOf(record))
        If stamped.Contains(record) Then Return
        pending.Add(record)
        Try
            writeDate(record, EffectiveDate)
        Catch ex As Exception
            Throw New MesSchedulingException(
                "MES Plan Effective Date could not be written for operation record " &
                record.ToString() & ". Scheduling stopped. The operation may already be placed. " &
                "Retry the rule to repair the date before continuing or saving for MES.", ex)
        End Try
        pending.Remove(record)
        stamped.Add(record)
    End Sub

    Friend Sub RetryPending(writeDate As Action(Of Integer, DateTime))
        For Each record As Integer In New List(Of Integer)(pending)
            Stamp(record, writeDate)
        Next
    End Sub
End Class

<ComVisible(False)>
Friend Module MesSchedulingDates
    Friend Const DefaultDateFieldName As String = "MES Plan Effective Date"

    Private NotInheritable Class Session
        ' Retain the RCW while its IUnknown address is used as a dictionary key.
        Friend Owner As PreactorObj
        Friend Scope As String
        Friend Cycle As MesSchedulingCycle
    End Class

    Private NotInheritable Class PreparedAction
        Friend Cycle As MesSchedulingCycle
        Friend WriteDate As Action(Of Integer, DateTime)
    End Class

    Private ReadOnly sessions As New Dictionary(Of IntPtr, Session)
    Private actions As New ConditionalWeakTable(Of IPreactor, PreparedAction)
    Private ReadOnly gate As New Object()

    ' Called once at each scheduling entry point, before any live mutation.
    Friend Function TryBeginAction(preactor As IPreactor,
                                   owner As PreactorObj,
                                   Optional changeDate As Boolean = False) As Boolean
        actions.Remove(preactor)
        Try
            Dim orders = preactor.FindFirstClassificationString("LAUNCH TIME")
            If Not orders.HasValue Then Throw New MesSchedulingException("The Orders format could not be resolved.")
            Dim format As Integer = orders.Value.FormatNumber
            Dim fieldName As String = ConfigurationManager.AppSettings("MesPlanEffectiveDateField")
            If String.IsNullOrWhiteSpace(fieldName) Then fieldName = DefaultDateFieldName
            Dim field As Integer = preactor.GetFieldNumber(format, fieldName)
            If field <= 0 OrElse preactor.GetFieldType(format, field) <> PreactorFieldType.DateTime Then
                Throw New MesSchedulingException("Configure an Orders DateTime field named '" & fieldName &
                                                 "' before running scheduling rules.")
            End If

            Dim scope As String = preactor.ParseShellString("{PATH}") & vbNullChar &
                                  preactor.CurrentFileName(format)
            Dim identity As IntPtr = Marshal.GetIUnknownForObject(owner)
            Dim session As Session = Nothing
            Try
                SyncLock gate
                    If Not sessions.TryGetValue(identity, session) OrElse
                       Not String.Equals(session.Scope, scope, StringComparison.Ordinal) Then
                        session = New Session With {.Owner = owner, .Scope = scope}
                        sessions(identity) = session
                    End If

                    Dim writeDate As Action(Of Integer, DateTime) =
                        Sub(record, value) preactor.WriteField(format, field, record, value)
                    ' Repair stamps with the original cycle date before changing dates.
                    If session.Cycle IsNot Nothing Then session.Cycle.RetryPending(writeDate)
                    If changeDate OrElse session.Cycle Is Nothing Then
                        Dim selectedDate As DateTime? = SelectDate()
                        If Not selectedDate.HasValue Then Return False
                        session.Cycle = New MesSchedulingCycle(selectedDate.Value)
                    End If
                    actions.Add(preactor, New PreparedAction With {
                        .Cycle = session.Cycle, .WriteDate = writeDate})
                End SyncLock
            Finally
                Marshal.Release(identity)
            End Try
            Return True
        Catch ex As Exception
            MessageBox.Show("Scheduling was not started." & Environment.NewLine & ex.Message,
                            "MES Plan Effective Date", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
    End Function

    Friend Sub RequirePreparedAction(preactor As IPreactor)
        GetAction(preactor)
    End Sub

    Private Function GetAction(preactor As IPreactor) As PreparedAction
        Dim action As PreparedAction = Nothing
        If Not actions.TryGetValue(preactor, action) Then
            Throw New MesSchedulingException("Start a scheduling rule and select an MES effective date before placing operations.")
        End If
        Return action
    End Function

    Friend Sub StampSuccessfulPlacement(preactor As IPreactor, record As Integer)
        Dim action As PreparedAction = GetAction(preactor)
        action.Cycle.Stamp(record, action.WriteDate)
    End Sub

    Friend Sub Invalidate(owner As PreactorObj)
        Dim identity As IntPtr = Marshal.GetIUnknownForObject(owner)
        Try
            SyncLock gate
                sessions.Remove(identity)
                ' A reload can reuse record numbers and the same IPreactor wrapper.
                ' Invalidating prepared actions prevents stale deduplication too.
                actions = New ConditionalWeakTable(Of IPreactor, PreparedAction)()
            End SyncLock
        Finally
            Marshal.Release(identity)
        End Try
    End Sub

    Private Function SelectDate() As DateTime?
        Using dialog As New Form With {
            .Text = "MES Plan Effective Date", .ClientSize = New Size(390, 148),
            .StartPosition = FormStartPosition.CenterScreen,
            .FormBorderStyle = FormBorderStyle.FixedDialog,
            .MinimizeBox = False, .MaximizeBox = False, .ShowInTaskbar = False}
            Dim label As New Label With {
                .Text = "Effective date for this scheduling cycle:",
                .AutoSize = True, .Location = New Point(16, 17)}
            Dim picker As New DateTimePicker With {
                .Format = DateTimePickerFormat.Custom, .CustomFormat = "dd MMM yyyy",
                .Value = DateTime.Today, .Location = New Point(16, 44), .Width = 355}
            Dim accept As New Button With {
                .Text = "OK", .DialogResult = DialogResult.OK,
                .Location = New Point(207, 99), .Width = 78}
            Dim cancel As New Button With {
                .Text = "Cancel", .DialogResult = DialogResult.Cancel,
                .Location = New Point(293, 99), .Width = 78}
            dialog.Controls.AddRange(New Control() {label, picker, accept, cancel})
            dialog.AcceptButton = accept
            dialog.CancelButton = cancel
            If dialog.ShowDialog() = DialogResult.OK Then Return picker.Value.Date
        End Using
        Return Nothing
    End Function
End Module

' Configure these as ordinary custom actions. InvalidateCycle must run after
' dataset load/reload/import/undo and before another scheduling action is allowed.
<ComVisible(True)>
<Microsoft.VisualBasic.ComClass("837ae098-2b84-4e9a-9ae1-4d31a947b58d", "9d3920af-f4eb-4f85-a040-9cfb1d11a850")>
Public Class MesSchedulingActions
    Public Function ChangeDate(ByRef preactorComObject As PreactorObj,
                               ByRef pespComObject As Object) As Integer
        Dim preactor = PreactorFactory.CreatePreactorObject(preactorComObject)
        Return If(MesSchedulingDates.TryBeginAction(preactor, preactorComObject, changeDate:=True), 0, -1)
    End Function

    Public Function InvalidateCycle(ByRef preactorComObject As PreactorObj,
                                    ByRef pespComObject As Object) As Integer
        MesSchedulingDates.Invalidate(preactorComObject)
        Return 0
    End Function
End Class
