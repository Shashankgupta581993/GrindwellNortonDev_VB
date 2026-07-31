Option Strict On
Option Explicit On

Imports System
Imports System.Collections.Generic
Imports System.Data
Imports System.Data.SqlClient
Imports System.Windows.Forms

Public Class DataViewExplorerForm

    Private _session As DbSession
    Private _repo As MetadataExplorerRepository

    Private ReadOnly _allowedTables As New HashSet(Of String)(
        StringComparer.OrdinalIgnoreCase)

    Private _currentSchema As String = "metadata"
    Private _currentTable As String
    Private _currentData As DataTable
    Private _currentAdapter As SqlDataAdapter
    Private _currentConnection As SqlConnection

    Private Sub DataViewExplorerForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Using lf As New LoginForm()
            If lf.ShowDialog() <> DialogResult.OK Then
                Me.Close()
                Return
            End If
            _session = New DbSession(lf.PasswordValue)
            _repo = New MetadataExplorerRepository(_session)
        End Using

        LoadTables()
    End Sub

    Private Sub LoadTables()
        Dim dt = _repo.GetAllDboTables()

        ' Example for TreeView:
        tvTables.Nodes.Clear()
        _allowedTables.Clear()
        Dim root = tvTables.Nodes.Add("Tables")

        For Each r As DataRow In dt.Rows
            Dim schema = r.Field(Of String)("SchemaName")  ' will be "userdata"
            Dim tableName = r.Field(Of String)("TableName")

            Dim n = root.Nodes.Add($"{schema}.{tableName}")
            n.Tag = schema & "|" & tableName
            _allowedTables.Add(GetTableKey(schema, tableName))
        Next
        root.Expand()
    End Sub

    Private Sub tvTables_AfterSelect(sender As Object, e As TreeViewEventArgs) Handles tvTables.AfterSelect
        If e.Node Is Nothing OrElse e.Node.Tag Is Nothing Then Return

        Dim parts = e.Node.Tag.ToString().Split("|"c)
        If parts.Length <> 2 OrElse
           Not _allowedTables.Contains(GetTableKey(parts(0), parts(1))) Then

            MessageBox.Show("The selected table is not in the loaded metadata list.")
            Return
        End If

        _currentSchema = parts(0)
        _currentTable = parts(1)

        ReloadCurrent()
    End Sub

    Private Sub ReloadCurrent()
        If Not _allowedTables.Contains(
            GetTableKey(_currentSchema, _currentTable)) Then
            Throw New InvalidOperationException(
                "The selected table is not in the loaded metadata list.")
        End If

        CleanupConnection()

        Dim result = _repo.LoadTable(_currentSchema, _currentTable)
        _currentData = result.Data
        _currentAdapter = result.Adapter
        _currentConnection = result.Connection

        dgvData.DataSource = _currentData
    End Sub

    Private Sub CleanupConnection()
        Try
            If _currentAdapter IsNot Nothing Then
                DisposeCommand(_currentAdapter.InsertCommand)
                DisposeCommand(_currentAdapter.UpdateCommand)
                DisposeCommand(_currentAdapter.DeleteCommand)
                DisposeCommand(_currentAdapter.SelectCommand)
                _currentAdapter.Dispose()
            End If

            If _currentConnection IsNot Nothing Then
                _currentConnection.Close()
                _currentConnection.Dispose()
            End If
        Catch
        End Try
        _currentAdapter = Nothing
        _currentConnection = Nothing
    End Sub

    Private Shared Sub DisposeCommand(command As SqlCommand)
        If command IsNot Nothing Then command.Dispose()
    End Sub

    Private Shared Function GetTableKey(schemaName As String,
                                        tableName As String) As String
        Return If(schemaName, String.Empty) & "|" &
            If(tableName, String.Empty)
    End Function

    Private Sub DataViewExplorerForm_FormClosed(
        sender As Object,
        e As FormClosedEventArgs) Handles MyBase.FormClosed

        CleanupConnection()
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        If _currentData Is Nothing OrElse _currentAdapter Is Nothing Then Return

        Try
            dgvData.EndEdit()
            _repo.SaveChanges(_currentData, _currentAdapter)
            MessageBox.Show("Saved successfully.")
            ReloadCurrent()
        Catch ex As Exception
            MessageBox.Show("Save failed: " & ex.Message)
        End Try
    End Sub

    Private Sub btnReload_Click(sender As Object, e As EventArgs) Handles btnReload.Click
        If String.IsNullOrWhiteSpace(_currentTable) Then Return
        ReloadCurrent()
    End Sub
    Private Sub btnDelete_Click(sender As Object, e As EventArgs) Handles btnDelete.Click
        If dgvData.CurrentRow Is Nothing Then Return
        dgvData.Rows.Remove(dgvData.CurrentRow)
    End Sub
    Private Sub btnApplyFilter_Click(sender As Object, e As EventArgs) Handles btnApplyFilter.Click
        If _currentData Is Nothing Then Return
        Dim view As DataView = _currentData.DefaultView
        Dim x = txtFilter.Text.Replace("'", "''")

        Dim parts As New List(Of String)
        For Each col As DataColumn In _currentData.Columns
            If col.DataType Is GetType(String) Then
                parts.Add($"[{col.ColumnName}] LIKE '%{x}%'")
            End If
        Next

        view.RowFilter = String.Join(" OR ", parts)
    End Sub

End Class
