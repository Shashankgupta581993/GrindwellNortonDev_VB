Option Strict On
Option Explicit On

Imports System
Imports System.Data
Imports System.Data.SqlClient

Public Class MetadataExplorerRepository
    Private Const MetadataSchemaName As String = "metadata"
    Private ReadOnly _session As DbSession

    Public Sub New(session As DbSession)
        _session = session
    End Sub

    Public Function GetAllDboTables() As DataTable
        Using cn = _session.OpenConnection()
            Dim sql = "
                SELECT s.name AS SchemaName, t.name AS TableName
                FROM sys.tables t
                JOIN sys.schemas s ON t.schema_id = s.schema_id
                WHERE t.is_ms_shipped = 0 AND s.name = 'metadata'
                ORDER BY t.name;"
            Using da As New SqlDataAdapter(sql, cn)
                Dim dt As New DataTable()
                da.Fill(dt)
                Return dt
            End Using
        End Using
    End Function

    Public Function LoadTable(schemaName As String,
                              tableName As String) As (
                                  Data As DataTable,
                                  Adapter As SqlDataAdapter,
                                  Connection As SqlConnection)

        Dim cn As SqlConnection = _session.OpenConnection()
        Dim da As SqlDataAdapter = Nothing

        Try
            ValidateMetadataTable(cn, schemaName, tableName)

            Dim fullName As String =
                QuoteIdentifier(schemaName) & "." & QuoteIdentifier(tableName)
            Dim sql As String = "SELECT * FROM " & fullName & ";"

            da = New SqlDataAdapter(sql, cn)
            Dim dt As New DataTable(tableName)
            da.Fill(dt)

            ' Generate CRUD commands while the builder is deterministically
            ' disposed. Tables without a usable key remain readable and retain
            ' the existing save-time failure behavior.
            Using builder As New SqlCommandBuilder(da)
                Try
                    da.InsertCommand = builder.GetInsertCommand(True)
                    da.UpdateCommand = builder.GetUpdateCommand(True)
                    da.DeleteCommand = builder.GetDeleteCommand(True)
                Catch ex As InvalidOperationException
                    ' Preserve read-only browsing when auto-CRUD is unavailable.
                End Try
            End Using

            Return (dt, da, cn)
        Catch
            If da IsNot Nothing Then da.Dispose()
            cn.Dispose()
            Throw
        End Try
    End Function

    Public Sub SaveChanges(dt As DataTable, da As SqlDataAdapter)
        da.Update(dt)
    End Sub

    Private Shared Sub ValidateMetadataTable(connection As SqlConnection,
                                             schemaName As String,
                                             tableName As String)
        If String.IsNullOrWhiteSpace(schemaName) OrElse
           String.IsNullOrWhiteSpace(tableName) Then

            Throw New ArgumentException(
                "A metadata schema and table name are required.")
        End If

        If Not schemaName.Equals(MetadataSchemaName,
                                 StringComparison.OrdinalIgnoreCase) Then
            Throw New InvalidOperationException(
                "Only tables from the metadata schema may be opened.")
        End If

        Const sql As String =
            "SELECT COUNT(1) " &
            "FROM sys.tables t " &
            "JOIN sys.schemas s ON t.schema_id = s.schema_id " &
            "WHERE t.is_ms_shipped = 0 " &
            "AND s.name = @schemaName " &
            "AND t.name = @tableName;"

        Using command As New SqlCommand(sql, connection)
            command.Parameters.Add("@schemaName",
                                   SqlDbType.NVarChar,
                                   128).Value = schemaName
            command.Parameters.Add("@tableName",
                                   SqlDbType.NVarChar,
                                   128).Value = tableName

            Dim matchCount As Integer =
                Convert.ToInt32(command.ExecuteScalar(),
                                Globalization.CultureInfo.InvariantCulture)
            If matchCount <> 1 Then
                Throw New InvalidOperationException(
                    "The selected metadata table is not in the loaded table list.")
            End If
        End Using
    End Sub

    Private Shared Function QuoteIdentifier(identifier As String) As String
        Return "[" & identifier.Replace("]", "]]") & "]"
    End Function
End Class
