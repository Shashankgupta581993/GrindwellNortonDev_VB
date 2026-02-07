Imports System.Data
Imports System.Data.SqlClient

Public Class MetadataExplorerRepository
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
                WHERE t.is_ms_shipped = 0 AND s.name = 'sampledata'
                ORDER BY t.name;"
            Using da As New SqlDataAdapter(sql, cn)
                Dim dt As New DataTable()
                da.Fill(dt)
                Return dt
            End Using
        End Using
    End Function

    Public Function LoadTable(schemaName As String, tableName As String) As (Data As DataTable, Adapter As SqlDataAdapter, Connection As SqlConnection)
        Dim cn = _session.OpenConnection()

        Dim fullName = $"[{schemaName}].[{tableName}]"
        Dim sql = $"SELECT * FROM {fullName};"

        Dim da As New SqlDataAdapter(sql, cn)
        Dim cb As New SqlCommandBuilder(da) ' auto CRUD (requires PK!)

        Dim dt As New DataTable(tableName)
        da.Fill(dt)

        Return (dt, da, cn)
    End Function

    Public Sub SaveChanges(dt As DataTable, da As SqlDataAdapter)
        da.Update(dt)
    End Sub
End Class
