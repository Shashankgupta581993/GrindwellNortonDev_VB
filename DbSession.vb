Imports System.Configuration
Imports System.Data.SqlClient

Public Class DbSession
    Public ReadOnly Property ConnectionString As String

    Public Sub New(sqlPassword As String)
        Dim cs = ConfigurationManager.ConnectionStrings("MetaDb")
        If cs Is Nothing OrElse String.IsNullOrWhiteSpace(cs.ConnectionString) Then
            Throw New ApplicationException(
        "Missing connection string 'MetaDb'. Config file used: " &
        AppDomain.CurrentDomain.SetupInformation.ConfigurationFile
    )
        End If

        Dim baseConn = cs.ConnectionString
        ConnectionString = baseConn & "Password=" & sqlPassword & ";"
    End Sub

    Public Function OpenConnection() As SqlConnection
        Dim cn As New SqlConnection(ConnectionString)
        cn.Open()
        Return cn
    End Function
End Class
