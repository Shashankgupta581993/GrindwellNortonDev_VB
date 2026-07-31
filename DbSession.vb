Option Strict On
Option Explicit On

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

        Dim builder As New SqlConnectionStringBuilder(cs.ConnectionString)
        builder.Password = If(sqlPassword, String.Empty)
        ConnectionString = builder.ConnectionString
    End Sub

    Public Function OpenConnection() As SqlConnection
        Dim cn As New SqlConnection(ConnectionString)
        cn.Open()
        Return cn
    End Function
End Class
