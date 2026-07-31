Option Strict On
Option Explicit On

Imports System.Data.SqlClient
Imports System.Text
Imports System.Windows.Forms
Imports System
Imports System.Runtime.InteropServices
Imports System.Security
Imports Preactor
Imports Preactor.Interop.PreactorObject
Imports System.Globalization
Imports System.IO
Imports System.Data
Imports System.Linq


Public Class DBConnection

    Private _connectionString As String

    Public Function GetConnectionString(ByRef preactorComObject As PreactorObj) As String

        Dim preactor As IPreactor = PreactorFactory.CreatePreactorObject(preactorComObject)

        ' Get the connection string
        Dim connectionString = preactor.ParseShellString("{DB CONNECT STRING}")
        _connectionString = connectionString
        Return connectionString
    End Function
    Public Function GetConnect(ByRef preactorComObject As PreactorObj) As Integer

        Dim preactor As IPreactor = PreactorFactory.CreatePreactorObject(preactorComObject)

        ' Get the connection string
        Dim connectionString = preactor.ParseShellString("{DB CONNECT STRING}")

        Dim result As New StringBuilder()

        Using connection As New SqlConnection(connectionString)
            connection.Open()

            Dim sql As String =
                "SELECT [Id], [Name], [Color], [Pattern], [Efficiency], " &
                "[CostFactor], [IsSetupAllowed] " &
                "FROM [Calendar].[CalendarStates]"

            Using command As New SqlCommand(sql, connection)
                Using reader As SqlDataReader = command.ExecuteReader()
                    Dim efficiencyOrdinal As Integer =
                        reader.GetOrdinal("Efficiency")
                    Dim nameOrdinal As Integer = reader.GetOrdinal("Name")

                    While reader.Read()
                        Dim stateName As String = reader.GetString(nameOrdinal)
                        Dim efficiency As Double =
                            reader.GetDouble(efficiencyOrdinal) * 100
                        Dim formatted As String =
                            String.Format(CultureInfo.InvariantCulture,
                                          "{0}/({1}%)",
                                          stateName,
                                          efficiency)
                        result.AppendLine(formatted)
                    End While
                End Using
            End Using
        End Using

        ' Display in a message box all of the states and their efficiencies
        MessageBox.Show(result.ToString())
        Return 0
    End Function

    ' 1) Generic: return results as DataTable
    Public Function ExecuteDataTable(sql As String, ByRef preactorComObject As PreactorObj,
                                     Optional parameters As IEnumerable(Of SqlParameter) = Nothing,
                                     Optional commandType As CommandType = CommandType.Text) As DataTable
        _connectionString = GetConnectionString(preactorComObject)

        Dim dt As New DataTable()

        Using conn As New SqlConnection(_connectionString)
            Using cmd As New SqlCommand(sql, conn)
                cmd.CommandType = commandType

                If parameters IsNot Nothing Then
                    cmd.Parameters.AddRange(parameters.ToArray())
                End If

                Using da As New SqlDataAdapter(cmd)
                    da.Fill(dt)
                End Using
            End Using
        End Using

        Return dt
    End Function


    ' 2) For INSERT/UPDATE/DELETE
    Public Function ExecuteNonQuery(sql As String,
                                    Optional parameters As IEnumerable(Of SqlParameter) = Nothing,
                                    Optional commandType As CommandType = CommandType.Text) As Integer

        Using conn As New SqlConnection(_connectionString)
            Using cmd As New SqlCommand(sql, conn)
                cmd.CommandType = commandType

                If parameters IsNot Nothing Then
                    cmd.Parameters.AddRange(parameters.ToArray())
                End If

                conn.Open()
                Return cmd.ExecuteNonQuery()
            End Using
        End Using
    End Function

    ' 3) For getting a single value (COUNT, MAX, etc.)
    Public Function ExecuteScalar(Of T)(sql As String,
                                        Optional parameters As IEnumerable(Of SqlParameter) = Nothing,
                                        Optional commandType As CommandType = CommandType.Text) As T

        Using conn As New SqlConnection(_connectionString)
            Using cmd As New SqlCommand(sql, conn)
                cmd.CommandType = commandType

                If parameters IsNot Nothing Then
                    cmd.Parameters.AddRange(parameters.ToArray())
                End If

                conn.Open()
                Dim result = cmd.ExecuteScalar()

                If result Is Nothing OrElse Convert.IsDBNull(result) Then
                    Return Nothing
                Else
                    Return CType(result, T)
                End If
            End Using
        End Using
    End Function

    ' 4) Optional: Execute and map each row to a custom object
    Public Function ExecuteReader(Of T)(sql As String,
                                        map As Func(Of SqlDataReader, T),
                                        Optional parameters As IEnumerable(Of SqlParameter) = Nothing,
                                        Optional commandType As CommandType = CommandType.Text) As List(Of T)

        Dim list As New List(Of T)()

        Using conn As New SqlConnection(_connectionString)
            Using cmd As New SqlCommand(sql, conn)
                cmd.CommandType = commandType

                If parameters IsNot Nothing Then
                    cmd.Parameters.AddRange(parameters.ToArray())
                End If

                conn.Open()
                Using reader As SqlDataReader = cmd.ExecuteReader()
                    While reader.Read()
                        list.Add(map(reader))
                    End While
                End Using
            End Using
        End Using

        Return list
    End Function

End Class
