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
    Public Function GetConnectionString(ByRef preactorComObject As PreactorObj) As String

        Dim preactor As IPreactor = PreactorFactory.CreatePreactorObject(preactorComObject)

        ' Get the connection string
        Dim connectionString = preactor.ParseShellString("{DB CONNECT STRING}")
        Return connectionString
    End Function
    Public Function GetConnect(ByRef preactorComObject As PreactorObj) As Integer

        Dim preactor As IPreactor = PreactorFactory.CreatePreactorObject(preactorComObject)

        ' Get the connection string
        Dim connectionString = preactor.ParseShellString("{DB CONNECT STRING}")

        ' Create a connection to the database
        Dim connection = New SqlConnection(connectionString)

        ' Open the connection
        connection.Open()

        ' Define the sql to select the calendar states
        Dim sql = "SELECT " +
        "[Id], [Name], [Color], [Pattern], [Efficiency], [CostFactor], [IsSetupAllowed] " +
        "FROM " +
        "[Calendar].[CalendarStates]"

        ' Create a new command
        Dim command = New SqlCommand(sql, connection)

        ' Execute the command and get a reader
        Dim reader = command.ExecuteReader()

        ' Get the ordinals for the fields we are interested in
        Dim efficiencyOrdinal = reader.GetOrdinal("Efficiency")
        Dim nameOrdinal = reader.GetOrdinal("Name")

        ' Create a new string builder
        Dim result = New StringBuilder()

        ' Loop through all of the rows
        While (reader.Read())

            ' Get the state name and efficiency
            Dim stateName = reader.GetString(nameOrdinal)
            Dim efficiency = reader.GetDouble(efficiencyOrdinal) * 100

            ' Create a string like: StateName (100%)
            Dim format = String.Format("{0}/({1}%)", stateName, efficiency)

            ' Add it to the result
            result.AppendLine(format)

        End While

        ' Close the connection
        connection.Close()

        ' Display in a message box all of the states and their efficiencies
        MessageBox.Show(result.ToString())
        Return 0
    End Function

    ' 1) Generic: return results as DataTable
    Public Function ExecuteDataTable(sql As String,
                                     Optional parameters As IEnumerable(Of SqlParameter) = Nothing,
                                     Optional commandType As CommandType = CommandType.Text) As DataTable


        Dim dt As New DataTable()

        Using conn As New SqlConnection(_connectionstring)
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
