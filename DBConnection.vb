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
    Public Function GetConnectionString(ByRef preactorComObject As PreactorObj) As Integer

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
            Dim format = String.Format("{0} ({1}%)", stateName, efficiency)

            ' Add it to the result
            result.AppendLine(format)

        End While

        ' Close the connection
        connection.Close()

        ' Display in a message box all of the states and their efficiencies
        MessageBox.Show(result.ToString())
        Return 0
    End Function
End Class
