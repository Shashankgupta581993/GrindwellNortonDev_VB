Option Strict On
Option Explicit On

Imports System
Imports System.Data
Imports System.Drawing
Imports System.Globalization
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports Preactor
Imports Preactor.Interop.PreactorObject

<ComVisible(True)>
<Guid("15670678-e476-4339-8b89-54a203928d66")>
Public Interface IKilnAvailabilityWindow

    Function OnOpen(ByRef preactorComObject As PreactorObj,
                    ByRef parameter As String) As Integer

    Function OnClose(ByRef preactorComObject As PreactorObj,
                     ByRef parameter As String) As Integer

End Interface


<ComVisible(True)>
<Guid("cd20b040-825b-4637-9474-3082310fc1b0")>
<ProgId("OpcenterAPSProject_VB.KilnAvailabilityWindow")>
<ClassInterface(ClassInterfaceType.None)>
Public Class KilnAvailabilityWindow
    Inherits UserControl
    Implements IKilnAvailabilityWindow

    Private preactor As IPreactor

    Private Const TABLE_NAME As String = "GN Kiln Availability"

    Private grid As DataGridView
    Private btnRefresh As Button
    Private btnSave As Button
    Private lblStatus As Label

    Public Sub New()
        MyBase.New()
        InitializeComponent()
        BuildTableUi()
    End Sub

    Public Function OnOpen(ByRef preactorComObject As PreactorObj,
                       ByRef parameter As String) As Integer _
                       Implements IKilnAvailabilityWindow.OnOpen

        Try
            preactor = PreactorFactory.CreatePreactorObject(preactorComObject)

            LoadKilnAvailabilityTable()

            Return 0

        Catch ex As Exception
            MessageBox.Show("Kiln Availability OnOpen failed: " & ex.Message,
                        "Kiln Availability",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error)
            Return -1
        End Try

    End Function
    Public Function OnClose(ByRef preactorComObject As PreactorObj,
                            ByRef parameter As String) As Integer _
                            Implements IKilnAvailabilityWindow.OnClose

        Return 0

    End Function

    Private Sub BuildTableUi()

        Me.Controls.Clear()

        Dim topPanel As New Panel()
        topPanel.Dock = DockStyle.Top
        topPanel.Height = 42

        btnRefresh = New Button()
        btnRefresh.Text = "Refresh"
        btnRefresh.Left = 8
        btnRefresh.Top = 8
        btnRefresh.Width = 90
        AddHandler btnRefresh.Click, AddressOf RefreshClicked

        btnSave = New Button()
        btnSave.Text = "Save"
        btnSave.Left = 105
        btnSave.Top = 8
        btnSave.Width = 90
        AddHandler btnSave.Click, AddressOf SaveClicked

        lblStatus = New Label()
        lblStatus.Left = 210
        lblStatus.Top = 13
        lblStatus.AutoSize = True
        lblStatus.Text = ""

        topPanel.Controls.Add(btnRefresh)
        topPanel.Controls.Add(btnSave)
        topPanel.Controls.Add(lblStatus)

        grid = New DataGridView()
        grid.Dock = DockStyle.Fill
        grid.AllowUserToAddRows = False
        grid.AllowUserToDeleteRows = False
        grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        grid.MultiSelect = False

        Me.Controls.Add(grid)
        Me.Controls.Add(topPanel)

    End Sub
    Private Sub LoadKilnAvailabilityTable()

        If preactor Is Nothing Then
            Throw New InvalidOperationException("Preactor object is not initialized.")
        End If

        Dim fmt As Integer = preactor.GetFormatNumber(TABLE_NAME)

        Dim fNumber As Integer = preactor.GetFieldNumber(fmt, "Number")
        Dim fKiln As Integer = preactor.GetFieldNumber(fmt, "Kiln")
        Dim fKilnType As Integer = preactor.GetFieldNumber(fmt, "Kiln Type")
        Dim fStatus As Integer = preactor.GetFieldNumber(fmt, "Availability Status")
        Dim fAvailableFrom As Integer = preactor.GetFieldNumber(fmt, "Available From")
        Dim fAvailableUntil As Integer = preactor.GetFieldNumber(fmt, "Available Until")
        Dim fOverrideStart As Integer = preactor.GetFieldNumber(fmt, "Override Start Time")
        Dim fOverrideEnd As Integer = preactor.GetFieldNumber(fmt, "Override End Time")
        Dim fMinOcc As Integer = preactor.GetFieldNumber(fmt, "Min Occupancy Override")
        Dim fMaxOcc As Integer = preactor.GetFieldNumber(fmt, "Max Occupancy Override")
        Dim fCartsPerDay As Integer = preactor.GetFieldNumber(fmt, "Carts Per Day Override")
        Dim fTotalCarts As Integer = preactor.GetFieldNumber(fmt, "Total Carts Override")
        Dim fReason As Integer = preactor.GetFieldNumber(fmt, "Reason")
        Dim fLastUpdatedBy As Integer = preactor.GetFieldNumber(fmt, "Last Updated By")
        Dim fLastUpdatedAt As Integer = preactor.GetFieldNumber(fmt, "Last Updated At")
        Dim fActive As Integer = preactor.GetFieldNumber(fmt, "Active")
        Dim fSortOrder As Integer = preactor.GetFieldNumber(fmt, "Sort Order")

        Dim dt As New DataTable(TABLE_NAME)

        dt.Columns.Add("__Record", GetType(Integer))
        dt.Columns.Add("Number", GetType(Integer))
        dt.Columns.Add("Kiln", GetType(String))
        dt.Columns.Add("Kiln Type", GetType(String))
        dt.Columns.Add("Availability Status", GetType(String))
        dt.Columns.Add("Available From", GetType(String))
        dt.Columns.Add("Available Until", GetType(String))
        dt.Columns.Add("Override Start Time", GetType(String))
        dt.Columns.Add("Override End Time", GetType(String))
        dt.Columns.Add("Min Occupancy Override", GetType(Double))
        dt.Columns.Add("Max Occupancy Override", GetType(Double))
        dt.Columns.Add("Carts Per Day Override", GetType(Double))
        dt.Columns.Add("Total Carts Override", GetType(Double))
        dt.Columns.Add("Reason", GetType(String))
        dt.Columns.Add("Last Updated By", GetType(String))
        dt.Columns.Add("Last Updated At", GetType(String))
        dt.Columns.Add("Active", GetType(Boolean))
        dt.Columns.Add("Sort Order", GetType(Integer))

        For rec As Integer = 1 To preactor.RecordCount(fmt)

            Dim row As DataRow = dt.NewRow()

            row("__Record") = rec
            row("Number") = preactor.ReadFieldInt(fmt, fNumber, rec)
            row("Kiln") = preactor.ReadFieldString(fmt, fKiln, rec)
            row("Kiln Type") = preactor.ReadFieldString(fmt, fKilnType, rec)
            row("Availability Status") = preactor.ReadFieldString(fmt, fStatus, rec)

            row("Available From") = FormatDateForGrid(fmt, fAvailableFrom, rec)
            row("Available Until") = FormatDateForGrid(fmt, fAvailableUntil, rec)
            row("Override Start Time") = FormatDateForGrid(fmt, fOverrideStart, rec)
            row("Override End Time") = FormatDateForGrid(fmt, fOverrideEnd, rec)

            row("Min Occupancy Override") = preactor.ReadFieldDouble(fmt, fMinOcc, rec)
            row("Max Occupancy Override") = preactor.ReadFieldDouble(fmt, fMaxOcc, rec)
            row("Carts Per Day Override") = preactor.ReadFieldDouble(fmt, fCartsPerDay, rec)
            row("Total Carts Override") = preactor.ReadFieldDouble(fmt, fTotalCarts, rec)

            row("Reason") = preactor.ReadFieldString(fmt, fReason, rec)
            row("Last Updated By") = preactor.ReadFieldString(fmt, fLastUpdatedBy, rec)
            row("Last Updated At") = FormatDateForGrid(fmt, fLastUpdatedAt, rec)
            row("Active") = preactor.ReadFieldBool(fmt, fActive, rec)
            row("Sort Order") = preactor.ReadFieldInt(fmt, fSortOrder, rec)

            dt.Rows.Add(row)

        Next

        grid.DataSource = dt

        grid.Columns("__Record").Visible = False
        grid.Columns("Number").ReadOnly = True
        grid.Columns("Kiln").ReadOnly = True
        grid.Columns("Kiln Type").ReadOnly = True
        grid.Columns("Last Updated By").ReadOnly = True
        grid.Columns("Last Updated At").ReadOnly = True

        lblStatus.Text = "Loaded " & dt.Rows.Count.ToString(CultureInfo.InvariantCulture) & " kiln availability records."

    End Sub
    Private Function FormatDateForGrid(fmt As Integer, fieldNumber As Integer, rec As Integer) As String

        Try
            Dim d As DateTime = preactor.ReadFieldDateTime(fmt, fieldNumber, rec)

            If d.Year <= 1900 Then
                Return ""
            End If

            Return d.ToString("dd-MM-yyyy HH:mm", CultureInfo.InvariantCulture)

        Catch
            Return ""
        End Try

    End Function

    Private Sub RefreshClicked(sender As Object, e As EventArgs)

        Try
            LoadKilnAvailabilityTable()
        Catch ex As Exception
            MessageBox.Show("Refresh failed: " & ex.Message,
                            "Kiln Availability",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error)
        End Try

    End Sub

    Private Sub SaveClicked(sender As Object, e As EventArgs)

        Try
            Dim dt As DataTable = TryCast(grid.DataSource, DataTable)

            If dt Is Nothing Then
                Return
            End If

            Dim fmt As Integer = preactor.GetFormatNumber(TABLE_NAME)

            Dim fStatus As Integer = preactor.GetFieldNumber(fmt, "Availability Status")
            Dim fAvailableFrom As Integer = preactor.GetFieldNumber(fmt, "Available From")
            Dim fAvailableUntil As Integer = preactor.GetFieldNumber(fmt, "Available Until")
            Dim fOverrideStart As Integer = preactor.GetFieldNumber(fmt, "Override Start Time")
            Dim fOverrideEnd As Integer = preactor.GetFieldNumber(fmt, "Override End Time")
            Dim fMinOcc As Integer = preactor.GetFieldNumber(fmt, "Min Occupancy Override")
            Dim fMaxOcc As Integer = preactor.GetFieldNumber(fmt, "Max Occupancy Override")
            Dim fCartsPerDay As Integer = preactor.GetFieldNumber(fmt, "Carts Per Day Override")
            Dim fTotalCarts As Integer = preactor.GetFieldNumber(fmt, "Total Carts Override")
            Dim fReason As Integer = preactor.GetFieldNumber(fmt, "Reason")
            Dim fLastUpdatedAt As Integer = preactor.GetFieldNumber(fmt, "Last Updated At")
            Dim fActive As Integer = preactor.GetFieldNumber(fmt, "Active")
            Dim fSortOrder As Integer = preactor.GetFieldNumber(fmt, "Sort Order")

            For Each row As DataRow In dt.Rows

                Dim rec As Integer = Convert.ToInt32(row("__Record"), CultureInfo.InvariantCulture)

                preactor.WriteField(fmt, fStatus, rec, Convert.ToString(row("Availability Status"), CultureInfo.InvariantCulture))

                WriteDateIfProvided(fmt, fAvailableFrom, rec, row("Available From"))
                WriteDateIfProvided(fmt, fAvailableUntil, rec, row("Available Until"))
                WriteDateIfProvided(fmt, fOverrideStart, rec, row("Override Start Time"))
                WriteDateIfProvided(fmt, fOverrideEnd, rec, row("Override End Time"))

                preactor.WriteField(fmt, fMinOcc, rec, SafeDouble(row("Min Occupancy Override")))
                preactor.WriteField(fmt, fMaxOcc, rec, SafeDouble(row("Max Occupancy Override")))
                preactor.WriteField(fmt, fCartsPerDay, rec, SafeDouble(row("Carts Per Day Override")))
                preactor.WriteField(fmt, fTotalCarts, rec, SafeDouble(row("Total Carts Override")))

                preactor.WriteField(fmt, fReason, rec, Convert.ToString(row("Reason"), CultureInfo.InvariantCulture))
                preactor.WriteField(fmt, fActive, rec, SafeBool(row("Active")))
                preactor.WriteField(fmt, fSortOrder, rec, SafeInt(row("Sort Order")))

                preactor.WriteField(fmt, fLastUpdatedAt, rec, DateTime.Now)

            Next

            lblStatus.Text = "Saved at " & DateTime.Now.ToString("HH:mm:ss", CultureInfo.InvariantCulture)

            MessageBox.Show("Kiln availability saved.",
                            "Kiln Availability",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information)

            LoadKilnAvailabilityTable()

        Catch ex As Exception
            MessageBox.Show("Save failed: " & ex.Message,
                            "Kiln Availability",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error)
        End Try

    End Sub

    Private Sub WriteDateIfProvided(fmt As Integer, fieldNumber As Integer, rec As Integer, value As Object)

        If value Is Nothing OrElse value Is DBNull.Value Then
            Return
        End If

        Dim raw As String = Convert.ToString(value, CultureInfo.InvariantCulture).Trim()

        If raw = "" Then
            Return
        End If

        Dim d As DateTime

        If DateTime.TryParseExact(raw,
                                  "dd-MM-yyyy HH:mm",
                                  CultureInfo.InvariantCulture,
                                  DateTimeStyles.None,
                                  d) Then

            preactor.WriteField(fmt, fieldNumber, rec, d)
            Return

        End If

        If DateTime.TryParse(raw, CultureInfo.InvariantCulture, DateTimeStyles.None, d) Then
            preactor.WriteField(fmt, fieldNumber, rec, d)
        End If

    End Sub

    Private Function SafeDouble(value As Object) As Double

        If value Is Nothing OrElse value Is DBNull.Value Then
            Return -1.0
        End If

        Dim d As Double

        If Double.TryParse(Convert.ToString(value, CultureInfo.InvariantCulture),
                           NumberStyles.Any,
                           CultureInfo.InvariantCulture,
                           d) Then
            Return d
        End If

        Return -1.0

    End Function

    Private Function SafeInt(value As Object) As Integer

        If value Is Nothing OrElse value Is DBNull.Value Then
            Return 0
        End If

        Dim i As Integer

        If Integer.TryParse(Convert.ToString(value, CultureInfo.InvariantCulture),
                            NumberStyles.Any,
                            CultureInfo.InvariantCulture,
                            i) Then
            Return i
        End If

        Return 0

    End Function

    Private Function SafeBool(value As Object) As Boolean

        If value Is Nothing OrElse value Is DBNull.Value Then
            Return False
        End If

        If TypeOf value Is Boolean Then
            Return DirectCast(value, Boolean)
        End If

        Dim s As String = Convert.ToString(value, CultureInfo.InvariantCulture).Trim()

        Return s.Equals("1") _
            OrElse s.Equals("true", StringComparison.OrdinalIgnoreCase) _
            OrElse s.Equals("yes", StringComparison.OrdinalIgnoreCase)

    End Function
End Class
