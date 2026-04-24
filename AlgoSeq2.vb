Option Strict On
Option Explicit On

Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Globalization
Imports System.IO
Imports System.Linq
Imports System.Runtime.InteropServices
Imports System.Security
Imports System.Text
Imports System.Windows.Forms
Imports Microsoft.VisualBasic.FileIO
Imports Preactor
Imports Preactor.Interop.PreactorObject


<ComVisible(True)>
<Microsoft.VisualBasic.ComClass("7e470a62-8dac-4ed5-a786-23a33df74c3e", "625abc3c-7727-4167-8b6e-ba6df010136e")>
Public Class AlgoSeq2


    Public Function Run(ByRef preactorComObject As PreactorObj, ByRef pespComObject As Object) As Integer

        Dim preactor As IPreactor = PreactorFactory.CreatePreactorObject(preactorComObject)
        Dim planningboard As IPlanningBoard = preactor.PlanningBoard

        Dim reader As New CsvRoutingReader()
        Dim routingDt As DataTable = reader.ReadRoutingCsv()
        reader.AddExpectedFiringStartDate(routingDt)
        reader.AddFiringWeekAndBatchColumns_V2_ByExpectedDateAndCycle(routingDt, 1, 0.2)
        reader.AddPressingFields(routingDt)
        ExportDataTableToCsv(routingDt, "output1")
        reader.CreatePressingBatches_PseudoSchedule(routingDt,
                                                     10, 2, 2, "300", "Batch", "ExpectedFiringStartDate")
        ExportDataTableToCsv(routingDt, "output2")


        ' Example: loop through rows
        For Each row As DataRow In routingDt.Rows
            Debug.WriteLine(row(0).ToString())
        Next


        Return 0
    End Function


    Public Sub PrintDataTable(dt As DataTable)
        Debug.WriteLine("------ DATA TABLE ------")

        ' Print header
        Dim header As String = String.Join(" | ", dt.Columns.Cast(Of DataColumn).Select(Function(c) c.ColumnName))
        Debug.WriteLine(header)

        ' Print rows
        For Each row As DataRow In dt.Rows
            Dim line As String = String.Join(" | ", row.ItemArray.Select(Function(v) v.ToString()))
            Debug.WriteLine(line)
        Next

        Debug.WriteLine("-------------------------")
    End Sub



    Public Function CalculateMakespans(preactor As IPreactor) _
                                       As List(Of Tuple(Of Integer, String, TimeSpan))

        Dim results As New List(Of Tuple(Of Integer, String, TimeSpan))()

        ' --- Resolve Orders table ---
        Dim ordersLaunch = preactor.FindFirstClassificationString("LAUNCH TIME")
        If Not ordersLaunch.HasValue Then
            Debug.WriteLine("Orders: LAUNCH TIME field not found. Abort.")
            Return results
        End If
        Dim ordersTable As Integer = ordersLaunch.Value.FormatNumber

        ' --- Get Orders.NUMBER field ---
        Dim ordersNumber = preactor.FindFirstClassificationString("NUMBER")
        If Not ordersNumber.HasValue Then
            Debug.WriteLine("Orders: NUMBER field not found. Abort.")
            Return results
        End If

        ' --- Ops foreign key to orders ---
        Dim opsOrderNo = preactor.FindFirstClassificationString("ORDER NUMBER")
        If Not opsOrderNo.HasValue Then
            opsOrderNo = preactor.FindFirstClassificationString("ORDER NO")
            If Not opsOrderNo.HasValue Then
                Debug.WriteLine("Operations: ORDER NUMBER / ORDER NO field not found. Abort.")
                Return results
            End If
        End If

        ' --- Ops duration field ---
        Dim opDuration = preactor.FindFirstClassificationString("Op. Time per Item")
        If Not opDuration.HasValue Then
            Debug.WriteLine("Operations: OPERATION TIME field not found. Abort.")
            Return results
        End If

        ' --- Loop through all orders ---
        Dim nOrders As Integer = preactor.RecordCount(ordersTable)
        For orderRec As Integer = 1 To nOrders
            Dim orderNo As String = preactor.ReadFieldString(ordersNumber.Value, orderRec)
            If String.IsNullOrWhiteSpace(orderNo) Then Continue For

            Dim total As TimeSpan = TimeSpan.Zero

            ' Find all ops linked to this order
            Dim opRec As Integer = preactor.FindMatchingRecord(opsOrderNo.Value, 0, orderNo)
            While opRec > 0
                ' Read OPERATION TIME (assumed to be numeric, in minutes or hours depending on dataset)
                Dim durVal As Double = preactor.ReadFieldDouble(opDuration.Value, opRec)
                total += TimeSpan.FromMinutes(durVal) ' adjust if your units are hours/seconds

                ' Next operation with same orderNo
                opRec = preactor.FindMatchingRecord(opsOrderNo.Value, opRec, orderNo)
            End While

            results.Add(Tuple.Create(orderRec, orderNo, total))
        Next

        Return results
    End Function

    Public Function getMakeSpan(ByRef preactorComObject As PreactorObj) As Integer
        Dim preactor As IPreactor = PreactorFactory.CreatePreactorObject(preactorComObject)
        Dim planningboard As IPlanningBoard = preactor.PlanningBoard

        Dim ordersParent As Preactor.FormatFieldPair
        Dim parentRecord As Integer
        Dim familyFields As IEnumerable(Of Preactor.FormatFieldPair)
        Dim ordersTable As Integer
        Dim orderno As String
        Dim duedate As DateTime
        Dim durationvalue As Double
        Dim duration As TimeSpan
        Dim firingrec As Integer
        Dim firingtime As Nullable(Of Preactor.OperationResourceTimes)
        Dim firop As Integer = 300
        Dim ordernoNxt As String
        Dim weekdayNumber As Integer
        Dim cal As Calendar = CultureInfo.InvariantCulture.Calendar
        Dim weekRule As CalendarWeekRule = CalendarWeekRule.FirstFourDayWeek
        Dim firstDayOfWeek As DayOfWeek = DayOfWeek.Monday
        Dim weekNumber As Integer
        Dim dt As New DataTable("Orderdata")
        Dim cycletype As String
        Dim volume As String
        Dim klintype As String

        dt.Columns.Add("OrderNo", GetType(String))
        dt.Columns.Add("Duration", GetType(String))
        dt.Columns.Add("FiringStart", GetType(String))
        dt.Columns.Add("Weekday", GetType(String))
        dt.Columns.Add("WeekNo", GetType(String))
        dt.Columns.Add("CycleType", GetType(String))
        dt.Columns.Add("Volume", GetType(String))
        dt.Columns.Add("Klintype", GetType(String))

        ordersParent = New FormatFieldPair()
        familyFields = preactor.FindClassificationString("FAMILY")
        ordersTable = preactor.FindFirstClassificationString("LAUNCH TIME").Value.FormatNumber
        For Each familyField In familyFields
            If (familyField.FormatNumber = ordersTable) Then
                ordersParent = familyField
            End If
        Next

        parentRecord = preactor.FindMatchingRecord(ordersParent, parentRecord, -1)
        While (parentRecord > 0)
            'MsgBox(parentRecord)
            orderno = preactor.ReadFieldString(ordersTable, "Order No.", parentRecord)
            duedate = preactor.ReadFieldDateTime(ordersTable, "Due Date", parentRecord)
            durationvalue = preactor.ReadFieldDouble(ordersTable, "Make Span", parentRecord)
            duration = TimeSpan.FromDays(durationvalue)
            firingrec = preactor.FindMatchingRecord(ordersTable, "Op. No.", parentRecord, firop)
            If (firingrec > 0) Then
                ordernoNxt = preactor.ReadFieldString(ordersTable, "Order No.", firingrec)
                If (ordernoNxt = orderno) Then
                    firingtime = planningboard.GetOperationTimes(firingrec).Value
                    cycletype = preactor.ReadFieldString("Orders", "String Attribute 3", firingrec)
                    volume = preactor.ReadFieldString("Orders", "String Attribute 4", firingrec)
                    klintype = preactor.ReadFieldString("Orders", "String Attribute 5", firingrec)
                    ' your code is here

                    weekdayNumber = (CInt(firingtime.Value.OperationTimes.ProcessStart.DayOfWeek) + 6) Mod 7 + 1
                    weekNumber = cal.GetWeekOfYear(firingtime.Value.OperationTimes.ProcessStart, weekRule, firstDayOfWeek)

                    ' your code is here
                    MsgBox("Order No.:" & orderno & vbNewLine & "duration total :" & durationvalue.ToString() & vbNewLine & "firingrec:" & firingtime.Value.OperationTimes.ProcessStart.ToString() & vbNewLine & "weekdaynumber:" & weekdayNumber & vbNewLine & "weeknumber:" & weekNumber)
                    dt.Rows.Add(orderno, durationvalue.ToString(), firingtime.Value.OperationTimes.ProcessStart.ToString(), weekdayNumber, weekNumber, cycletype, volume, klintype)

                End If
            End If
            parentRecord = preactor.FindMatchingRecord(ordersParent, parentRecord, -1)
        End While
        'For Each row As DataRow In dt.Rows
        '    Console.WriteLine($"{row("OrderNo")} | {row("Duration")} | {row("FiringStart")} | {row("Weekday")}| {row("WeekNo")}")
        'Next

        'Console.ReadLine()
        ExportDataTableToCsv(dt, "ExportData")
        Return 0
    End Function

    Public Function schedule(ByRef preactorComObject As PreactorObj) As Integer
        Dim preactor As IPreactor = PreactorFactory.CreatePreactorObject(preactorComObject)
        Dim planningboard As IPlanningBoard = preactor.PlanningBoard

        Dim ordersParent As Preactor.FormatFieldPair
        Dim dueDateField As Nullable(Of Preactor.FormatFieldPair)
        Dim priorityField As Nullable(Of Preactor.FormatFieldPair)
        Dim parentRecord As Integer
        Dim SequenceMode As Preactor.SequenceMode
        Dim familyFields As IEnumerable(Of Preactor.FormatFieldPair)
        Dim ordersTable As Integer
        Dim opRec As Integer
        Dim orderno As String
        Dim ResRec As Integer
        Dim ResRecs As IEnumerable(Of Integer)
        Dim opTimes As Double = 0
        'Nullable(Of Preactor.OperationTimes)
        Dim duedate As DateTime
        Dim duration As TimeSpan

        ordersParent = New FormatFieldPair()
        familyFields = preactor.FindClassificationString("FAMILY")
        ordersTable = preactor.FindFirstClassificationString("LAUNCH TIME").Value.FormatNumber
        For Each familyField In familyFields
            If (familyField.FormatNumber = ordersTable) Then
                ordersParent = familyField
            End If
        Next

        parentRecord = preactor.FindMatchingRecord(ordersParent, parentRecord, -1)
        While (parentRecord > 0)
            'MsgBox(parentRecord)
            orderno = preactor.ReadFieldString(ordersTable, "Order No.", parentRecord)
            duedate = preactor.ReadFieldDateTime(ordersTable, "Due Date", parentRecord)
            'MsgBox(orderno)
            opRec = parentRecord
            While (opRec > 0)
                'MsgBox(opRec)
                opTimes += preactor.ReadFieldDouble("Orders", "Op. Time per Item", opRec)
                duration += TimeSpan.FromDays(opTimes)
                MsgBox(orderno & "    RecNo:" & opRec & "    Duration:" & duration.ToString())
                opRec = preactor.FindMatchingRecord("Orders", "Order No.", opRec, orderno)
            End While
            'opTimes = opTime
            MsgBox(orderno & "duration total:" & duration.ToString())
            parentRecord = preactor.FindMatchingRecord(ordersParent, parentRecord, -1)
        End While

        Return 0
    End Function

    ' Reads a CSV file into a DataTable with flexible columns and safe value handling.
    ' If the file has headers, they are used and de-duplicated; otherwise columns are auto-generated.
    Public Function ReadCsvFlexible(filePath As String,
                                    Optional delimiter As Char = ","c,
                                    Optional hasHeader As Boolean = True,
                                    Optional encoding As Encoding = Nothing,
                                    Optional skipMalformedLines As Boolean = True) As DataTable

        If String.IsNullOrWhiteSpace(filePath) Then
            Throw New ArgumentException("filePath is required.", NameOf(filePath))
        End If

        If encoding Is Nothing Then
            encoding = Encoding.UTF8
        End If

        Dim dt As New DataTable(Path.GetFileNameWithoutExtension(filePath))

        Using parser As New TextFieldParser(filePath, encoding)
            parser.SetDelimiters(delimiter.ToString())
            parser.HasFieldsEnclosedInQuotes = True
            parser.TrimWhiteSpace = True

            Dim isFirstRow As Boolean = True
            While Not parser.EndOfData
                Dim fields As String() = Nothing
                Try
                    fields = parser.ReadFields()
                Catch ex As MalformedLineException
                    If skipMalformedLines Then
                        Continue While
                    End If
                    Throw
                End Try

                If fields Is Nothing Then
                    Continue While
                End If

                If isFirstRow AndAlso hasHeader Then
                    Dim headerNames As List(Of String) = BuildHeaderNames(fields)
                    For Each name As String In headerNames
                        dt.Columns.Add(name, GetType(String))
                    Next
                    isFirstRow = False
                    Continue While
                End If

                If isFirstRow Then
                    EnsureColumns(dt, fields.Length)
                    isFirstRow = False
                Else
                    EnsureColumns(dt, fields.Length)
                End If

                Dim row As DataRow = dt.NewRow()
                For i As Integer = 0 To fields.Length - 1
                    row(i) = If(fields(i), String.Empty)
                Next
                dt.Rows.Add(row)
            End While
        End Using

        Return dt
    End Function

    Private Sub EnsureColumns(dt As DataTable, requiredCount As Integer)
        While dt.Columns.Count < requiredCount
            Dim colName As String = "Column" & (dt.Columns.Count + 1).ToString(CultureInfo.InvariantCulture)
            dt.Columns.Add(colName, GetType(String))
        End While
    End Sub

    Private Function BuildHeaderNames(headers As String()) As List(Of String)
        Dim names As New List(Of String)(headers.Length)
        Dim seen As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)

        For i As Integer = 0 To headers.Length - 1
            Dim baseName As String = If(headers(i), String.Empty).Trim()
            If baseName.Length = 0 Then
                baseName = "Column" & (i + 1).ToString(CultureInfo.InvariantCulture)
            End If

            Dim finalName As String = baseName
            If seen.ContainsKey(baseName) Then
                seen(baseName) += 1
                finalName = baseName & "_" & seen(baseName).ToString(CultureInfo.InvariantCulture)
            Else
                seen(baseName) = 1
            End If

            names.Add(finalName)
        Next

        Return names
    End Function

    ' Exports a DataTable to a CSV file in the local directory as "DataTableExport.csv".
    Public Sub ExportDataTableToCsv(dt As DataTable, fileName As String)
        fileName &= ".csv"
        Dim filePath As String = Path.Combine(Directory.GetCurrentDirectory(), fileName)
        Using writer As New StreamWriter(filePath)
            ' Write headers
            Dim columnNames = dt.Columns.Cast(Of DataColumn)().Select(Function(col) col.ColumnName)
            writer.WriteLine(String.Join(",", columnNames))
            ' Write rows
            For Each row As DataRow In dt.Rows
                Dim fields = row.ItemArray.Select(Function(field) field.ToString().Replace(",", " "))
                writer.WriteLine(String.Join(",", fields))
            Next
        End Using
    End Sub

End Class


