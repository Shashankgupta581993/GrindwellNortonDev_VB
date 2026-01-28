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
Imports Preactor
Imports Preactor.Interop.PreactorObject


<ComVisible(True)>
<Microsoft.VisualBasic.ComClass("7e470a62-8dac-4ed5-a786-23a33df74c3e", "625abc3c-7727-4167-8b6e-ba6df010136e")>
Public Class AlgoSeq2


    Public Function Run(ByRef preactorComObject As PreactorObj, ByRef pespComObject As Object) As Integer

        Dim preactor As IPreactor = PreactorFactory.CreatePreactorObject(preactorComObject)
        Dim planningboard As IPlanningBoard = preactor.PlanningBoard

        'getMakeSpan(preactorComObject)


        'Dim connectionString = preactor.ParseShellString("{DB CONNECT STRING}")

        '' Create a connection to the database
        'Dim connection = New SqlConnection(connectionString)

        '' Open the connection
        'connection.Open()

        '' Define the sql to select the calendar states
        'Dim sql = "SELECT " +
        '"[Id], [Name], [Color], [Pattern], [Efficiency], [CostFactor], [IsSetupAllowed] " +
        '"FROM " +
        '"[Calendar].[CalendarStates]"

        '' Create a new command
        'Dim command = New SqlCommand(sql, connection)

        '' Execute the command and get a reader
        'Dim reader = command.ExecuteReader()

        '' Get the ordinals for the fields we are interested in
        'Dim efficiencyOrdinal = reader.GetOrdinal("Efficiency")
        'Dim nameOrdinal = reader.GetOrdinal("Name")

        '' Create a new string builder
        'Dim result = New StringBuilder()

        '' Loop through all of the rows
        'While (reader.Read())

        '    ' Get the state name and efficiency
        '    Dim stateName = reader.GetString(nameOrdinal)
        '    Dim efficiency = reader.GetDouble(efficiencyOrdinal) * 100

        '    ' Create a string like: StateName (100%)
        '    Dim format = String.Format("{0} ({1}%)", stateName, efficiency)

        '    ' Add it to the result
        '    result.AppendLine(format)

        'End While

        '' Close the connection
        'connection.Close()

        '' Display in a message box all of the states and their efficiencies
        'MessageBox.Show(result.ToString())

        ' ### this is connection to the DBconnection using class instance method'
        'Dim dbconnectioninstance As New DBConnection()
        'Dim result1 As Integer = dbconnectioninstance.GetConnect(preactorComObject)

        '''--------
        'Dim db As New DBConnection()

        'Dim sql As String = "Select OrdersId, Orderno, PartNo, Product,  OpNo, OperationName, SetupTime, IIF(BatchTime is not null,ROUND(BatchTime*24,1),ROUND(OpTimePerItem*24 ,1)) as [CalculatedTime]" +
        '                    ", ROUND(OpTimePerItem*24,1) as OpTimePerTime, ROUND(BatchTime*24,1) as [BatchTime]" +
        '                    ", DueDate, Quantity, Priority, DeliveryBuffer, StringAttribute3 from UserData.Orders where DatasetId=15"
        ''       Select Case OrdersId, Orderno, PartNo, Product,  OpNo, OperationName, SetupTime, IIF(BatchTime Is Not null,ROUND(BatchTime*24,1),ROUND(OpTimePerItem*24,1)) As [CalculatedTime]
        '', ROUND(OpTimePerItem*24,1) as OpTimePerTime, ROUND(BatchTime*24,1) as [BatchTime]
        '',(sum(IIF(BatchTime Is Not null,ROUND(EffectiveOpTime*24,1),ROUND(OpTimePerItem*24,1))) over (partition by OrderNo))/24 as [TotalTime]
        '', DueDate, Quantity, Priority, DeliveryBuffer, StringAttribute3 from UserData.Orders where DatasetId=15

        'Dim parameters As New List(Of SqlParameter) From {
        '    New SqlParameter("@IsActive", SqlDbType.Bit) With {.Value = True}
        '}

        'Dim dt As DataTable = db.ExecuteDataTable(sql, preactorComObject)
        'PrintDataTable(dt)
        '' dt is your "table structure" that you can use anywhere (bind to grid, loop, etc.)
        'ExportDataTableToCsv(dt, "DataExportTable2")

        'CsvTemplateGenerator.CreateCyclePlannerCsvTemplates("C:\Temp\VitrifiedCycleTest")
        '''--------

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


