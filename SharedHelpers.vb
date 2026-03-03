Option Strict On
Option Explicit On

Imports System
Imports System.Globalization
Imports System.IO
Imports System.Text
Imports Preactor

Public Module SharedHelpers

    Public Sub RequireColumn(dt As DataTable, name As String)
        If Not dt.Columns.Contains(name) Then Throw New ArgumentException($"Missing required column: '{name}'")
    End Sub

    Public Function SafeInt(o As Object) As Integer
        If o Is Nothing Then Return 0
        If TypeOf o Is Integer Then Return CInt(o)
        Dim s As String = o.ToString().Trim()
        Dim v As Integer
        If Integer.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, v) Then Return v
        Return 0
    End Function

    Public Function SafeDbl(o As Object) As Double
        If o Is Nothing Then Return 0
        Dim v As Double
        If Double.TryParse(o.ToString().Trim(), NumberStyles.Any, CultureInfo.InvariantCulture, v) Then Return v
        Return 0
    End Function

    Public Function SafeDate(o As Object) As DateTime
        If o Is Nothing Then Return DateTime.MinValue
        If TypeOf o Is DateTime Then Return CType(o, DateTime)
        Dim d As DateTime
        If DateTime.TryParse(o.ToString(), CultureInfo.InvariantCulture, DateTimeStyles.None, d) Then Return d
        Return DateTime.MinValue
    End Function

    Public Function SafeStr(o As Object) As String
        If o Is Nothing Then Return ""
        Return o.ToString()
    End Function

    Public Function SafeBool(o As Object) As Boolean
        If o Is Nothing Then Return False
        Dim s As String = o.ToString().Trim().ToUpperInvariant()
        Return s = "TRUE" OrElse s = "T" OrElse s = "1" OrElse s = "YES" OrElse s = "Y"
    End Function

    Public Function SafeArray(arr As String(), idx As Integer) As String
        If arr Is Nothing Then Return ""
        If idx < 0 OrElse idx >= arr.Length Then Return ""
        Return If(arr(idx), "")
    End Function

    Public Function IsTruthy(s As String) As Boolean
        If s Is Nothing Then Return False
        Dim u As String = s.Trim().ToUpperInvariant()
        Return u = "1" OrElse u = "TRUE" OrElse u = "T" OrElse u = "YES" OrElse u = "Y"
    End Function

    Public Function Csv(value As String) As String
        If value Is Nothing Then value = ""
        Dim mustQuote As Boolean = value.Contains(","c) OrElse value.Contains(""""c) OrElse value.Contains(ControlChars.Cr) OrElse value.Contains(ControlChars.Lf)
        If value.Contains(""""c) Then value = value.Replace("""", """""")
        If mustQuote Then Return """" & value & """"
        Return value
    End Function

    Public Function FormatDateOrBlank(d As DateTime) As String
        If d = DateTime.MinValue Then Return ""
        Return d.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture)
    End Function

    Public Function GetOrDefault(Of TKey, TValue)(dict As Dictionary(Of TKey, TValue), key As TKey, defaultValue As TValue) As TValue
        If dict Is Nothing Then Return defaultValue
        Dim v As TValue = defaultValue
        If dict.TryGetValue(key, v) Then Return v
        Return defaultValue
    End Function

    Public Function GetOrEmpty(Of TKey)(dict As Dictionary(Of TKey, String), key As TKey) As String
        If dict Is Nothing Then Return ""
        Dim v As String = ""
        If dict.TryGetValue(key, v) Then Return If(v, "")
        Return ""
    End Function

    Public Function GetOrEmptyDate(Of TKey)(dict As Dictionary(Of TKey, DateTime), key As TKey) As String
        If dict Is Nothing Then Return ""
        Dim v As DateTime
        If dict.TryGetValue(key, v) Then Return v.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture)
        Return ""
    End Function

    Public Function ParseDateDdMmYyyy(s As String) As DateTime
        If String.IsNullOrWhiteSpace(s) Then Return DateTime.MinValue
        Dim formats As String() = {"dd-MM-yyyy", "d-M-yyyy", "dd-M-yyyy", "d-MM-yyyy"}
        Dim dt As DateTime
        If DateTime.TryParseExact(s.Trim(), formats, CultureInfo.InvariantCulture, DateTimeStyles.None, dt) Then
            Return dt.Date
        End If
        Throw New FormatException("Invalid date: " & s)
    End Function

    Public Function ParseDueAsEndOfDay(o As Object) As DateTime
        Dim s As String = SafeStr(o).Trim()
        If s = "" Then Return DateTime.MinValue

        Dim d As DateTime
        If DateTime.TryParseExact(s,
                                  "dd-MM-yyyy",
                                  CultureInfo.InvariantCulture,
                                  DateTimeStyles.None,
                                  d) Then
            Return d.Date.AddDays(1).AddTicks(-1) ' end of day
        End If

        If DateTime.TryParse(s, CultureInfo.InvariantCulture, DateTimeStyles.None, d) Then
            Return d.Date.AddDays(1).AddTicks(-1)
        End If

        Return DateTime.MinValue
    End Function



    ' ----------------------
    ' Queue helpers
    ' ----------------------
    Public Function GetQueueSnapshot(ByVal planningboard As IPlanningBoard, ByVal queueName As String) As List(Of Integer)
        Dim snapshot As New List(Of Integer)()
        Dim pos As Integer = 1
        Dim opRec As Integer = 0
        While planningboard.GetOperationInQueue(queueName, pos, opRec)
            snapshot.Add(opRec)
            pos += 1
        End While
        Return snapshot
    End Function

    ' ----------------------
    ' small helper to access format field pair(s)
    ' ----------------------
    Public Function getformatfieldpair(ByVal preactor As IPreactor, Optional ByVal field As String = "Field", Optional ByVal format As String = "Format") As Preactor.FormatFieldPair
        Dim ffp As Preactor.FormatFieldPair = Nothing
        Dim ordersTable As Integer
        Dim fields As IEnumerable(Of Preactor.FormatFieldPair)

        Select Case field
            Case "DUE DATE", "PRIORITY", "EARLIEST START DATE"
                Return CType(preactor.FindFirstClassificationString(field), FormatFieldPair)
            Case "Operation Name", "Product", "OP NO", "STRING ATTRIBUTE 1", "STRING ATTRIBUTE 2", "ORDER NO", "QUANTITY", "TABLE ATTRIBUTE 1", "TABLE ATTRIBUTE 2", "TABLE ATTRIBUTE 3", "RESOURCE", "RESOURCE GROUP", "SETUP TIME", "OP TIME PER ITEM", "DATE ATTRIBUTE 1", "PART NO"
                ordersTable = preactor.FindFirstClassificationString("LAUNCH TIME").Value.FormatNumber
                fields = preactor.FindClassificationString(field)


                For Each field1 In fields
                    If (field1.FormatNumber = ordersTable) Then
                        Return field1
                    End If
                Next
            Case Else
                If format = "ORDERS" Then
                    Return CType(preactor.FindFirstClassificationString("LAUNCH TIME"), FormatFieldPair)
                End If
        End Select
        Return ffp
    End Function

    ' Creating the datastructure for routing information
    Public Function BuildRoutingSchema() As DataTable
        Dim dt As New DataTable("RoutingFromOpcenter")
        Dim cols As String() = {
            "OrdersID", "Order No", "Part Number", "Part Name", "Operation Number", "Operation Name",
            "Resource Group", "Required Resource", "Setup Time", "Time Per Item", "Sales Order", "Quantity",
            "Due Date", "Batch Time", "Process Time Type", "Tonnage", "Cycle Type", "Volume Occupancy",
            "Klin Type", "Firing buffer", "MTS/MTO", "MTS/MTO priority", "Que Time", "Pressing buffer",
            "Wheel Dia", "Wheel thickness", "Week start", "Pressing earliest start", "Pressing Due date",
            "Constaint Usage", "Constraint Qty", "firing earliest start date", "firing due date", "scheduled_start_time", "scheduled_end_time", "is_scheduled", "parent_record", "prev_op_is_scheduled"
        }
        For Each c In cols
            dt.Columns.Add(New DataColumn(c, GetType(Object)))
        Next
        Return dt
    End Function

    Public Function readOrderTable(ByVal preactor As IPreactor) As DataTable

        Dim planningboard As IPlanningBoard = preactor.PlanningBoard


        Dim dt As DataTable = BuildRoutingSchema()
        Dim ordersTable = preactor.GetFormatNumber("Orders")
        Dim orderNo = preactor.GetFieldNumber(ordersTable, "Order No.")
        Dim partNo = preactor.GetFieldNumber(ordersTable, "Part No.")
        Dim product = preactor.GetFieldNumber(ordersTable, "Product")
        Dim opNo = preactor.GetFieldNumber(ordersTable, "Op. No.")
        Dim opName = preactor.GetFieldNumber(ordersTable, "Operation Name")
        Dim resGroup = preactor.GetFieldNumber(ordersTable, "Resource Group")
        Dim res = preactor.GetFieldNumber(ordersTable, "Required Resource")
        Dim stpTime = preactor.GetFieldNumber(ordersTable, "Setup Time")
        Dim timePerItem = preactor.GetFieldNumber(ordersTable, "Op. Time per Item")
        Dim salesOrder = preactor.GetFieldNumber(ordersTable, "Operation Name")
        Dim Qty = preactor.GetFieldNumber(ordersTable, "Quantity")
        Dim dueDate = preactor.GetFieldNumber(ordersTable, "Due Date")
        Dim batchTime = preactor.GetFieldNumber(ordersTable, "Batch Time")
        Dim prsTimeType = preactor.GetFieldNumber(ordersTable, "Process Time Type")
        Dim tonnage = preactor.GetFieldNumber(ordersTable, "Numerical Attribute 4")
        Dim cycleType = preactor.GetFieldNumber(ordersTable, "Table Attribute 2")
        Dim klnType = preactor.GetFieldNumber(ordersTable, "Table Attribute 3")
        Dim volumeOcc = preactor.GetFieldNumber(ordersTable, "Numerical Attribute 5")
        Dim presEarlyStart = preactor.GetFieldNumber(ordersTable, "Date Attribute 1")
        Dim presDue = preactor.GetFieldNumber(ordersTable, "Date Attribute 2")
        Dim firingDue = preactor.GetFieldNumber(ordersTable, "Date Attribute 3")
        Dim mts = preactor.GetFieldNumber(ordersTable, "Table Attribute 1")
        Dim wheelDia = preactor.GetFieldNumber(ordersTable, "String Attribute 5")
        Dim wheelThck = preactor.GetFieldNumber(ordersTable, "String Attribute 4")
        Dim wheelPin = preactor.GetFieldNumber(ordersTable, "String Attribute 3")
        Dim schStart = preactor.GetFieldNumber(ordersTable, "Start Time")
        Dim schEnd = preactor.GetFieldNumber(ordersTable, "End Time")
        'Dim parentRecord = preactor.GetFieldNumber(ordersTable, "Belongs to Order No.")
        Dim rowCount = preactor.RecordCount(ordersTable)
        For rec As Integer = 1 To rowCount
            Dim r As DataRow = dt.NewRow()
            r("OrdersID") = rec
            r("Order No") = preactor.ReadFieldString(ordersTable, orderNo, rec)
            r("Part Number") = preactor.ReadFieldString(ordersTable, partNo, rec)
            r("Part Name") = preactor.ReadFieldString(ordersTable, product, rec)
            r("Operation Number") = preactor.ReadFieldInt(ordersTable, opNo, rec)
            r("Operation Name") = preactor.ReadFieldString(ordersTable, opName, rec)
            r("Resource Group") = preactor.ReadFieldString(ordersTable, resGroup, rec)
            r("Required Resource") = preactor.ReadFieldString(ordersTable, res, rec)
            r("Setup Time") = preactor.ReadFieldDouble(ordersTable, stpTime, rec) * 1440
            'r("Time Per Item")
            'r("Sales Order")
            r("Quantity") = preactor.ReadFieldInt(ordersTable, Qty, rec)
            r("Due Date") = preactor.ReadFieldDateTime(ordersTable, dueDate, rec)
            r("Batch Time") = preactor.ReadFieldDouble(ordersTable, batchTime, rec) * 1440
            r("Process Time Type") = preactor.ReadFieldString(ordersTable, prsTimeType, rec)
            r("Tonnage") = preactor.ReadFieldInt(ordersTable, Qty, rec)
            r("Cycle Type") = preactor.ReadFieldString(ordersTable, cycleType, rec)
            r("Volume Occupancy") = preactor.ReadFieldDouble(ordersTable, volumeOcc, rec)
            r("Klin Type") = preactor.ReadFieldInt(ordersTable, klnType, rec)
            'r("Firing buffer") = preactor.ReadFieldInt(ordersTable, , rec)
            r("MTS/MTO") = preactor.ReadFieldInt(ordersTable, mts, rec)
            'r("MTS/MTO priority") = preactor.ReadFieldInt(ordersTable, Qty, rec)
            'r("Que Time") = preactor.ReadFieldInt(ordersTable, Qty, rec)
            'r("Pressing buffer") = preactor.ReadFieldInt(ordersTable, Qty, rec)
            r("Wheel Dia") = preactor.ReadFieldString(ordersTable, wheelDia, rec)
            r("Wheel thickness") = preactor.ReadFieldString(ordersTable, wheelThck, rec)
            'r("Week start") = preactor.ReadFieldString(ordersTable, wheelPin, rec)
            r("Pressing earliest start") = preactor.ReadFieldDateTime(ordersTable, presEarlyStart, rec)
            r("Pressing Due date") = preactor.ReadFieldDateTime(ordersTable, presDue, rec)
            'r("Constaint Usage") = preactor.ReadFieldInt(ordersTable, Qty, rec)
            'r("Constraint Qty") = preactor.ReadFieldInt(ordersTable, Qty, rec)
            'r("firing earliest start date") = preactor.ReadFieldInt(ordersTable, Qty, rec)
            r("firing due date") = preactor.ReadFieldDateTime(ordersTable, firingDue, rec)
            If planningboard.IsOperationScheduled(rec) Then
                r("scheduled_start_time") = preactor.ReadFieldDateTime(ordersTable, schStart, rec)
                r("scheduled_end_time") = preactor.ReadFieldDateTime(ordersTable, schEnd, rec)
                r("is_scheduled") = True
            Else
                r("is_scheduled") = False
            End If
            If rec > 1 Then
                r("parent_record") = preactor.FindMatchingRecord(ordersTable, 1, rec, -1, SearchDirection.Backwards)
            Else r("parent_record") = 1
            End If
            If rec > 1 Then
                If preactor.ReadFieldInt(ordersTable, opNo, rec) = 300 Then
                    If planningboard.IsOperationScheduled(rec - 1) Then
                        r("prev_op_is_scheduled") = True
                    ElseIf planningboard.IsOperationScheduled(rec - 2) Then
                        r("prev_op_is_scheduled") = True
                    End If
                End If
            End If

            dt.Rows.Add(r)
        Next

        Return dt
    End Function

    ' Returns the end time of the last scheduled operation on a given resource.
    ' If nothing is scheduled on that resource, returns Nothing (you can swap to ScheduleHorizon.Start, Now, etc.)

    Public Function GetResourceLastScheduledEnd(
                                               preactor As IPreactor,
                                               planningboard As IPlanningBoard,
                                               resourceRec As Integer) As Nullable(Of DateTime)

        Dim ordersFmt As Integer = preactor.GetFormatNumber("Orders")

        ' NOTE: field name depends on your dataset (commonly "Required Resource").
        ' Use your PRTDF/field list for the exact name.
        Dim reqResFieldNo As Integer = preactor.GetFieldNumber(ordersFmt, "Resource")

        Dim lastEnd As Nullable(Of DateTime) = Nothing

        For opRec As Integer = 1 To preactor.RecordCount(ordersFmt)

            ' Filter: scheduled only
            If Not planningboard.IsOperationScheduled(opRec) Then Continue For

            ' Filter: operation belongs to this resource
            Dim opResRec As Integer = preactor.ReadFieldInt(ordersFmt, reqResFieldNo, opRec)
            If opResRec <> resourceRec Then Continue For

            ' Get scheduled timing
            Dim times As Nullable(Of Preactor.OperationResourceTimes) = planningboard.GetOperationTimes(opRec)
            If Not times.HasValue Then Continue For

            Dim opEnd As DateTime = times.Value.OperationTimes.ProcessEnd

            If (Not lastEnd.HasValue) OrElse (opEnd > lastEnd.Value) Then
                lastEnd = opEnd
            End If
        Next

        Return lastEnd
    End Function

    ' Minimal CSV escape: wrap in quotes if it contains comma or quote; double quotes inside.
    Public Function CsvEscape(value As String) As String
        If value Is Nothing Then Return ""
        Dim mustQuote = value.Contains(",") OrElse value.Contains("""") OrElse value.Contains(vbCr) OrElse value.Contains(vbLf)
        If value.Contains("""") Then value = value.Replace("""", """""")
        If mustQuote Then Return $"""{value}"""
        Return value
    End Function
End Module