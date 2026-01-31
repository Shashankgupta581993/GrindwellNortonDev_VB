Option Strict On
Option Explicit On

Imports System
Imports System.Globalization
Imports System.IO
Imports System.Runtime.InteropServices
Imports Microsoft.VisualBasic.FileIO
Imports Preactor
Imports Preactor.Interop.PreactorObject

<ComVisible(True)>
<Microsoft.VisualBasic.ComClass("4196dd4d-4e89-45a5-9ca5-4fc6dcf10308", "ef5b2382-ab81-47a5-9c8d-0826dcc85a0a")>
Public Class AlgoSeq4

    '========================
    ' USER-TUNABLE CONSTANTS
    '========================
    Private Const PRESS_OP_NUMBER As Integer = 200
    Private Const QUEUE_NAME As String = "Pressing200Queue"

    Private Shared ReadOnly CyclePriority As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase) From {
        {"150VT", 3},
        {"102VT", 2},
        {"65VT", 1}
    }

    ' ----------------------
    ' Public entry point
    ' ----------------------
    Public Function Run(ByRef preactorComObject As PreactorObj, ByRef pespComObject As Object) As Integer
        Dim preactor As IPreactor = PreactorFactory.CreatePreactorObject(preactorComObject)
        Dim planningboard As IPlanningBoard = preactor.PlanningBoard

        Dim ordersTable As Integer = preactor.FindFirstClassificationString("LAUNCH TIME").Value.FormatNumber

        ' Example: import a CSV, build pressing queue, create ranked queue and schedule
        Dim filePath As String = "D:\Documents\Opcenter\Cases\Grindwell Norton\Opcenter SC - Dev\Files\Templates\Routing.csv"
        Dim routingDt As DataTable = ImportRoutingCsvToDataTable(filePath)
        Dim currentDate As New System.DateTime(2025, 8, 1, 0, 0, 0)

        Dim pressingQueue As List(Of Integer) = BuildPressing200Queue(routingDt, currentDate)
        CreateRankedOperationQueue(preactor, planningboard, ordersTable, "JobsQueue", pressingQueue)

        ' Snapshot for debugging
        Dim jobsQueueSnapshot As List(Of Integer) = GetQueueSnapshot(planningboard, "JobsQueue")

        ' Simple scheduling loop: pop queue and place operation on earliest feasible resource
        Dim pos As Integer = 1
        Dim opRec As Integer = 0
        While planningboard.GetOperationInQueue("JobsQueue", pos, opRec)
            planningboard.RemoveOperationFromQueue("JobsQueue", opRec)

            Dim bestRes As Integer = 0
            Dim bestOpTimes As Nullable(Of OperationTimes) = Nothing

            For Each res In planningboard.FindResources(opRec)
                Dim ot = planningboard.TestOperationOnResource(opRec, res, planningboard.TerminatorTime)
                If ot.HasValue Then
                    If Not bestOpTimes.HasValue OrElse ot.Value.ChangeStart < bestOpTimes.Value.ChangeStart Then
                        bestRes = res
                        bestOpTimes = ot
                    End If
                End If
            Next

            If bestOpTimes.HasValue AndAlso bestRes > 0 Then
                planningboard.PutOperationOnResource(opRec, bestRes, bestOpTimes.Value.ChangeStart)
            End If

            pos += 1
        End While

        ' Append schedule times from board for a few operation numbers (example)
        routingDt = AppendOperationTimesFromBoard(routingDt, preactor, planningboard, 100)
        routingDt = AppendOperationTimesFromBoard(routingDt, preactor, planningboard, 200)

        ' Build firing plan using firing optimizer (external class)
        Dim minOcc As Double = 0.8
        Dim maxOcc As Double = 1.0
        Dim firingObj As New firingOptimizer_vf
        Dim plan = firingObj.BuildBatchKilnPlan(routingDt, "D:\Documents\Opcenter\Cases\Grindwell Norton\Opcenter SC - Dev\Files\kilndata.csv", currentDate, minOcc, maxOcc, batchStartDelayMins:=60)

        ' 1) iterate firing queue (these are op 300 record numbers)
        For Each firingOpRec As Integer In plan.QueueFiringOpRecs

            ' 2) get batch metadata
            Dim batchNo As Integer = plan.BatchNoByFiringOpRec(firingOpRec)
            Dim batchStart As DateTime = plan.BatchStartByBatchNo(batchNo).AddMinutes(60)
            Dim batchEnd As DateTime = plan.BatchEndByBatchNo(batchNo)
            Dim kilnName As String = plan.KilnByBatchNo(batchNo)
            Dim batchKind As String = plan.BatchKindByBatchNo(batchNo)

            Select Case (kilnName)
                Case "AKLN"
                    planningboard.TestOperationOnResource(firingOpRec, 66, batchStart)
                    planningboard.PutOperationOnResource(firingOpRec, 66, batchStart)
                    'planningboard.LockOperation(firingOpRec, OperationSelection.ThisOperation, True)
                    'Dim test = planningboard.TestOperationOnResource(planningboard.GetNextOperation(firingOpRec, 1), 64, batchEnd)
                    'planningboard.PutOperationOnResource(planningboard.GetNextOperation(firingOpRec, 1), 64, test.Value.ProcessStart)
                Case "BKLN"
                    planningboard.PutOperationOnResource(firingOpRec, 67, batchStart)
                    planningboard.PutOperationOnResource(planningboard.GetNextOperation(firingOpRec, 1), 64, batchEnd.AddMinutes(1))

                Case "CKLN"
                    planningboard.PutOperationOnResource(firingOpRec, 68, batchStart)
                    planningboard.PutOperationOnResource(planningboard.GetNextOperation(firingOpRec, 1), 64, batchEnd.AddMinutes(1))

                Case "DKLN"
                    planningboard.PutOperationOnResource(firingOpRec, 69, batchStart)
                    planningboard.PutOperationOnResource(planningboard.GetNextOperation(firingOpRec, 1), 64, batchEnd.AddMinutes(1))

                Case "RKLN"
                    planningboard.PutOperationOnResource(firingOpRec, 70, batchStart)
                    planningboard.PutOperationOnResource(planningboard.GetNextOperation(firingOpRec, 1), 64, batchEnd.AddMinutes(1))

                Case "NKLN"
                    planningboard.PutOperationOnResource(firingOpRec, 71, batchStart)
                    planningboard.PutOperationOnResource(planningboard.GetNextOperation(firingOpRec, 1), 64, batchEnd.AddMinutes(1))
            End Select


            ' 3) schedule operations in your preferred order:
            ' - loading (290/291) before batchStart (you’ll locate the correct opRec and assign)
            ' - firing (300) at batchStart on kilnName
            ' - unloading (310/311) after batchEnd
            '
            ' You can still validate kiln assignment here with TestOperationOnResource()
        Next

        firingObj.ExportPlanToCsv(plan, "D:\Documents\Opcenter\Cases\Grindwell Norton\Opcenter SC - Dev\Files\")



        ' Export plan
        firingObj.ExportPlanToCsv(plan, "D:\Documents\Opcenter\Cases\Grindwell Norton\Opcenter SC - Dev\Files\")
        Dim resrec As Integer = planningboard.GetResourceNumber("ULDBICK")

        Return 0
    End Function

    ' ----------------------
    ' Queue helpers
    ' ----------------------
    Private Function GetQueueSnapshot(ByVal planningboard As IPlanningBoard, ByVal queueName As String) As List(Of Integer)
        Dim snapshot As New List(Of Integer)()
        Dim pos As Integer = 1
        Dim opRec As Integer = 0
        While planningboard.GetOperationInQueue(queueName, pos, opRec)
            snapshot.Add(opRec)
            pos += 1
        End While
        Return snapshot
    End Function

    Private Function CreateRankedOperationQueue(ByRef preactor As IPreactor, ByVal planningboard As IPlanningBoard,
                                                ByVal ordersTable As Integer, ByVal QName As String, ByVal queue As List(Of Integer)) As Integer
        planningboard.CreateQueue(QName)
        For Each q In queue
            planningboard.AddOperationToQueue(QName, q, QueuePosition.End)
        Next
        Return 0
    End Function

    ' ----------------------
    ' Pressing 200 queue builder
    ' ----------------------
    Public Function BuildPressing200Queue(dt As DataTable, currentDate As DateTime, Optional approachingDays As Integer = 2) As List(Of Integer)
        If dt Is Nothing Then Throw New ArgumentNullException(NameOf(dt))

        SharedHelpers.RequireColumn(dt, "OrdersID")
        SharedHelpers.RequireColumn(dt, "Operation Number")
        SharedHelpers.RequireColumn(dt, "Pressing earliest start")
        SharedHelpers.RequireColumn(dt, "pressing due date")
        SharedHelpers.RequireColumn(dt, "Wheel Dia")
        SharedHelpers.RequireColumn(dt, "Wheel thickness")
        SharedHelpers.RequireColumn(dt, "Cycle Type")

        Dim today As DateTime = currentDate.Date
        Dim approachCutoff As DateTime = today.AddDays(approachingDays)

        Dim candidates As New List(Of Candidate)()

        For Each r As DataRow In dt.Rows
            Dim opNo = SharedHelpers.SafeInt(r("Operation Number"))
            If opNo <> PRESS_OP_NUMBER Then Continue For

            Dim orderId = SharedHelpers.SafeInt(r("OrdersID"))
            If orderId <= 0 Then Continue For

            Dim earliest = SharedHelpers.SafeDate(r("Pressing earliest start")).Date
            Dim due = SharedHelpers.SafeDate(r("Pressing Due date")).Date

            Dim missingEarliest = (earliest = DateTime.MinValue)
            Dim missingDue = (due = DateTime.MinValue)

            Dim tier As Integer
            If Not missingEarliest AndAlso Not missingDue AndAlso earliest <= approachCutoff AndAlso due >= today Then
                tier = 0
            ElseIf Not missingDue AndAlso due < today Then
                tier = 1
            Else
                tier = 2
            End If

            Dim wheelDia = SharedHelpers.SafeStr(r("Wheel Dia")).Trim()
            Dim wheelPin = SharedHelpers.SafeStr(r("Wheel thickness")).Trim()
            Dim cycleType = SharedHelpers.SafeStr(r("Cycle Type")).Trim()
            Dim cycleRank = GetCycleRank(cycleType)

            candidates.Add(New Candidate With {
                .OrdersID = orderId,
                .Earliest = earliest,
                .Due = due,
                .Tier = tier,
                .WheelDia = wheelDia,
                .WheelPin = wheelPin,
                .TypeKey = wheelDia & "|" & wheelPin,
                .CycleRank = cycleRank,
                .MissingEarliest = missingEarliest,
                .MissingDue = missingDue
            })
        Next

        Dim sorted = candidates.OrderBy(Function(c) c.Tier) _
                                   .ThenBy(Function(c) If(c.MissingDue, DateTime.MaxValue, c.Due)) _
                                   .ThenBy(Function(c) If(c.MissingEarliest, DateTime.MaxValue, c.Earliest)) _
                                   .ThenByDescending(Function(c) c.CycleRank) _
                                   .ThenBy(Function(c) c.TypeKey) _
                                   .ThenBy(Function(c) c.OrdersID) _
                                   .ToList()

        Dim batched = GreedyTypeBatchingWithinTier(sorted, lookahead:=50)
        Return batched.Select(Function(c) c.OrdersID).Distinct().ToList()
    End Function

    Private Function GreedyTypeBatchingWithinTier(sorted As List(Of Candidate), Optional lookahead As Integer = 50) As List(Of Candidate)
        If sorted.Count <= 2 Then Return sorted

        Dim work As New List(Of Candidate)(sorted)
        Dim result As New List(Of Candidate)(work.Count)

        Dim i As Integer = 0
        While i < work.Count
            Dim cur = work(i)
            result.Add(cur)

            If Not String.IsNullOrEmpty(cur.TypeKey) Then
                Dim pulled As Integer = 0
                Dim j As Integer = i + 1
                While j < work.Count AndAlso pulled < lookahead
                    If work(j).Tier = cur.Tier AndAlso work(j).TypeKey = cur.TypeKey AndAlso work(j).Due = cur.Due Then
                        result.Add(work(j))
                        work.RemoveAt(j)
                        pulled += 1
                        Continue While
                    End If
                    j += 1
                End While
            End If

            i += 1
        End While

        Return result
    End Function

    Private Function GetCycleRank(cycleType As String) As Integer
        If String.IsNullOrWhiteSpace(cycleType) Then Return 0
        Select Case cycleType.Trim().ToUpperInvariant()
            Case "150VT" : Return 3
            Case "102VT" : Return 2
            Case "65VT" : Return 1
            Case Else : Return 0
        End Select
    End Function

    ' ----------------------
    ' CSV import helper
    ' ----------------------
    Public Function ImportRoutingCsvToDataTable(csvPath As String) As DataTable
        If String.IsNullOrWhiteSpace(csvPath) Then Throw New ArgumentException("csvPath is empty.")
        If Not IO.File.Exists(csvPath) Then Throw New IO.FileNotFoundException("CSV not found.", csvPath)

        Dim dt As New DataTable("Routing")
        Using parser As New TextFieldParser(csvPath)
            parser.TextFieldType = FieldType.Delimited
            parser.SetDelimiters(",")
            parser.HasFieldsEnclosedInQuotes = True
            parser.TrimWhiteSpace = True

            If parser.EndOfData Then Return dt

            Dim headers = parser.ReadFields()
            If headers Is Nothing OrElse headers.Length = 0 Then Return dt

            For Each h In headers
                Dim colName = h.Trim()
                Dim colType As Type = GetType(String)
                Select Case colName
                    Case "OrdersID", "Order No", "Part Number", "Sales Order", "Quantity", "Operation Number"
                        colType = GetType(Integer)
                    Case "Setup Time", "Time Per Item", "Batch Time", "Tonnage", "Volume Occupancy", "Firing buffer", "MTS/MTO", "MTS/MTO priority", "Que Time", "Pressing buffer", "Wheel thickness", "Week start"
                        colType = GetType(Double)
                    Case "Due Date", "Pressing earliest start", "pressing due date"
                        colType = GetType(DateTime)
                    Case Else
                        colType = GetType(String)
                End Select
                dt.Columns.Add(New DataColumn(colName, colType))
            Next

            While Not parser.EndOfData
                Dim fields = parser.ReadFields()
                If fields Is Nothing Then Continue While
                Dim row = dt.NewRow()
                For i As Integer = 0 To dt.Columns.Count - 1
                    Dim raw As String = If(i < fields.Length, fields(i), "")
                    raw = If(raw, "").Trim()
                    Dim col = dt.Columns(i)
                    If String.IsNullOrEmpty(raw) Then
                        row(i) = GetDefaultValue(col.DataType)
                        Continue For
                    End If
                    Try
                        If col.DataType Is GetType(Integer) Then
                            row(i) = Integer.Parse(raw, CultureInfo.InvariantCulture)
                        ElseIf col.DataType Is GetType(Double) Then
                            row(i) = Double.Parse(raw, CultureInfo.InvariantCulture)
                        ElseIf col.DataType Is GetType(DateTime) Then
                            row(i) = SharedHelpers.ParseDateDdMmYyyy(raw)
                        Else
                            row(i) = raw
                        End If
                    Catch
                        row(i) = GetDefaultValue(col.DataType)
                    End Try
                Next
                dt.Rows.Add(row)
            End While
        End Using
        Return dt
    End Function

    Private Function GetDefaultValue(t As Type) As Object
        If t Is GetType(String) Then Return ""
        If t Is GetType(Integer) Then Return 0
        If t Is GetType(Double) Then Return 0.0R
        If t Is GetType(DateTime) Then Return DateTime.MinValue
        Return Nothing
    End Function

    ' ----------------------
    ' Append operation times from planning board into CSV datatable
    ' ----------------------
    Public Function AppendOperationTimesFromBoard(dt As DataTable,
                                                 preactor As IPreactor,
                                                 planningboard As IPlanningBoard,
                                                 opNumber As Integer) As DataTable
        If dt Is Nothing Then Throw New ArgumentNullException(NameOf(dt))
        If preactor Is Nothing Then Throw New ArgumentNullException(NameOf(preactor))
        If planningboard Is Nothing Then Throw New ArgumentNullException(NameOf(planningboard))

        SharedHelpers.RequireColumn(dt, "OrdersID")
        SharedHelpers.RequireColumn(dt, "Operation Number")

        Dim startColName As String = "scheduled_start_time"
        Dim endColName As String = "scheduled_end_time"
        Dim schColName As String = "is_scheduled"

        If Not dt.Columns.Contains(startColName) Then dt.Columns.Add(startColName, GetType(DateTime))
        If Not dt.Columns.Contains(endColName) Then dt.Columns.Add(endColName, GetType(DateTime))
        If Not dt.Columns.Contains(schColName) Then dt.Columns.Add(schColName, GetType(Boolean))

        Dim ordersTable As Integer = preactor.FindFirstClassificationString("LAUNCH TIME").Value.FormatNumber
        Dim opNoFields = preactor.FindClassificationString("OP NO")
        Dim ordersOpNoField As Preactor.FormatFieldPair = Nothing
        For Each f In opNoFields
            If f.FormatNumber = ordersTable Then ordersOpNoField = f
        Next
        '        If ordersOpNoField Is Nothing Then Return dt

        Dim opCount As Integer = preactor.RecordCount(ordersOpNoField.FormatNumber)

        Dim idColumn As DataColumn = dt.Columns("OrdersID")
        dt.PrimaryKey = New DataColumn() {idColumn}

        For opRec As Integer = 1 To opCount
            Dim oNo As Integer
            Try
                oNo = preactor.ReadFieldInt(ordersOpNoField.FormatNumber, ordersOpNoField.FieldNumber, opRec)
            Catch
                Continue For
            End Try
            If oNo <> opNumber Then Continue For

            ' find corresponding CSV row(s) by OrdersID (order record assumed stored in OrdersID column)
            Dim rowtoupdate As DataRow = Nothing
            Try
                rowtoupdate = dt.Rows.Find(opRec)
            Catch
            End Try

            If rowtoupdate IsNot Nothing Then
                Dim ot = planningboard.GetOperationTimes(opRec)
                If ot.HasValue Then
                    rowtoupdate(startColName) = ot.Value.OperationTimes.ProcessStart
                    rowtoupdate(endColName) = ot.Value.OperationTimes.ProcessEnd
                    rowtoupdate(schColName) = True
                End If
            End If
        Next

        Return dt
    End Function

    ' ----------------------
    ' small helper to access format field pair(s)
    ' ----------------------
    Private Function getformatfieldpair(ByVal preactor As IPreactor, Optional ByVal field As String = "Field", Optional ByVal format As String = "Format") As Preactor.FormatFieldPair?
        Dim ffp As Preactor.FormatFieldPair = Nothing
        Select Case field.ToUpperInvariant()
            Case "DUE DATE", "PRIORITY", "EARLIEST START DATE"
                Return preactor.FindFirstClassificationString(field)
            Case Else
                If format = "ORDERS" Then
                    Return preactor.FindFirstClassificationString("LAUNCH TIME")
                End If
        End Select
        Return ffp
    End Function

    ' ----------------------
    ' Candidate DTO
    ' ----------------------
    Private Class Candidate
        Public Property OrdersID As Integer
        Public Property Earliest As DateTime
        Public Property Due As DateTime
        Public Property Tier As Integer
        Public Property WheelDia As String
        Public Property WheelPin As String
        Public Property TypeKey As String
        Public Property CycleRank As Integer
        Public Property MissingEarliest As Boolean
        Public Property MissingDue As Boolean
    End Class

End Class
