Option Strict On
Option Explicit On

Imports System
Imports System.Globalization
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Threading
Imports Microsoft.VisualBasic.FileIO
Imports Preactor
Imports Preactor.Interop.PreactorObject
Imports System.Windows.Forms

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

        'Dim routingDt As DataTable = ImportRoutingCsvToDataTable(filePath)
        Dim routingDt As DataTable = readOrderTable(preactor)

        Dim currentDate As New System.DateTime(2025, 8, 1, 0, 0, 0)


        ' Append schedule times from board for a few operation numbers (example)
        routingDt = AppendOperationTimesFromBoard(routingDt, preactor, planningboard, 100)
        routingDt = AppendOperationTimesFromBoard(routingDt, preactor, planningboard, 200)
        routingDt = AppendOperationTimesFromBoard(routingDt, preactor, planningboard, 290)
        routingDt = AppendOperationTimesFromBoard(routingDt, preactor, planningboard, 291)


        ' Build firing plan using firing optimizer (external class)
        Dim minOcc As Double = 0.8
        Dim maxOcc As Double = 1.0
        Dim firingObj As New firingOptimizer_vf
        Dim plan = firingObj.BuildBatchKilnPlan(routingDt, "D:\Documents\Opcenter\Cases\Grindwell Norton\Opcenter SC - Dev\Files\kilndata.csv", currentDate, minOcc, maxOcc, batchStartDelayMins:=60)

        ' 1) iterate firing queue (these are op 300 record numbers)
        For Each firingOpRec As Integer In plan.QueueFiringOpRecs

            ' 2) get batch metadata
            Dim batchNo As Integer = plan.BatchNoByFiringOpRec(firingOpRec)
            Dim batchStart As DateTime = plan.BatchStartByBatchNo(batchNo)
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

        ' Export plan
        firingObj.ExportPlanToCsv(plan, "D:\Documents\Opcenter\Cases\Grindwell Norton\Opcenter SC - Dev\Files\")

        Return 0
    End Function

    Public Function Run2(ByRef preactorComObject As PreactorObj, ByRef pespComObject As Object) As Integer
        Dim preactor As IPreactor = PreactorFactory.CreatePreactorObject(preactorComObject)
        Dim planningboard As IPlanningBoard = preactor.PlanningBoard
        Dim dt As DataTable
        Dim filePath As String = "D:\Documents\Opcenter\Cases\Grindwell Norton\Opcenter SC - Dev\Files\Templates\Routing.csv"
        Dim routingDt As DataTable = ImportRoutingCsvToDataTable(filePath)
        dt = readOrderTable(preactor)

        'Dim placeDate As New System.DateTime(2025, 8, 7, 0, 0, 0)
        'planningboard.PutOperationOnResource(10, 5, placeDate)



        'Try
        '    ' ... your normal rule setup here ...

        '    ShowExplorerUI()   ' <-- call this when you want the UI

        '    Return 0
        'Catch ex As Exception
        '    ' log ex.Message somewhere
        '    Return -1
        'End Try
        Return 0
    End Function

    Public Function Run3(ByRef preactorComObject As PreactorObj, ByRef pespComObject As Object) As Integer
        Dim preactor As IPreactor = PreactorFactory.CreatePreactorObject(preactorComObject)
        Dim planningboard As IPlanningBoard = preactor.PlanningBoard
        Dim opRec As Integer
        Dim routingdt As DataTable
        Dim ResRec As Integer
        Dim ResRecs As IEnumerable(Of Integer)
        Dim opTimes As Nullable(Of Preactor.OperationTimes)
        Dim ordersTable As Integer = preactor.FindFirstClassificationString("LAUNCH TIME").Value.FormatNumber
        Dim currentDate As New System.DateTime(2025, 8, 1, 0, 0, 0)
        routingdt = readOrderTable(preactor)
        Dim pressingQueue As List(Of Integer) = BuildPressing200Queue(routingdt, currentDate)
        CreateRankedOperationQueue(preactor, planningboard, ordersTable, "JobsQueue", pressingQueue)

        ' Snapshot for debugging
        Dim jobsQueueSnapshot As List(Of Integer) = GetQueueSnapshot(planningboard, "JobsQueue")

        While (planningboard.GetOperationInQueue("JobsQueue", 1, opRec))

            ' Take the next operation out of the ranked queue so we can decide where to load it.
            planningboard.RemoveOperationFromQueue("JobsQueue", opRec)

            ' Inner loop: schedule this operation and then walk to subsequent operations
            ' (your "family" / routing chain) using GetNextOperation.
            While (opRec > 0 And preactor.ReadFieldInt(114, 57, opRec) < 200)

                ' Find all valid alternate resources for this operation.
                ResRecs = planningboard.FindResources(opRec)

                ' Track the best (earliest) feasible candidate we find.
                Dim bestResRec As Integer = 0
                Dim bestOpTimes As Nullable(Of Preactor.OperationTimes) = Nothing

                ' Loop through *all* alternate resources and test feasibility on each.
                For Each ResRec In ResRecs

                    ' Test if the operation can be placed on this resource, and get the timing result.
                    ' TerminatorTime is the boundary between schedule history and schedule future;
                    ' using it here aligns with "schedule as soon as possible" in the future horizon. :contentReference[oaicite:3]{index=3}
                    opTimes = planningboard.TestOperationOnResource(opRec, ResRec, planningboard.TerminatorTime)

                    If opTimes.HasValue Then
                        ' This resource is feasible. Compare it to the current best candidate.
                        ' We want the earliest possible start time (ChangeStart).
                        If (Not bestOpTimes.HasValue) Then
                            ' First feasible candidate becomes the best by default.
                            bestResRec = ResRec
                            bestOpTimes = opTimes
                        Else
                            ' Replace best candidate if this one starts earlier.
                            If opTimes.Value.ChangeStart < bestOpTimes.Value.ChangeStart Then
                                bestResRec = ResRec
                                bestOpTimes = opTimes
                            End If
                        End If
                    End If

                Next ' evaluate next alternate resource

                ' After scanning all alternates:
                If bestOpTimes.HasValue AndAlso bestResRec > 0 Then
                    ' Load the operation onto the resource that gives the earliest feasible start.
                    planningboard.PutOperationOnResource(opRec, bestResRec, bestOpTimes.Value.ChangeStart)
                Else
                    ' No feasible resource was found.
                    ' Practical meaning:
                    '   - This operation cannot be scheduled on any alternate resource at/after the terminator boundary
                    '     under current constraints (calendars, setups, secondary constraints, etc.).
                    ' Leave it unscheduled (or handle with a custom queue / reason code if your design requires).
                End If

                ' Move to the next operation in the routing chain.
                opRec = planningboard.GetNextOperation(opRec, 1) ' API-supported routing traversal:contentReference[oaicite:4]{index=4}

            End While ' next operation in chain

        End While ' next op in JobsQueue
        Return 0
    End Function

    Public Function Runpress(ByRef preactorComObject As PreactorObj, ByRef pespComObject As Object) As Integer
        Dim preactor As IPreactor = PreactorFactory.CreatePreactorObject(preactorComObject)
        Dim planningboard As IPlanningBoard = preactor.PlanningBoard
        Dim opRec As Integer
        Dim routingdt As DataTable
        Dim ResRec As Integer
        Dim ResRecs As IEnumerable(Of Integer)
        Dim opTimes As Nullable(Of Preactor.OperationTimes)
        Dim ordersTable As Integer = preactor.FindFirstClassificationString("LAUNCH TIME").Value.FormatNumber
        Dim currentDate As New System.DateTime(2025, 8, 1, 0, 0, 0)
        routingdt = readOrderTable(preactor)
        Dim pressingQueue As List(Of Integer) = BuildPressing200Queue(routingdt, currentDate)
        CreateRankedOperationQueue(preactor, planningboard, ordersTable, "JobsQueue", pressingQueue)

        ' Snapshot for debugging
        Dim jobsQueueSnapshot As List(Of Integer) = GetQueueSnapshot(planningboard, "JobsQueue")

        While (planningboard.GetOperationInQueue("JobsQueue", 1, opRec))

            ' Take the next operation out of the ranked queue so we can decide where to load it.
            planningboard.RemoveOperationFromQueue("JobsQueue", opRec)

            ' Inner loop: schedule this operation and then walk to subsequent operations
            ' (your "family" / routing chain) using GetNextOperation.
            While (opRec > 0) ' this condition is wrong
                If preactor.ReadFieldInt(114, 57, opRec) = 200 Then
                    ' Find all valid alternate resources for this operation.
                    ResRecs = planningboard.FindResources(opRec)

                    ' Track the best (earliest) feasible candidate we find.
                    Dim bestResRec As Integer = 0
                    Dim bestOpTimes As Nullable(Of Preactor.OperationTimes) = Nothing

                    ' Loop through *all* alternate resources and test feasibility on each.
                    For Each ResRec In ResRecs

                        ' Test if the operation can be placed on this resource, and get the timing result.
                        ' TerminatorTime is the boundary between schedule history and schedule future;
                        ' using it here aligns with "schedule as soon as possible" in the future horizon. :contentReference[oaicite:3]{index=3}
                        opTimes = planningboard.TestOperationOnResource(opRec, ResRec, planningboard.TerminatorTime)

                        If opTimes.HasValue Then
                            ' This resource is feasible. Compare it to the current best candidate.
                            ' We want the earliest possible start time (ChangeStart).
                            If (Not bestOpTimes.HasValue) Then
                                ' First feasible candidate becomes the best by default.
                                bestResRec = ResRec
                                bestOpTimes = opTimes
                            Else
                                ' Replace best candidate if this one starts earlier.
                                If opTimes.Value.ChangeStart < bestOpTimes.Value.ChangeStart Then
                                    bestResRec = ResRec
                                    bestOpTimes = opTimes
                                End If
                            End If
                        End If

                    Next ' evaluate next alternate resource

                    ' After scanning all alternates:
                    If bestOpTimes.HasValue AndAlso bestResRec > 0 Then
                        ' Load the operation onto the resource that gives the earliest feasible start.
                        planningboard.PutOperationOnResource(opRec, bestResRec, bestOpTimes.Value.ChangeStart)
                    Else
                        ' No feasible resource was found.
                        ' Practical meaning:
                        '   - This operation cannot be scheduled on any alternate resource at/after the terminator boundary
                        '     under current constraints (calendars, setups, secondary constraints, etc.).
                        ' Leave it unscheduled (or handle with a custom queue / reason code if your design requires).
                    End If
                End If
                ' Move to the next operation in the routing chain.
                opRec = planningboard.GetNextOperation(opRec, 1) ' API-supported routing traversal:contentReference[oaicite:4]{index=4}

            End While ' next operation in chain

        End While ' next op in JobsQueue

        Return 0
    End Function
    Public Function RunPresstoFiring(ByRef preactorComObject As PreactorObj, ByRef pespComObject As Object) As Integer
        Dim preactor As IPreactor = PreactorFactory.CreatePreactorObject(preactorComObject)
        Dim planningboard As IPlanningBoard = preactor.PlanningBoard

        Dim ordersTable As Integer = preactor.FindFirstClassificationString("LAUNCH TIME").Value.FormatNumber

        ' Example: import a CSV, build pressing queue, create ranked queue and schedule
        'Dim filePath As String = "D:\Documents\Opcenter\Cases\Grindwell Norton\Opcenter SC - Dev\Files\Templates\Routing.csv"
        'Dim routingDt As DataTable = ImportRoutingCsvToDataTable(filePath)
        Dim currentDate As New System.DateTime(2025, 8, 1, 0, 0, 0)
        Dim ResRecs As IEnumerable(Of Integer)
        Dim opTimes As Nullable(Of Preactor.OperationTimes)
        Dim routingdt As DataTable = readOrderTable(preactor)
        Dim pressingQueue As List(Of Integer) = BuildPressing200Queue(routingDt, currentDate)
        CreateRankedOperationQueue(preactor, planningboard, ordersTable, "JobsQueue", pressingQueue)

        ' Snapshot for debugging
        Dim jobsQueueSnapshot As List(Of Integer) = GetQueueSnapshot(planningboard, "JobsQueue")

        ' Simple scheduling loop: pop queue and place operation on earliest feasible resource
        Dim pos As Integer = 1
        Dim opRec As Integer = 1

        '------------------------------------------------------------
        ' Select the resource that gives the earliest feasible
        ' ChangeStart time for the operation.
        '
        ' Key idea:
        '   - Test the operation on ALL valid alternate resources
        '   - Choose the candidate with the minimum ChangeStart
        '   - Then load the operation on that chosen resource
        '------------------------------------------------------------

        '        While (planningboard.GetOperationInQueue("JobsQueue", 1, opRec))

        ' Take the next operation out of the ranked queue so we can decide where to load it.
        '        planningboard.RemoveOperationFromQueue("JobsQueue", opRec)

        ' Inner loop: schedule this operation and then walk to subsequent operations
        ' (your "family" / routing chain) using GetNextOperation.
        While (opRec > 0 And opRec < 7444)
            If preactor.ReadFieldInt(114, 57, opRec) > 200 And preactor.ReadFieldInt(114, 57, opRec) < 300 Then
                ' Find all valid alternate resources for this operation.
                ResRecs = planningboard.FindResources(opRec)

                ' Track the best (earliest) feasible candidate we find.
                Dim bestResRec As Integer = 0
                Dim bestOpTimes As Nullable(Of Preactor.OperationTimes) = Nothing

                ' Loop through *all* alternate resources and test feasibility on each.
                For Each ResRec In ResRecs

                    ' Test if the operation can be placed on this resource, and get the timing result.
                    ' TerminatorTime is the boundary between schedule history and schedule future;
                    ' using it here aligns with "schedule as soon as possible" in the future horizon. :contentReference[oaicite:3]{index=3}
                    opTimes = planningboard.TestOperationOnResource(opRec, ResRec, planningboard.TerminatorTime)

                    If opTimes.HasValue Then
                        ' This resource is feasible. Compare it to the current best candidate.
                        ' We want the earliest possible start time (ChangeStart).
                        If (Not bestOpTimes.HasValue) Then
                            ' First feasible candidate becomes the best by default.
                            bestResRec = ResRec
                            bestOpTimes = opTimes
                        Else
                            ' Replace best candidate if this one starts earlier.
                            If opTimes.Value.ChangeStart < bestOpTimes.Value.ChangeStart Then
                                bestResRec = ResRec
                                bestOpTimes = opTimes
                            End If
                        End If
                    End If

                Next ' evaluate next alternate resource

                ' After scanning all alternates:
                If bestOpTimes.HasValue AndAlso bestResRec > 0 Then
                    ' Load the operation onto the resource that gives the earliest feasible start.
                    planningboard.PutOperationOnResource(opRec, bestResRec, bestOpTimes.Value.ChangeStart)
                Else
                    ' No feasible resource was found.
                    ' Practical meaning:
                    '   - This operation cannot be scheduled on any alternate resource at/after the terminator boundary
                    '     under current constraints (calendars, setups, secondary constraints, etc.).
                    ' Leave it unscheduled (or handle with a custom queue / reason code if your design requires).
                End If
            End If
            ' Move to the next operation in the routing chain.
            'opRec = planningboard.GetNextOperation(opRec, 1)
            opRec += 1 ' API-supported routing traversal:contentReference[oaicite:4]{index=4}

        End While ' next operation in chain

        '      End While ' next op in JobsQueue
        Return 0
    End Function



    Private Sub ShowExplorerUI()
        ' If Run() is not on an STA thread, create one for WinForms
        Dim t As New Thread(Sub()
                                Application.EnableVisualStyles()
                                Application.SetCompatibleTextRenderingDefault(False)

                                Using frm As New DataViewExplorerForm()
                                    frm.StartPosition = FormStartPosition.CenterScreen
                                    frm.ShowDialog() ' blocks until closed
                                End Using
                            End Sub)

        t.SetApartmentState(ApartmentState.STA)
        t.IsBackground = True
        t.Start()
        t.Join() ' Wait until user closes form (remove Join() if you do NOT want to block)
    End Sub


    Private Function CreateRankedOperationQueue(ByRef preactor As IPreactor, ByVal planningboard As IPlanningBoard,
                                                ByVal ordersTable As Integer, ByVal QName As String, ByVal queue As List(Of Integer)) As Integer
        planningboard.CreateQueue(QName)
        For Each q In queue
            planningboard.AddOperationToQueue(QName, q, QueuePosition.End)
        Next
        Return 0
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

        Dim ordersTable As Integer = getformatfieldpair(preactor, format:="ORDERS").FormatNumber
        Dim opNoFields = preactor.FindClassificationString("OP NO")
        Dim ordersOpNoField As Preactor.FormatFieldPair = Nothing
        For Each f In opNoFields
            If f.FormatNumber = ordersTable Then ordersOpNoField = f
        Next
        'If ordersOpNoField Is Nothing Then Return dt

        Dim opCount As Integer = preactor.RecordCount(ordersTable)

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

    'Old pressing optimizer
    ' ----------------------
    ' Candidate DTO
    ' ----------------------
    'Private Class Candidate
    '    Public Property OrdersID As Integer
    '    Public Property Earliest As DateTime
    '    Public Property Due As DateTime
    '    Public Property Tier As Integer
    '    Public Property WheelDia As String
    '    Public Property WheelPin As String
    '    Public Property TypeKey As String
    '    Public Property CycleRank As Integer
    '    Public Property MissingEarliest As Boolean
    '    Public Property MissingDue As Boolean
    'End Class


    '' ----------------------
    '' Pressing 200 queue builder
    '' ----------------------
    'Public Function BuildPressing200Queue(dt As DataTable, currentDate As DateTime, Optional approachingDays As Integer = 2) As List(Of Integer)
    '    If dt Is Nothing Then Throw New ArgumentNullException(NameOf(dt))

    '    SharedHelpers.RequireColumn(dt, "OrdersID")
    '    SharedHelpers.RequireColumn(dt, "Operation Number")
    '    SharedHelpers.RequireColumn(dt, "Pressing earliest start")
    '    SharedHelpers.RequireColumn(dt, "pressing due date")
    '    SharedHelpers.RequireColumn(dt, "Wheel Dia")
    '    SharedHelpers.RequireColumn(dt, "Wheel thickness")
    '    SharedHelpers.RequireColumn(dt, "Cycle Type")

    '    Dim today As DateTime = currentDate.Date
    '    Dim approachCutoff As DateTime = today.AddDays(approachingDays)

    '    Dim candidates As New List(Of Candidate)()

    '    For Each r As DataRow In dt.Rows
    '        Dim opNo = SharedHelpers.SafeInt(r("Operation Number"))
    '        If opNo <> PRESS_OP_NUMBER Then Continue For

    '        Dim orderId = SharedHelpers.SafeInt(r("OrdersID"))
    '        If orderId <= 0 Then Continue For

    '        Dim earliest = SharedHelpers.SafeDate(r("Pressing earliest start")).Date
    '        Dim due = SharedHelpers.SafeDate(r("Pressing Due date")).Date

    '        Dim missingEarliest = (earliest = DateTime.MinValue)
    '        Dim missingDue = (due = DateTime.MinValue)

    '        Dim tier As Integer
    '        If Not missingEarliest AndAlso Not missingDue AndAlso earliest <= approachCutoff AndAlso due >= today Then
    '            tier = 0
    '        ElseIf Not missingDue AndAlso due < today Then
    '            tier = 1
    '        Else
    '            tier = 2
    '        End If

    '        Dim wheelDia = SharedHelpers.SafeStr(r("Wheel Dia")).Trim()
    '        Dim wheelPin = SharedHelpers.SafeStr(r("Wheel thickness")).Trim()
    '        Dim cycleType = SharedHelpers.SafeStr(r("Cycle Type")).Trim()
    '        Dim cycleRank = GetCycleRank(cycleType)

    '        candidates.Add(New Candidate With {
    '            .OrdersID = orderId,
    '            .Earliest = earliest,
    '            .Due = due,
    '            .Tier = tier,
    '            .WheelDia = wheelDia,
    '            .WheelPin = wheelPin,
    '            .TypeKey = wheelDia & "|" & wheelPin,
    '            .CycleRank = cycleRank,
    '            .MissingEarliest = missingEarliest,
    '            .MissingDue = missingDue
    '        })
    '    Next

    '    Dim sorted = candidates.OrderBy(Function(c) c.Tier) _
    '                               .ThenBy(Function(c) If(c.MissingDue, DateTime.MaxValue, c.Due)) _
    '                               .ThenBy(Function(c) If(c.MissingEarliest, DateTime.MaxValue, c.Earliest)) _
    '                               .ThenByDescending(Function(c) c.CycleRank) _
    '                               .ThenBy(Function(c) c.TypeKey) _
    '                               .ThenBy(Function(c) c.OrdersID) _
    '                               .ToList()

    '    Dim batched = GreedyTypeBatchingWithinTier(sorted, lookahead:=50)
    '    Return batched.Select(Function(c) c.OrdersID).Distinct().ToList()
    'End Function

    'Private Function GreedyTypeBatchingWithinTier(sorted As List(Of Candidate), Optional lookahead As Integer = 50) As List(Of Candidate)
    '    If sorted.Count <= 2 Then Return sorted

    '    Dim work As New List(Of Candidate)(sorted)
    '    Dim result As New List(Of Candidate)(work.Count)

    '    Dim i As Integer = 0
    '    While i < work.Count
    '        Dim cur = work(i)
    '        result.Add(cur)

    '        If Not String.IsNullOrEmpty(cur.TypeKey) Then
    '            Dim pulled As Integer = 0
    '            Dim j As Integer = i + 1
    '            While j < work.Count AndAlso pulled < lookahead
    '                If work(j).Tier = cur.Tier AndAlso work(j).TypeKey = cur.TypeKey AndAlso work(j).Due = cur.Due Then
    '                    result.Add(work(j))
    '                    work.RemoveAt(j)
    '                    pulled += 1
    '                    Continue While
    '                End If
    '                j += 1
    '            End While
    '        End If

    '        i += 1
    '    End While

    '    Return result
    'End Function

    'Private Function GetCycleRank(cycleType As String) As Integer
    '    If String.IsNullOrWhiteSpace(cycleType) Then Return 0
    '    Select Case cycleType.Trim().ToUpperInvariant()
    '        Case "150VT" : Return 3
    '        Case "102VT" : Return 2
    '        Case "65VT" : Return 1
    '        Case Else : Return 0
    '    End Select
    'End Function

    ' New pressing optimizer
    ' ----------------------
    ' Pressing 200 queue builder (STRICT UNSCHEDULED) returning PARENT RECORD
    ' Minimal changes from your original:
    '   1) Adds is_scheduled + parent_record columns
    '   2) Filters only op=PRESS_OP_NUMBER AND is_scheduled=False
    '   3) Uses ParentRecord as the queue item instead of OrdersID
    '   4) Cycle ranking remains via GetCycleRank(), but you can later replace it with a map/dictionary
    ' ----------------------

    ' Candidate structure used for sorting/batching
    Private Class Candidate
        Public Property ParentRecord As Integer          ' <-- NEW: queue key to return
        Public Property Earliest As DateTime
        Public Property Due As DateTime
        Public Property Tier As Integer
        Public Property WheelDia As String
        Public Property WheelPin As String
        Public Property TypeKey As String                ' WheelDia|WheelThickness
        Public Property CycleRank As Integer
        Public Property MissingEarliest As Boolean
        Public Property MissingDue As Boolean
    End Class

    ' ----------------------
    ' Pressing 200 queue builder
    ' ----------------------
    Public Function BuildPressing200Queue(dt As DataTable,
                                      currentDate As DateTime,
                                      Optional approachingDays As Integer = 2) As List(Of Integer)

        If dt Is Nothing Then Throw New ArgumentNullException(NameOf(dt))

        ' ---- Required columns (these must exist in dt) ----
        ' NOTE: Keep these names EXACTLY matching your DataTable column names.
        SharedHelpers.RequireColumn(dt, "parent_record")               ' <-- NEW
        SharedHelpers.RequireColumn(dt, "is_scheduled")                ' <-- NEW

        SharedHelpers.RequireColumn(dt, "Operation Number")
        SharedHelpers.RequireColumn(dt, "Pressing earliest start")
        SharedHelpers.RequireColumn(dt, "Pressing Due date")           ' (fixes earlier mismatch)
        SharedHelpers.RequireColumn(dt, "Wheel Dia")
        SharedHelpers.RequireColumn(dt, "Wheel thickness")
        SharedHelpers.RequireColumn(dt, "Cycle Type")

        ' ---- Date boundaries used for tiering ----
        Dim today As DateTime = currentDate.Date
        Dim approachCutoff As DateTime = today.AddDays(approachingDays)

        Dim candidates As New List(Of Candidate)()

        ' ---- Extract candidates from table rows ----
        For Each r As DataRow In dt.Rows

            ' 1) Filter: only pressing operation (typically 200)
            Dim opNo As Integer = SharedHelpers.SafeInt(r("Operation Number"))
            If opNo <> PRESS_OP_NUMBER Then Continue For

            ' 2) Filter: STRICT UNSCHEDULED only
            ' is_scheduled is expected to be reliably boolean-like based on your confirmation.
            Dim isScheduled As Boolean = SharedHelpers.SafeBool(r("is_scheduled"))
            If isScheduled Then Continue For

            ' 3) Queue key: parent_record must be valid
            Dim parentRec As Integer = SharedHelpers.SafeInt(r("parent_record"))
            If parentRec <= 0 Then Continue For

            ' 4) Read pressing dates (date-only logic)
            Dim earliest As DateTime = SharedHelpers.SafeDate(r("Pressing earliest start")).Date
            Dim due As DateTime = SharedHelpers.SafeDate(r("Pressing Due date")).Date

            ' Treat missing/parse-failed dates as MinValue (consistent with your existing SafeDate behavior)
            Dim missingEarliest As Boolean = (earliest = DateTime.MinValue)
            Dim missingDue As Boolean = (due = DateTime.MinValue)

            ' 5) Tiering logic:
            ' Tier 0: approaching soon and not late (Earliest <= cutoff AND Due >= today)
            ' Tier 1: already late (Due < today)
            ' Tier 2: everything else (includes missing values)
            Dim tier As Integer
            If Not missingEarliest AndAlso Not missingDue AndAlso earliest <= approachCutoff AndAlso due >= today Then
                tier = 0
            ElseIf Not missingDue AndAlso due < today Then
                tier = 1
            Else
                tier = 2
            End If

            ' 6) Read batching/type attributes
            Dim wheelDia As String = SharedHelpers.SafeStr(r("Wheel Dia")).Trim()
            Dim wheelPin As String = SharedHelpers.SafeStr(r("Wheel thickness")).Trim()
            Dim cycleType As String = SharedHelpers.SafeStr(r("Cycle Type")).Trim()

            ' CycleRank currently uses GetCycleRank() (hardcoded list).
            ' Later you can replace this call with a dictionary/map lookup.
            Dim cycleRank As Integer = GetCycleRank(cycleType)

            candidates.Add(New Candidate With {
            .ParentRecord = parentRec,
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

        ' ---- Sorting: the exact queue priority order ----
        ' 1) Tier asc (0,1,2)
        ' 2) Due asc (missing due goes last)
        ' 3) Earliest asc (missing earliest goes last)
        ' 4) CycleRank desc (higher rank first)
        ' 5) TypeKey asc (clusters same Dia|Thickness)
        ' 6) ParentRecord asc (stable tie-breaker)
        Dim sorted As List(Of Candidate) =
        candidates.OrderBy(Function(c) c.Tier) _
                  .ThenBy(Function(c) If(c.MissingDue, DateTime.MaxValue, c.Due)) _
                  .ThenBy(Function(c) If(c.MissingEarliest, DateTime.MaxValue, c.Earliest)) _
                  .ThenByDescending(Function(c) c.CycleRank) _
                  .ThenBy(Function(c) c.TypeKey) _
                  .ThenBy(Function(c) c.ParentRecord) _
                  .ToList()

        ' ---- Soft batching: greedy pull-forward of same Tier + TypeKey + Due ----
        Dim batched As List(Of Candidate) = GreedyTypeBatchingWithinTier(sorted, lookahead:=50)

        ' ---- Output: return parent_record list (distinct to be safe) ----
        Return batched.Select(Function(c) c.ParentRecord).Distinct().ToList()
    End Function

    ' ----------------------
    ' Greedy batching (unchanged logic, but now works on ParentRecord candidates)
    ' Pull-forward rules:
    '   - same Tier
    '   - same TypeKey (WheelDia|WheelThickness)
    '   - same Due date
    ' This is a "soft clustering" step; it does NOT change tier ordering.
    ' ----------------------
    Private Function GreedyTypeBatchingWithinTier(sorted As List(Of Candidate),
                                             Optional lookahead As Integer = 50) As List(Of Candidate)

        If sorted Is Nothing OrElse sorted.Count <= 2 Then Return If(sorted, New List(Of Candidate)())

        Dim work As New List(Of Candidate)(sorted)
        Dim result As New List(Of Candidate)(work.Count)

        Dim i As Integer = 0
        While i < work.Count
            Dim cur As Candidate = work(i)
            result.Add(cur)

            ' Only attempt pull-forward if TypeKey is meaningful
            If Not String.IsNullOrEmpty(cur.TypeKey) Then
                Dim pulled As Integer = 0
                Dim j As Integer = i + 1

                ' Scan forward but stop after lookahead pulls
                While j < work.Count AndAlso pulled < lookahead

                    ' Match only within same tier, same type, same due date
                    If work(j).Tier = cur.Tier AndAlso
                   work(j).TypeKey = cur.TypeKey AndAlso
                   work(j).Due = cur.Due Then

                        result.Add(work(j))
                        work.RemoveAt(j)     ' remove from work list so it doesn’t appear again later
                        pulled += 1
                        Continue While       ' keep scanning from same j index (since list shifted)
                    End If

                    j += 1
                End While
            End If

            i += 1
        End While

        Return result
    End Function

    ' ----------------------
    ' Cycle ranking (PRESSING ONLY)
    ' Current version: hardcoded fallback list.
    ' Next step (per your requirement): replace this with a rank map (e.g., Dictionary loaded from CSV).
    ' ----------------------
    Private Function GetCycleRank(cycleType As String) As Integer
        If String.IsNullOrWhiteSpace(cycleType) Then Return 0

        Select Case cycleType.Trim().ToUpperInvariant()
            Case "150VT" : Return 3
            Case "102VT" : Return 2
            Case "65VT" : Return 1
            Case Else : Return 0
        End Select
    End Function


End Class
