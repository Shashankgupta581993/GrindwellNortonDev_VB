Option Strict On
Option Explicit On

Imports System
Imports System.Globalization
Imports System.IO
Imports System.Text
Imports Preactor

Public Module SharedHelpers

    Public Class WipInfo
        Public Property ParentRecord As Integer
        Public Property CurrentOpRec As Integer
        Public Property CurrentOpNo As Integer

        Public Property CurrentOpScheduled As Boolean
        Public Property CurrentOpStarted As Boolean

        Public Property PrevOpRec As Integer
        Public Property PrevOpNo As Integer
        Public Property PrevOpScheduled As Boolean
        Public Property PrevOpEndTime As DateTime

        Public Property HasAnyPriorScheduled As Boolean
        Public Property LastPriorScheduledOpNo As Integer
        Public Property HasFutureScheduledOp As Boolean

        Public Property ReadyTime As DateTime
        Public Property WipScore As Integer

        Public Property CandidateStatus As String
        Public Property RejectReason As String

        Public Property CurrentOpCompleted As Boolean
        Public Property CurrentOpActualized As Boolean

        Public Property PrevOpReleased As Boolean
        Public Property PrevOpReleaseTime As DateTime

        Public Property CurrentOpReleased As Boolean
        Public Property CurrentOpReleaseTime As DateTime

        Public Property HasAnyPriorReleased As Boolean
        Public Property LastPriorReleasedOpNo As Integer
        Public Property LastPriorReleasedOpRec As Integer

        Public Property ExecutionStatus As String
        Public Property StatusConflict As Boolean
        Public Property StatusReason As String
    End Class
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
    Public Function TryGetFieldNumber(preactor As IPreactor,
                                  formatNo As Integer,
                                  fieldName As String) As Integer
        If preactor Is Nothing Then Return 0
        If formatNo <= 0 Then Return 0
        If String.IsNullOrWhiteSpace(fieldName) Then Return 0

        Try
            Return preactor.GetFieldNumber(formatNo, fieldName)
        Catch
            Return 0
        End Try
    End Function

    Public Function ResolveFirstExistingField(preactor As IPreactor,
                                          formatNo As Integer,
                                          fieldNames As String()) As Integer
        If fieldNames Is Nothing Then Return 0

        For Each fieldName As String In fieldNames
            Dim fieldNo As Integer = TryGetFieldNumber(preactor, formatNo, fieldName)
            If fieldNo > 0 Then Return fieldNo
        Next

        Return 0
    End Function

    Public Function ReadBoolField(preactor As IPreactor,
                              formatNo As Integer,
                              fieldNo As Integer,
                              recNo As Integer) As Boolean
        If fieldNo <= 0 Then Return False

        Try
            Return preactor.ReadFieldBool(formatNo, fieldNo, recNo)
        Catch
        End Try

        Try
            Return preactor.ReadFieldInt(formatNo, fieldNo, recNo) <> 0
        Catch
        End Try

        Try
            Return SafeBool(preactor.ReadFieldString(formatNo, fieldNo, recNo))
        Catch
        End Try

        Return False
    End Function

    Public Function ReadDateField(preactor As IPreactor,
                              formatNo As Integer,
                              fieldNo As Integer,
                              recNo As Integer) As DateTime
        If fieldNo <= 0 Then Return DateTime.MinValue

        Try
            Return preactor.ReadFieldDateTime(formatNo, fieldNo, recNo)
        Catch
            Return DateTime.MinValue
        End Try
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
        'Dim cols As String() = {
        '    "OrdersID", "Order No", "Part Number", "Part Name", "Operation Number", "Operation Name",
        '    "Resource Group", "Required Resource", "Setup Time", "Time Per Item", "Sales Order", "Quantity",
        '    "Due Date", "Batch Time", "Process Time Type", "Tonnage", "Cycle Type", "Volume Occupancy",
        '    "Kiln Type", "Firing buffer", "MTS/MTO", "MTS/MTO priority", "Que Time", "Pressing buffer",
        '    "Wheel Dia", "Wheel thickness", "Week start", "Pressing earliest start", "Pressing Due date",
        '    "Constaint Usage", "Constraint Qty", "firing earliest start date", "firing due date", "scheduled_start_time", "scheduled_end_time", "is_scheduled", "parent_record", "prev_op_is_scheduled"
        '}
        Dim cols As String() = {
    "OrdersID", "Order No", "Part Number", "Part Name", "Operation Number", "Operation Name",
    "Resource Group", "Required Resource", "Setup Time", "Time Per Item", "Sales Order", "Quantity",
    "Due Date", "Batch Time", "Process Time Type", "Tonnage", "Cycle Type", "Volume Occupancy",
    "Kiln Type", "Firing buffer", "MTS/MTO", "MTS/MTO priority", "Que Time", "Pressing buffer",
    "Wheel Dia", "Wheel thickness", "Week start", "Pressing earliest start", "Pressing Due date",
    "Constaint Usage", "Constraint Qty", "firing earliest start date", "firing due date",
    "scheduled_start_time", "scheduled_end_time", "is_scheduled", "parent_record", "prev_op_is_scheduled",
    "source_is_completed",
    "opcenter_use_actual",
    "actual_start_time",
    "actual_end_time",
    "order_last_completed_op_no",
    "order_last_completed_op_rec",
    "order_last_completed_release_time",
    "operation_effective_completed",
    "operation_execution_status",
    "operation_releases_next",
    "operation_release_time",
    "operation_status_conflict",
    "operation_status_reason",
    "wip_prev_op_released",
    "wip_prev_op_release_time",
    "wip_any_prior_released",
    "wip_last_prior_released_op_no",
    "wip_last_prior_released_op_rec",
    "wip_current_op_scheduled",
    "wip_current_op_started",
    "wip_prev_op_rec",
    "wip_prev_op_no",
    "wip_prev_op_scheduled",
    "wip_prev_op_end_time",
    "wip_any_prior_scheduled",
    "wip_last_prior_scheduled_op_no",
    "wip_has_future_scheduled_op",
    "wip_ready_time",
    "wip_score",
    "wip_status",
    "wip_reject_reason"
        }
        For Each c In cols
            dt.Columns.Add(New DataColumn(c, GetType(Object)))
        Next
        Return dt
    End Function

    Public Function readOrderTable(ByVal preactor As IPreactor) As DataTable

        Dim planningboard As IPlanningBoard = preactor.PlanningBoard
        Dim dt As DataTable = BuildRoutingSchema()

        ' 1. Suspend indexing, events, and constraints for bulk insert performance
        dt.BeginLoadData()

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
        ' Source completion flag imported into Opcenter.
        ' Per current design: Toggle Attribute 1 = is_completed.
        Dim sourceCompletedField As Integer =
    TryGetFieldNumber(preactor, ordersTable, "Toggle Attribute 1")

        ' Opcenter actualization flag.
        Dim useActualField As Integer =
    ResolveFirstExistingField(preactor,
                              ordersTable,
                              New String() {
                                  "Use Actual",
                                  "Use Actual Times",
                                  "USE ACTUAL TIMES",
                                  "Use actual"
                              })

        ' Actual time fields. If your actual end has a different name,
        ' add it to this list before "End Time".
        Dim actualStartField As Integer =
    ResolveFirstExistingField(preactor,
                              ordersTable,
                              New String() {
                                  "Actual Start Time",
                                  "Actual Start",
                                  "Actual Start Date",
                                  "Start Time Actual"
                              })

        Dim actualEndField As Integer =
    ResolveFirstExistingField(preactor,
                              ordersTable,
                              New String() {
                                  "Actual End Time",
                                  "Actual End",
                                  "Actual Finish Time",
                                  "Actual End Date",
                                  "End Time Actual"
                              })
        'Dim parentRecord = preactor.GetFieldNumber(ordersTable, "Belongs to Order No.")
        Dim rowCount = preactor.RecordCount(ordersTable)
        Dim parentRecordByOrderNo As New Dictionary(Of String, Integer)(
            StringComparer.OrdinalIgnoreCase)

        For rec As Integer = 1 To rowCount
            Dim r As DataRow = dt.NewRow()

            ' 2. Cache values used multiple times to avoid redundant API reads
            Dim currentOpNo As Integer = preactor.ReadFieldInt(ordersTable, opNo, rec)
            Dim isScheduled As Boolean = planningboard.IsOperationScheduled(rec)
            Dim currentOrderNo As String =
                preactor.ReadFieldString(ordersTable, orderNo, rec).Trim()

            r("OrdersID") = rec
            r("Order No") = currentOrderNo
            'r("Part Number") = preactor.ReadFieldString(ordersTable, partNo, rec)
            'r("Part Name") = preactor.ReadFieldString(ordersTable, product, rec)
            r("Operation Number") = preactor.ReadFieldInt(ordersTable, opNo, rec)
            r("Operation Name") = preactor.ReadFieldString(ordersTable, opName, rec)
            r("Resource Group") = preactor.ReadFieldString(ordersTable, resGroup, rec)
            r("Required Resource") = preactor.ReadFieldString(ordersTable, res, rec)
            'r("Setup Time") = preactor.ReadFieldDouble(ordersTable, stpTime, rec) * 1440
            'r("Time Per Item")
            'r("Sales Order")
            r("Quantity") = preactor.ReadFieldInt(ordersTable, Qty, rec)
            r("Due Date") = preactor.ReadFieldDateTime(ordersTable, dueDate, rec)
            r("Batch Time") = preactor.ReadFieldDouble(ordersTable, batchTime, rec) * 1440
            'r("Process Time Type") = preactor.ReadFieldString(ordersTable, prsTimeType, rec)
            r("Tonnage") = preactor.ReadFieldDouble(ordersTable, tonnage, rec)
            r("Cycle Type") = preactor.ReadFieldString(ordersTable, cycleType, rec)
            r("Volume Occupancy") = preactor.ReadFieldDouble(ordersTable, volumeOcc, rec)
            r("Kiln Type") = preactor.ReadFieldInt(ordersTable, klnType, rec)
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

            ' 3. Streamlined scheduling assignment
            r("is_scheduled") = isScheduled
            If isScheduled Then
                r("scheduled_start_time") = preactor.ReadFieldDateTime(ordersTable, schStart, rec)
                r("scheduled_end_time") = preactor.ReadFieldDateTime(ordersTable, schEnd, rec)
            End If
            Dim sourceCompleted As Boolean =
            ReadBoolField(preactor, ordersTable, sourceCompletedField, rec)

            Dim useActual As Boolean =
            ReadBoolField(preactor, ordersTable, useActualField, rec)

            Dim actualStartValue As DateTime =
            ReadDateField(preactor, ordersTable, actualStartField, rec)

            Dim actualEndValue As DateTime =
            ReadDateField(preactor, ordersTable, actualEndField, rec)

            ' Fallback:
            ' If Opcenter uses Start Time / End Time as actual time when Use Actual is checked,
            ' read End Time as actual end only when Use Actual is true.
            If actualEndValue = DateTime.MinValue AndAlso useActual AndAlso schEnd > 0 Then
                actualEndValue = ReadDateField(preactor, ordersTable, schEnd, rec)
            End If

            If actualStartValue = DateTime.MinValue AndAlso useActual AndAlso schStart > 0 Then
                actualStartValue = ReadDateField(preactor, ordersTable, schStart, rec)
            End If

            r("source_is_completed") = sourceCompleted
            r("opcenter_use_actual") = useActual

            If actualStartValue <> DateTime.MinValue Then
                r("actual_start_time") = actualStartValue
            End If

            If actualEndValue <> DateTime.MinValue Then
                r("actual_end_time") = actualEndValue
            End If
            ' 4. Cache the first operation record for each order. This replaces
            ' one FindMatchingRecord COM call per operation.
            Dim parentRecord As Integer

            If currentOrderNo.Length = 0 Then
                parentRecord = rec
            ElseIf Not parentRecordByOrderNo.TryGetValue(currentOrderNo, parentRecord) Then
                parentRecord = rec
                parentRecordByOrderNo.Add(currentOrderNo, parentRecord)
            End If

            r("parent_record") = parentRecord

            Dim prevOpRec As Integer = 0

            Try
                prevOpRec = planningboard.GetPreviousOperation(rec, 1)

                If prevOpRec > 0 Then
                    r("prev_op_is_scheduled") = planningboard.IsOperationScheduled(prevOpRec)
                End If

            Catch ex As Exception
                ' Keep default False if previous-operation lookup fails.
                ' Optional: add logging later.
                r("prev_op_is_scheduled") = False
            End Try

            dt.Rows.Add(r)
        Next
        dt.EndLoadData()
        PopulateWipColumns(dt, planningboard, planningboard.TerminatorTime)
        Return dt
    End Function
    Public Function GetWipInfo(dt As DataTable,
                           planningboard As IPlanningBoard,
                           targetRow As DataRow,
                           terminatorTime As DateTime,
                           readyBufferMinutes As Integer,
                           requirePrevScheduled As Boolean) As WipInfo

        If dt Is Nothing Then Throw New ArgumentNullException(NameOf(dt))
        If targetRow Is Nothing Then Throw New ArgumentNullException(NameOf(targetRow))

        Dim parentRecord As Integer = GetEffectiveParentRecord(targetRow)
        Dim orderRows As List(Of DataRow) =
        dt.AsEnumerable().
            Where(Function(x) GetEffectiveParentRecord(x) = parentRecord).
            OrderBy(Function(x) SafeInt(x("Operation Number"))).
            ThenBy(Function(x) SafeInt(x("OrdersID"))).
            ToList()

        Dim currentOpNo As Integer = SafeInt(targetRow("Operation Number"))
        Dim prevRow As DataRow = Nothing
        Dim hasAnyPriorScheduled As Boolean = False
        Dim lastPriorScheduledOpNo As Integer = 0
        Dim hasFutureScheduledOp As Boolean = False

        For Each row As DataRow In orderRows
            Dim opNo As Integer = SafeInt(row("Operation Number"))

            If opNo < currentOpNo Then
                prevRow = row

                If SafeBool(row("is_scheduled")) AndAlso
                   SafeDate(row("scheduled_end_time")) <> DateTime.MinValue Then

                    hasAnyPriorScheduled = True
                    lastPriorScheduledOpNo = opNo
                End If
            ElseIf opNo > currentOpNo AndAlso SafeBool(row("is_scheduled")) Then
                hasFutureScheduledOp = True
            End If
        Next

        Return CreateWipInfo(targetRow,
                             prevRow,
                             hasAnyPriorScheduled,
                             lastPriorScheduledOpNo,
                             hasFutureScheduledOp,
                             terminatorTime,
                             readyBufferMinutes,
                             requirePrevScheduled)

    End Function

    Private Function GetEffectiveParentRecord(row As DataRow) As Integer

        Dim parentRecord As Integer = SafeInt(row("parent_record"))
        If parentRecord > 0 Then Return parentRecord

        Return SafeInt(row("OrdersID"))

    End Function

    'Private Function CreateWipInfo(targetRow As DataRow,
    '                               prevRow As DataRow,
    '                               hasAnyPriorScheduled As Boolean,
    '                               lastPriorScheduledOpNo As Integer,
    '                               hasFutureScheduledOp As Boolean,
    '                               terminatorTime As DateTime,
    '                               readyBufferMinutes As Integer,
    '                               requirePrevScheduled As Boolean) As WipInfo

    '    Dim result As New WipInfo With {
    '        .CurrentOpRec = SafeInt(targetRow("OrdersID")),
    '        .CurrentOpNo = SafeInt(targetRow("Operation Number")),
    '        .ParentRecord = GetEffectiveParentRecord(targetRow),
    '        .CurrentOpScheduled = SafeBool(targetRow("is_scheduled")),
    '        .CurrentOpStarted = False,
    '        .HasAnyPriorScheduled = hasAnyPriorScheduled,
    '        .LastPriorScheduledOpNo = lastPriorScheduledOpNo,
    '        .HasFutureScheduledOp = hasFutureScheduledOp
    '    }

    '    Dim startT As DateTime = SafeDate(targetRow("scheduled_start_time"))
    '    If startT <> DateTime.MinValue AndAlso startT <= terminatorTime Then
    '        result.CurrentOpStarted = True
    '    End If

    '    If prevRow IsNot Nothing Then
    '        result.PrevOpRec = SafeInt(prevRow("OrdersID"))
    '        result.PrevOpNo = SafeInt(prevRow("Operation Number"))
    '        result.PrevOpScheduled = SafeBool(prevRow("is_scheduled"))

    '        If result.PrevOpScheduled Then
    '            result.PrevOpEndTime = SafeDate(prevRow("scheduled_end_time"))
    '        End If
    '    End If

    '    If result.PrevOpScheduled AndAlso result.PrevOpEndTime <> DateTime.MinValue Then
    '        result.ReadyTime = result.PrevOpEndTime.AddMinutes(readyBufferMinutes)
    '    Else
    '        result.ReadyTime = DateTime.MinValue
    '    End If

    '    If result.LastPriorScheduledOpNo > 0 Then
    '        result.WipScore = 1000 + result.LastPriorScheduledOpNo
    '    Else
    '        result.WipScore = 0
    '    End If

    '    result.CandidateStatus = "Candidate"
    '    result.RejectReason = ""

    '    If result.CurrentOpScheduled Then
    '        result.CandidateStatus = "Rejected"
    '        result.RejectReason = "Current operation already scheduled"
    '    ElseIf result.CurrentOpStarted Then
    '        result.CandidateStatus = "Rejected"
    '        result.RejectReason = "Current operation already started or historical"
    '    ElseIf result.HasFutureScheduledOp Then
    '        result.CandidateStatus = "Rejected"
    '        result.RejectReason = "Future operation already scheduled"
    '    ElseIf requirePrevScheduled AndAlso result.PrevOpRec <= 0 Then
    '        result.CandidateStatus = "Rejected"
    '        result.RejectReason = "No previous operation found"
    '    ElseIf requirePrevScheduled AndAlso Not result.PrevOpScheduled Then
    '        result.CandidateStatus = "Rejected"
    '        result.RejectReason = "Previous operation not scheduled"
    '    ElseIf requirePrevScheduled AndAlso result.PrevOpEndTime = DateTime.MinValue Then
    '        result.CandidateStatus = "Rejected"
    '        result.RejectReason = "Previous operation has no valid end time"
    '    End If

    '    Return result

    'End Function
    Private Function CreateWipInfo(targetRow As DataRow,
                               prevRow As DataRow,
                               hasAnyPriorScheduled As Boolean,
                               lastPriorScheduledOpNo As Integer,
                               hasFutureScheduledOp As Boolean,
                               terminatorTime As DateTime,
                               readyBufferMinutes As Integer,
                               requirePrevScheduled As Boolean,
                               Optional lastPriorReleasedOpRec As Integer = 0) As WipInfo

        Dim result As New WipInfo With {
        .CurrentOpRec = SafeInt(targetRow("OrdersID")),
        .CurrentOpNo = SafeInt(targetRow("Operation Number")),
        .ParentRecord = GetEffectiveParentRecord(targetRow),
        .CurrentOpScheduled = SafeBool(targetRow("is_scheduled")),
        .CurrentOpStarted = False,
        .CurrentOpCompleted = SafeBool(targetRow("operation_effective_completed")),
        .CurrentOpActualized = SafeBool(targetRow("opcenter_use_actual")) AndAlso
                              SafeDate(targetRow("actual_end_time")) <> DateTime.MinValue,
        .CurrentOpReleased = SafeBool(targetRow("operation_releases_next")),
        .CurrentOpReleaseTime = SafeDate(targetRow("operation_release_time")),
        .HasAnyPriorReleased = hasAnyPriorScheduled,
        .LastPriorReleasedOpNo = lastPriorScheduledOpNo,
        .LastPriorReleasedOpRec = lastPriorReleasedOpRec,
        .HasAnyPriorScheduled = hasAnyPriorScheduled,
        .LastPriorScheduledOpNo = lastPriorScheduledOpNo,
        .HasFutureScheduledOp = hasFutureScheduledOp,
        .ExecutionStatus = SafeStr(targetRow("operation_execution_status")),
        .StatusConflict = SafeBool(targetRow("operation_status_conflict")),
        .StatusReason = SafeStr(targetRow("operation_status_reason"))
    }

        Dim startT As DateTime = SafeDate(targetRow("scheduled_start_time"))

        If startT <> DateTime.MinValue AndAlso startT <= terminatorTime Then
            result.CurrentOpStarted = True
        End If

        If prevRow IsNot Nothing Then

            result.PrevOpRec = SafeInt(prevRow("OrdersID"))
            result.PrevOpNo = SafeInt(prevRow("Operation Number"))
            result.PrevOpScheduled = SafeBool(prevRow("is_scheduled"))

            If result.PrevOpScheduled Then
                result.PrevOpEndTime = SafeDate(prevRow("scheduled_end_time"))
            End If

            result.PrevOpReleased = SafeBool(prevRow("operation_releases_next"))
            result.PrevOpReleaseTime = SafeDate(prevRow("operation_release_time"))

        End If

        If result.PrevOpReleased AndAlso result.PrevOpReleaseTime <> DateTime.MinValue Then
            result.ReadyTime = result.PrevOpReleaseTime.AddMinutes(readyBufferMinutes)
        Else
            result.ReadyTime = DateTime.MinValue
        End If

        If result.LastPriorReleasedOpNo > 0 Then
            result.WipScore = 1000 + result.LastPriorReleasedOpNo
        Else
            result.WipScore = 0
        End If

        result.CandidateStatus = "Candidate"
        result.RejectReason = ""

        If result.CurrentOpCompleted OrElse result.CurrentOpActualized Then

            result.CandidateStatus = "Rejected"
            result.RejectReason = "Current operation already completed/actualized"

        ElseIf result.CurrentOpScheduled Then

            result.CandidateStatus = "Rejected"
            result.RejectReason = "Current operation already scheduled"

        ElseIf result.CurrentOpStarted Then

            result.CandidateStatus = "Rejected"
            result.RejectReason = "Current operation already started or historical"

        ElseIf result.HasFutureScheduledOp Then

            result.CandidateStatus = "Rejected"
            result.RejectReason = "Future operation already scheduled or completed"

        ElseIf requirePrevScheduled AndAlso result.PrevOpRec <= 0 Then

            result.CandidateStatus = "Rejected"
            result.RejectReason = "No previous operation found"

        ElseIf requirePrevScheduled AndAlso Not result.PrevOpReleased Then

            result.CandidateStatus = "Rejected"
            result.RejectReason = "Previous operation not released"

        ElseIf requirePrevScheduled AndAlso result.ReadyTime = DateTime.MinValue Then

            result.CandidateStatus = "Rejected"
            result.RejectReason = "Previous operation has no valid release time"

        End If

        Return result

    End Function

    Public Function CanScheduleCandidate(wip As WipInfo) As Boolean
        Return wip IsNot Nothing AndAlso
           wip.CandidateStatus.Equals("Candidate", StringComparison.OrdinalIgnoreCase)
    End Function

    'Private Sub WriteWipColumns(row As DataRow, wip As WipInfo)

    '    row("wip_current_op_scheduled") = wip.CurrentOpScheduled
    '    row("wip_current_op_started") = wip.CurrentOpStarted
    '    row("wip_prev_op_rec") = wip.PrevOpRec
    '    row("wip_prev_op_no") = wip.PrevOpNo
    '    row("wip_prev_op_scheduled") = wip.PrevOpScheduled
    '    row("wip_prev_op_end_time") =
    '        If(wip.PrevOpEndTime = DateTime.MinValue,
    '           CType(DBNull.Value, Object),
    '           CType(wip.PrevOpEndTime, Object))
    '    row("wip_any_prior_scheduled") = wip.HasAnyPriorScheduled
    '    row("wip_last_prior_scheduled_op_no") = wip.LastPriorScheduledOpNo
    '    row("wip_has_future_scheduled_op") = wip.HasFutureScheduledOp
    '    row("wip_ready_time") =
    '        If(wip.ReadyTime = DateTime.MinValue,
    '           CType(DBNull.Value, Object),
    '           CType(wip.ReadyTime, Object))
    '    row("wip_score") = wip.WipScore
    '    row("wip_status") = wip.CandidateStatus
    '    row("wip_reject_reason") = wip.RejectReason

    'End Sub
    Private Sub WriteWipColumns(row As DataRow, wip As WipInfo)

        row("wip_current_op_scheduled") = wip.CurrentOpScheduled
        row("wip_current_op_started") = wip.CurrentOpStarted

        row("wip_prev_op_rec") = wip.PrevOpRec
        row("wip_prev_op_no") = wip.PrevOpNo
        row("wip_prev_op_scheduled") = wip.PrevOpScheduled

        row("wip_prev_op_end_time") =
        If(wip.PrevOpEndTime = DateTime.MinValue,
           CType(DBNull.Value, Object),
           CType(wip.PrevOpEndTime, Object))

        ' Backward compatibility:
        ' Existing optimizers still read these scheduled-column names.
        ' Now they represent released WIP depth.
        row("wip_any_prior_scheduled") = wip.HasAnyPriorReleased
        row("wip_last_prior_scheduled_op_no") = wip.LastPriorReleasedOpNo

        row("wip_prev_op_released") = wip.PrevOpReleased

        row("wip_prev_op_release_time") =
        If(wip.PrevOpReleaseTime = DateTime.MinValue,
           CType(DBNull.Value, Object),
           CType(wip.PrevOpReleaseTime, Object))

        row("wip_any_prior_released") = wip.HasAnyPriorReleased
        row("wip_last_prior_released_op_no") = wip.LastPriorReleasedOpNo
        row("wip_last_prior_released_op_rec") = wip.LastPriorReleasedOpRec

        row("wip_has_future_scheduled_op") = wip.HasFutureScheduledOp

        row("wip_ready_time") =
        If(wip.ReadyTime = DateTime.MinValue,
           CType(DBNull.Value, Object),
           CType(wip.ReadyTime, Object))

        row("wip_score") = wip.WipScore
        row("wip_status") = wip.CandidateStatus
        row("wip_reject_reason") = wip.RejectReason

    End Sub
    Private Function IsProgressMarker(row As DataRow) As Boolean

        If row Is Nothing Then Return False

        ' Source Y is authoritative for imported progress boundary.
        If SafeBool(row("source_is_completed")) Then Return True

        ' Use Actual with actual end is also progress evidence.
        If SafeBool(row("opcenter_use_actual")) AndAlso
       SafeDate(row("actual_end_time")) <> DateTime.MinValue Then
            Return True
        End If

        Return False

    End Function

    Private Function ResolveProgressReleaseTime(row As DataRow,
                                            terminatorTime As DateTime) As DateTime

        If row Is Nothing Then Return DateTime.MinValue

        Dim actualEnd As DateTime = SafeDate(row("actual_end_time"))
        If actualEnd <> DateTime.MinValue Then Return actualEnd

        Dim scheduledEnd As DateTime = SafeDate(row("scheduled_end_time"))
        If scheduledEnd <> DateTime.MinValue Then Return scheduledEnd

        If SafeBool(row("source_is_completed")) Then
            Return terminatorTime
        End If

        Return DateTime.MinValue

    End Function

    Private Function ResolveScheduledReleaseTime(row As DataRow) As DateTime

        If row Is Nothing Then Return DateTime.MinValue

        If Not SafeBool(row("is_scheduled")) Then Return DateTime.MinValue

        Return SafeDate(row("scheduled_end_time"))

    End Function

    Private Sub WriteOperationProgressColumns(row As DataRow,
                                          boundaryOpNo As Integer,
                                          boundaryOpRec As Integer,
                                          boundaryReleaseTime As DateTime,
                                          terminatorTime As DateTime)

        Dim opNo As Integer = SafeInt(row("Operation Number"))

        row("order_last_completed_op_no") = boundaryOpNo
        row("order_last_completed_op_rec") = boundaryOpRec

        row("order_last_completed_release_time") =
        If(boundaryReleaseTime = DateTime.MinValue,
           CType(DBNull.Value, Object),
           CType(boundaryReleaseTime, Object))

        Dim effectiveCompleted As Boolean =
        boundaryOpNo > 0 AndAlso opNo > 0 AndAlso opNo <= boundaryOpNo

        row("operation_effective_completed") = effectiveCompleted

        Dim releaseTime As DateTime = DateTime.MinValue
        Dim releasesNext As Boolean = False
        Dim status As String = "Pending"
        Dim conflict As Boolean = False
        Dim reason As String = ""

        If effectiveCompleted Then

            releasesNext = True

            If opNo = boundaryOpNo Then
                releaseTime = boundaryReleaseTime
                status = "CompletedBoundary"
            Else
                releaseTime = ResolveProgressReleaseTime(row, terminatorTime)
                If releaseTime = DateTime.MinValue Then releaseTime = terminatorTime

                status = "CompletedByBoundaryInference"

                If Not IsProgressMarker(row) Then
                    conflict = True
                    reason = "Completion inferred because a later operation is completed"
                End If
            End If

            If releaseTime = DateTime.MinValue Then releaseTime = terminatorTime

        Else

            Dim scheduledRelease As DateTime =
            ResolveScheduledReleaseTime(row)

            If scheduledRelease <> DateTime.MinValue Then
                releasesNext = True
                releaseTime = scheduledRelease
                status = "PlannedScheduled"
            Else
                releasesNext = False
                releaseTime = DateTime.MinValue
                status = "Pending"
            End If

        End If

        row("operation_execution_status") = status
        row("operation_releases_next") = releasesNext

        row("operation_release_time") =
        If(releaseTime = DateTime.MinValue,
           CType(DBNull.Value, Object),
           CType(releaseTime, Object))

        row("operation_status_conflict") = conflict
        row("operation_status_reason") = reason

    End Sub

    Public Function IsCompletedOrActualizedRow(row As DataRow) As Boolean
        If row Is Nothing Then Return False

        If row.Table.Columns.Contains("operation_effective_completed") AndAlso
       SafeBool(row("operation_effective_completed")) Then
            Return True
        End If

        If row.Table.Columns.Contains("source_is_completed") AndAlso
       SafeBool(row("source_is_completed")) Then
            Return True
        End If

        If row.Table.Columns.Contains("opcenter_use_actual") AndAlso
       SafeBool(row("opcenter_use_actual")) AndAlso
       row.Table.Columns.Contains("actual_end_time") AndAlso
       SafeDate(row("actual_end_time")) <> DateTime.MinValue Then
            Return True
        End If

        Return False
    End Function

    Public Function IsCompletedOrActualizedOp(routingDt As DataTable,
                                          opRec As Integer) As Boolean

        If routingDt Is Nothing OrElse opRec <= 0 Then Return False

        For Each r As DataRow In routingDt.Rows
            If SafeInt(r("OrdersID")) = opRec Then
                Return IsCompletedOrActualizedRow(r)
            End If
        Next

        Return False

    End Function

    Public Function GetOperationReleaseTime(routingDt As DataTable,
                                        opRec As Integer) As DateTime

        If routingDt Is Nothing OrElse opRec <= 0 Then Return DateTime.MinValue

        For Each r As DataRow In routingDt.Rows
            If SafeInt(r("OrdersID")) = opRec Then
                Return SafeDate(r("operation_release_time"))
            End If
        Next

        Return DateTime.MinValue

    End Function
    Public Sub PopulateWipColumns(dt As DataTable, planningboard As IPlanningBoard, terminatorTime As DateTime)

        If dt Is Nothing Then Throw New ArgumentNullException(NameOf(dt))
        If planningboard Is Nothing Then Throw New ArgumentNullException(NameOf(planningboard))

        ' Build and sort each order's routing once. Calling GetWipInfo for every
        ' row would otherwise scan and sort the entire DataTable repeatedly.
        Dim rowsByParent As New Dictionary(Of Integer, List(Of DataRow))()

        For Each r As DataRow In dt.Rows
            Dim parentRecord As Integer = GetEffectiveParentRecord(r)
            Dim orderRows As List(Of DataRow) = Nothing

            If Not rowsByParent.TryGetValue(parentRecord, orderRows) Then
                orderRows = New List(Of DataRow)()
                rowsByParent.Add(parentRecord, orderRows)
            End If

            orderRows.Add(r)
        Next

        For Each orderRows As List(Of DataRow) In rowsByParent.Values
            orderRows.Sort(
                Function(leftRow As DataRow, rightRow As DataRow) As Integer
                    Dim compareOpNo As Integer =
                        SafeInt(leftRow("Operation Number")).CompareTo(
                            SafeInt(rightRow("Operation Number")))

                    If compareOpNo <> 0 Then Return compareOpNo

                    Return SafeInt(leftRow("OrdersID")).CompareTo(
                        SafeInt(rightRow("OrdersID")))
                End Function)
        Next

        For Each orderRows As List(Of DataRow) In rowsByParent.Values
            PopulateOrderWipColumns(orderRows, terminatorTime)
        Next

    End Sub

    'Private Sub PopulateOrderWipColumns(orderRows As List(Of DataRow),
    '                                    terminatorTime As DateTime)

    '    If orderRows.Count = 0 Then Return

    '    ' A future operation is one with a strictly greater operation number.
    '    ' Compute that state once per operation-number group.
    '    Dim hasFutureScheduled(orderRows.Count - 1) As Boolean
    '    Dim futureScheduled As Boolean = False
    '    Dim groupEnd As Integer = orderRows.Count - 1

    '    While groupEnd >= 0
    '        Dim opNo As Integer = SafeInt(orderRows(groupEnd)("Operation Number"))
    '        Dim groupStart As Integer = groupEnd

    '        While groupStart > 0 AndAlso
    '              SafeInt(orderRows(groupStart - 1)("Operation Number")) = opNo
    '            groupStart -= 1
    '        End While

    '        For i As Integer = groupStart To groupEnd
    '            hasFutureScheduled(i) = futureScheduled
    '        Next

    '        For i As Integer = groupStart To groupEnd
    '            If SafeBool(orderRows(i)("is_scheduled")) Then
    '                futureScheduled = True
    '                Exit For
    '            End If
    '        Next

    '        groupEnd = groupStart - 1
    '    End While

    '    Dim prevRow As DataRow = Nothing
    '    Dim hasAnyPriorScheduled As Boolean = False
    '    Dim lastPriorScheduledOpNo As Integer = 0
    '    Dim groupStartForward As Integer = 0

    '    While groupStartForward < orderRows.Count
    '        Dim opNo As Integer =
    '            SafeInt(orderRows(groupStartForward)("Operation Number"))
    '        Dim groupEndForward As Integer = groupStartForward

    '        While groupEndForward + 1 < orderRows.Count AndAlso
    '              SafeInt(orderRows(groupEndForward + 1)("Operation Number")) = opNo
    '            groupEndForward += 1
    '        End While

    '        For i As Integer = groupStartForward To groupEndForward
    '            Dim wip As WipInfo =
    '                CreateWipInfo(orderRows(i),
    '                              prevRow,
    '                              hasAnyPriorScheduled,
    '                              lastPriorScheduledOpNo,
    '                              hasFutureScheduled(i),
    '                              terminatorTime,
    '                              0,
    '                              False)

    '            WriteWipColumns(orderRows(i), wip)
    '        Next

    '        For i As Integer = groupStartForward To groupEndForward
    '            If SafeBool(orderRows(i)("is_scheduled")) AndAlso
    '               SafeDate(orderRows(i)("scheduled_end_time")) <> DateTime.MinValue Then

    '                hasAnyPriorScheduled = True
    '                lastPriorScheduledOpNo = opNo
    '            End If
    '        Next

    '        ' Rows are sorted by operation number and record ID, so the final
    '        ' row in this group matches the old previous-operation tie-break.
    '        prevRow = orderRows(groupEndForward)
    '        groupStartForward = groupEndForward + 1
    '    End While

    'End Sub

    Private Sub PopulateOrderWipColumns(orderRows As List(Of DataRow),
                                    terminatorTime As DateTime)

        If orderRows Is Nothing OrElse orderRows.Count = 0 Then Return

        ' ------------------------------------------------------------
        ' STEP 1:
        ' Find the order progress boundary.
        '
        ' Business rule:
        ' If op 240 = Y, then operations up to 240 are complete,
        ' even if op 200 = N.
        ' ------------------------------------------------------------
        Dim boundaryOpNo As Integer = 0
        Dim boundaryOpRec As Integer = 0
        Dim boundaryReleaseTime As DateTime = DateTime.MinValue

        For Each r As DataRow In orderRows

            Dim opNo As Integer = SafeInt(r("Operation Number"))
            If opNo <= 0 Then Continue For

            If Not IsProgressMarker(r) Then Continue For

            Dim releaseTime As DateTime =
            ResolveProgressReleaseTime(r, terminatorTime)

            If releaseTime = DateTime.MinValue Then
                releaseTime = terminatorTime
            End If

            If opNo > boundaryOpNo OrElse
           (opNo = boundaryOpNo AndAlso releaseTime > boundaryReleaseTime) Then

                boundaryOpNo = opNo
                boundaryOpRec = SafeInt(r("OrdersID"))
                boundaryReleaseTime = releaseTime

            End If

        Next

        ' ------------------------------------------------------------
        ' STEP 2:
        ' Write operation-level execution/release columns.
        ' ------------------------------------------------------------
        For Each r As DataRow In orderRows
            WriteOperationProgressColumns(r,
                                      boundaryOpNo,
                                      boundaryOpRec,
                                      boundaryReleaseTime,
                                      terminatorTime)
        Next

        ' ------------------------------------------------------------
        ' STEP 3:
        ' Future block calculation.
        ' Future means a later operation is already scheduled
        ' OR completed/actualized by the progress boundary.
        ' ------------------------------------------------------------
        Dim hasFutureScheduled(orderRows.Count - 1) As Boolean
        Dim futureBlocked As Boolean = False

        Dim groupEnd As Integer = orderRows.Count - 1

        While groupEnd >= 0

            Dim opNo As Integer =
            SafeInt(orderRows(groupEnd)("Operation Number"))

            Dim groupStart As Integer = groupEnd

            While groupStart > 0 AndAlso
              SafeInt(orderRows(groupStart - 1)("Operation Number")) = opNo
                groupStart -= 1
            End While

            For i As Integer = groupStart To groupEnd
                hasFutureScheduled(i) = futureBlocked
            Next

            For i As Integer = groupStart To groupEnd
                If SafeBool(orderRows(i)("is_scheduled")) OrElse
               SafeBool(orderRows(i)("operation_effective_completed")) Then

                    futureBlocked = True
                    Exit For

                End If
            Next

            groupEnd = groupStart - 1

        End While

        ' ------------------------------------------------------------
        ' STEP 4:
        ' Build WIP columns using released operations, not scheduled-only.
        ' ------------------------------------------------------------
        Dim prevRow As DataRow = Nothing

        Dim hasAnyPriorReleased As Boolean = False
        Dim lastPriorReleasedOpNo As Integer = 0
        Dim lastPriorReleasedOpRec As Integer = 0

        Dim groupStartForward As Integer = 0

        While groupStartForward < orderRows.Count

            Dim opNo As Integer =
            SafeInt(orderRows(groupStartForward)("Operation Number"))

            Dim groupEndForward As Integer = groupStartForward

            While groupEndForward + 1 < orderRows.Count AndAlso
              SafeInt(orderRows(groupEndForward + 1)("Operation Number")) = opNo
                groupEndForward += 1
            End While

            For i As Integer = groupStartForward To groupEndForward

                Dim wip As WipInfo =
                CreateWipInfo(orderRows(i),
                              prevRow,
                              hasAnyPriorReleased,
                              lastPriorReleasedOpNo,
                              hasFutureScheduled(i),
                              terminatorTime,
                              0,
                              False,
                              lastPriorReleasedOpRec)

                WriteWipColumns(orderRows(i), wip)

            Next

            ' After writing WIP for this group, decide whether this group
            ' releases the next group.
            Dim groupReleased As Boolean = False
            Dim groupReleaseTime As DateTime = DateTime.MinValue
            Dim groupReleaseRec As Integer = 0

            For i As Integer = groupStartForward To groupEndForward

                If SafeBool(orderRows(i)("operation_releases_next")) Then

                    Dim releaseTime As DateTime =
                    SafeDate(orderRows(i)("operation_release_time"))

                    If releaseTime <> DateTime.MinValue Then

                        groupReleased = True

                        If releaseTime > groupReleaseTime Then
                            groupReleaseTime = releaseTime
                            groupReleaseRec = SafeInt(orderRows(i)("OrdersID"))
                        End If

                    End If

                End If

            Next

            If groupReleased Then
                hasAnyPriorReleased = True
                lastPriorReleasedOpNo = opNo
                lastPriorReleasedOpRec = groupReleaseRec
            End If

            ' Rows are sorted by operation number and record ID.
            prevRow = orderRows(groupEndForward)

            groupStartForward = groupEndForward + 1

        End While

    End Sub
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
    Public Function BuildEffectiveStartByResource(preactor As IPreactor,
                                              planningboard As IPlanningBoard,
                                              resourceNames As IEnumerable(Of String),
                                              Optional metadataDates As Dictionary(Of String, DateTime) = Nothing) As Dictionary(Of String, DateTime)

        Dim result As New Dictionary(Of String, DateTime)(StringComparer.OrdinalIgnoreCase)

        For Each resourceName As String In resourceNames

            If String.IsNullOrWhiteSpace(resourceName) Then Continue For

            Dim metadataDate As DateTime = DateTime.MinValue

            If metadataDates IsNot Nothing AndAlso metadataDates.ContainsKey(resourceName) Then
                metadataDate = metadataDates(resourceName)
            End If

            result(resourceName) =
            GetEffectiveResourceStart(preactor, planningboard, resourceName, metadataDate)

        Next

        Return result

    End Function
    Public Function GetEffectiveResourceStart(preactor As IPreactor,
                                          planningboard As IPlanningBoard,
                                          resourceName As String,
                                          Optional metadataAvailableFrom As DateTime = Nothing) As DateTime

        If preactor Is Nothing Then Throw New ArgumentNullException(NameOf(preactor))
        If planningboard Is Nothing Then Throw New ArgumentNullException(NameOf(planningboard))
        If String.IsNullOrWhiteSpace(resourceName) Then Throw New ArgumentException("Resource name is blank.")

        Dim terminator As DateTime = planningboard.TerminatorTime

        Dim resourceRec As Integer = planningboard.GetResourceNumber(resourceName)
        If resourceRec <= 0 Then
            Throw New Exception("Resource not found: " & resourceName)
        End If

        Dim lastScheduledEnd As DateTime = DateTime.MinValue

        Dim lastEndNullable As Nullable(Of DateTime) =
        GetResourceLastScheduledEnd(preactor, planningboard, resourceRec)

        If lastEndNullable.HasValue Then
            lastScheduledEnd = lastEndNullable.Value
        End If

        Dim metadataDate As DateTime = metadataAvailableFrom

        Dim effective As DateTime =
        MaxDate(terminator, metadataDate, lastScheduledEnd)

        System.Diagnostics.Debug.WriteLine(
        "Effective Resource Start | Resource=" & resourceName &
        " | Terminator=" & FormatDateOrBlank(terminator) &
        " | Metadata=" & FormatDateOrBlank(metadataDate) &
        " | LastScheduledEnd=" & FormatDateOrBlank(lastScheduledEnd) &
        " | Effective=" & FormatDateOrBlank(effective)
    )

        Return effective

    End Function
    Public Function ReadOptimizerSettingDate(preactor As IPreactor,
                                         parameterName As String,
                                         Optional defaultValue As DateTime = Nothing) As DateTime

        If preactor Is Nothing Then Throw New ArgumentNullException(NameOf(preactor))
        If String.IsNullOrWhiteSpace(parameterName) Then Return defaultValue

        Dim settingsFmt As Integer = preactor.GetFormatNumber("GN Optimizer Settings")
        If settingsFmt <= 0 Then Return defaultValue

        Dim parameterField As Integer = preactor.GetFieldNumber(settingsFmt, "Parameter")
        Dim dateField As Integer = preactor.GetFieldNumber(settingsFmt, "Date Value")

        If parameterField <= 0 OrElse dateField <= 0 Then Return defaultValue

        For rec As Integer = 1 To preactor.RecordCount(settingsFmt)

            Dim p As String = preactor.ReadFieldString(settingsFmt, parameterField, rec).Trim()

            If p.Equals(parameterName, StringComparison.OrdinalIgnoreCase) Then

                Dim d As DateTime = preactor.ReadFieldDateTime(settingsFmt, dateField, rec)

                If d = DateTime.MinValue Then Return defaultValue
                Return d

            End If

        Next

        Return defaultValue

    End Function
    Public Class FiringReadinessInfo
        Public Property OrderNo As String
        Public Property ReadyTime As DateTime
        Public Property LastReleaseOpNo As Integer
        Public Property LastReleaseOpRec As Integer
        Public Property LoadingAlreadyReleased As Boolean
        Public Property WipScore As Integer
    End Class

    Public Function BuildFiringReadinessByOrder(dt As DataTable) _
    As Dictionary(Of String, FiringReadinessInfo)

        Dim result As New Dictionary(Of String, FiringReadinessInfo)(StringComparer.OrdinalIgnoreCase)

        If dt Is Nothing OrElse dt.Rows.Count = 0 Then Return result

        Dim rowsByOrder As New Dictionary(Of String, List(Of DataRow))(StringComparer.OrdinalIgnoreCase)

        For Each r As DataRow In dt.Rows

            Dim orderNo As String = SafeStr(r("Order No")).Trim()
            If orderNo = "" Then Continue For

            Dim rows As List(Of DataRow) = Nothing

            If Not rowsByOrder.TryGetValue(orderNo, rows) Then
                rows = New List(Of DataRow)()
                rowsByOrder.Add(orderNo, rows)
            End If

            rows.Add(r)

        Next

        For Each kvp In rowsByOrder

            Dim orderNo As String = kvp.Key
            Dim rows As List(Of DataRow) = kvp.Value

            rows.Sort(
            Function(a As DataRow, b As DataRow) As Integer

                Dim opCompare As Integer =
                    SafeInt(a("Operation Number")).CompareTo(
                        SafeInt(b("Operation Number")))

                If opCompare <> 0 Then Return opCompare

                Return SafeInt(a("OrdersID")).CompareTo(
                    SafeInt(b("OrdersID")))

            End Function)

            ' --------------------------------------------------------
            ' If loading op 290/291 is already released, use that
            ' as firing readiness and do not add loading time again.
            ' --------------------------------------------------------
            Dim loadingReadyTime As DateTime = DateTime.MinValue
            Dim loadingReadyOpNo As Integer = 0
            Dim loadingReadyOpRec As Integer = 0

            For Each r As DataRow In rows

                Dim opNo As Integer = SafeInt(r("Operation Number"))

                If opNo <> 290 AndAlso opNo <> 291 Then Continue For

                If Not SafeBool(r("operation_releases_next")) Then Continue For

                Dim releaseTime As DateTime =
                SafeDate(r("operation_release_time"))

                If releaseTime = DateTime.MinValue Then Continue For

                If releaseTime > loadingReadyTime Then
                    loadingReadyTime = releaseTime
                    loadingReadyOpNo = opNo
                    loadingReadyOpRec = SafeInt(r("OrdersID"))
                End If

            Next

            If loadingReadyTime <> DateTime.MinValue Then

                result(orderNo) = New FiringReadinessInfo With {
                .OrderNo = orderNo,
                .ReadyTime = loadingReadyTime,
                .LastReleaseOpNo = loadingReadyOpNo,
                .LastReleaseOpRec = loadingReadyOpRec,
                .LoadingAlreadyReleased = True,
                .WipScore = 1000 + loadingReadyOpNo
            }

                Continue For

            End If

            ' --------------------------------------------------------
            ' Normal case:
            ' Firing readiness comes from the last released operation
            ' before loading 290.
            ' --------------------------------------------------------
            Dim lastPre290OpNo As Integer = 0
            Dim lastPre290ReadyTime As DateTime = DateTime.MinValue
            Dim lastPre290OpRec As Integer = 0

            For Each r As DataRow In rows

                Dim opNo As Integer = SafeInt(r("Operation Number"))

                If opNo <= 0 OrElse opNo >= 290 Then Continue For

                ' This helper requires the last operation before 290
                ' to be released. If a later pre-290 operation is pending,
                ' firing will not be released yet.
                If opNo > lastPre290OpNo Then
                    lastPre290OpNo = opNo
                    lastPre290ReadyTime = DateTime.MinValue
                    lastPre290OpRec = SafeInt(r("OrdersID"))
                End If

                If opNo = lastPre290OpNo AndAlso
               SafeBool(r("operation_releases_next")) Then

                    Dim releaseTime As DateTime =
                    SafeDate(r("operation_release_time"))

                    If releaseTime <> DateTime.MinValue AndAlso
                   releaseTime > lastPre290ReadyTime Then

                        lastPre290ReadyTime = releaseTime
                        lastPre290OpRec = SafeInt(r("OrdersID"))

                    End If

                End If

            Next

            If lastPre290OpNo > 0 AndAlso
           lastPre290ReadyTime <> DateTime.MinValue Then

                result(orderNo) = New FiringReadinessInfo With {
                .OrderNo = orderNo,
                .ReadyTime = lastPre290ReadyTime,
                .LastReleaseOpNo = lastPre290OpNo,
                .LastReleaseOpRec = lastPre290OpRec,
                .LoadingAlreadyReleased = False,
                .WipScore = 1000 + lastPre290OpNo
            }

            End If

        Next

        Return result

    End Function
    Public Function BuildMetadataAvailabilityByResource(preactor As IPreactor,
                                                    resourceNames As IEnumerable(Of String)) As Dictionary(Of String, DateTime)

        Dim result As New Dictionary(Of String, DateTime)(StringComparer.OrdinalIgnoreCase)

        For Each resourceName As String In resourceNames

            If String.IsNullOrWhiteSpace(resourceName) Then Continue For

            Dim d As DateTime =
            ReadOptimizerSettingDate(preactor,
                                     resourceName & " Available From",
                                     DateTime.MinValue)

            If d <> DateTime.MinValue Then
                result(resourceName) = d
            End If

        Next

        Return result

    End Function
    ' Minimal CSV escape: wrap in quotes if it contains comma or quote; double quotes inside.
    Public Function CsvEscape(value As String) As String
        If value Is Nothing Then Return ""
        Dim mustQuote = value.Contains(",") OrElse value.Contains("""") OrElse value.Contains(vbCr) OrElse value.Contains(vbLf)
        If value.Contains("""") Then value = value.Replace("""", """""")
        If mustQuote Then Return $"""{value}"""
        Return value
    End Function

    Public Function MaxDate(ParamArray dates() As DateTime) As DateTime

        Dim result As DateTime = DateTime.MinValue

        For Each d As DateTime In dates
            If d > result Then result = d
        Next

        Return result

    End Function

    Public Function BuildGnKilnToResourceMap(preactor As IPreactor) As Dictionary(Of String, String)

        Dim result As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

        Dim fmt As Integer = preactor.GetFormatNumber("GN Kilns")
        If fmt <= 0 Then Return result

        Dim nameField As Integer = preactor.GetFieldNumber(fmt, "Name")
        Dim resourceField As Integer = preactor.GetFieldNumber(fmt, "Resource Name")
        Dim activeField As Integer = preactor.GetFieldNumber(fmt, "Active")

        For rec As Integer = 1 To preactor.RecordCount(fmt)

            If activeField > 0 AndAlso preactor.ReadFieldInt(fmt, activeField, rec) = 0 Then
                Continue For
            End If

            Dim kilnName As String = preactor.ReadFieldString(fmt, nameField, rec).Trim()
            Dim resourceName As String = preactor.ReadFieldString(fmt, resourceField, rec).Trim()

            If kilnName = "" Then Continue For
            If resourceName = "" Then resourceName = kilnName

            result(kilnName) = resourceName
            result(resourceName) = resourceName

        Next

        Return result

    End Function
    Public Function BuildMetadataAvailabilityFromGnKilnAvailability(preactor As IPreactor,
                                                                resourceNames As IEnumerable(Of String),
                                                                baseTime As DateTime) As Dictionary(Of String, DateTime)

        Dim result As New Dictionary(Of String, DateTime)(StringComparer.OrdinalIgnoreCase)

        Dim targetResources As New HashSet(Of String)(resourceNames, StringComparer.OrdinalIgnoreCase)
        Dim kilnToResource As Dictionary(Of String, String) = BuildGnKilnToResourceMap(preactor)

        Dim fmt As Integer = preactor.GetFormatNumber("GN Kiln Availability")
        If fmt <= 0 Then Return result

        Dim kilnField As Integer = preactor.GetFieldNumber(fmt, "Kiln")
        Dim statusField As Integer = preactor.GetFieldNumber(fmt, "Availability Status")
        Dim availableFromField As Integer = preactor.GetFieldNumber(fmt, "Available From")
        Dim availableUntilField As Integer = preactor.GetFieldNumber(fmt, "Available Until")
        Dim overrideStartField As Integer = preactor.GetFieldNumber(fmt, "Override Start Time")
        Dim overrideEndField As Integer = preactor.GetFieldNumber(fmt, "Override End Time")
        Dim activeField As Integer = preactor.GetFieldNumber(fmt, "Active")

        For rec As Integer = 1 To preactor.RecordCount(fmt)

            If activeField > 0 AndAlso preactor.ReadFieldInt(fmt, activeField, rec) = 0 Then
                Continue For
            End If

            Dim kilnName As String = preactor.ReadFieldString(fmt, kilnField, rec).Trim()
            If kilnName = "" Then Continue For

            Dim resourceName As String = kilnName

            If kilnToResource.ContainsKey(kilnName) Then
                resourceName = kilnToResource(kilnName)
            End If

            If Not targetResources.Contains(resourceName) Then Continue For

            Dim status As String = preactor.ReadFieldString(fmt, statusField, rec).Trim().ToUpperInvariant()

            Dim availableFrom As DateTime = preactor.ReadFieldDateTime(fmt, availableFromField, rec)
            Dim availableUntil As DateTime = preactor.ReadFieldDateTime(fmt, availableUntilField, rec)
            Dim overrideStart As DateTime = preactor.ReadFieldDateTime(fmt, overrideStartField, rec)
            Dim overrideEnd As DateTime = preactor.ReadFieldDateTime(fmt, overrideEndField, rec)

            Dim metadataStart As DateTime =
                ResolveGnKilnAvailabilityStart(status,
                                               availableFrom,
                                               availableUntil,
                                               overrideStart,
                                               overrideEnd,
                                               baseTime)

            If metadataStart = DateTime.MinValue Then Continue For

            If Not result.ContainsKey(resourceName) OrElse metadataStart > result(resourceName) Then
                result(resourceName) = metadataStart
            End If

        Next

        Return result

    End Function
    Public Function ResolveGnKilnAvailabilityStart(status As String,
                                               availableFrom As DateTime,
                                               availableUntil As DateTime,
                                               overrideStart As DateTime,
                                               overrideEnd As DateTime,
                                               baseTime As DateTime) As DateTime

        Dim result As DateTime = DateTime.MinValue

        Dim normalizedStatus As String = If(status, "").Trim().ToUpperInvariant()

        ' 1. Normal available-from date.
        If availableFrom <> DateTime.MinValue AndAlso availableFrom > baseTime Then
            result = MaxDate(result, availableFrom)
        End If

        ' 2. Manual override start behaves as a stronger start anchor.
        If overrideStart <> DateTime.MinValue AndAlso overrideStart > baseTime Then
            result = MaxDate(result, overrideStart)
        End If

        ' 3. If resource is currently unavailable/down/maintenance,
        ' release it from Available Until or Override End.
        If normalizedStatus <> "" AndAlso normalizedStatus <> "AVAILABLE" Then

            If availableUntil <> DateTime.MinValue AndAlso availableUntil > baseTime Then
                result = MaxDate(result, availableUntil)
            End If

            If overrideEnd <> DateTime.MinValue AndAlso overrideEnd > baseTime Then
                result = MaxDate(result, overrideEnd)
            End If

        End If

        ' 4. If we are currently inside an override window, release at override end.
        If overrideStart <> DateTime.MinValue AndAlso
           overrideEnd <> DateTime.MinValue AndAlso
           overrideStart <= baseTime AndAlso
           overrideEnd > baseTime Then

            result = MaxDate(result, overrideEnd)

        End If

        Return result

    End Function
    Public Function BuildEffectiveStartByResourceFromGnKilnAvailability(preactor As IPreactor,
                                                                    planningboard As IPlanningBoard,
                                                                    resourceNames As IEnumerable(Of String)) As Dictionary(Of String, DateTime)

        Dim result As New Dictionary(Of String, DateTime)(StringComparer.OrdinalIgnoreCase)

        Dim terminator As DateTime = planningboard.TerminatorTime

        Dim metadataDates As Dictionary(Of String, DateTime) =
            BuildMetadataAvailabilityFromGnKilnAvailability(preactor,
                                                            resourceNames,
                                                            terminator)

        For Each resourceName As String In resourceNames

            If String.IsNullOrWhiteSpace(resourceName) Then Continue For

            Dim resourceRec As Integer = planningboard.GetResourceNumber(resourceName)
            If resourceRec <= 0 Then
                Throw New Exception("Resource not found: " & resourceName)
            End If

            Dim metadataDate As DateTime = DateTime.MinValue
            If metadataDates.ContainsKey(resourceName) Then
                metadataDate = metadataDates(resourceName)
            End If

            Dim lastScheduledEnd As DateTime = DateTime.MinValue

            Dim lastEndNullable As Nullable(Of DateTime) =
                GetResourceLastScheduledEnd(preactor, planningboard, resourceRec)

            If lastEndNullable.HasValue Then
                lastScheduledEnd = lastEndNullable.Value
            End If

            Dim effectiveStart As DateTime =
                MaxDate(terminator, metadataDate, lastScheduledEnd)

            result(resourceName) = effectiveStart

            System.Diagnostics.Debug.WriteLine(
                "GN Availability | Resource=" & resourceName &
                " | Terminator=" & FormatDateOrBlank(terminator) &
                " | Metadata=" & FormatDateOrBlank(metadataDate) &
                " | LastScheduledEnd=" & FormatDateOrBlank(lastScheduledEnd) &
                " | EffectiveStart=" & FormatDateOrBlank(effectiveStart)
            )

        Next

        Return result

    End Function
    Public Function GetEffectiveStartFromGnKilnAvailability(preactor As IPreactor,
                                                        planningboard As IPlanningBoard,
                                                        resourceName As String) As DateTime

        Dim names As New List(Of String) From {resourceName}

        Dim dict As Dictionary(Of String, DateTime) =
            BuildEffectiveStartByResourceFromGnKilnAvailability(preactor,
                                                                planningboard,
                                                                names)

        Return dict(resourceName)

    End Function
End Module
