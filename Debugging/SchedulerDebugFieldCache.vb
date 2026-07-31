Option Strict On
Option Explicit On

Imports System
Imports System.Collections.Generic
Imports Preactor

Public Class SchedulerDebugFieldCache
    Private Class FieldEntry
        Public Property LogicalName As String
        Public Property RequestedName As String
        Public Property ResolvedName As String
        Public Property FieldNumber As Integer
        Public Property UsedFallback As Boolean
    End Class

    Private ReadOnly _preactor As IPreactor
    Private ReadOnly _fields As New Dictionary(Of String, FieldEntry)(StringComparer.OrdinalIgnoreCase)
    Private ReadOnly _fieldMapRows As New List(Of DebugFieldMapRow)
    Public ReadOnly Property OrdersFormatNumber As Integer
    Public ReadOnly Property FieldMapRows As List(Of DebugFieldMapRow)
        Get
            Return _fieldMapRows
        End Get
    End Property

    Public Sub New(preactor As IPreactor)
        _preactor = preactor
        OrdersFormatNumber = ResolveOrdersFormat(preactor)

        Resolve("Order No", New String() {"Order No", "Order No."})
        Resolve("Operation Number", New String() {"Operation Number", "Op. No."})
        Resolve("Operation Name", New String() {"Operation Name"})
        Resolve("Belongs to Order No.", New String() {"Belongs to Order No."})
        Resolve("Required Resource", New String() {"Required Resource"})
        Resolve("Resource Group", New String() {"Resource Group"})
        Resolve("Start Time", New String() {"Start Time"})
        Resolve("End Time", New String() {"End Time"})
        Resolve("Due Date", New String() {"Due Date"})
        Resolve("Earliest Start Date", New String() {"Earliest Start Date", "Date Attribute 1"})
        Resolve("Quantity", New String() {"Quantity"})
        Resolve("Setup Time", New String() {"Setup Time"})
        Resolve("Time Per Item", New String() {"Time Per Item", "Op. Time per Item"})
        Resolve("Batch Time", New String() {"Batch Time"})
        Resolve("Process Time Type", New String() {"Process Time Type"})
        Resolve("Cycle Type", New String() {"Cycle Type", "Table Attribute 2"})
        Resolve("Kiln Type", New String() {"Kiln Type", "Klin Type", "Table Attribute 3"})
        Resolve("Volume Occupancy", New String() {"Volume Occupancy", "Numerical Attribute 5"})
        Resolve("is_scheduled", New String() {"is_scheduled"})
        Resolve("scheduled_start_time", New String() {"scheduled_start_time", "Start Time"})
        Resolve("scheduled_end_time", New String() {"scheduled_end_time", "End Time"})
        Resolve("prev_op_is_scheduled", New String() {"prev_op_is_scheduled"})
        Resolve("parent_record", New String() {"parent_record"})
        Resolve("Show", New String() {"Show"})
        Resolve("Disable Op", New String() {"Disable Op"})
        Resolve("Complete", New String() {"Complete", "Toggle Attribute 1"})
        Resolve("Table Attribute 1", New String() {"Table Attribute 1"})
        Resolve("Table Attribute 2", New String() {"Table Attribute 2"})
        Resolve("Table Attribute 3", New String() {"Table Attribute 3"})
        Resolve("Wheel Dia", New String() {"Wheel Dia", "String Attribute 5"})
        Resolve("Wheel thickness", New String() {"Wheel thickness", "Wheel Thickness", "String Attribute 4"})
    End Sub

    Public Function HasField(logicalFieldName As String) As Boolean
        Dim entry As FieldEntry = Nothing
        Return _fields.TryGetValue(logicalFieldName, entry) AndAlso entry.FieldNumber > 0
    End Function

    Public Function ReadString(recordNo As Integer, logicalFieldName As String) As String
        Dim fieldNo As Integer = GetFieldNo(logicalFieldName)
        If fieldNo <= 0 Then Return ""
        Try
            Return _preactor.ReadFieldString(OrdersFormatNumber, fieldNo, recordNo)
        Catch
            Return ""
        End Try
    End Function

    Public Function ReadInt(recordNo As Integer, logicalFieldName As String) As Integer
        Dim fieldNo As Integer = GetFieldNo(logicalFieldName)
        If fieldNo <= 0 Then Return 0
        Try
            Return _preactor.ReadFieldInt(OrdersFormatNumber, fieldNo, recordNo)
        Catch
            Dim value As Integer
            If Integer.TryParse(ReadString(recordNo, logicalFieldName), value) Then Return value
            Return 0
        End Try
    End Function

    Public Function ReadDouble(recordNo As Integer, logicalFieldName As String) As Double
        Dim fieldNo As Integer = GetFieldNo(logicalFieldName)
        If fieldNo <= 0 Then Return 0
        Try
            Return _preactor.ReadFieldDouble(OrdersFormatNumber, fieldNo, recordNo)
        Catch
            Dim value As Double
            If Double.TryParse(ReadString(recordNo, logicalFieldName), Globalization.NumberStyles.Any,
                               Globalization.CultureInfo.InvariantCulture, value) Then Return value
            Return 0
        End Try
    End Function

    Public Function ReadBool(recordNo As Integer, logicalFieldName As String) As Boolean
        Dim fieldNo As Integer = GetFieldNo(logicalFieldName)
        If fieldNo <= 0 Then Return False
        Try
            Return _preactor.ReadFieldBool(OrdersFormatNumber, fieldNo, recordNo)
        Catch
            Try
                Return _preactor.ReadFieldInt(OrdersFormatNumber, fieldNo, recordNo) <> 0
            Catch
                Return SharedHelpers.SafeBool(ReadString(recordNo, logicalFieldName))
            End Try
        End Try
    End Function

    Public Function ReadDateNullable(recordNo As Integer, logicalFieldName As String) As DateTime?
        Dim fieldNo As Integer = GetFieldNo(logicalFieldName)
        If fieldNo <= 0 Then Return Nothing
        Try
            Dim value As DateTime = _preactor.ReadFieldDateTime(OrdersFormatNumber, fieldNo, recordNo)
            If value = DateTime.MinValue Then Return Nothing
            Return value
        Catch
            Return Nothing
        End Try
    End Function

    Private Function ResolveOrdersFormat(preactor As IPreactor) As Integer
        Try
            Dim formatNo As Integer = preactor.GetFormatNumber("Orders")
            If formatNo > 0 Then Return formatNo
        Catch
        End Try
        Try
            Return preactor.FindFirstClassificationString("LAUNCH TIME").Value.FormatNumber
        Catch
            Return 0
        End Try
    End Function

    Private Sub Resolve(logicalName As String, names As String())
        Dim entry As New FieldEntry With {.LogicalName = logicalName, .RequestedName = names(0)}
        For i As Integer = 0 To names.Length - 1
            Try
                Dim fieldNo As Integer = _preactor.GetFieldNumber(OrdersFormatNumber, names(i))
                If fieldNo > 0 Then
                    entry.FieldNumber = fieldNo
                    entry.ResolvedName = names(i)
                    entry.UsedFallback = i > 0
                    Exit For
                End If
            Catch
            End Try
        Next
        _fields(logicalName) = entry
        _fieldMapRows.Add(New DebugFieldMapRow With {
            .LogicalFieldName = logicalName,
            .RequestedFieldName = entry.RequestedName,
            .ResolvedFieldName = entry.ResolvedName,
            .FieldNumber = entry.FieldNumber,
            .Exists = entry.FieldNumber > 0,
            .UsedFallback = entry.UsedFallback,
            .Detail = If(entry.FieldNumber > 0, If(entry.UsedFallback, "Fallback field used.", "Primary field used."), "Field not found.")
        })
    End Sub

    Private Function GetFieldNo(logicalName As String) As Integer
        Dim entry As FieldEntry = Nothing
        If _fields.TryGetValue(logicalName, entry) Then Return entry.FieldNumber
        Return 0
    End Function
End Class

Friend NotInheritable Class SchedulerRunLookupCache
    Private ReadOnly _classificationFormats As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
    Private ReadOnly _formatNumbers As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
    Private ReadOnly _fieldNumbers As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
    Private ReadOnly _resourceNumbers As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
    Private ReadOnly _operationNumbers As New Dictionary(Of Integer, Integer)()
    Private ReadOnly _scheduledStates As New Dictionary(Of Integer, Boolean)()
    Private ReadOnly _operationTimes As New Dictionary(Of Integer, Nullable(Of OperationResourceTimes))()
    Private ReadOnly _previousOperations As New Dictionary(Of Long, Integer)()
    Private ReadOnly _nextOperations As New Dictionary(Of Long, Integer)()

    Friend Property Metrics As SchedulerActionMetricsRow

    Friend Function GetClassificationFormatNumber(preactor As IPreactor,
                                                  classificationName As String) As Integer
        Dim formatNo As Integer
        If _classificationFormats.TryGetValue(classificationName, formatNo) Then
            If Metrics IsNot Nothing Then Metrics.FormatLookupCacheHits += 1
            Return formatNo
        End If

        If Metrics IsNot Nothing Then Metrics.FormatLookupCalls += 1
        formatNo = preactor.FindFirstClassificationString(classificationName).Value.FormatNumber
        _classificationFormats.Add(classificationName, formatNo)
        Return formatNo
    End Function

    Friend Function GetFormatNumber(preactor As IPreactor,
                                    formatName As String) As Integer
        Dim formatNo As Integer
        If _formatNumbers.TryGetValue(formatName, formatNo) Then
            If Metrics IsNot Nothing Then Metrics.FormatLookupCacheHits += 1
            Return formatNo
        End If

        If Metrics IsNot Nothing Then Metrics.FormatLookupCalls += 1
        formatNo = preactor.GetFormatNumber(formatName)
        _formatNumbers.Add(formatName, formatNo)
        Return formatNo
    End Function

    Friend Function GetFieldNumber(preactor As IPreactor,
                                   formatNo As Integer,
                                   fieldName As String) As Integer
        Dim key As String = formatNo.ToString(Globalization.CultureInfo.InvariantCulture) &
                            ControlChars.Tab &
                            fieldName
        Dim fieldNo As Integer
        If _fieldNumbers.TryGetValue(key, fieldNo) Then
            If Metrics IsNot Nothing Then Metrics.FieldLookupCacheHits += 1
            Return fieldNo
        End If

        If Metrics IsNot Nothing Then Metrics.FieldLookupCalls += 1
        fieldNo = preactor.GetFieldNumber(formatNo, fieldName)
        _fieldNumbers.Add(key, fieldNo)
        Return fieldNo
    End Function

    Friend Function GetResourceNumber(planningboard As IPlanningBoard,
                                      resourceName As String) As Integer
        Dim resourceNo As Integer
        If _resourceNumbers.TryGetValue(resourceName, resourceNo) Then
            If Metrics IsNot Nothing Then Metrics.ResourceLookupCacheHits += 1
            Return resourceNo
        End If

        If Metrics IsNot Nothing Then Metrics.ResourceLookupCalls += 1
        resourceNo = planningboard.GetResourceNumber(resourceName)
        _resourceNumbers.Add(resourceName, resourceNo)
        Return resourceNo
    End Function

    Friend Function ReadOperationNumber(preactor As IPreactor,
                                        ordersFormatNo As Integer,
                                        operationNumberFieldNo As Integer,
                                        recordNo As Integer) As Integer
        Dim operationNo As Integer
        If _operationNumbers.TryGetValue(recordNo, operationNo) Then
            If Metrics IsNot Nothing Then Metrics.OperationNumberCacheHits += 1
            Return operationNo
        End If

        If Metrics IsNot Nothing Then Metrics.ReadOperationNumberCalls += 1
        operationNo = preactor.ReadFieldInt(ordersFormatNo, operationNumberFieldNo, recordNo)
        _operationNumbers.Add(recordNo, operationNo)
        Return operationNo
    End Function

    Friend Function IsOperationScheduled(planningboard As IPlanningBoard,
                                         recordNo As Integer,
                                         Optional forceRefresh As Boolean = False) As Boolean
        Dim isScheduled As Boolean
        If Not forceRefresh AndAlso _scheduledStates.TryGetValue(recordNo, isScheduled) Then
            If Metrics IsNot Nothing Then Metrics.ScheduledStateCacheHits += 1
            Return isScheduled
        End If

        If Metrics IsNot Nothing Then Metrics.IsOperationScheduledCalls += 1
        isScheduled = planningboard.IsOperationScheduled(recordNo)
        _scheduledStates(recordNo) = isScheduled
        Return isScheduled
    End Function

    Friend Function GetOperationTimes(planningboard As IPlanningBoard,
                                      recordNo As Integer,
                                      Optional forceRefresh As Boolean = False) As Nullable(Of OperationResourceTimes)
        Dim times As Nullable(Of OperationResourceTimes) = Nothing
        If Not forceRefresh AndAlso _operationTimes.TryGetValue(recordNo, times) Then
            If Metrics IsNot Nothing Then Metrics.OperationTimesCacheHits += 1
            Return times
        End If

        If Metrics IsNot Nothing Then Metrics.GetOperationTimesCalls += 1
        times = planningboard.GetOperationTimes(recordNo)
        _operationTimes(recordNo) = times
        Return times
    End Function

    Friend Function GetPreviousOperation(planningboard As IPlanningBoard,
                                         recordNo As Integer,
                                         routeIndex As Integer) As Integer
        Dim key As Long = CreateRouteKey(recordNo, routeIndex)
        Dim previousRecord As Integer
        If _previousOperations.TryGetValue(key, previousRecord) Then
            If Metrics IsNot Nothing Then Metrics.PreviousOperationCacheHits += 1
            Return previousRecord
        End If

        If Metrics IsNot Nothing Then Metrics.GetPreviousOperationCalls += 1
        previousRecord = planningboard.GetPreviousOperation(recordNo, routeIndex)
        _previousOperations.Add(key, previousRecord)
        Return previousRecord
    End Function

    Friend Function GetNextOperation(planningboard As IPlanningBoard,
                                     recordNo As Integer,
                                     routeIndex As Integer) As Integer
        Dim key As Long = CreateRouteKey(recordNo, routeIndex)
        Dim nextRecord As Integer
        If _nextOperations.TryGetValue(key, nextRecord) Then
            If Metrics IsNot Nothing Then Metrics.NextOperationCacheHits += 1
            Return nextRecord
        End If

        If Metrics IsNot Nothing Then Metrics.GetNextOperationCalls += 1
        nextRecord = planningboard.GetNextOperation(recordNo, routeIndex)
        _nextOperations.Add(key, nextRecord)
        Return nextRecord
    End Function

    Friend Sub MarkOperationPlaced(recordNo As Integer)
        _scheduledStates(recordNo) = True
        _operationTimes.Remove(recordNo)
    End Sub

    Friend Sub InvalidateOperation(recordNo As Integer)
        _scheduledStates.Remove(recordNo)
        _operationTimes.Remove(recordNo)
    End Sub

    Private Shared Function CreateRouteKey(recordNo As Integer,
                                           routeIndex As Integer) As Long
        Return (CLng(recordNo) << 32) Or
               (CLng(routeIndex) And &HFFFFFFFFL)
    End Function
End Class
