Option Strict On
Option Explicit On

Imports System.Configuration
Imports System.Collections.Specialized
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Globalization
Imports System.IO
Imports System.Linq
Imports Microsoft.VisualBasic.FileIO
Imports Preactor

Friend Structure PressSetupRates
    Friend RingMinutes As Double?
    Friend BorePinMinutes As Double?

    Friend Shared Function HasIdleGap(previousEnd As DateTime, setupStart As DateTime) As Boolean
        Return setupStart > previousEnd AndAlso Not SameScheduleInstant(previousEnd, setupStart)
    End Function

    Friend Shared Function SameScheduleInstant(first As DateTime, second As DateTime) As Boolean
        If first = second Then Return True
        If Math.Abs(first.Ticks - second.Ticks) >= 4 * TimeSpan.TicksPerMillisecond Then Return False
        ' The Orders timestamps are stored as SQL datetime. A difference that
        ' rounds to the same saved instant does not create a separate idle slot.
        Return New SqlDateTime(first).Value = New SqlDateTime(second).Value
    End Function

    Friend Function Calculate(current As PressCampaignSlot, previous As PressCampaignSlot,
                              start As DateTime, repair As Boolean) As Double?
        If Not RingMinutes.HasValue AndAlso Not BorePinMinutes.HasValue Then Return Nothing
        Dim ring = PressCampaignScheduler.Key(current.Ring)
        Dim pin = PressCampaignScheduler.Key(current.Pin)
        Dim idle = previous Is Nothing OrElse (Not repair AndAlso HasIdleGap(previous.Finish, start))
        Dim previousRing = If(idle, "", PressCampaignScheduler.Key(previous.Ring))
        Dim previousPin = If(idle, "", PressCampaignScheduler.Key(previous.Pin))

        If Not idle AndAlso ring <> "" AndAlso pin <> "" AndAlso
           ring = previousRing AndAlso pin = previousPin Then Return 0

        Dim ringRate = If(ring <> "" AndAlso (idle OrElse (previousRing <> "" AndAlso ring <> previousRing)),
                          RingMinutes, Nothing)
        Dim pinRate = If(pin <> "" AndAlso (idle OrElse (previousPin <> "" AndAlso pin <> previousPin)),
                         BorePinMinutes, Nothing)
        If Not ringRate.HasValue AndAlso Not pinRate.HasValue Then Return Nothing
        Return ringRate.GetValueOrDefault() + pinRate.GetValueOrDefault()
    End Function
End Structure

Friend NotInheritable Class PressSetupBoard
    Private Const SetupFieldName As String = "Setup Time"
    ' The site's imported base Setup Time for operation 200 is zero.
    ' A duration written by a previous scheduling run is not a baseline.
    Private Const BaseSetupDays As Double = 0
    Private Const MinutesPerDay As Double = 1440
    Private Const SnapshotSlot As Integer = 200
    Private Shared ReadOnly Headers As String() = {"Workstation", "RingSetupMinutes", "BorePinSetupMinutes"}

    Private ReadOnly preactor As IPreactor
    Private ReadOnly board As IPlanningBoard
    Private ReadOnly ordersFormat As Integer
    Private ReadOnly setupField As Integer
    Private ReadOnly rates As Dictionary(Of Integer, PressSetupRates)
    Private ReadOnly horizon As DateTimeSpan
    Private ReadOnly timelines As New Dictionary(Of Integer, List(Of PressCampaignSlot))
    Private ReadOnly savedSetupTimes As New Dictionary(Of Integer, Double)
    Friend ReadOnly RingFieldName As String
    Friend ReadOnly BorePinFieldName As String

    Private Sub New(preactor As IPreactor, path As String, ringFieldName As String, borePinFieldName As String)
        Me.preactor = preactor
        board = preactor.PlanningBoard
        ordersFormat = preactor.GetFormatNumber("Orders")
        setupField = RequireField(SetupFieldName)
        If preactor.GetFieldType(ordersFormat, setupField) <> PreactorFieldType.Duration Then
            Throw New InvalidOperationException("Orders." & SetupFieldName & " must be a duration.")
        End If
        horizon = board.ScheduleHorizon
        Me.RingFieldName = ringFieldName
        Me.BorePinFieldName = borePinFieldName
        Dim ringField = RequireField(ringFieldName)
        Dim pinField = RequireField(borePinFieldName)
        If ringField = pinField Then Throw New InvalidOperationException("Press ring and bore/pin must use different Orders fields.")

        Dim resourceFormat = preactor.GetFormatNumber("Resources")
        Dim capacityField = preactor.GetFieldNumber(resourceFormat, "Finite or Infinite")
        If capacityField <= 0 Then Throw New InvalidOperationException("Resources.Finite or Infinite is required for press setup.")
        rates = ReadRates(path,
            Function(name)
                Dim resource = board.GetResourceNumber(name)
                If resource <= 0 OrElse Not String.Equals(
                    preactor.ReadFieldString(resourceFormat, capacityField, resource).Trim(),
                    "Finite", StringComparison.OrdinalIgnoreCase) Then
                    Throw New ArgumentException("Expected a finite resource.")
                End If
                Return resource
            End Function)
    End Sub

    Friend Shared Function Load(preactor As IPreactor) As PressSetupBoard
        Dim host = ConfigurationManager.AppSettings
        Dim dllConfig = GetType(PressSetupBoard).Assembly.Location & ".config"
        Dim library As KeyValueConfigurationCollection = Nothing
        If File.Exists(dllConfig) Then
            Dim map As New ExeConfigurationFileMap With {.ExeConfigFilename = dllConfig}
            library = ConfigurationManager.OpenMappedExeConfiguration(map, ConfigurationUserLevel.None).AppSettings.Settings
        End If

        Dim path = ReadSetting(host, library, "PressSetupCsvPath", "")
        If path = "" Then Return Nothing
        If Not IO.Path.IsPathRooted(path) Then path = IO.Path.Combine(preactor.ParseShellString("{PATH}"), path)
        Dim ringField = ReadSetting(host, library, "PressRingField", "String Attribute 5")
        Dim pinField = ReadSetting(host, library, "PressBorePinField", "String Attribute 4")
        Return New PressSetupBoard(preactor, path, ringField, pinField)
    End Function

    Private Shared Function ReadSetting(host As NameValueCollection, library As KeyValueConfigurationCollection,
                                        name As String, fallback As String) As String
        Dim value = host(name)
        If value Is Nothing AndAlso library IsNot Nothing Then value = library(name)?.Value
        Return If(value, fallback).Trim()
    End Function

    Private Function RequireField(name As String) As Integer
        Dim field = preactor.GetFieldNumber(ordersFormat, name)
        If field <= 0 Then Throw New InvalidOperationException("Press setup requires Orders." & name & ".")
        Return field
    End Function

    Friend Shared Function ReadRates(path As String, resolveResource As Func(Of String, Integer)) As Dictionary(Of Integer, PressSetupRates)
        Dim result As New Dictionary(Of Integer, PressSetupRates)
        Dim errors As New List(Of String)
        Using parser As New TextFieldParser(path)
            parser.SetDelimiters(",")
            If parser.EndOfData OrElse Not Headers.SequenceEqual(parser.ReadFields()) Then
                Throw New InvalidDataException("Expected CSV columns " & String.Join(",", Headers) & " in " & path)
            End If
            While Not parser.EndOfData
                Dim line = parser.LineNumber
                Try
                    Dim fields = parser.ReadFields()
                    If fields.Length <> Headers.Length OrElse String.IsNullOrWhiteSpace(fields(0)) Then
                        Throw New FormatException("Expected a workstation and two duration cells.")
                    End If
                    Dim workstation = fields(0).Trim()
                    Dim resource As Integer
                    Try
                        resource = resolveResource(workstation)
                    Catch ex As ArgumentException
                        Throw New FormatException("Unknown or non-finite workstation '" & workstation & "'.", ex)
                    End Try
                    If resource <= 0 Then Throw New FormatException("Unknown workstation '" & workstation & "'.")
                    Dim entry As New PressSetupRates With {
                        .RingMinutes = ReadMinutes(fields(1), Headers(1)),
                        .BorePinMinutes = ReadMinutes(fields(2), Headers(2))}
                    If result.ContainsKey(resource) Then Throw New FormatException("Duplicate workstation '" & workstation & "'.")
                    result.Add(resource, entry)
                Catch ex As Exception When TypeOf ex Is FormatException OrElse TypeOf ex Is MalformedLineException
                    errors.Add("Line " & line.ToString(CultureInfo.InvariantCulture) & " — " & ex.Message)
                End Try
            End While
        End Using
        If errors.Count > 0 Then Throw New InvalidDataException(path & Environment.NewLine & String.Join(Environment.NewLine, errors))
        Return result
    End Function

    Private Shared Function ReadMinutes(value As String, column As String) As Double?
        If String.IsNullOrWhiteSpace(value) Then Return Nothing
        Dim minutes As Double
        If Not Double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, minutes) OrElse
           Double.IsNaN(minutes) OrElse minutes < 0 OrElse minutes >= TimeSpan.MaxValue.TotalMinutes / 2 Then
            Throw New FormatException(column & " requires a nonnegative duration in minutes or a blank cell.")
        End If
        Return minutes
    End Function

    Friend Shared Function ReadSlot(record As Integer, operationNumber As Integer, allocation As OperationResourceTimes,
                                   ring As String, pin As String) As PressCampaignSlot
        Return New PressCampaignSlot With {
            .RecordNo = record, .Resource = allocation.Resource, .IsPress = operationNumber = 200,
            .Start = allocation.OperationTimes.ChangeStart, .ProcessStart = allocation.OperationTimes.ProcessStart,
            .Finish = allocation.OperationTimes.ProcessEnd,
            .Ring = If(operationNumber = 200, PressCampaignScheduler.Key(ring), ""),
            .Pin = If(operationNumber = 200, PressCampaignScheduler.Key(pin), "")}
    End Function

    Friend Sub LoadTimeline(allocations As IEnumerable(Of PressCampaignSlot), resources As IEnumerable(Of Integer))
        For Each resource In resources.Distinct()
            timelines.Add(resource, New List(Of PressCampaignSlot))
        Next
        For Each slot In allocations
            Dim slots As List(Of PressCampaignSlot) = Nothing
            If timelines.TryGetValue(slot.Resource, slots) Then slots.Add(slot)
        Next
        For Each slots In timelines.Values
            slots.Sort(Function(left, right)
                           Dim order = left.Start.CompareTo(right.Start)
                           Return If(order <> 0, order, left.RecordNo.CompareTo(right.RecordNo))
                       End Function)
            For index = 0 To slots.Count - 1
                If slots(index).Finish < slots(index).Start OrElse
                   (index > 0 AndAlso slots(index).Start < slots(index - 1).Finish) Then
                    Throw New InvalidOperationException("Press has an invalid or overlapping fixed allocation at operation " &
                                                        slots(index).RecordNo.ToString())
                End If
            Next
        Next
    End Sub

    Friend Function Repair(routing As DataTable) As Integer
        Dim lockField = RequireField("Lock Operation")
        Dim actualSetupField = RequireField("Actual Setup Start")
        Dim totalSetupField = RequireField("Total Setup Time")
        Dim original As New Dictionary(Of Integer, PressCampaignSlot)
        Dim orderNumbers As New Dictionary(Of Integer, String)
        Dim unallocated As New List(Of Integer)
        Dim earliestByRecord As New Dictionary(Of Integer, DateTime)
        Dim boundary = board.TerminatorTime

        For Each row As DataRow In routing.Rows
            Dim record = SharedHelpers.SafeInt(row("OrdersID"))
            orderNumbers.Add(record, SharedHelpers.SafeStr(row("Order No")))
            If Not SharedHelpers.SafeBool(row("is_scheduled")) Then
                unallocated.Add(record)
                Continue For
            End If

            Dim times = board.GetOperationTimes(record)
            If Not times.HasValue OrElse times.Value.Resource <= 0 Then
                Throw New InvalidOperationException("Cannot snapshot allocated order " & orderNumbers(record) &
                                                    " (record " & record.ToString() & ").")
            End If
            Dim slot = ReadSlot(record, SharedHelpers.SafeInt(row("Operation Number")), times.Value,
                                SharedHelpers.SafeStr(row("Wheel Dia")), SharedHelpers.SafeStr(row("Wheel thickness")))
            original.Add(record, slot)
            If slot.IsPress AndAlso slot.Start >= boundary AndAlso
               Not SharedHelpers.SafeBool(row("source_is_completed")) AndAlso
               Not SharedHelpers.SafeBool(row("opcenter_use_actual")) AndAlso
               SharedHelpers.SafeDate(row("actual_start_time")) = DateTime.MinValue AndAlso
               SharedHelpers.SafeDate(row("actual_end_time")) = DateTime.MinValue AndAlso
               Not preactor.ReadFieldBool(ordersFormat, lockField, record) AndAlso
               SharedHelpers.SafeDate(preactor.ReadFieldDateTime(ordersFormat, actualSetupField, record)) = DateTime.MinValue Then
                earliestByRecord.Add(record, MaxDate(SharedHelpers.SafeDate(row("Pressing earliest start")).Date,
                                                     SharedHelpers.SafeDate(row("wip_ready_time"))))
            End If
        Next

        Dim originalSetupTimes As New Dictionary(Of Integer, Double)
        Dim desiredSetup As New Dictionary(Of Integer, Double)
        Dim affected As New Dictionary(Of Integer, Integer)
        Dim eligiblePresses As New HashSet(Of Integer)(earliestByRecord.Keys.Select(Function(record) original(record).Resource))
        Dim sequences = original.Values.Where(Function(slot) eligiblePresses.Contains(slot.Resource)).
            GroupBy(Function(slot) slot.Resource).ToDictionary(
            Function(group) group.Key,
            Function(group) group.OrderBy(Function(slot) slot.Start).ThenBy(Function(slot) slot.RecordNo).ToList())
        For Each slots In sequences.Values
            Dim previous As PressCampaignSlot = Nothing
            For Each slot In slots
                If earliestByRecord.ContainsKey(slot.RecordNo) Then
                    Dim desired = CalculateDays(slot, previous, slot.Start, repair:=True)
                    desiredSetup.Add(slot.RecordNo, desired)
                    Dim stored = preactor.ReadFieldDouble(ordersFormat, setupField, slot.RecordNo)
                    If stored <> desired Then originalSetupTimes.Add(slot.RecordNo, stored)
                    If stored <> desired OrElse Not NativeSetupMatches(slot, desired,
                            preactor.ReadFieldDouble(ordersFormat, totalSetupField, slot.RecordNo)) Then
                        affected(slot.Resource) = slot.RecordNo
                    End If
                End If
                previous = slot
            Next
        Next
        If affected.Count = 0 Then Return 0

        Dim dependents As New HashSet(Of Integer)
        For Each record In earliestByRecord.Keys
            If Not affected.ContainsKey(original(record).Resource) Then Continue For
            Dim index = 1
            Do
                Dim dependent = board.GetNextOperation(record, index)
                If dependent <= 0 Then Exit Do
                If original.ContainsKey(dependent) Then dependents.Add(dependent)
                index += 1
            Loop
        Next
        Dim validDependents = dependents.Where(Function(record) board.CanOperationStart(record, original(record).Start)).ToList()

        Dim stage = "write Setup Time"
        Dim currentRecord = 0
        Dim attemptedWrites As New List(Of Integer)
        Try
            board.SaveScheduleState(SnapshotSlot)
        Catch ex As Exception
            Throw New InvalidOperationException("RepairPress could not save the schedule state. No setup writes were attempted. " & ex.Message, ex)
        End Try
        Try
            For Each record In originalSetupTimes.Keys
                currentRecord = record
                attemptedWrites.Add(record)
                WriteSetupTime(record, desiredSetup(record))
            Next

            stage = "recalculate press"
            For Each press In affected
                currentRecord = press.Value
                board.RecalculateOperationTimes(press.Key)
            Next

            stage = "verify allocations"
            Dim actual As New Dictionary(Of Integer, PressCampaignSlot)
            For Each slot In original.Values
                currentRecord = slot.RecordNo
                Dim times = board.GetOperationTimes(slot.RecordNo)
                If Not times.HasValue OrElse times.Value.Resource <> slot.Resource Then
                    Throw New InvalidOperationException("The original allocation is missing or has changed resource.")
                End If
                If Not affected.ContainsKey(slot.Resource) OrElse Not earliestByRecord.ContainsKey(slot.RecordNo) Then
                    If Not Matches(slot, times) Then Throw New InvalidOperationException("A fixed or unrelated allocation moved.")
                Else
                    actual.Add(slot.RecordNo, ReadSlot(slot.RecordNo, 200, times.Value, slot.Ring, slot.Pin))
                End If
            Next
            For Each record In unallocated
                currentRecord = record
                If board.IsOperationScheduled(record) Then Throw New InvalidOperationException("An unallocated operation was scheduled.")
            Next

            stage = "verify press timing and setup"
            For Each resource In affected.Keys
                Dim before = sequences(resource)
                Dim after = before.Select(Function(slot) If(actual.ContainsKey(slot.RecordNo), actual(slot.RecordNo), slot)).ToList()
                Dim previous As PressCampaignSlot = Nothing
                For Each slot In after
                    currentRecord = slot.RecordNo
                    If slot.ProcessStart < slot.Start OrElse slot.Finish < slot.ProcessStart OrElse
                       (previous IsNot Nothing AndAlso slot.Start < previous.Finish) Then
                        Throw New InvalidOperationException("Opcenter returned an invalid interval or changed the press sequence.")
                    End If
                    If actual.ContainsKey(slot.RecordNo) Then
                        Dim expected = CalculateDays(slot, previous, slot.Start, repair:=True)
                        Dim stored = preactor.ReadFieldDouble(ordersFormat, setupField, slot.RecordNo)
                        Dim total = preactor.ReadFieldDouble(ordersFormat, totalSetupField, slot.RecordNo)
                        If stored <> expected OrElse Not NativeSetupMatches(slot, expected, total) Then
                            Throw New InvalidOperationException(FormattableString.Invariant(
                                $"Expected {expected * MinutesPerDay} setup minutes. Stored {stored * MinutesPerDay} minutes. Native total {total * MinutesPerDay} minutes. Scheduled setup interval {(slot.ProcessStart - slot.Start).TotalMinutes} minutes."))
                        End If
                        Dim priorInterval = original(slot.RecordNo)
                        If slot.Start <> priorInterval.Start OrElse slot.ProcessStart <> priorInterval.ProcessStart OrElse
                           slot.Finish <> priorInterval.Finish Then
                            If slot.Start < boundary OrElse slot.Start < earliestByRecord(slot.RecordNo) OrElse
                               slot.Start < horizon.Start OrElse slot.Finish > horizon.End OrElse
                               Not board.CanOperationStart(slot.RecordNo, slot.Start) Then
                                Throw New InvalidOperationException("The recalculated interval violates readiness or the schedule boundary.")
                            End If
                            Dim processState = board.GetCurrentCalendarState(resource, slot.ProcessStart)
                            Dim setupState = If(expected > 0, board.GetCurrentCalendarState(resource, slot.Start), Nothing)
                            If (processState.HasValue AndAlso processState.Value.Efficiency <= 0) OrElse
                               (setupState.HasValue AndAlso Not setupState.Value.SetupAllowed) Then
                                Throw New InvalidOperationException("The recalculated interval starts in an unavailable calendar state.")
                            End If
                        End If
                    End If
                    previous = slot
                Next
                Dim oldViolations = PressCampaignScheduler.Violations(before.Where(Function(slot) slot.IsPress))
                For Each issue In PressCampaignScheduler.Violations(after.Where(Function(slot) slot.IsPress))
                    Dim oldLimit As DateTime
                    If Not oldViolations.TryGetValue(issue.Key, oldLimit) OrElse issue.Value > oldLimit Then
                        currentRecord = issue.Key
                        Throw New InvalidOperationException("The recalculated interval worsens a ring cooldown violation.")
                    End If
                Next
            Next

            stage = "verify downstream readiness"
            For Each record In validDependents
                currentRecord = record
                Dim slot As PressCampaignSlot = Nothing
                If Not actual.TryGetValue(record, slot) Then slot = original(record)
                If Not board.CanOperationStart(record, slot.Start) Then
                    Throw New InvalidOperationException("The recalculation invalidated an existing downstream operation.")
                End If
            Next
            Return originalSetupTimes.Count
        Catch repairError As Exception
            Try
                RestoreRepair(attemptedWrites, originalSetupTimes, original.Values, unallocated)
            Catch rollbackError As Exception
                Throw New InvalidOperationException(RepairFailure(stage, currentRecord, orderNumbers, original, repairError) &
                    " Restoration could not be verified. " & rollbackError.Message,
                    New AggregateException(repairError, rollbackError))
            End Try
            Throw New InvalidOperationException(RepairFailure(stage, currentRecord, orderNumbers, original, repairError) &
                                                " Original allocations and Setup Time were restored and verified.", repairError)
        End Try
    End Function

    Private Shared Function MaxDate(first As DateTime, second As DateTime) As DateTime
        Return If(first > second, first, second)
    End Function

    Private Shared Function NativeSetupMatches(slot As PressCampaignSlot, expected As Double, total As Double) As Boolean
        Dim hasSetupInterval = Not PressSetupRates.SameScheduleInstant(slot.Start, slot.ProcessStart)
        Return Math.Abs(total - expected) <= TimeSpan.FromMilliseconds(1).TotalDays AndAlso
               (expected > 0) = hasSetupInterval
    End Function

    Private Function RepairFailure(stage As String, record As Integer, orderNumbers As Dictionary(Of Integer, String),
                                   original As Dictionary(Of Integer, PressCampaignSlot), errorDetail As Exception) As String
        Dim orderNumber As String = Nothing
        orderNumbers.TryGetValue(record, orderNumber)
        Dim slot As PressCampaignSlot = Nothing
        Dim resource = "unallocated"
        If original.TryGetValue(record, slot) Then
            resource = slot.Resource.ToString(CultureInfo.InvariantCulture)
            Try
                resource = board.GetResourceName(slot.Resource)
            Catch
                ' Preserve the original exception and the numeric resource identifier.
            End Try
        End If
        Return $"RepairPress failed at {stage} for order {orderNumber} (record {record}) on {resource}. {errorDetail.Message}"
    End Function

    Private Sub RestoreRepair(attemptedWrites As IEnumerable(Of Integer), originalSetupTimes As Dictionary(Of Integer, Double),
                              originalSlots As IEnumerable(Of PressCampaignSlot), unallocated As IEnumerable(Of Integer))
        Dim failures As New List(Of Exception)
        For Each record In attemptedWrites
            Try
                WriteSetupTime(record, originalSetupTimes(record))
            Catch ex As Exception
                failures.Add(ex)
            End Try
        Next
        Try
            board.RestoreScheduleState(SnapshotSlot)
        Catch ex As Exception
            failures.Add(ex)
        End Try
        For Each slot In originalSlots
            If Not Matches(slot, board.GetOperationTimes(slot.RecordNo)) Then
                failures.Add(New InvalidOperationException("Could not restore allocation for record " & slot.RecordNo.ToString()))
            End If
        Next
        For Each record In unallocated
            If board.IsOperationScheduled(record) Then
                failures.Add(New InvalidOperationException("Could not restore unallocated record " & record.ToString()))
            End If
        Next
        For Each entry In originalSetupTimes
            If preactor.ReadFieldDouble(ordersFormat, setupField, entry.Key) <> entry.Value Then
                failures.Add(New InvalidOperationException("Could not restore Setup Time for record " & entry.Key.ToString()))
            End If
        Next
        If failures.Count > 0 Then Throw New AggregateException(failures(0).Message, failures)
    End Sub

    Private Shared Function WindowIndex(slots As List(Of PressCampaignSlot), start As DateTime) As Integer
        Dim low = 0
        Dim high = slots.Count
        While low < high
            Dim middle = low + (high - low) \ 2
            If slots(middle).Finish <= start Then
                low = middle + 1
            Else
                high = middle
            End If
        End While
        Return low
    End Function

    Private Function CalculateDays(current As PressCampaignSlot, previous As PressCampaignSlot,
                                   start As DateTime, Optional repair As Boolean = False) As Double
        Dim rate As PressSetupRates
        If Not rates.TryGetValue(current.Resource, rate) Then Return BaseSetupDays
        Dim minutes = rate.Calculate(current, previous, start, repair)
        Return If(minutes.HasValue, minutes.GetValueOrDefault() / MinutesPerDay, BaseSetupDays)
    End Function

    Private Function ReadSetupTime(record As Integer) As Double
        Dim value As Double
        If Not savedSetupTimes.TryGetValue(record, value) Then
            value = preactor.ReadFieldDouble(ordersFormat, setupField, record)
            savedSetupTimes.Add(record, value)
        End If
        Return value
    End Function

    Private Sub WriteSetupTime(record As Integer, value As Double)
        preactor.WriteField(ordersFormat, setupField, record, value)
    End Sub

    Private Sub VerifySetupTime(slot As PressCampaignSlot, previous As PressCampaignSlot)
        Dim expected = CalculateDays(slot, previous, slot.Start)
        Dim stored = preactor.ReadFieldDouble(ordersFormat, setupField, slot.RecordNo)
        If stored = expected AndAlso
           (expected <> 0 OrElse Not PressSetupRates.HasIdleGap(slot.Start, slot.ProcessStart)) Then Return

        Dim predecessor = If(previous Is Nothing, "none", previous.RecordNo.ToString(CultureInfo.InvariantCulture))
        Dim gap = If(previous Is Nothing, "none", (slot.Start - previous.Finish).TotalMilliseconds.ToString(CultureInfo.InvariantCulture))
        Throw New InvalidOperationException(FormattableString.Invariant(
            $"Press setup mismatch for record {slot.RecordNo} on resource {slot.Resource}. Predecessor: {predecessor}. Gap (ms): {gap}. Expected: {expected * MinutesPerDay} min. Stored: {stored * MinutesPerDay} min. Scheduled setup interval: {(slot.ProcessStart - slot.Start).TotalMinutes} min."))
    End Sub

    Friend Function Query(order As PressCampaignOrder, resource As Integer, earliest As DateTime,
                          Optional trial As PressCampaignSlot = Nothing) As PressCampaignSlot
        Dim limit = horizon.End
        If earliest < horizon.Start Then earliest = horizon.Start
        If earliest >= limit Then Return Nothing
        Dim slots = timelines(resource)
        Dim candidate As New PressCampaignSlot With {
            .RecordNo = order.RecordNo, .Resource = resource, .Ring = order.Ring, .Pin = order.Pin, .IsPress = True}
        Dim original = ReadSetupTime(order.RecordNo)
        Dim applied = original
        Try
            For index = WindowIndex(slots, earliest) To slots.Count
                Dim prior = If(index = 0, Nothing, slots(index - 1))
                If trial IsNot Nothing AndAlso trial.Finish <= earliest AndAlso
                   (prior Is Nothing OrElse trial.Finish > prior.Finish) Then prior = trial
                Dim successor = If(index = slots.Count, Nothing, slots(index))
                Dim start = If(prior Is Nothing OrElse earliest >= prior.Finish, earliest, prior.Finish)
                Dim windowEnd = If(successor IsNot Nothing AndAlso successor.Start < limit, successor.Start, limit)
                While start < windowEnd
                    Dim value = CalculateDays(candidate, prior, start)
                    If value <> applied Then
                        applied = value
                        WriteSetupTime(order.RecordNo, value)
                    End If
                    Dim times = board.QueryOperationOnResource(order.RecordNo, resource, start)
                    If Not times.HasValue Then Exit While
                    Dim result = times.Value
                    If result.ChangeStart < start OrElse result.ProcessStart < result.ChangeStart OrElse
                       result.ProcessEnd < result.ProcessStart OrElse result.ProcessEnd <= result.ChangeStart Then
                        Throw New InvalidOperationException("Opcenter returned an invalid press interval for operation " & order.RecordNo.ToString())
                    End If
                    If result.ProcessEnd > windowEnd OrElse result.ChangeStart >= windowEnd Then Exit While
                    If value <> CalculateDays(candidate, prior, result.ChangeStart) Then
                        start = result.ChangeStart
                        Continue While
                    End If
                    candidate.Start = result.ChangeStart
                    candidate.ProcessStart = result.ProcessStart
                    candidate.Finish = result.ProcessEnd
                    candidate.SetupDays = value
                    If successor IsNot Nothing AndAlso successor.IsPress AndAlso
                       CalculateDays(successor, prior, successor.Start) <>
                       CalculateDays(successor, candidate, successor.Start) Then Exit While
                    Return candidate
                End While
                If successor Is Nothing OrElse successor.Finish >= limit Then Exit For
                earliest = If(successor.Finish > earliest, successor.Finish, earliest)
            Next
            Return Nothing
        Finally
            If applied <> original Then WriteSetupTime(order.RecordNo, original)
        End Try
    End Function

    Friend Sub VerifyAndRecord(slot As PressCampaignSlot)
        Dim actual = board.GetOperationTimes(slot.RecordNo)
        If Not Matches(slot, actual) Then
            Throw New InvalidOperationException("Opcenter changed the verified press interval for operation " & slot.RecordNo.ToString())
        End If
        Dim slots = timelines(slot.Resource)
        Dim index = WindowIndex(slots, slot.Start)
        VerifySetupTime(slot, If(index = 0, Nothing, slots(index - 1)))
        For position = index To slots.Count - 1
            Dim fixedSlot = slots(position)
            If Not Matches(fixedSlot, board.GetOperationTimes(fixedSlot.RecordNo)) Then
                Throw New InvalidOperationException("Press placement moved existing operation " & fixedSlot.RecordNo.ToString())
            End If
        Next
        slots.Insert(index, slot)
        savedSetupTimes(slot.RecordNo) = slot.SetupDays
    End Sub

    Private Shared Function Matches(expected As PressCampaignSlot, actual As OperationResourceTimes?) As Boolean
        Return actual.HasValue AndAlso actual.Value.Resource = expected.Resource AndAlso
            actual.Value.OperationTimes.ChangeStart = expected.Start AndAlso
            actual.Value.OperationTimes.ProcessStart = expected.ProcessStart AndAlso
            actual.Value.OperationTimes.ProcessEnd = expected.Finish
    End Function

    Friend Function Commit(slot As PressCampaignSlot, put As Action(Of Action)) As PressCampaignSlot
        Dim slots = timelines(slot.Resource)
        Dim index = WindowIndex(slots, slot.Start)
        If index < slots.Count Then board.SaveScheduleState(SnapshotSlot)
        Dim original = ReadSetupTime(slot.RecordNo)
        Dim verified = False
        Try
            If slot.SetupDays <> original Then WriteSetupTime(slot.RecordNo, slot.SetupDays)
            put(Sub()
                    VerifyAndRecord(slot)
                    verified = True
                End Sub)
            Return If(verified, slot, Nothing)
        Finally
            If Not verified Then RestorePlacement(slot.RecordNo, original, slots, index)
        End Try
    End Function

    Private Sub RestorePlacement(record As Integer, original As Double,
                                 slots As List(Of PressCampaignSlot), firstFixed As Integer)
        Try
            WriteSetupTime(record, original)
        Finally
            If firstFixed < slots.Count Then
                board.RestoreScheduleState(SnapshotSlot)
            ElseIf board.IsOperationScheduled(record) Then
                board.UnallocateOperation(record, OperationSelection.ThisOperation)
            End If
        End Try
        If board.IsOperationScheduled(record) OrElse preactor.ReadFieldDouble(ordersFormat, setupField, record) <> original Then
            Throw New InvalidOperationException("Could not verify rollback of pressing operation " & record.ToString())
        End If
        For index = firstFixed To slots.Count - 1
            Dim slot = slots(index)
            If Not Matches(slot, board.GetOperationTimes(slot.RecordNo)) Then
                Throw New InvalidOperationException("Could not restore existing operation " & slot.RecordNo.ToString())
            End If
        Next
    End Sub
End Class
