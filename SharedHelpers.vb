Option Strict On
Option Explicit On

Imports System
Imports System.Globalization
Imports System.IO
Imports System.Text
Imports Preactor

Public Module SharedHelpers

    Private ReadOnly MinimumOperationalDate As New DateTime(1900, 1, 1)
    Private Const OperationRowIndexPropertyName As String =
        "GN.SharedHelpers.OperationRowIndex"
    Private Const OperationRowIndexRowCountPropertyName As String =
        "GN.SharedHelpers.OperationRowIndex.RowCount"

    ' ========================================================================
    ' OPTIMIZER SETTINGS POINTER CATALOG
    '
    ' USER REPLACEMENT INSTRUCTIONS:
    '   1. Confirm the display name, record number, field type, units,
    '      conversion, valid range, fallback, and consumers in
    '      GetPointerDefinitions below.
    '   2. Change only the matching constants in OptimizerSettingsCatalog
    '      after the authoritative GN Optimizer Settings mapping is known.
    '   3. Record number 0 means that no authoritative dataset pointer exists
    '      yet. Those optimizer inputs intentionally remain hardcoded.
    '   4. Do not wire a record-0 placeholder into a live read merely because
    '      a likely record is discovered; validate it in Opcenter first.
    '
    ' Compatibility notes that must remain visible:
    '   - Batch records 3 and 4 are read today, but runFiring still passes
    '     True and 60 explicitly to firingOptimizer_vf.
    '   - Tunnel record 8 is read both as minimum preferred occupancy and as
    '     drying-to-firing buffer minutes divided by 60.
    ' ========================================================================
    Friend NotInheritable Class OptimizerSettingPointerDefinition
        Friend Sub New(key As String,
                       displayName As String,
                       recordNumber As Integer,
                       fieldType As String,
                       unitAndConversion As String,
                       validRange As String,
                       fallbackBehavior As String,
                       consumingActions As String,
                       currentEffectiveValue As String)

            Me.Key = key
            Me.DisplayName = displayName
            Me.RecordNumber = recordNumber
            Me.FieldType = fieldType
            Me.UnitAndConversion = unitAndConversion
            Me.ValidRange = validRange
            Me.FallbackBehavior = fallbackBehavior
            Me.ConsumingActions = consumingActions
            Me.CurrentEffectiveValue = currentEffectiveValue
        End Sub

        Friend ReadOnly Property Key As String
        Friend ReadOnly Property DisplayName As String
        Friend ReadOnly Property RecordNumber As Integer
        Friend ReadOnly Property FieldType As String
        Friend ReadOnly Property UnitAndConversion As String
        Friend ReadOnly Property ValidRange As String
        Friend ReadOnly Property FallbackBehavior As String
        Friend ReadOnly Property ConsumingActions As String
        Friend ReadOnly Property CurrentEffectiveValue As String
    End Class

    Friend NotInheritable Class OptimizerSettingsCatalog
        Private Sub New()
        End Sub

        Friend Const FormatName As String = "GN Optimizer Settings"
        Friend Const NumericFieldName As String = "Numeric Value"
        Friend Const ToggleFieldName As String = "Toggle Value"
        Friend Const ParameterFieldName As String = "Parameter"
        Friend Const DateFieldName As String = "Date Value"
        Friend Const AvailabilityParameterRecordNumber As Integer = 0
        Friend Const AvailabilityParameterSuffix As String = " Available From"

        Friend Const BatchMinOccupancyRecordNumber As Integer = 1
        Friend Const BatchMaxOccupancyRecordNumber As Integer = 2
        Friend Const BatchAllowUnderfilledTailRecordNumber As Integer = 3
        Friend Const BatchStartDelayMinutesRecordNumber As Integer = 4

        Friend Const TunnelCartsPerDayRecordNumber As Integer = 5
        Friend Const TunnelTotalCartsRecordNumber As Integer = 7
        Friend Const TunnelMinOccupancyPreferredRecordNumber As Integer = 8
        Friend Const TunnelDryingBufferMinutesRecordNumber As Integer = 8
        Friend Const TunnelMaxOccupancyRecordNumber As Integer = 9
        Friend Const MinutesPerHour As Double = 60.0

        ' These effective values are deliberately separate from pointer
        ' records. They preserve the arguments currently passed by AlgoSeq4.
        Friend Const BatchEffectiveAllowUnderfilledTail As Boolean = True
        Friend Const BatchEffectiveStartDelayMinutes As Integer = 60
        Friend Const BatchEffectiveMaxBatchesPerDay As Integer = 2
        Friend Const BatchKilnMatrixFileName As String = "kilndata.csv"

        Friend Const SwkMinTonnageRecordNumber As Integer = 0
        Friend Const SwkMaxTonnageRecordNumber As Integer = 0
        Friend Const SwkDailyBatchLimitRecordNumber As Integer = 0
        Friend Const SwkStartDelayMinutesRecordNumber As Integer = 0
        Friend Const SwkAllowUnderfilledTailRecordNumber As Integer = 0
        Friend Const SwkEffectiveMinTonnage As Double = 0.8
        Friend Const SwkEffectiveMaxTonnage As Double = 1.0
        Friend Const SwkEffectiveDailyBatchLimit As Integer = 2
        Friend Const SwkEffectiveStartDelayMinutes As Integer = 60
        Friend Const SwkEffectiveAllowUnderfilledTail As Boolean = True
        Friend Const SwkResourceName As String = "SWBKILN"

        Friend Const PressPrioritizePreviousOperationRecordNumber As Integer = 0
        Friend Const PressEffectivePrioritizePreviousOperation As Boolean = True
        Friend Const PressApproachingDaysRecordNumber As Integer = 0
        Friend Const PressEffectiveApproachingDays As Integer = 2

        ' This catalog is documentation and replacement guidance. It is not
        ' instantiated by scheduler hot paths.
        Friend Shared Function GetPointerDefinitions() _
            As OptimizerSettingPointerDefinition()

            Return New OptimizerSettingPointerDefinition() {
                New OptimizerSettingPointerDefinition(
                    "BatchMinOccupancy",
                    "TODO: confirm Batch Minimum Occupancy display name",
                    BatchMinOccupancyRecordNumber,
                    "Numeric Value / Double",
                    "Occupancy units; no conversion",
                    "> 0 and <= BatchMaxOccupancy",
                    "No fallback; the current COM read failure propagates",
                    "AlgoSeq4.runFiring -> firingOptimizer_vf",
                    "Live value from record 1"),
                New OptimizerSettingPointerDefinition(
                    "BatchMaxOccupancy",
                    "TODO: confirm Batch Maximum Occupancy display name",
                    BatchMaxOccupancyRecordNumber,
                    "Numeric Value / Double",
                    "Occupancy units; no conversion",
                    ">= BatchMinOccupancy",
                    "No fallback; the current COM read failure propagates",
                    "AlgoSeq4.runFiring -> firingOptimizer_vf",
                    "Live value from record 2"),
                New OptimizerSettingPointerDefinition(
                    "BatchAllowUnderfilledTailConfigured",
                    "TODO: confirm Batch Allow Underfilled Tail display name",
                    BatchAllowUnderfilledTailRecordNumber,
                    "Toggle Value / Integer interpreted as Boolean",
                    "1=True; every other value=False",
                    "Boolean",
                    "The read is retained, but the optimizer currently receives hardcoded True",
                    "AlgoSeq4.runFiring",
                    "True"),
                New OptimizerSettingPointerDefinition(
                    "BatchStartDelayMinutesConfigured",
                    "TODO: confirm Batch Start Delay display name",
                    BatchStartDelayMinutesRecordNumber,
                    "Numeric Value / Integer",
                    "Minutes; no conversion",
                    "Current optimizer clamps negative values to zero",
                    "The read is retained, but the optimizer currently receives hardcoded 60",
                    "AlgoSeq4.runFiring",
                    "60 minutes"),
                New OptimizerSettingPointerDefinition(
                    "BatchMaxBatchesPerDay",
                    "TODO: supply Batch Maximum Batches Per Day pointer",
                    0,
                    "TODO: Numeric Value / Integer",
                    "Batches per day",
                    "> 0",
                    "No live pointer; preserve hardcoded 2",
                    "AlgoSeq4.runFiring -> firingOptimizer_vf",
                    "2"),
                New OptimizerSettingPointerDefinition(
                    "BatchKilnMatrixFile",
                    "TODO: confirm whether the kiln matrix filename is configurable",
                    0,
                    "Hardcoded String",
                    "Filename relative to the Opcenter configuration path",
                    "Existing readable CSV filename",
                    "No live pointer; preserve kilndata.csv",
                    "AlgoSeq4.runFiring -> firingOptimizer_vf",
                    "kilndata.csv"),
                New OptimizerSettingPointerDefinition(
                    "ResourceAvailableFrom",
                    "Resource name plus ' Available From'",
                    AvailabilityParameterRecordNumber,
                    "Parameter / String plus Date Value / DateTime",
                    "Opcenter DateTime; no conversion",
                    "DateTime.MinValue or a valid Opcenter date",
                    "Lookup is by parameter name; missing/blank values fall back through current kiln availability logic",
                    "AlgoSeq4.runFiring, runSWKFiring, and runFiring2",
                    "Dynamic parameter-name lookup for each kiln resource"),
                New OptimizerSettingPointerDefinition(
                    "TunnelCartsPerDay",
                    "TODO: confirm Tunnel Carts Per Day display name",
                    TunnelCartsPerDayRecordNumber,
                    "Numeric Value / Double",
                    "Carts per day; no conversion",
                    "> 0",
                    "No fallback; the current COM read failure propagates",
                    "AlgoSeq4.runFiring2 -> tunnelOptimizer_vf",
                    "Live value from record 5"),
                New OptimizerSettingPointerDefinition(
                    "TunnelTotalCarts",
                    "TODO: confirm Tunnel Total Carts display name",
                    TunnelTotalCartsRecordNumber,
                    "Numeric Value / Integer",
                    "Cart count; no conversion",
                    "> 0",
                    "No fallback; the current COM read failure propagates",
                    "AlgoSeq4.runFiring2 -> tunnelOptimizer_vf",
                    "Live value from record 7"),
                New OptimizerSettingPointerDefinition(
                    "TunnelMinOccupancyPreferred",
                    "TODO: confirm Tunnel Minimum Preferred Occupancy display name",
                    TunnelMinOccupancyPreferredRecordNumber,
                    "Numeric Value / Double",
                    "Occupancy units; no conversion",
                    ">= 0 and <= TunnelMaxOccupancy",
                    "No fallback; the current COM read failure propagates",
                    "AlgoSeq4.runFiring2 -> tunnelOptimizer_vf",
                    "Live value from record 8"),
                New OptimizerSettingPointerDefinition(
                    "TunnelDryingToFiringBuffer",
                    "TODO: supply or confirm Tunnel Drying-to-Firing Buffer display name",
                    TunnelDryingBufferMinutesRecordNumber,
                    "Numeric Value / Double",
                    "Stored minutes divided by 60 to produce hours",
                    "TODO: confirm nonnegative range",
                    "No fallback; record 8 is intentionally shared with minimum occupancy",
                    "AlgoSeq4.runFiring2 -> tunnelOptimizer_vf",
                    "Record 8 divided by 60"),
                New OptimizerSettingPointerDefinition(
                    "TunnelMaxOccupancy",
                    "TODO: confirm Tunnel Maximum Occupancy display name",
                    TunnelMaxOccupancyRecordNumber,
                    "Numeric Value / Double",
                    "Occupancy units; no conversion",
                    "> 0 and >= TunnelMinOccupancyPreferred",
                    "No fallback; the current COM read failure propagates",
                    "AlgoSeq4.runFiring2 -> tunnelOptimizer_vf",
                    "Live value from record 9"),
                New OptimizerSettingPointerDefinition(
                    "SwkMinTonnage",
                    "TODO: supply SWK Minimum Tonnage pointer",
                    SwkMinTonnageRecordNumber,
                    "TODO: Numeric Value / Double",
                    "Tonnage; no conversion",
                    "> 0 and <= SWK maximum tonnage",
                    "No live pointer; preserve hardcoded 0.8",
                    "AlgoSeq4.runSWKFiring -> swkOptimizer_vf",
                    "0.8"),
                New OptimizerSettingPointerDefinition(
                    "SwkMaxTonnage",
                    "TODO: supply SWK Maximum Tonnage pointer",
                    SwkMaxTonnageRecordNumber,
                    "TODO: Numeric Value / Double",
                    "Tonnage; no conversion",
                    "> 0 and >= SWK minimum tonnage",
                    "No live pointer; preserve hardcoded 1.0",
                    "AlgoSeq4.runSWKFiring -> swkOptimizer_vf",
                    "1.0"),
                New OptimizerSettingPointerDefinition(
                    "SwkDailyBatchLimit",
                    "TODO: supply SWK Daily Batch Limit pointer",
                    SwkDailyBatchLimitRecordNumber,
                    "TODO: Numeric Value / Integer",
                    "Batches per day",
                    "> 0",
                    "No live pointer; preserve hardcoded 2",
                    "AlgoSeq4.runSWKFiring -> swkOptimizer_vf",
                    "2"),
                New OptimizerSettingPointerDefinition(
                    "SwkBatchStartDelayMinutes",
                    "TODO: supply SWK Batch Start Delay pointer",
                    SwkStartDelayMinutesRecordNumber,
                    "TODO: Numeric Value / Integer",
                    "Minutes; no conversion",
                    ">= 0",
                    "No live pointer; preserve hardcoded 60",
                    "AlgoSeq4.runSWKFiring -> swkOptimizer_vf",
                    "60 minutes"),
                New OptimizerSettingPointerDefinition(
                    "SwkAllowUnderfilledTail",
                    "TODO: supply SWK Allow Underfilled Tail pointer",
                    SwkAllowUnderfilledTailRecordNumber,
                    "TODO: Toggle Value / Boolean",
                    "Boolean",
                    "Boolean",
                    "No live pointer; preserve hardcoded True",
                    "AlgoSeq4.runSWKFiring -> swkOptimizer_vf",
                    "True"),
                New OptimizerSettingPointerDefinition(
                    "SwkResourceName",
                    "TODO: confirm whether the SWK resource name is configurable",
                    0,
                    "Hardcoded String",
                    "Opcenter resource name",
                    "Existing SWK resource",
                    "No live pointer; preserve SWBKILN",
                    "AlgoSeq4.runSWKFiring -> swkOptimizer_vf",
                    "SWBKILN"),
                New OptimizerSettingPointerDefinition(
                    "PressPrioritizePreviousOperation",
                    "TODO: supply Press Prioritize Previous Operation pointer",
                    PressPrioritizePreviousOperationRecordNumber,
                    "TODO: Toggle Value / Boolean",
                    "Boolean",
                    "Boolean",
                    "No live pointer; preserve hardcoded True",
                    "AlgoSeq4.runPress -> pressingOptimizer_vf",
                    "True"),
                New OptimizerSettingPointerDefinition(
                    "PressApproachingDays",
                    "TODO: supply Press Approaching Days pointer",
                    PressApproachingDaysRecordNumber,
                    "TODO: Numeric Value / Integer",
                    "Calendar days; no conversion",
                    "TODO: confirm nonnegative range",
                    "No live pointer; preserve effective default 2",
                    "AlgoSeq4.untilPress and AlgoSeq4.runPress",
                    "2 days")
            }
        End Function
    End Class

    ' Numeric project dates are interpreted deterministically. Day-first
    ' formats intentionally precede US slash formats so ambiguous values such
    ' as 07/12/2026 mean 7 December, matching the Opcenter data convention.
    Private ReadOnly KnownDateFormats As String() = {
        "dd-MM-yyyy HH:mm:ss",
        "d-M-yyyy H:mm:ss",
        "dd-MM-yyyy H:mm:ss",
        "d-M-yyyy HH:mm:ss",
        "dd-MM-yyyy HH:mm:ss.FFFFFFF",
        "d-M-yyyy H:mm:ss.FFFFFFF",
        "dd-MM-yyyy",
        "d-M-yyyy",
        "dd/MM/yyyy HH:mm:ss",
        "d/M/yyyy H:mm:ss",
        "dd/MM/yyyy H:mm:ss",
        "d/M/yyyy HH:mm:ss",
        "dd/MM/yyyy HH:mm:ss.FFFFFFF",
        "d/M/yyyy H:mm:ss.FFFFFFF",
        "dd/MM/yyyy",
        "d/M/yyyy",
        "yyyy-MM-dd HH:mm:ss",
        "yyyy-M-d H:mm:ss",
        "yyyy-MM-dd HH:mm:ss.FFFFFFF",
        "yyyy-M-d H:mm:ss.FFFFFFF",
        "yyyy-MM-dd'T'HH:mm:ss",
        "yyyy-MM-dd'T'HH:mm:ss.FFFFFFF",
        "yyyy-MM-dd",
        "yyyy-M-d",
        "MM/dd/yyyy HH:mm:ss",
        "M/d/yyyy H:mm:ss",
        "MM/dd/yyyy",
        "M/d/yyyy"
    }

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

    'Public Function SafeDate(o As Object) As DateTime
    '    If o Is Nothing Then Return DateTime.MinValue
    '    If TypeOf o Is DateTime Then Return CType(o, DateTime)
    '    Dim d As DateTime
    '    If DateTime.TryParse(o.ToString(), CultureInfo.InvariantCulture, DateTimeStyles.None, d) Then Return d
    '    Return DateTime.MinValue
    'End Function
    Public Function SafeDate(o As Object) As DateTime
        Dim parsed As DateTime
        If TryParseDateValue(o, parsed) Then Return parsed
        Return DateTime.MinValue
    End Function

    Public Function TryParseDateValue(value As Object,
                                      ByRef parsed As DateTime) As Boolean
        parsed = DateTime.MinValue

        If value Is Nothing OrElse Object.ReferenceEquals(value, DBNull.Value) Then
            Return False
        End If

        If TypeOf value Is DateTime Then
            parsed = DirectCast(value, DateTime)
            Return parsed > MinimumOperationalDate
        End If

        If TypeOf value Is DateTimeOffset Then
            parsed = DirectCast(value, DateTimeOffset).DateTime
            Return parsed > MinimumOperationalDate
        End If

        Dim s As String = Convert.ToString(value, CultureInfo.InvariantCulture).Trim()
        If s.Length = 0 Then Return False

        If DateTime.TryParseExact(s,
                                  KnownDateFormats,
                                  CultureInfo.InvariantCulture,
                                  DateTimeStyles.AllowWhiteSpaces,
                                  parsed) Then
            Return parsed > MinimumOperationalDate
        End If

        ' Do not let the machine culture reinterpret an unrecognized numeric
        ' date by swapping its day and month.
        If LooksLikeNumericDate(s) Then Return False

        If DateTime.TryParse(s,
                             CultureInfo.CurrentCulture,
                             DateTimeStyles.AllowWhiteSpaces,
                             parsed) Then
            Return parsed > MinimumOperationalDate
        End If

        If DateTime.TryParse(s,
                             CultureInfo.InvariantCulture,
                             DateTimeStyles.AllowWhiteSpaces,
                             parsed) Then
            Return parsed > MinimumOperationalDate
        End If

        Return False
    End Function

    Private Function LooksLikeNumericDate(value As String) As Boolean
        Dim datePart As String = value
        Dim separatorIndex As Integer = value.IndexOfAny(New Char() {" "c, "T"c})
        If separatorIndex >= 0 Then datePart = value.Substring(0, separatorIndex)

        If datePart.Length = 0 Then Return False

        For Each character As Char In datePart
            If Not Char.IsDigit(character) AndAlso
               character <> "-"c AndAlso
               character <> "/"c AndAlso
               character <> "."c Then
                Return False
            End If
        Next

        Return datePart.IndexOf("-"c) >= 0 OrElse
               datePart.IndexOf("/"c) >= 0 OrElse
               datePart.IndexOf("."c) >= 0
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

    Private Function TryGetFieldNumber(preactor As IPreactor,
                                       formatNo As Integer,
                                       fieldName As String,
                                       lookupCache As SchedulerRunLookupCache) As Integer
        If preactor Is Nothing Then Return 0
        If formatNo <= 0 Then Return 0
        If String.IsNullOrWhiteSpace(fieldName) Then Return 0

        Try
            Return lookupCache.GetFieldNumber(preactor, formatNo, fieldName)
        Catch
            Return 0
        End Try
    End Function

    Private Function ResolveFirstExistingField(preactor As IPreactor,
                                               formatNo As Integer,
                                               fieldNames As String(),
                                               lookupCache As SchedulerRunLookupCache) As Integer
        If fieldNames Is Nothing Then Return 0

        For Each fieldName As String In fieldNames
            Dim fieldNo As Integer =
                TryGetFieldNumber(preactor,
                                  formatNo,
                                  fieldName,
                                  lookupCache)
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

    'Public Function ParseDueAsEndOfDay(o As Object) As DateTime
    '    Dim s As String = SafeStr(o).Trim()
    '    If s = "" Then Return DateTime.MinValue

    '    Dim d As DateTime
    '    If DateTime.TryParseExact(s,
    '                              "dd-MM-yyyy",
    '                              CultureInfo.InvariantCulture,
    '                              DateTimeStyles.None,
    '                              d) Then
    '        Return d.Date.AddDays(1).AddTicks(-1) ' end of day
    '    End If

    '    If DateTime.TryParse(s, CultureInfo.InvariantCulture, DateTimeStyles.None, d) Then
    '        Return d.Date.AddDays(1).AddTicks(-1)
    '    End If

    '    Return DateTime.MinValue
    'End Function
    Public Function ParseDueAsEndOfDay(o As Object) As DateTime
        Dim parsed As DateTime = SafeDate(o)
        If parsed = DateTime.MinValue Then Return DateTime.MinValue
        Return parsed.Date.AddDays(1).AddTicks(-1)
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
        Return ReadOrderTableCore(preactor, Nothing)
    End Function

    Friend Function ReadOrderTableWithCache(preactor As IPreactor,
                                            lookupCache As SchedulerRunLookupCache) As DataTable
        Return ReadOrderTableCore(preactor, lookupCache)
    End Function

    Private Function ReadOrderTableCore(preactor As IPreactor,
                                        lookupCache As SchedulerRunLookupCache) As DataTable
        Dim planningboard As IPlanningBoard = preactor.PlanningBoard
        Dim dt As DataTable = BuildRoutingSchema()
        Dim cache As SchedulerRunLookupCache = lookupCache
        If cache Is Nothing Then cache = New SchedulerRunLookupCache()

        ' Snapshot COM calls are represented by RoutingSnapshotMilliseconds.
        ' Temporarily suppress action-traversal counters while retaining values
        ' in the same invocation-local cache for later scheduler reuse.
        Dim actionMetrics As SchedulerActionMetricsRow = cache.Metrics
        cache.Metrics = Nothing

        Try
            ' Suspend indexing, events, and constraints for bulk insert performance.
            dt.BeginLoadData()
            Try
                Dim ordersTable As Integer =
                    cache.GetFormatNumber(preactor, "Orders")
                Dim orderNo As Integer =
                    cache.GetFieldNumber(preactor, ordersTable, "Order No.")
                Dim partNo As Integer =
                    cache.GetFieldNumber(preactor, ordersTable, "Part No.")
                Dim product As Integer =
                    cache.GetFieldNumber(preactor, ordersTable, "Product")
                Dim opNo As Integer =
                    cache.GetFieldNumber(preactor, ordersTable, "Op. No.")
                Dim opName As Integer =
                    cache.GetFieldNumber(preactor, ordersTable, "Operation Name")
                Dim resGroup As Integer =
                    cache.GetFieldNumber(preactor, ordersTable, "Resource Group")
                Dim res As Integer =
                    cache.GetFieldNumber(preactor, ordersTable, "Required Resource")
                Dim stpTime As Integer =
                    cache.GetFieldNumber(preactor, ordersTable, "Setup Time")
                Dim timePerItem As Integer =
                    cache.GetFieldNumber(preactor, ordersTable, "Op. Time per Item")
                Dim salesOrder As Integer =
                    cache.GetFieldNumber(preactor, ordersTable, "Operation Name")
                Dim qty As Integer =
                    cache.GetFieldNumber(preactor, ordersTable, "Quantity")
                Dim dueDate As Integer =
                    cache.GetFieldNumber(preactor, ordersTable, "Due Date")
                Dim batchTime As Integer =
                    cache.GetFieldNumber(preactor, ordersTable, "Batch Time")
                Dim prsTimeType As Integer =
                    cache.GetFieldNumber(preactor, ordersTable, "Process Time Type")
                Dim tonnage As Integer =
                    cache.GetFieldNumber(preactor, ordersTable, "Numerical Attribute 4")
                Dim cycleType As Integer =
                    cache.GetFieldNumber(preactor, ordersTable, "Table Attribute 2")
                Dim klnType As Integer =
                    cache.GetFieldNumber(preactor, ordersTable, "Table Attribute 3")
                Dim volumeOcc As Integer =
                    cache.GetFieldNumber(preactor, ordersTable, "Numerical Attribute 5")
                Dim presEarlyStart As Integer =
                    cache.GetFieldNumber(preactor, ordersTable, "Date Attribute 1")
                Dim presDue As Integer =
                    cache.GetFieldNumber(preactor, ordersTable, "Date Attribute 2")
                Dim firingDue As Integer =
                    cache.GetFieldNumber(preactor, ordersTable, "Date Attribute 3")
                Dim mts As Integer =
                    cache.GetFieldNumber(preactor, ordersTable, "Table Attribute 1")
                Dim wheelDia As Integer =
                    cache.GetFieldNumber(preactor, ordersTable, "String Attribute 5")
                Dim wheelThck As Integer =
                    cache.GetFieldNumber(preactor, ordersTable, "String Attribute 4")
                Dim wheelPin As Integer =
                    cache.GetFieldNumber(preactor, ordersTable, "String Attribute 3")
                Dim schStart As Integer =
                    cache.GetFieldNumber(preactor, ordersTable, "Start Time")
                Dim schEnd As Integer =
                    cache.GetFieldNumber(preactor, ordersTable, "End Time")

                ' Toggle Attribute 1 remains the source is_completed flag.
                Dim sourceCompletedField As Integer =
                    TryGetFieldNumber(preactor,
                                      ordersTable,
                                      "Toggle Attribute 1",
                                      cache)

                Dim useActualField As Integer =
                    ResolveFirstExistingField(preactor,
                                              ordersTable,
                                              New String() {
                                                  "Use Actual",
                                                  "Use Actual Times",
                                                  "USE ACTUAL TIMES",
                                                  "Use actual"
                                              },
                                              cache)

                Dim actualStartField As Integer =
                    ResolveFirstExistingField(preactor,
                                              ordersTable,
                                              New String() {
                                                  "Actual Start Time",
                                                  "Actual Start",
                                                  "Actual Start Date",
                                                  "Start Time Actual"
                                              },
                                              cache)

                Dim actualEndField As Integer =
                    ResolveFirstExistingField(preactor,
                                              ordersTable,
                                              New String() {
                                                  "Actual End Time",
                                                  "Actual End",
                                                  "Actual Finish Time",
                                                  "Actual End Date",
                                                  "End Time Actual"
                                              },
                                              cache)

                Dim rowCount As Integer = preactor.RecordCount(ordersTable)
                Dim scheduledByRecord(rowCount) As Boolean

                ' Capture scheduled state exactly once per normal Orders record.
                For rec As Integer = 1 To rowCount
                    scheduledByRecord(rec) =
                        cache.IsOperationScheduled(planningboard, rec)
                Next

                Dim parentRecordByOrderNo As New Dictionary(Of String, Integer)(
                    StringComparer.OrdinalIgnoreCase)
                Dim operationRows As New Dictionary(Of Integer, DataRow)(rowCount)

                For rec As Integer = 1 To rowCount
                    Dim r As DataRow = dt.NewRow()
                    Dim currentOpNo As Integer =
                        cache.ReadOperationNumber(preactor,
                                                  ordersTable,
                                                  opNo,
                                                  rec)
                    Dim isScheduled As Boolean = scheduledByRecord(rec)
                    Dim currentOrderNo As String =
                        preactor.ReadFieldString(ordersTable, orderNo, rec).Trim()

                    r("OrdersID") = rec
                    r("Order No") = currentOrderNo
                    r("Operation Number") = currentOpNo
                    r("Operation Name") = preactor.ReadFieldString(ordersTable, opName, rec)
                    r("Resource Group") = preactor.ReadFieldString(ordersTable, resGroup, rec)
                    r("Required Resource") = preactor.ReadFieldString(ordersTable, res, rec)
                    r("Quantity") = preactor.ReadFieldInt(ordersTable, qty, rec)
                    r("Due Date") = preactor.ReadFieldDateTime(ordersTable, dueDate, rec)
                    r("Batch Time") = preactor.ReadFieldDouble(ordersTable, batchTime, rec) * 1440
                    r("Tonnage") = preactor.ReadFieldDouble(ordersTable, tonnage, rec)
                    r("Cycle Type") = preactor.ReadFieldString(ordersTable, cycleType, rec)
                    r("Volume Occupancy") = preactor.ReadFieldDouble(ordersTable, volumeOcc, rec)
                    r("Kiln Type") = preactor.ReadFieldInt(ordersTable, klnType, rec)
                    r("MTS/MTO") = preactor.ReadFieldInt(ordersTable, mts, rec)
                    r("Wheel Dia") = preactor.ReadFieldString(ordersTable, wheelDia, rec)
                    r("Wheel thickness") = preactor.ReadFieldString(ordersTable, wheelThck, rec)
                    r("Pressing earliest start") =
                        preactor.ReadFieldDateTime(ordersTable, presEarlyStart, rec)
                    r("Pressing Due date") =
                        preactor.ReadFieldDateTime(ordersTable, presDue, rec)
                    r("firing due date") =
                        preactor.ReadFieldDateTime(ordersTable, firingDue, rec)
                    r("is_scheduled") = isScheduled

                    If isScheduled Then
                        Dim liveTimes As Nullable(Of Preactor.OperationResourceTimes) = Nothing
                        Try
                            liveTimes = cache.GetOperationTimes(planningboard, rec)
                        Catch
                            liveTimes = Nothing
                        End Try

                        If liveTimes.HasValue Then
                            r("scheduled_start_time") =
                                liveTimes.Value.OperationTimes.ProcessStart
                            r("scheduled_end_time") =
                                liveTimes.Value.OperationTimes.ProcessEnd
                        Else
                            r("scheduled_start_time") =
                                preactor.ReadFieldDateTime(ordersTable, schStart, rec)
                            r("scheduled_end_time") =
                                preactor.ReadFieldDateTime(ordersTable, schEnd, rec)
                        End If
                    End If

                    Dim sourceCompleted As Boolean =
                        ReadBoolField(preactor,
                                      ordersTable,
                                      sourceCompletedField,
                                      rec)
                    Dim useActual As Boolean =
                        ReadBoolField(preactor,
                                      ordersTable,
                                      useActualField,
                                      rec)
                    Dim actualStartValue As DateTime =
                        ReadDateField(preactor,
                                      ordersTable,
                                      actualStartField,
                                      rec)
                    Dim actualEndValue As DateTime =
                        ReadDateField(preactor,
                                      ordersTable,
                                      actualEndField,
                                      rec)

                    If actualEndValue = DateTime.MinValue AndAlso
                       useActual AndAlso
                       schEnd > 0 Then

                        actualEndValue =
                            ReadDateField(preactor, ordersTable, schEnd, rec)
                    End If

                    If actualStartValue = DateTime.MinValue AndAlso
                       useActual AndAlso
                       schStart > 0 Then

                        actualStartValue =
                            ReadDateField(preactor, ordersTable, schStart, rec)
                    End If

                    r("source_is_completed") = sourceCompleted
                    r("opcenter_use_actual") = useActual
                    If actualStartValue <> DateTime.MinValue Then
                        r("actual_start_time") = actualStartValue
                    End If
                    If actualEndValue <> DateTime.MinValue Then
                        r("actual_end_time") = actualEndValue
                    End If

                    Dim parentRecord As Integer
                    If currentOrderNo.Length = 0 Then
                        parentRecord = rec
                    ElseIf Not parentRecordByOrderNo.TryGetValue(currentOrderNo,
                                                                  parentRecord) Then
                        parentRecord = rec
                        parentRecordByOrderNo.Add(currentOrderNo, parentRecord)
                    End If
                    r("parent_record") = parentRecord

                    Try
                        Dim prevOpRec As Integer =
                            cache.GetPreviousOperation(planningboard, rec, 1)

                        If prevOpRec > 0 Then
                            If prevOpRec <= rowCount Then
                                r("prev_op_is_scheduled") =
                                    scheduledByRecord(prevOpRec)
                            Else
                                r("prev_op_is_scheduled") =
                                    cache.IsOperationScheduled(planningboard,
                                                               prevOpRec)
                            End If
                        End If
                    Catch
                        r("prev_op_is_scheduled") = False
                    End Try

                    dt.Rows.Add(r)
                    If Not operationRows.ContainsKey(rec) Then
                        operationRows.Add(rec, r)
                    End If
                Next

                SetOperationRowIndex(dt, operationRows)
            Finally
                dt.EndLoadData()
            End Try

            Dim terminatorTime As DateTime = planningboard.TerminatorTime
            PopulateWipColumns(dt, planningboard, terminatorTime)
            Return dt
        Finally
            cache.Metrics = actionMetrics
        End Try
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

    Public Function GetEffectiveParentRecord(row As DataRow) As Integer

        Dim parentRecord As Integer = SafeInt(row("parent_record"))
        If parentRecord > 0 Then Return parentRecord

        Return SafeInt(row("OrdersID"))

    End Function

    Public Class TunnelReleasedContinuationInfo
        Public Property GroupKey As String
        Public Property ParentRecord As Integer
        Public Property OrderNo As String
        Public Property OpRec As Integer
        Public Property OpNo As Integer
        Public Property ReleaseTime As DateTime
        Public Property StartStageIndex As Integer
    End Class

    Public Class ReleaseBoundaryInfo
        Public Property GroupKey As String
        Public Property ParentRecord As Integer
        Public Property OrderNo As String
        Public Property OpRec As Integer
        Public Property OpNo As Integer
        Public Property ReleaseTime As DateTime
        Public Property WipScore As Integer
    End Class

    Public Function BuildTunnelOrderLookup(dt As DataTable) As Dictionary(Of String, Boolean)

        Dim result As New Dictionary(Of String, Boolean)(StringComparer.OrdinalIgnoreCase)
        If dt Is Nothing Then Return result

        For Each row As DataRow In dt.Rows
            If SafeInt(row("Kiln Type")) <> 2 Then Continue For

            For Each key As String In GetTunnelOrderKeys(row)
                If Not result.ContainsKey(key) Then result.Add(key, True)
            Next
        Next

        Return result

    End Function

    Public Function IsTunnelOrderRow(row As DataRow,
                                     tunnelOrderLookup As IDictionary(Of String, Boolean)) As Boolean

        If row Is Nothing OrElse tunnelOrderLookup Is Nothing Then Return False

        For Each key As String In GetTunnelOrderKeys(row)
            If tunnelOrderLookup.ContainsKey(key) Then Return True
        Next

        Return False

    End Function

    Public Function BuildLatestReleaseBoundaries(dt As DataTable) _
        As List(Of ReleaseBoundaryInfo)

        Dim result As New List(Of ReleaseBoundaryInfo)()
        If dt Is Nothing Then Return result

        Dim latestByGroup As New Dictionary(Of String, ReleaseBoundaryInfo)(
            StringComparer.OrdinalIgnoreCase)

        For Each row As DataRow In dt.Rows

            If Not SafeBool(row("operation_releases_next")) Then Continue For

            Dim opRec As Integer = SafeInt(row("OrdersID"))
            Dim opNo As Integer = SafeInt(row("Operation Number"))
            Dim releaseTime As DateTime = SafeDate(row("operation_release_time"))
            If opRec <= 0 OrElse opNo <= 0 Then Continue For
            If releaseTime = DateTime.MinValue Then Continue For

            Dim groupKey As String = GetPreferredReleaseBoundaryKey(row)
            If groupKey = "" Then Continue For

            Dim candidate As New ReleaseBoundaryInfo With {
                .GroupKey = groupKey,
                .ParentRecord = GetEffectiveParentRecord(row),
                .OrderNo = SafeStr(row("Order No")).Trim(),
                .OpRec = opRec,
                .OpNo = opNo,
                .ReleaseTime = releaseTime,
                .WipScore = SafeInt(row("wip_score"))
            }

            Dim existing As ReleaseBoundaryInfo = Nothing
            If latestByGroup.TryGetValue(groupKey, existing) Then
                If candidate.OpNo < existing.OpNo Then Continue For
                If candidate.OpNo = existing.OpNo AndAlso
                   candidate.ReleaseTime <= existing.ReleaseTime Then Continue For
            End If

            latestByGroup(groupKey) = candidate

        Next

        result.AddRange(latestByGroup.Values)
        Return result

    End Function

    Public Function BuildTunnelReleasedContinuations(dt As DataTable) _
        As List(Of TunnelReleasedContinuationInfo)

        Dim result As New List(Of TunnelReleasedContinuationInfo)()
        If dt Is Nothing Then Return result

        Dim tunnelOrderLookup As Dictionary(Of String, Boolean) =
            BuildTunnelOrderLookup(dt)

        Dim latestByGroup As New Dictionary(Of String, TunnelReleasedContinuationInfo)(
            StringComparer.OrdinalIgnoreCase)

        For Each row As DataRow In dt.Rows

            If Not IsTunnelOrderRow(row, tunnelOrderLookup) Then Continue For
            If Not SafeBool(row("operation_releases_next")) Then Continue For

            Dim opNo As Integer = SafeInt(row("Operation Number"))
            Dim startStageIndex As Integer = GetTunnelContinuationStartStageIndex(opNo)
            If startStageIndex < 0 Then Continue For

            Dim releaseTime As DateTime = SafeDate(row("operation_release_time"))
            If releaseTime = DateTime.MinValue Then Continue For

            Dim groupKey As String = GetPreferredTunnelOrderKey(row)
            If groupKey = "" Then Continue For

            Dim candidate As New TunnelReleasedContinuationInfo With {
                .GroupKey = groupKey,
                .ParentRecord = GetEffectiveParentRecord(row),
                .OrderNo = SafeStr(row("Order No")).Trim(),
                .OpRec = SafeInt(row("OrdersID")),
                .OpNo = opNo,
                .ReleaseTime = releaseTime,
                .StartStageIndex = startStageIndex
            }

            Dim existing As TunnelReleasedContinuationInfo = Nothing
            If latestByGroup.TryGetValue(groupKey, existing) Then
                If candidate.OpNo < existing.OpNo Then Continue For
                If candidate.OpNo = existing.OpNo AndAlso
                   candidate.ReleaseTime <= existing.ReleaseTime Then Continue For
            End If

            latestByGroup(groupKey) = candidate

        Next

        result.AddRange(latestByGroup.Values)
        Return result

    End Function

    Public Function GetTunnelContinuationStartStageIndex(opNo As Integer) As Integer

        Select Case opNo
            Case 300
                Return 0
            Case 310
                Return 1
            Case 320
                Return 2
            Case 390
                Return 3
            Case Else
                Return -1
        End Select

    End Function

    Private Function GetPreferredTunnelOrderKey(row As DataRow) As String

        Dim parentRecord As Integer = GetEffectiveParentRecord(row)
        If parentRecord > 0 Then
            Return "P:" & parentRecord.ToString(CultureInfo.InvariantCulture)
        End If

        Dim orderNo As String = SafeStr(row("Order No")).Trim()
        If orderNo <> "" Then Return "O:" & orderNo

        Dim opRec As Integer = SafeInt(row("OrdersID"))
        If opRec > 0 Then Return "R:" & opRec.ToString(CultureInfo.InvariantCulture)

        Return ""

    End Function

    Private Function GetPreferredReleaseBoundaryKey(row As DataRow) As String

        Dim parentRecord As Integer = GetEffectiveParentRecord(row)
        If parentRecord > 0 Then
            Return "P:" & parentRecord.ToString(CultureInfo.InvariantCulture)
        End If

        Dim orderNo As String = SafeStr(row("Order No")).Trim()
        If orderNo <> "" Then Return "O:" & orderNo

        Dim opRec As Integer = SafeInt(row("OrdersID"))
        If opRec > 0 Then Return "R:" & opRec.ToString(CultureInfo.InvariantCulture)

        Return ""

    End Function

    Private Function GetTunnelOrderKeys(row As DataRow) As List(Of String)

        Dim result As New List(Of String)()
        If row Is Nothing Then Return result

        Dim parentRecord As Integer = GetEffectiveParentRecord(row)
        If parentRecord > 0 Then
            result.Add("P:" & parentRecord.ToString(CultureInfo.InvariantCulture))
        End If

        Dim orderNo As String = SafeStr(row("Order No")).Trim()
        If orderNo <> "" Then result.Add("O:" & orderNo)

        Dim opRec As Integer = SafeInt(row("OrdersID"))
        If opRec > 0 AndAlso result.Count = 0 Then
            result.Add("R:" & opRec.ToString(CultureInfo.InvariantCulture))
        End If

        Return result

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

    Public Function BuildOperationRowIndex(routingDt As DataTable) As Dictionary(Of Integer, DataRow)
        Dim result As New Dictionary(Of Integer, DataRow)()
        If routingDt Is Nothing Then Return result

        Dim cached As Dictionary(Of Integer, DataRow) =
            TryCast(routingDt.ExtendedProperties(OperationRowIndexPropertyName),
                    Dictionary(Of Integer, DataRow))
        Dim cachedRowCount As Integer = -1
        Dim cachedCountValue As Object =
            routingDt.ExtendedProperties(OperationRowIndexRowCountPropertyName)

        If cachedCountValue IsNot Nothing Then
            Integer.TryParse(cachedCountValue.ToString(), cachedRowCount)
        End If

        Dim cacheBelongsToTable As Boolean =
            cached IsNot Nothing AndAlso cached.Count = 0
        If cached IsNot Nothing AndAlso cached.Count > 0 Then
            For Each cachedRow As DataRow In cached.Values
                cacheBelongsToTable =
                    Object.ReferenceEquals(cachedRow.Table, routingDt)
                Exit For
            Next
        End If

        If cached IsNot Nothing AndAlso
           cacheBelongsToTable AndAlso
           cachedRowCount = routingDt.Rows.Count Then

            Return cached
        End If

        For Each row As DataRow In routingDt.Rows
            Dim opRec As Integer = SafeInt(row("OrdersID"))
            If opRec > 0 AndAlso Not result.ContainsKey(opRec) Then
                result.Add(opRec, row)
            End If
        Next

        SetOperationRowIndex(routingDt, result)
        Return result
    End Function

    Private Sub SetOperationRowIndex(routingDt As DataTable,
                                     operationRows As Dictionary(Of Integer, DataRow))
        If routingDt Is Nothing OrElse operationRows Is Nothing Then Return

        routingDt.ExtendedProperties(OperationRowIndexPropertyName) =
            operationRows
        routingDt.ExtendedProperties(OperationRowIndexRowCountPropertyName) =
            routingDt.Rows.Count
    End Sub

    Public Function IsCompletedOrActualizedOp(operationRows As IDictionary(Of Integer, DataRow),
                                               opRec As Integer) As Boolean
        If operationRows Is Nothing OrElse opRec <= 0 Then Return False

        Dim row As DataRow = Nothing
        If operationRows.TryGetValue(opRec, row) Then
            Return IsCompletedOrActualizedRow(row)
        End If

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

    Public Function GetOperationReleaseTime(operationRows As IDictionary(Of Integer, DataRow),
                                            opRec As Integer) As DateTime
        If operationRows Is Nothing OrElse opRec <= 0 Then Return DateTime.MinValue

        Dim row As DataRow = Nothing
        If operationRows.TryGetValue(opRec, row) Then
            Return SafeDate(row("operation_release_time"))
        End If

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

    Public Function GetResourceLastScheduledEnds(
                                                preactor As IPreactor,
                                                planningboard As IPlanningBoard,
                                                resourceRecs As IEnumerable(Of Integer)) As Dictionary(Of Integer, DateTime)

        Dim result As New Dictionary(Of Integer, DateTime)()
        Dim requestedResources As New HashSet(Of Integer)(
            resourceRecs.Where(Function(resourceRec) resourceRec > 0))

        If requestedResources.Count = 0 Then Return result

        Dim ordersFmt As Integer = preactor.GetFormatNumber("Orders")
        Dim reqResFieldNo As Integer = preactor.GetFieldNumber(ordersFmt, "Resource")
        Dim recordCount As Integer = preactor.RecordCount(ordersFmt)

        For opRec As Integer = 1 To recordCount
            If Not planningboard.IsOperationScheduled(opRec) Then Continue For

            Dim opResRec As Integer =
                preactor.ReadFieldInt(ordersFmt, reqResFieldNo, opRec)
            If Not requestedResources.Contains(opResRec) Then Continue For

            Dim times As Nullable(Of Preactor.OperationResourceTimes) =
                planningboard.GetOperationTimes(opRec)
            If Not times.HasValue Then Continue For

            Dim opEnd As DateTime = times.Value.OperationTimes.ProcessEnd
            Dim lastEnd As DateTime
            If Not result.TryGetValue(opResRec, lastEnd) OrElse opEnd > lastEnd Then
                result(opResRec) = opEnd
            End If
        Next

        Return result
    End Function

    Public Function BuildEffectiveStartByResource(preactor As IPreactor,
                                               planningboard As IPlanningBoard,
                                               resourceNames As IEnumerable(Of String),
                                               Optional metadataDates As Dictionary(Of String, DateTime) = Nothing) As Dictionary(Of String, DateTime)

        Dim result As New Dictionary(Of String, DateTime)(StringComparer.OrdinalIgnoreCase)
        Dim names As List(Of String) = resourceNames.ToList()
        Dim resourceRecByName As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)

        For Each resourceName As String In names
            If String.IsNullOrWhiteSpace(resourceName) Then Continue For

            Dim resourceRec As Integer = planningboard.GetResourceNumber(resourceName)
            If resourceRec <= 0 Then
                Throw New Exception("Resource not found: " & resourceName)
            End If
            resourceRecByName(resourceName) = resourceRec
        Next

        Dim lastEndByResource As Dictionary(Of Integer, DateTime) =
            GetResourceLastScheduledEnds(preactor,
                                         planningboard,
                                         resourceRecByName.Values)

        For Each resourceName As String In names

            If String.IsNullOrWhiteSpace(resourceName) Then Continue For

            Dim metadataDate As DateTime = DateTime.MinValue

            If metadataDates IsNot Nothing Then
                metadataDates.TryGetValue(resourceName, metadataDate)
            End If

            Dim lastScheduledEnd As DateTime = DateTime.MinValue
            lastEndByResource.TryGetValue(resourceRecByName(resourceName), lastScheduledEnd)

            Dim effective As DateTime =
                MaxDate(planningboard.TerminatorTime, metadataDate, lastScheduledEnd)
            result(resourceName) = effective

            System.Diagnostics.Debug.WriteLine(
                "Effective Resource Start | Resource=" & resourceName &
                " | Terminator=" & FormatDateOrBlank(planningboard.TerminatorTime) &
                " | Metadata=" & FormatDateOrBlank(metadataDate) &
                " | LastScheduledEnd=" & FormatDateOrBlank(lastScheduledEnd) &
                " | Effective=" & FormatDateOrBlank(effective)
            )

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

        Dim settingsFmt As Integer =
            preactor.GetFormatNumber(OptimizerSettingsCatalog.FormatName)
        If settingsFmt <= 0 Then Return defaultValue

        Dim parameterField As Integer =
            preactor.GetFieldNumber(
                settingsFmt,
                OptimizerSettingsCatalog.ParameterFieldName)
        Dim dateField As Integer =
            preactor.GetFieldNumber(
                settingsFmt,
                OptimizerSettingsCatalog.DateFieldName)

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
            ' Firing readiness comes from the latest released operation
            ' before loading 290. This intentionally supports direct
            ' 200 -> 290 routings where no drying operations exist.
            ' --------------------------------------------------------
            Dim lastPre290OpNo As Integer = 0
            Dim lastPre290ReadyTime As DateTime = DateTime.MinValue
            Dim lastPre290OpRec As Integer = 0

            For Each r As DataRow In rows

                Dim opNo As Integer = SafeInt(r("Operation Number"))

                If opNo <= 0 OrElse opNo >= 290 Then Continue For
                If Not SafeBool(r("operation_releases_next")) Then Continue For

                Dim releaseTime As DateTime =
                SafeDate(r("operation_release_time"))

                If releaseTime = DateTime.MinValue Then Continue For

                If opNo > lastPre290OpNo OrElse
               (opNo = lastPre290OpNo AndAlso releaseTime > lastPre290ReadyTime) Then

                    lastPre290OpNo = opNo
                    lastPre290ReadyTime = releaseTime
                    lastPre290OpRec = SafeInt(r("OrdersID"))

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
                                     resourceName &
                                         OptimizerSettingsCatalog.AvailabilityParameterSuffix,
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
        Dim names As List(Of String) = resourceNames.ToList()

        Dim metadataDates As Dictionary(Of String, DateTime) =
            BuildMetadataAvailabilityFromGnKilnAvailability(preactor,
                                                            names,
                                                            terminator)

        Dim resourceRecByName As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)

        For Each resourceName As String In names

            If String.IsNullOrWhiteSpace(resourceName) Then Continue For

            Dim resourceRec As Integer = planningboard.GetResourceNumber(resourceName)
            If resourceRec <= 0 Then
                Throw New Exception("Resource not found: " & resourceName)
            End If
            resourceRecByName(resourceName) = resourceRec
        Next

        Dim lastEndByResource As Dictionary(Of Integer, DateTime) =
            GetResourceLastScheduledEnds(preactor,
                                         planningboard,
                                         resourceRecByName.Values)

        For Each resourceName As String In names

            If String.IsNullOrWhiteSpace(resourceName) Then Continue For

            Dim resourceRec As Integer = resourceRecByName(resourceName)

            Dim metadataDate As DateTime = DateTime.MinValue
            metadataDates.TryGetValue(resourceName, metadataDate)

            Dim lastScheduledEnd As DateTime = DateTime.MinValue

            lastEndByResource.TryGetValue(resourceRec, lastScheduledEnd)

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
