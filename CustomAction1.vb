Option Strict On
Option Explicit On

Imports System
Imports System.Runtime.InteropServices
Imports Preactor
Imports Preactor.Interop.PreactorObject
Imports System.Data
Imports System.Linq
Imports System.Collections.Generic


<ComVisible(True)> _
<Microsoft.VisualBasic.ComClass("1e22e401-29d2-4449-b9d1-f2105e02aae7", "70501438-bce8-4d9e-9c68-a440cde930aa")> _
Public Class CustomAction1
    Public Function Run(ByRef preactorComObject As PreactorObj, ByRef pespComObject As Object) As Integer

        Dim preactor As IPreactor = PreactorFactory.CreatePreactorObject(preactorComObject)

        'TODO : Your code goes here

        Return 0
    End Function
End Class

' =====================================================================================
'  VITRIFIED FIRING CYCLE BUILDER + DAY ASSIGNER (NO SCHEDULING)
'
'  What this code does:
'   1) Takes Vitrified orders (already filtered/bucketed/week-keyed outside, or you can do it upstream)
'   2) Builds "cycles" by accumulating compatible orders until capacity
'   3) Enforces minimum fill threshold (e.g., >= 60%) by trying to "club" with matching cycles
'   4) Assigns each cycle to an allowed firing day (e.g., Mon/Wed/Fri) using a greedy load balancing
'
'  What this code does NOT do:
'   - No operation scheduling
'   - No resource assignment in Opcenter
'   - No start/end times or calendars
'
'  Output:
'   - CycleHeader DataTable (one row per cycle)
'   - CycleLines DataTable (one row per order assigned to a cycle)
'   - CycleDayAssignment DataTable (one row per cycle with DayIndex/DayName)
'
'  Assumptions (tables/columns exist):
'   Orders table contains:
'     OrderId (String)
'     WeekKey (Date or String)  ' e.g., Monday anchor date
'     DueDate (Date)
'     Priority (Integer)        ' higher means more important
'     EquipType (String)        ' "TUNNEL" or "BATCH"
'     FiringCycleCode (String)  ' e.g., "65", "102", etc.
'     Qty (Decimal)
'     UnitTonnage (Decimal)     ' tonnage per unit (or per qty unit)
'     UnitVolume (Decimal)      ' volume per unit (or per qty unit)
'
'   MatchingCycle table contains:
'     CycleA (String), CycleB (String), IsAllowed (Boolean)
'     (Symmetry not guaranteed; we will treat as symmetric by checking both directions)
'
'   Capacity table contains:
'     EquipType (String)
'     WeekKey (Date or String) optional (if capacity varies by week); else can ignore
'     MaxTonnage (Decimal)
'     MaxVolume (Decimal)
'
'   Day capacity (optional) can also be passed; if not, we load-balance by total load only.
' =====================================================================================

Public Class VitrifiedCyclePlanner

    ' ---------------------------
    ' Types / Enums for clarity
    ' ---------------------------
    Public Enum LoadMetric
        TONNAGE
        VOLUME
    End Enum

    Public Class CycleBuildParams
        Public Property Metric As LoadMetric = LoadMetric.TONNAGE
        Public Property FillThreshold As Decimal = 0.6D   ' 60% minimum fill
        Public Property MaxAttemptsToClub As Integer = 500 ' safety to avoid infinite loops
        Public Property AllowMixingDifferentCycleCodesInBatch As Boolean = False
        ' If batch cycles must not mix different cycle codes, keep False.
        ' If business allows some mixing, set True.
    End Class

    ' ---------------------------
    ' Public entry point
    ' ---------------------------
    Public Shared Function BuildCyclesAndAssignDays(
        orders As DataTable,
        matchingCycle As DataTable,
        capacity As DataTable,
        allowedDays As List(Of Integer), ' ex: Monday=1, Wednesday=3, Friday=5 (1..7)
        weekKey As Object,               ' the current week bucket key you are processing
        equipType As String,             ' "TUNNEL" or "BATCH"
        p As CycleBuildParams,
        Optional dayCapacity As DataTable = Nothing, ' optional: columns DayIndex, MaxTonnage, MaxVolume
        Optional intermediateCsvPrefix As String = Nothing
    ) As DataSet

        ' Validate
        If orders Is Nothing Then Throw New ArgumentNullException(NameOf(orders))
        If matchingCycle Is Nothing Then Throw New ArgumentNullException(NameOf(matchingCycle))
        If capacity Is Nothing Then Throw New ArgumentNullException(NameOf(capacity))
        If allowedDays Is Nothing OrElse allowedDays.Count = 0 Then Throw New ArgumentException("allowedDays must be provided.")

        ' 1) Filter the input orders for this week + equipment
        Dim wkOrders = FilterOrdersForWeekAndEquip(orders, weekKey, equipType)

        ' 2) Build cycles (header + lines)
        Dim dsCycles = BuildCycles(wkOrders, matchingCycle, capacity, weekKey, equipType, p)

        Dim dtCycleHeader = dsCycles.Tables("CycleHeader")
        Dim dtCycleLines = dsCycles.Tables("CycleLines")

        ' Optional intermediate export
        If Not String.IsNullOrWhiteSpace(intermediateCsvPrefix) Then
            ' You said you already have a CSV export/import function that takes DataTable.
            ' Call your existing export methods here (examples commented out):
            ' ExportDataTableToCsv(dtCycleHeader, $"{intermediateCsvPrefix}_CycleHeader.csv")
            ' ExportDataTableToCsv(dtCycleLines,  $"{intermediateCsvPrefix}_CycleLines.csv")
        End If

        ' 3) Assign days to cycles (no scheduling, only assignment output)
        Dim dtDayAssign = AssignDaysToCycles(dtCycleHeader, allowedDays, p.Metric, dayCapacity)

        If Not String.IsNullOrWhiteSpace(intermediateCsvPrefix) Then
            ' ExportDataTableToCsv(dtDayAssign, $"{intermediateCsvPrefix}_CycleDayAssignment.csv")
        End If

        ' Package output as a DataSet for convenience
        Dim outDs As New DataSet("CyclePlan")
        outDs.Tables.Add(dtCycleHeader.Copy())
        outDs.Tables.Add(dtCycleLines.Copy())
        outDs.Tables.Add(dtDayAssign.Copy())
        Return outDs
    End Function

    ' =====================================================================================
    '  PART A: CYCLE BUILDER
    ' =====================================================================================

    Public Shared Function BuildCycles(
        wkOrders As DataTable,
        matchingCycle As DataTable,
        capacity As DataTable,
        weekKey As Object,
        equipType As String,
        p As CycleBuildParams
    ) As DataSet

        ' Prepare output tables
        Dim dtCycleHeader = CreateCycleHeaderTable()
        Dim dtCycleLines = CreateCycleLinesTable()

        ' Capacity for this equipment
        Dim cap = GetCapacity(capacity, equipType, weekKey)
        Dim maxTonnage As Decimal = cap.MaxTonnage
        Dim maxVolume As Decimal = cap.MaxVolume

        ' Sort candidate orders: earliest due date first, then higher priority
        Dim candidates = wkOrders.AsEnumerable().
            OrderBy(Function(r) SafeDate(r("DueDate"))).
            ThenByDescending(Function(r) SafeInt(r("Priority"))).
            ToList()

        Dim cycleSeq As Integer = 0

        ' We will repeatedly pick a "seed" order and try to fill a cycle around it
        While candidates.Any()
            cycleSeq += 1
            Dim cycleId As String = $"{equipType}_{weekKey}_{cycleSeq}"

            ' Create cycle header row
            Dim hdr = dtCycleHeader.NewRow()
            hdr("CycleId") = cycleId
            hdr("WeekKey") = weekKey
            hdr("EquipType") = equipType
            hdr("SeedCycleCode") = SafeStr(candidates(0)("FiringCycleCode"))
            hdr("MaxTonnage") = maxTonnage
            hdr("MaxVolume") = maxVolume
            hdr("LoadMetric") = p.Metric.ToString()
            hdr("FillThreshold") = p.FillThreshold
            hdr("TotalTonnage") = 0D
            hdr("TotalVolume") = 0D
            hdr("FillRatio") = 0D
            dtCycleHeader.Rows.Add(hdr)

            ' Create cycle line list (we will keep track in-memory too)
            Dim cycleOrderIds As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

            ' 1) Seed the cycle with the "best next" candidate (first in sorted list)
            Dim seed = candidates(0)
            AddOrderToCycle(cycleId, seed, dtCycleLines, hdr, cycleOrderIds, p.Metric)
            candidates.RemoveAt(0)

            ' 2) Greedily add compatible orders while we have capacity
            '    NOTE: "compatible" means:
            '      - For TUNNEL: must be allowed by MatchingCycle rules (club table)
            '      - For BATCH: by default do NOT mix different firing cycle codes (unless allowed by param)
            Dim madeProgress As Boolean = True
            While madeProgress
                madeProgress = False

                Dim nextCandidateIndex As Integer = FindBestNextCompatibleOrderIndex(
                    candidates, hdr, matchingCycle, equipType, p
                )

                If nextCandidateIndex >= 0 Then
                    Dim cand = candidates(nextCandidateIndex)

                    ' Capacity check: don't exceed max for selected metric
                    If Not WouldExceedCapacity(hdr, cand, p.Metric, maxTonnage, maxVolume) Then
                        AddOrderToCycle(cycleId, cand, dtCycleLines, hdr, cycleOrderIds, p.Metric)
                        candidates.RemoveAt(nextCandidateIndex)
                        madeProgress = True
                    End If
                End If
            End While

            ' 3) If cycle fill < threshold, attempt "clubbing" (pull extra compatible jobs more aggressively)
            '    This is the key business logic: avoid running low-fill cycles unless forced.
            Dim attempts As Integer = 0
            While GetFillRatio(hdr, p.Metric) < p.FillThreshold AndAlso candidates.Any() AndAlso attempts < p.MaxAttemptsToClub
                attempts += 1

                Dim idx As Integer = FindAnyCompatibleOrderIndex(candidates, hdr, matchingCycle, equipType, p)

                If idx < 0 Then
                    ' No compatible order exists. Stop clubbing for this cycle.
                    Exit While
                End If

                Dim cand = candidates(idx)
                If WouldExceedCapacity(hdr, cand, p.Metric, maxTonnage, maxVolume) Then
                    ' If it exceeds capacity, we cannot add this job. Remove it from consideration for this cycle
                    ' but do NOT delete it overall; we just skip for now.
                    ' We'll try a different compatible job next.
                    ' (To avoid infinite loops, we can temporarily mark and continue.)
                    candidates.RemoveAt(idx)
                    candidates.Add(cand) ' rotate to back
                Else
                    AddOrderToCycle(cycleId, cand, dtCycleLines, hdr, cycleOrderIds, p.Metric)
                    candidates.RemoveAt(idx)
                End If
            End While

            ' 4) Finalize fill ratio
            hdr("FillRatio") = GetFillRatio(hdr, p.Metric)
        End While

        Dim ds As New DataSet("Cycles")
        ds.Tables.Add(dtCycleHeader)
        ds.Tables.Add(dtCycleLines)
        Return ds
    End Function

    ' Finds "best next" compatible order, trying to reduce setups and respect due dates.
    ' Strategy:
    '   1) Prefer same FiringCycleCode as seed (min setups)
    '   2) Then any code compatible by matching rules (tunnel) or allowed mixing rules (batch)
    '   3) Within those, choose earliest due date, then highest priority
    Private Shared Function FindBestNextCompatibleOrderIndex(
        candidates As List(Of DataRow),
        cycleHeaderRow As DataRow,
        matchingCycle As DataTable,
        equipType As String,
        p As CycleBuildParams
    ) As Integer

        If candidates Is Nothing OrElse candidates.Count = 0 Then Return -1

        Dim seedCode As String = SafeStr(cycleHeaderRow("SeedCycleCode"))

        Dim bestIdx As Integer = -1
        Dim bestScore As Decimal = Decimal.MaxValue

        For i As Integer = 0 To candidates.Count - 1
            Dim r = candidates(i)
            If Not IsCompatible(r, cycleHeaderRow, matchingCycle, equipType, p) Then Continue For

            ' Score: lower is better
            ' Primary: prefer same cycle code (score bonus)
            Dim code As String = SafeStr(r("FiringCycleCode"))
            Dim isSameCode As Boolean = String.Equals(code, seedCode, StringComparison.OrdinalIgnoreCase)

            Dim due As Date = SafeDate(r("DueDate"))
            Dim pr As Integer = SafeInt(r("Priority"))

            ' Weighted score:
            '   - due date as base
            '   - subtract small amount for higher priority
            '   - subtract bigger amount if same code (to keep cycle homogeneous)
            Dim score As Decimal = CType(due.ToOADate(), Decimal)
            score -= (pr * 0.0001D)
            If isSameCode Then score -= 0.5D

            If score < bestScore Then
                bestScore = score
                bestIdx = i
            End If
        Next

        Return bestIdx
    End Function

    ' More permissive search used during clubbing:
    '   - returns first compatible candidate in sorted list (candidates already due/pr sorted)
    Private Shared Function FindAnyCompatibleOrderIndex(
        candidates As List(Of DataRow),
        cycleHeaderRow As DataRow,
        matchingCycle As DataTable,
        equipType As String,
        p As CycleBuildParams
    ) As Integer

        For i As Integer = 0 To candidates.Count - 1
            If IsCompatible(candidates(i), cycleHeaderRow, matchingCycle, equipType, p) Then
                Return i
            End If
        Next
        Return -1
    End Function

    ' Compatibility rules:
    '  - Tunnel: use MatchingCycle table between SeedCycleCode and candidate's code
    '  - Batch: default is same cycle code only, unless AllowMixingDifferentCycleCodesInBatch = True
    Private Shared Function IsCompatible(
        orderRow As DataRow,
        cycleHeaderRow As DataRow,
        matchingCycle As DataTable,
        equipType As String,
        p As CycleBuildParams
    ) As Boolean

        Dim seedCode As String = SafeStr(cycleHeaderRow("SeedCycleCode"))
        Dim candCode As String = SafeStr(orderRow("FiringCycleCode"))

        If String.Equals(equipType, "TUNNEL", StringComparison.OrdinalIgnoreCase) Then
            ' Tunnel: require matching rule (including same code if table doesn’t contain it)
            If String.Equals(seedCode, candCode, StringComparison.OrdinalIgnoreCase) Then Return True
            Return IsMatchingAllowed(matchingCycle, seedCode, candCode)
        End If

        ' Batch:
        If p.AllowMixingDifferentCycleCodesInBatch Then
            Return True
        Else
            Return String.Equals(seedCode, candCode, StringComparison.OrdinalIgnoreCase)
        End If
    End Function

    ' Reads MatchingCycle table to see if (A,B) or (B,A) is allowed
    Private Shared Function IsMatchingAllowed(matchingCycle As DataTable, a As String, b As String) As Boolean
        If String.IsNullOrWhiteSpace(a) OrElse String.IsNullOrWhiteSpace(b) Then Return False

        Dim rows = matchingCycle.AsEnumerable().Where(Function(r)
                                                          Dim ca = SafeStr(r("CycleA"))
                                                          Dim cb = SafeStr(r("CycleB"))
                                                          Dim ok = SafeBool(r("IsAllowed"))
                                                          If Not ok Then Return False
                                                          Return (String.Equals(ca, a, StringComparison.OrdinalIgnoreCase) AndAlso
                                                                  String.Equals(cb, b, StringComparison.OrdinalIgnoreCase)) _
                                                              OrElse
                                                                 (String.Equals(ca, b, StringComparison.OrdinalIgnoreCase) AndAlso
                                                                  String.Equals(cb, a, StringComparison.OrdinalIgnoreCase))
                                                      End Function)

        Return rows.Any()
    End Function

    ' Adds order to cycle lines and updates cycle header aggregates
    Private Shared Sub AddOrderToCycle(
        cycleId As String,
        orderRow As DataRow,
        dtCycleLines As DataTable,
        hdr As DataRow,
        cycleOrderIds As HashSet(Of String),
        metric As LoadMetric
    )
        Dim orderId As String = SafeStr(orderRow("OrderId"))
        If cycleOrderIds.Contains(orderId) Then Exit Sub ' safety

        cycleOrderIds.Add(orderId)

        Dim qty As Decimal = SafeDec(orderRow("Qty"))
        Dim unitTon As Decimal = SafeDec(orderRow("UnitTonnage"))
        Dim unitVol As Decimal = SafeDec(orderRow("UnitVolume"))

        Dim orderTon As Decimal = qty * unitTon
        Dim orderVol As Decimal = qty * unitVol

        ' Add a line
        Dim ln = dtCycleLines.NewRow()
        ln("CycleId") = cycleId
        ln("OrderId") = orderId
        ln("FiringCycleCode") = SafeStr(orderRow("FiringCycleCode"))
        ln("DueDate") = SafeDate(orderRow("DueDate"))
        ln("Priority") = SafeInt(orderRow("Priority"))
        ln("Qty") = qty
        ln("OrderTonnage") = orderTon
        ln("OrderVolume") = orderVol
        dtCycleLines.Rows.Add(ln)

        ' Update header totals
        hdr("TotalTonnage") = SafeDec(hdr("TotalTonnage")) + orderTon
        hdr("TotalVolume") = SafeDec(hdr("TotalVolume")) + orderVol

        hdr("FillRatio") = GetFillRatio(hdr, metric)
    End Sub

    Private Shared Function WouldExceedCapacity(
        hdr As DataRow,
        cand As DataRow,
        metric As LoadMetric,
        maxTonnage As Decimal,
        maxVolume As Decimal
    ) As Boolean

        Dim qty As Decimal = SafeDec(cand("Qty"))
        Dim candTon As Decimal = qty * SafeDec(cand("UnitTonnage"))
        Dim candVol As Decimal = qty * SafeDec(cand("UnitVolume"))

        Dim currentTon As Decimal = SafeDec(hdr("TotalTonnage"))
        Dim currentVol As Decimal = SafeDec(hdr("TotalVolume"))

        If metric = LoadMetric.TONNAGE Then
            Return (currentTon + candTon) > maxTonnage
        Else
            Return (currentVol + candVol) > maxVolume
        End If
    End Function

    Private Shared Function GetFillRatio(hdr As DataRow, metric As LoadMetric) As Decimal
        Dim totTon As Decimal = SafeDec(hdr("TotalTonnage"))
        Dim totVol As Decimal = SafeDec(hdr("TotalVolume"))
        Dim maxTon As Decimal = SafeDec(hdr("MaxTonnage"))
        Dim maxVol As Decimal = SafeDec(hdr("MaxVolume"))

        If metric = LoadMetric.TONNAGE Then
            If maxTon <= 0D Then Return 0D
            Return totTon / maxTon
        Else
            If maxVol <= 0D Then Return 0D
            Return totVol / maxVol
        End If
    End Function

    ' =====================================================================================
    '  PART B: DAY ASSIGNMENT (NO SCHEDULING)
    ' =====================================================================================

    Public Shared Function AssignDaysToCycles(
        dtCycleHeader As DataTable,
        allowedDays As List(Of Integer), ' 1..7 (Mon..Sun) or any index you use
        metric As LoadMetric,
        Optional dayCapacity As DataTable = Nothing
    ) As DataTable

        ' Output:
        ' CycleId, DayIndex, DayName, AssignedLoad, Notes
        Dim dt As New DataTable("CycleDayAssignment")
        dt.Columns.Add("CycleId", GetType(String))
        dt.Columns.Add("DayIndex", GetType(Integer))
        dt.Columns.Add("DayName", GetType(String))
        dt.Columns.Add("AssignedLoad", GetType(Decimal))
        dt.Columns.Add("Notes", GetType(String))

        ' Track running load per day (to balance)
        Dim dayLoad As New Dictionary(Of Integer, Decimal)
        For Each d In allowedDays
            dayLoad(d) = 0D
        Next

        ' If dayCapacity is provided, we can also avoid exceeding day max
        ' Expected columns: DayIndex, MaxTonnage, MaxVolume
        Dim hasDayCap As Boolean = (dayCapacity IsNot Nothing AndAlso dayCapacity.Columns.Contains("DayIndex"))

        ' Sort cycles: schedule earlier due cycles first, and higher priority first
        ' We approximate due date of a cycle as min due date from its lines, but since we only
        ' have header here, we fallback to FillRatio/TotalLoad sorting.
        ' If you want true due-date sorting, pass in a "CycleMinDueDate" column on header.
        Dim cycles = dtCycleHeader.AsEnumerable().
            OrderByDescending(Function(r) SafeDec(r("FillRatio"))). ' fuller cycles first
            ThenByDescending(Function(r) GetCycleLoad(r, metric)).
            ToList()

        For Each c In cycles
            Dim cycleId As String = SafeStr(c("CycleId"))
            Dim load As Decimal = GetCycleLoad(c, metric)

            ' Choose the best day:
            '   1) the day with minimum current load (greedy balance)
            '   2) optionally, ensure day capacity not exceeded
            Dim chosenDay As Integer = ChooseDayGreedy(dayLoad, allowedDays, load, metric, dayCapacity)

            Dim row = dt.NewRow()
            row("CycleId") = cycleId
            row("DayIndex") = chosenDay
            row("DayName") = DayIndexToName(chosenDay)
            row("AssignedLoad") = load
            row("Notes") = If(hasDayCap, "Greedy load-balance with day capacity check", "Greedy load-balance")
            dt.Rows.Add(row)

            dayLoad(chosenDay) = dayLoad(chosenDay) + load
        Next

        Return dt
    End Function

    Private Shared Function ChooseDayGreedy(
        dayLoad As Dictionary(Of Integer, Decimal),
        allowedDays As List(Of Integer),
        cycleLoad As Decimal,
        metric As LoadMetric,
        dayCapacity As DataTable
    ) As Integer

        Dim hasCap As Boolean = (dayCapacity IsNot Nothing AndAlso dayCapacity.Columns.Contains("DayIndex"))

        ' Candidate days ordered by current smallest load
        Dim orderedDays = allowedDays.OrderBy(Function(d) dayLoad(d)).ToList()

        If Not hasCap Then
            Return orderedDays.First()
        End If

        ' With day capacity:
        ' pick the first day where dayLoad + cycleLoad <= dayMax, else fallback to least loaded day
        For Each d In orderedDays
            Dim dayMax = GetDayMax(dayCapacity, d, metric)
            If dayMax <= 0D Then
                ' If max not defined, treat as unlimited
                Return d
            End If

            If (dayLoad(d) + cycleLoad) <= dayMax Then
                Return d
            End If
        Next

        ' If all days exceed capacity, still assign to least-loaded day
        ' (Your downstream scheduling / repair logic can push to next week if needed)
        Return orderedDays.First()
    End Function

    Private Shared Function GetDayMax(dayCapacity As DataTable, dayIndex As Integer, metric As LoadMetric) As Decimal
        If dayCapacity Is Nothing Then Return 0D

        Dim row = dayCapacity.AsEnumerable().
            FirstOrDefault(Function(r) SafeInt(r("DayIndex")) = dayIndex)

        If row Is Nothing Then Return 0D

        If metric = LoadMetric.TONNAGE AndAlso dayCapacity.Columns.Contains("MaxTonnage") Then
            Return SafeDec(row("MaxTonnage"))
        ElseIf metric = LoadMetric.VOLUME AndAlso dayCapacity.Columns.Contains("MaxVolume") Then
            Return SafeDec(row("MaxVolume"))
        End If

        Return 0D
    End Function

    Private Shared Function GetCycleLoad(cycleHeaderRow As DataRow, metric As LoadMetric) As Decimal
        If metric = LoadMetric.TONNAGE Then
            Return SafeDec(cycleHeaderRow("TotalTonnage"))
        Else
            Return SafeDec(cycleHeaderRow("TotalVolume"))
        End If
    End Function

    Private Shared Function DayIndexToName(dayIndex As Integer) As String
        ' Adjust if you use different encoding.
        ' Here: 1=Mon ... 7=Sun
        Select Case dayIndex
            Case 1 : Return "Mon"
            Case 2 : Return "Tue"
            Case 3 : Return "Wed"
            Case 4 : Return "Thu"
            Case 5 : Return "Fri"
            Case 6 : Return "Sat"
            Case 7 : Return "Sun"
            Case Else : Return $"Day{dayIndex}"
        End Select
    End Function

    ' =====================================================================================
    '  HELPERS: Filtering, Capacity lookup, Output table schemas, Safe converters
    ' =====================================================================================

    Private Shared Function FilterOrdersForWeekAndEquip(orders As DataTable, weekKey As Object, equipType As String) As DataTable
        Dim dt = orders.Clone()

        Dim rows = orders.AsEnumerable().
            Where(Function(r)
                      Dim wkOk = Object.Equals(r("WeekKey"), weekKey)
                      Dim eqOk = String.Equals(SafeStr(r("EquipType")), equipType, StringComparison.OrdinalIgnoreCase)
                      Return wkOk AndAlso eqOk
                  End Function)

        For Each r In rows
            dt.ImportRow(r)
        Next

        Return dt
    End Function

    Private Class CapacityInfo
        Public Property MaxTonnage As Decimal
        Public Property MaxVolume As Decimal
    End Class

    Private Shared Function GetCapacity(capacity As DataTable, equipType As String, weekKey As Object) As CapacityInfo
        ' If capacity is week-specific, include WeekKey in the lookup.
        ' If not, just match equipType.
        Dim hasWeek As Boolean = capacity.Columns.Contains("WeekKey")

        Dim row = capacity.AsEnumerable().
            FirstOrDefault(Function(r)
                               Dim eqOk = String.Equals(SafeStr(r("EquipType")), equipType, StringComparison.OrdinalIgnoreCase)
                               If Not eqOk Then Return False
                               If hasWeek Then
                                   Return Object.Equals(r("WeekKey"), weekKey)
                               End If
                               Return True
                           End Function)

        If row Is Nothing Then
            ' Default safe capacity to avoid crash (but you should populate your table properly)
            Return New CapacityInfo With {.MaxTonnage = 0D, .MaxVolume = 0D}
        End If

        Dim info As New CapacityInfo()
        info.MaxTonnage = If(capacity.Columns.Contains("MaxTonnage"), SafeDec(row("MaxTonnage")), 0D)
        info.MaxVolume = If(capacity.Columns.Contains("MaxVolume"), SafeDec(row("MaxVolume")), 0D)
        Return info
    End Function

    Private Shared Function CreateCycleHeaderTable() As DataTable
        Dim dt As New DataTable("CycleHeader")
        dt.Columns.Add("CycleId", GetType(String))
        dt.Columns.Add("WeekKey", GetType(Object))
        dt.Columns.Add("EquipType", GetType(String))
        dt.Columns.Add("SeedCycleCode", GetType(String))
        dt.Columns.Add("MaxTonnage", GetType(Decimal))
        dt.Columns.Add("MaxVolume", GetType(Decimal))
        dt.Columns.Add("LoadMetric", GetType(String))
        dt.Columns.Add("FillThreshold", GetType(Decimal))
        dt.Columns.Add("TotalTonnage", GetType(Decimal))
        dt.Columns.Add("TotalVolume", GetType(Decimal))
        dt.Columns.Add("FillRatio", GetType(Decimal))
        Return dt
    End Function

    Private Shared Function CreateCycleLinesTable() As DataTable
        Dim dt As New DataTable("CycleLines")
        dt.Columns.Add("CycleId", GetType(String))
        dt.Columns.Add("OrderId", GetType(String))
        dt.Columns.Add("FiringCycleCode", GetType(String))
        dt.Columns.Add("DueDate", GetType(Date))
        dt.Columns.Add("Priority", GetType(Integer))
        dt.Columns.Add("Qty", GetType(Decimal))
        dt.Columns.Add("OrderTonnage", GetType(Decimal))
        dt.Columns.Add("OrderVolume", GetType(Decimal))
        Return dt
    End Function

    ' ---- Safe converters (avoid DBNull problems) ----
    Private Shared Function SafeStr(v As Object) As String
        If v Is Nothing OrElse v Is DBNull.Value Then Return ""
        Return Convert.ToString(v)
    End Function

    Private Shared Function SafeInt(v As Object) As Integer
        If v Is Nothing OrElse v Is DBNull.Value Then Return 0
        Dim i As Integer
        If Integer.TryParse(Convert.ToString(v), i) Then Return i
        Return 0
    End Function

    Private Shared Function SafeDec(v As Object) As Decimal
        If v Is Nothing OrElse v Is DBNull.Value Then Return 0D
        Dim d As Decimal
        If Decimal.TryParse(Convert.ToString(v), d) Then Return d
        Return 0D
    End Function

    Private Shared Function SafeBool(v As Object) As Boolean
        If v Is Nothing OrElse v Is DBNull.Value Then Return False
        Dim b As Boolean
        If Boolean.TryParse(Convert.ToString(v), b) Then Return b
        ' Allow 0/1
        Dim i As Integer
        If Integer.TryParse(Convert.ToString(v), i) Then Return (i <> 0)
        Return False
    End Function

    Private Shared Function SafeDate(v As Object) As Date
        If v Is Nothing OrElse v Is DBNull.Value Then Return Date.MinValue
        Dim dt As Date
        If Date.TryParse(Convert.ToString(v), dt) Then Return dt
        Return Date.MinValue
    End Function

End Class
