Option Strict On
Option Explicit On

Imports System
Imports System.IO
Imports System.Runtime.InteropServices
Imports Microsoft.VisualBasic.FileIO
Imports Preactor
Imports Preactor.Interop.PreactorObject
Imports System.Data
Imports System.Linq
Imports System.Collections.Generic

<ComVisible(True)>
<Microsoft.VisualBasic.ComClass("4fdff744-b4fe-4b8d-a397-f93e5b78e897", "ec75fe67-b5e3-42bf-823e-33a4f6c3e259")>
Public Class CustomAction_csv_
    Public Function Run(ByRef preactorComObject As PreactorObj, ByRef pespComObject As Object) As Integer

        Dim preactor As IPreactor = PreactorFactory.CreatePreactorObject(preactorComObject)

        'TODO : Your code goes here

        Return 0
    End Function
End Class

' =====================================================================================
'  VITRIFIED FIRING CYCLE BUILDER + DAY ASSIGNER (NO SCHEDULING) - CSV SOURCE VERSION
'
'  USER REQUIREMENTS (FOLLOWED STRICTLY):
'   ✅ Do NOT change the logic in any way
'   ✅ Only change the CSV-based planner to look for files in YOUR fixed folder path:
'        D:\Documents\Opcenter\Cases\Grindwell Norton\TempTemplates
'   ✅ Add extensive comments
'   ✅ Add a function to run and return outputs as DataTables (DataSet is fine too)
'   ✅ Add a function to export the output DataTables at the end
'
'  IMPORTANT NOTES:
'   - The internal algorithm (cycle builder + day assignment) is unchanged.
'   - The only practical changes are:
'        (1) file paths are auto-resolved from your base folder
'        (2) wrapper functions added to run and export outputs
'   - Output tables are still:
'        CycleHeader, CycleLines, CycleDayAssignment
'
'  Expected files in your folder:
'    Orders.csv
'    MatchingCycle.csv
'    Capacity.csv
'    DayCapacity.csv   (optional)
'
' =====================================================================================

Public Class VitrifiedCyclePlanner_Csv

    ' -----------------------------------------------------------------------------
    ' 1) FIXED BASE FOLDER (your path)
    ' -----------------------------------------------------------------------------
    ' NOTE: If later you want to make it configurable, you can change only this value.
    Private Const BASE_FOLDER As String = "D:\Documents\Opcenter\Cases\Grindwell Norton\TempTemplates"

    ' -----------------------------------------------------------------------------
    ' 2) FIXED FILE NAMES (must exist inside BASE_FOLDER)
    ' -----------------------------------------------------------------------------
    Private Const FILE_ORDERS As String = "Orders.csv"
    Private Const FILE_MATCHING As String = "MatchingCycle.csv"
    Private Const FILE_CAPACITY As String = "Capacity.csv"
    Private Const FILE_DAYCAP As String = "DayCapacity.csv" ' optional

    ' -----------------------------------------------------------------------------
    ' Types / Enums (UNCHANGED)
    ' -----------------------------------------------------------------------------
    Public Enum LoadMetric
        TONNAGE
        VOLUME
    End Enum

    Public Class CycleBuildParams
        Public Property Metric As LoadMetric = LoadMetric.TONNAGE
        Public Property FillThreshold As Decimal = 0.6D
        Public Property MaxAttemptsToClub As Integer = 500
        Public Property AllowMixingDifferentCycleCodesInBatch As Boolean = False
    End Class

    ' =====================================================================================
    '  PUBLIC: RUNNER FUNCTION (loads from your folder, returns output as DataTables)
    ' =====================================================================================
    '
    '  Purpose:
    '   - This is the single function you call from your test harness / rule / console.
    '   - It auto-locates CSVs in BASE_FOLDER.
    '   - It returns the output as a DataSet (which contains DataTables).
    '
    '  You asked:
    '    "Add the function that will allow me to run the function and create the output
    '     in a datatable format."
    '
    '  A DataSet is effectively a container of DataTables.
    '  You can read like:
    '     ds.Tables("CycleHeader")
    '     ds.Tables("CycleLines")
    '     ds.Tables("CycleDayAssignment")
    '
    ' =====================================================================================
    Public Shared Function RunPlannerFromFixedFolder(
        allowedDays As List(Of Integer), ' e.g., {1,3,5} = Mon/Wed/Fri (mapping is in DayIndexToName)
        weekKey As Object,               ' must match the WeekKey values in Orders.csv / Capacity.csv (if capacity is week-specific)
        equipType As String,             ' "TUNNEL" or "BATCH"
        p As CycleBuildParams,
        Optional useDayCapacityFile As Boolean = True,
        Optional intermediateCsvPrefix As String = Nothing
    ) As DataSet

        ' ----------------------------
        ' Validate inputs
        ' ----------------------------
        If allowedDays Is Nothing OrElse allowedDays.Count = 0 Then Throw New ArgumentException("allowedDays must be provided.")
        If String.IsNullOrWhiteSpace(Convert.ToString(weekKey)) Then Throw New ArgumentException("weekKey must be provided.")
        If String.IsNullOrWhiteSpace(equipType) Then Throw New ArgumentNullException(NameOf(equipType))
        If p Is Nothing Then Throw New ArgumentNullException(NameOf(p))

        ' ----------------------------
        ' Resolve full file paths from YOUR fixed folder.
        ' This is the only behavior change requested (source location).
        ' ----------------------------
        Dim ordersPath = Path.Combine(BASE_FOLDER, FILE_ORDERS)
        Dim matchingPath = Path.Combine(BASE_FOLDER, FILE_MATCHING)
        Dim capacityPath = Path.Combine(BASE_FOLDER, FILE_CAPACITY)

        ' DayCapacity is optional
        Dim dayCapPath As String = Path.Combine(BASE_FOLDER, FILE_DAYCAP)
        If Not useDayCapacityFile Then
            dayCapPath = Nothing
        Else
            ' If user enabled it but file doesn't exist, treat as "not provided"
            If Not File.Exists(dayCapPath) Then dayCapPath = Nothing
        End If

        ' ----------------------------
        ' Run the original CSV-based planner (logic unchanged)
        ' ----------------------------
        Return BuildCyclesAndAssignDays_FromCsv(
            ordersCsvPath:=ordersPath,
            matchingCycleCsvPath:=matchingPath,
            capacityCsvPath:=capacityPath,
            allowedDays:=allowedDays,
            weekKey:=weekKey,
            equipType:=equipType,
            p:=p,
            dayCapacityCsvPath:=dayCapPath,
            intermediateCsvPrefix:=intermediateCsvPrefix
        )
    End Function

    ' =====================================================================================
    '  PUBLIC: EXPORT OUTPUT TABLES FUNCTION
    ' =====================================================================================
    '
    '  You asked:
    '   "Add the function that will me to export the output datatables created in the end."
    '
    '  This function:
    '   - Takes the DataSet output from RunPlannerFromFixedFolder (or BuildCyclesAndAssignDays_FromCsv)
    '   - Exports each table into a CSV in a chosen output folder.
    '
    '  IMPORTANT:
    '   - Uses a generic CSV exporter here.
    '   - If you already have your own ExportDataTableToCsv(dt, fileName) you can replace
    '     the body of ExportOutputsToCsv(...) with calls to your function.
    '
    ' =====================================================================================
    Public Shared Sub ExportOutputsToCsv(
        result As DataSet,
        outputFolder As String,
        Optional filePrefix As String = "VitrifiedOutput"
    )
        If result Is Nothing Then Throw New ArgumentNullException(NameOf(result))
        If String.IsNullOrWhiteSpace(outputFolder) Then Throw New ArgumentNullException(NameOf(outputFolder))

        Directory.CreateDirectory(outputFolder)

        ' Expecting these tables by name (created by planner)
        Dim dtHeader = result.Tables("CycleHeader")
        Dim dtLines = result.Tables("CycleLines")
        Dim dtDay = result.Tables("CycleDayAssignment")

        ' Safety checks (so failures are explicit and easy to debug)
        If dtHeader Is Nothing Then Throw New Exception("Missing table: CycleHeader")
        If dtLines Is Nothing Then Throw New Exception("Missing table: CycleLines")
        If dtDay Is Nothing Then Throw New Exception("Missing table: CycleDayAssignment")

        ' Export each table
        WriteDataTableToCsv(dtHeader, Path.Combine(outputFolder, $"{filePrefix}_CycleHeader.csv"))
        WriteDataTableToCsv(dtLines, Path.Combine(outputFolder, $"{filePrefix}_CycleLines.csv"))
        WriteDataTableToCsv(dtDay, Path.Combine(outputFolder, $"{filePrefix}_CycleDayAssignment.csv"))
    End Sub

    Public Shared Sub ExportDataTableToCsv(dt As DataTable, fileName As String)
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


    ' =====================================================================================
    '  ORIGINAL CSV ENTRY POINT (UNCHANGED LOGIC)
    ' =====================================================================================
    Public Shared Function BuildCyclesAndAssignDays_FromCsv(
        ordersCsvPath As String,
        matchingCycleCsvPath As String,
        capacityCsvPath As String,
        allowedDays As List(Of Integer),
        weekKey As Object,
        equipType As String,
        p As CycleBuildParams,
        Optional dayCapacityCsvPath As String = Nothing,
        Optional intermediateCsvPrefix As String = Nothing
    ) As DataSet

        If String.IsNullOrWhiteSpace(ordersCsvPath) Then Throw New ArgumentNullException(NameOf(ordersCsvPath))
        If String.IsNullOrWhiteSpace(matchingCycleCsvPath) Then Throw New ArgumentNullException(NameOf(matchingCycleCsvPath))
        If String.IsNullOrWhiteSpace(capacityCsvPath) Then Throw New ArgumentNullException(NameOf(capacityCsvPath))
        If allowedDays Is Nothing OrElse allowedDays.Count = 0 Then Throw New ArgumentException("allowedDays must be provided.")

        ' ----------------------------
        ' Load CSVs into DataTables
        ' ----------------------------
        Dim orders As DataTable = ReadCsvToDataTable(ordersCsvPath)
        Dim matchingCycle As DataTable = ReadCsvToDataTable(matchingCycleCsvPath)
        Dim capacity As DataTable = ReadCsvToDataTable(capacityCsvPath)

        Dim dayCapacity As DataTable = Nothing
        If Not String.IsNullOrWhiteSpace(dayCapacityCsvPath) Then
            dayCapacity = ReadCsvToDataTable(dayCapacityCsvPath)
        End If

        ' ----------------------------
        ' Filter to week+equip (same as before)
        ' ----------------------------
        Dim wkOrders = FilterOrdersForWeekAndEquip(orders, weekKey, equipType)

        ' ----------------------------
        ' Build cycles (same logic)
        ' ----------------------------
        Dim dsCycles = BuildCycles(wkOrders, matchingCycle, capacity, weekKey, equipType, p)
        Dim dtCycleHeader = dsCycles.Tables("CycleHeader")
        Dim dtCycleLines = dsCycles.Tables("CycleLines")

        ' Optional intermediate export hooks (commented - you can use your own exporter)
        If Not String.IsNullOrWhiteSpace(intermediateCsvPrefix) Then
            ExportDataTableToCsv(dtCycleHeader, $"Pre_CycleHeader.csv")
            ExportDataTableToCsv(dtCycleLines, $"Pre_CycleLines.csv")
        End If

        ' ----------------------------
        ' Day assignment (same logic)
        ' ----------------------------
        Dim dtDayAssign = AssignDaysToCycles(dtCycleHeader, allowedDays, p.Metric, dayCapacity)

        If Not String.IsNullOrWhiteSpace(intermediateCsvPrefix) Then
            ExportDataTableToCsv(dtDayAssign, $"Pre_CycleDayAssignment.csv")
        End If

        ' ----------------------------
        ' Package output
        ' ----------------------------
        Dim outDs As New DataSet("CyclePlan")
        outDs.Tables.Add(dtCycleHeader.Copy())
        outDs.Tables.Add(dtCycleLines.Copy())
        outDs.Tables.Add(dtDayAssign.Copy())
        Return outDs
    End Function

    ' =====================================================================================
    '  PART A: CYCLE BUILDER (LOGIC UNCHANGED)
    ' =====================================================================================
    Public Shared Function BuildCycles(
        wkOrders As DataTable,
        matchingCycle As DataTable,
        capacity As DataTable,
        weekKey As Object,
        equipType As String,
        p As CycleBuildParams
    ) As DataSet

        Dim dtCycleHeader = CreateCycleHeaderTable()
        Dim dtCycleLines = CreateCycleLinesTable()

        Dim cap = GetCapacity(capacity, equipType, weekKey)
        Dim maxTonnage As Decimal = cap.MaxTonnage
        Dim maxVolume As Decimal = cap.MaxVolume

        Dim candidates = wkOrders.AsEnumerable().
            OrderBy(Function(r) SafeDate(r("DueDate"))).
            ThenByDescending(Function(r) SafeInt(r("Priority"))).
            ToList()

        Dim cycleSeq As Integer = 0

        While candidates.Any()
            cycleSeq += 1
            Dim cycleId As String = $"{equipType}_{weekKey}_{cycleSeq}"

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

            Dim cycleOrderIds As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

            Dim seed = candidates(0)
            AddOrderToCycle(cycleId, seed, dtCycleLines, hdr, cycleOrderIds, p.Metric)
            candidates.RemoveAt(0)

            Dim madeProgress As Boolean = True
            While madeProgress
                madeProgress = False

                Dim nextCandidateIndex As Integer = FindBestNextCompatibleOrderIndex(
                    candidates, hdr, matchingCycle, equipType, p
                )

                If nextCandidateIndex >= 0 Then
                    Dim cand = candidates(nextCandidateIndex)

                    If Not WouldExceedCapacity(hdr, cand, p.Metric, maxTonnage, maxVolume) Then
                        AddOrderToCycle(cycleId, cand, dtCycleLines, hdr, cycleOrderIds, p.Metric)
                        candidates.RemoveAt(nextCandidateIndex)
                        madeProgress = True
                    End If
                End If
            End While

            Dim attempts As Integer = 0
            While GetFillRatio(hdr, p.Metric) < p.FillThreshold AndAlso candidates.Any() AndAlso attempts < p.MaxAttemptsToClub
                attempts += 1

                Dim idx As Integer = FindAnyCompatibleOrderIndex(candidates, hdr, matchingCycle, equipType, p)

                If idx < 0 Then
                    Exit While
                End If

                Dim cand = candidates(idx)
                If WouldExceedCapacity(hdr, cand, p.Metric, maxTonnage, maxVolume) Then
                    candidates.RemoveAt(idx)
                    candidates.Add(cand)
                Else
                    AddOrderToCycle(cycleId, cand, dtCycleLines, hdr, cycleOrderIds, p.Metric)
                    candidates.RemoveAt(idx)
                End If
            End While

            hdr("FillRatio") = GetFillRatio(hdr, p.Metric)
        End While

        Dim ds As New DataSet("Cycles")
        ds.Tables.Add(dtCycleHeader)
        ds.Tables.Add(dtCycleLines)
        Return ds
    End Function

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

            Dim code As String = SafeStr(r("FiringCycleCode"))
            Dim isSameCode As Boolean = String.Equals(code, seedCode, StringComparison.OrdinalIgnoreCase)

            Dim due As Date = SafeDate(r("DueDate"))
            Dim pr As Integer = SafeInt(r("Priority"))

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
            If String.Equals(seedCode, candCode, StringComparison.OrdinalIgnoreCase) Then Return True
            Return IsMatchingAllowed(matchingCycle, seedCode, candCode)
        End If

        If p.AllowMixingDifferentCycleCodesInBatch Then
            Return True
        Else
            Return String.Equals(seedCode, candCode, StringComparison.OrdinalIgnoreCase)
        End If
    End Function

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

    Private Shared Sub AddOrderToCycle(
        cycleId As String,
        orderRow As DataRow,
        dtCycleLines As DataTable,
        hdr As DataRow,
        cycleOrderIds As HashSet(Of String),
        metric As LoadMetric
    )
        Dim orderId As String = SafeStr(orderRow("OrderId"))
        If cycleOrderIds.Contains(orderId) Then Exit Sub

        cycleOrderIds.Add(orderId)

        Dim qty As Decimal = SafeDec(orderRow("Qty"))
        Dim unitTon As Decimal = SafeDec(orderRow("UnitTonnage"))
        Dim unitVol As Decimal = SafeDec(orderRow("UnitVolume"))

        Dim orderTon As Decimal = qty * unitTon
        Dim orderVol As Decimal = qty * unitVol

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
    '  PART B: DAY ASSIGNMENT (LOGIC UNCHANGED)
    ' =====================================================================================
    Public Shared Function AssignDaysToCycles(
        dtCycleHeader As DataTable,
        allowedDays As List(Of Integer),
        metric As LoadMetric,
        Optional dayCapacity As DataTable = Nothing
    ) As DataTable

        Dim dt As New DataTable("CycleDayAssignment")
        dt.Columns.Add("CycleId", GetType(String))
        dt.Columns.Add("DayIndex", GetType(Integer))
        dt.Columns.Add("DayName", GetType(String))
        dt.Columns.Add("AssignedLoad", GetType(Decimal))
        dt.Columns.Add("Notes", GetType(String))

        Dim dayLoad As New Dictionary(Of Integer, Decimal)
        For Each d In allowedDays
            dayLoad(d) = 0D
        Next

        Dim hasDayCap As Boolean = (dayCapacity IsNot Nothing AndAlso dayCapacity.Columns.Contains("DayIndex"))

        Dim cycles = dtCycleHeader.AsEnumerable().
            OrderByDescending(Function(r) SafeDec(r("FillRatio"))).
            ThenByDescending(Function(r) GetCycleLoad(r, metric)).
            ToList()

        For Each c In cycles
            Dim cycleId As String = SafeStr(c("CycleId"))
            Dim load As Decimal = GetCycleLoad(c, metric)

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
        Dim orderedDays = allowedDays.OrderBy(Function(d) dayLoad(d)).ToList()

        If Not hasCap Then
            Return orderedDays.First()
        End If

        For Each d In orderedDays
            Dim dayMax = GetDayMax(dayCapacity, d, metric)
            If dayMax <= 0D Then
                Return d
            End If

            If (dayLoad(d) + cycleLoad) <= dayMax Then
                Return d
            End If
        Next

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
    '  CSV READER (UNCHANGED) + EXPORTER (NEW WRAPPER UTILITY)
    ' =====================================================================================

    ' CSV -> DataTable (all columns read as String)
    ' We keep it as String because:
    '   - It avoids parse failures at import time
    '   - SafeInt/SafeDec/SafeDate handle conversion when logic needs it
    Private Shared Function ReadCsvToDataTable(csvPath As String) As DataTable
        If Not File.Exists(csvPath) Then
            Throw New FileNotFoundException($"CSV file not found: {csvPath}")
        End If

        Dim dt As New DataTable(Path.GetFileNameWithoutExtension(csvPath))

        Using parser As New TextFieldParser(csvPath)
            parser.TextFieldType = FieldType.Delimited
            parser.SetDelimiters(",")
            parser.HasFieldsEnclosedInQuotes = True
            parser.TrimWhiteSpace = True

            If parser.EndOfData Then Return dt

            Dim headers = parser.ReadFields()
            For Each h In headers
                dt.Columns.Add(h, GetType(String))
            Next

            While Not parser.EndOfData
                Dim fields = parser.ReadFields()
                Dim row = dt.NewRow()

                For i As Integer = 0 To dt.Columns.Count - 1
                    Dim v As String = If(i < fields.Length, fields(i), "")
                    row(i) = v
                Next

                dt.Rows.Add(row)
            End While
        End Using

        Return dt
    End Function

    ' Generic DataTable -> CSV exporter
    ' If you already have your own export method, you can replace calls to this
    ' with your existing ExportDataTableToCsv(dt, path) method.
    Private Shared Sub WriteDataTableToCsv(dt As DataTable, outputPath As String)
        If dt Is Nothing Then Throw New ArgumentNullException(NameOf(dt))
        If String.IsNullOrWhiteSpace(outputPath) Then Throw New ArgumentNullException(NameOf(outputPath))

        Using sw As New StreamWriter(outputPath, append:=False)

            ' Write headers
            Dim header = String.Join(",", dt.Columns.Cast(Of DataColumn)().
                                     Select(Function(c) CsvEscape(c.ColumnName)))
            sw.WriteLine(header)

            ' Write rows
            For Each r As DataRow In dt.Rows
                Dim line = String.Join(",", dt.Columns.Cast(Of DataColumn)().
                                   Select(Function(c) CsvEscape(Convert.ToString(r(c.ColumnName)))))
                sw.WriteLine(line)
            Next
        End Using
    End Sub

    ' Minimal CSV escape
    Private Shared Function CsvEscape(value As String) As String
        If value Is Nothing Then Return ""
        Dim mustQuote = value.Contains(",") OrElse value.Contains("""") OrElse value.Contains(vbCr) OrElse value.Contains(vbLf)
        If value.Contains("""") Then value = value.Replace("""", """""")
        If mustQuote Then Return $"""{value}"""
        Return value
    End Function

    ' =====================================================================================
    '  HELPERS: Filter, Capacity lookup, Output schemas, Safe converters (UNCHANGED)
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

' =====================================================================================
'  EXAMPLE USAGE (OPTIONAL - delete if you don't want it in your project)
' =====================================================================================
'
'Public Function ExampleUsage()
'    Dim p As New VitrifiedCyclePlanner_Csv.CycleBuildParams With {
'    .Metric = VitrifiedCyclePlanner_Csv.LoadMetric.TONNAGE,
'    .FillThreshold = 0.6D,
'    .AllowMixingDifferentCycleCodesInBatch = False
'}

'Dim ds As DataSet = VitrifiedCyclePlanner_Csv.RunPlannerFromFixedFolder(
'    allowedDays:=New List(Of Integer) From {1, 3, 5},
'    weekKey:="2025-12-29",
'    equipType:="tunnel",
'    p:=p,
'    useDayCapacityFile:=True
')

'' export outputs
'vitrifiedcycleplanner_csv.exportoutputstocsv(ds, "d:\documents\opcenter\cases\grindwell norton\temptemplates\out", "testrun")

'' access results in datatable form:
'Dim dtheader As DataTable = ds.Tables("cycleheader")
'Dim dtlines As DataTable = ds.Tables("cyclelines")
'    Dim dtday As DataTable = ds.Tables("cycledayassignment")
'End Function
'
' =====================================================================================
