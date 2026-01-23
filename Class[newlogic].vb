Imports System.Data
Imports System.Globalization
Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Diagnostics

Public Class CsvRoutingReader

    '===========================================================
    ' FUNCTION 1 : ReadRoutingCsv
    ' PURPOSE       : Reads the Routing.csv file and loads
    '                 ALL data into a DataTable.
    '                 No business logic applied.
    ' INPUT         : None
    ' OUTPUT        : DataTable containing Routing.csv data
    '===========================================================
    Public Function ReadRoutingCsv() As DataTable

        ' Absolute path as per business requirement
        Dim filePath As String = "D:\Documents\Opcenter\Cases\Grindwell Norton\Opcenter SC - Dev\Files\Templates\Routing.csv"

        ' Validate file existence
        If Not File.Exists(filePath) Then
            Throw New FileNotFoundException("Routing.csv not found at specified path.", filePath)
        End If

        Dim routingTable As New DataTable("Routing")

        Using reader As New StreamReader(filePath, Encoding.UTF8)

            '--------------------c-------------------------------
            ' STEP 1: Read Header Line
            '---------------------------------------------------
            If reader.EndOfStream Then
                Throw New Exception("Routing.csv is empty.")
            End If

            Dim headerLine As String = reader.ReadLine()
            Dim headers() As String = headerLine.Split(","c)

            ' Create DataTable columns exactly as CSV headers
            For Each header As String In headers
                routingTable.Columns.Add(header.Trim())
            Next

            '---------------------------------------------------
            ' STEP 2: Read Data Rows
            '---------------------------------------------------
            While Not reader.EndOfStream
                Dim line As String = reader.ReadLine()

                ' Skip empty lines (safe guard)
                If String.IsNullOrWhiteSpace(line) Then Continue While

                Dim fields() As String = line.Split(","c)

                Dim row As DataRow = routingTable.NewRow()

                ' Populate row values
                For i As Integer = 0 To headers.Length - 1
                    If i < fields.Length Then
                        row(i) = fields(i).Trim()
                    Else
                        row(i) = String.Empty
                    End If
                Next

                routingTable.Rows.Add(row)
            End While

        End Using

        Return routingTable

    End Function

    '===========================================================
    ' FUNCTION 2 : AddExpectedFiringStartDate
    ' PURPOSE       : Adds a new column "ExpectedFiringStartDate"
    '                 to the existing Routing DataTable.
    '
    ' LOGIC         : ExpectedFiringStartDate =
    '                 DueDate - FiringBufferDays
    '
    ' DATE FORMAT   : dd-mm-yyyy
    '
    ' INPUT         : Existing DataTable (already populated)
    ' OUTPUT        : Updated DataTable
    '===========================================================
    Public Function AddExpectedFiringStartDate(routingTable As DataTable) As DataTable

        '-------------------------------------------------------
        ' Column names (change here ONLY if CSV headers differ)
        '-------------------------------------------------------
        Dim dueDateColumn As String = "Due Date"
        Dim bufferDaysColumn As String = "Firing buffer"
        Dim outputColumn As String = "ExpectedFiringStartDate"

        '-------------------------------------------------------
        ' Add column ONLY if it does not already exist
        '-------------------------------------------------------
        If Not routingTable.Columns.Contains(outputColumn) Then
            routingTable.Columns.Add(outputColumn, GetType(String))
        End If

        ' Date parsing format
        Dim dateFormat As String = "dd-MM-yyyy"
        Dim culture As CultureInfo = CultureInfo.InvariantCulture

        '-------------------------------------------------------
        ' Loop through each row and calculate date
        '-------------------------------------------------------
        For Each row As DataRow In routingTable.Rows

            ' Safely read Due Date
            Dim dueDateString As String = row(dueDateColumn).ToString().Trim()
            Dim bufferDaysString As String = row(bufferDaysColumn).ToString().Trim()

            ' Default empty output
            row(outputColumn) = String.Empty

            ' Validate input values
            If String.IsNullOrWhiteSpace(dueDateString) _
                OrElse String.IsNullOrWhiteSpace(bufferDaysString) Then
                Continue For
            End If

            Dim dueDate As DateTime
            Dim bufferDays As Integer

            ' Parse Due Date (dd-mm-yyyy)
            If Not DateTime.TryParseExact(dueDateString, dateFormat, culture,
                                          DateTimeStyles.None, dueDate) Then
                Continue For
            End If

            ' Parse buffer days
            If Not Integer.TryParse(bufferDaysString, bufferDays) Then
                Continue For
            End If

            ' Calculate Expected Firing Start Date
            Dim expectedFiringDate As DateTime = dueDate.AddDays(-bufferDays)

            ' Write back in dd-mm-yyyy format
            row(outputColumn) = expectedFiringDate.ToString(dateFormat)

        Next

        'helper function to import process mapping
        'Dim op_matrix = Helper.LoadFromCsv(filePath:="D:\Documents\Opcenter\Cases\Grindwell Norton\Opcenter SC - Dev\Files\Templates\Process_Mapping.csv", strict:=True)


        Return routingTable

    End Function

    ''===========================================================
    '' FUNCTION 3 : AddFiringWeekAndBatchColumns
    '' PURPOSE       :
    ''   Uses output from AddExpectedFiringStartDate() and adds:
    ''     1) "firing week"  -> week number of the year
    ''     2) "week start"   -> Monday date for that week (dd-mm-yyyy)
    ''     3) "cycle+batch"  -> "cycle type_batch no."
    ''
    '' UPDATED LOGIC (ADDED WITHOUT MAJOR CHANGES):
    ''   In addition to previous clubbing by (WeekStart + KlinType + CycleType),
    ''   we now further split into batches based on:
    ''       Sum(Volume Occupancy at Operation 300) between:
    ''           minOccupancy and maxOccupancy
    ''
    '' CLUBBING SEQUENCE (as requested):
    ''   - Filter Klin Type
    ''   - Filter Cycle Type
    ''   - Group by ExpectedFiringStartDate (into week buckets via week start)
    ''   - Satisfy min/max occupancy for volume occupancy at operation 300
    ''
    '' OPTIMIZATION GOAL:
    ''   "Take as much volume occupancy as early as possible within that week"
    ''   Implemented using a greedy filling approach:
    ''     - Consider orders in earliest expected date order
    ''     - For each new batch, keep adding orders (preferring larger occupancy)
    ''       until maxOccupancy is reached (or no more fit)
    ''
    '' INPUTS:
    ''   routingTable   : DataTable already containing ExpectedFiringStartDate
    ''   minOccupancy   : minimum allowed total occupancy for a batch (e.g., 0.8)
    ''   maxOccupancy   : maximum allowed total occupancy for a batch (e.g., 1.0)
    ''
    '' OUTPUT:
    ''   Updated DataTable (same reference) with new columns filled
    ''==========================================================

    ''===========================================================
    '' FUNCTION NAME : AddFiringWeekAndBatchColumns
    ''
    '' PURPOSE:
    ''   - Uses output from AddExpectedFiringStartDate()
    ''   - Adds/Populates:
    ''       "firing week"       -> week number (Monday-based)
    ''       "week start"        -> Monday date for that week (dd-MM-yyyy)
    ''       "cycle+batch"       -> "PRIMARY_CYCLETYPE_BatchNo"
    ''       "batch+occupancy"   -> batch total occupancy
    ''       "batch+orders"      -> concatenated orders in that batch (order1_order2...)
    ''
    '' UPDATED BATCHING REQUIREMENTS:
    ''   Scope: per (WeekStart + Klin Type) -> batch numbers restart here.
    ''   Min occupancy = 0.8 (required)
    ''   Max occupancy = 1.0 (required)
    ''
    ''   Phase 1 (Same-cycle only):
    ''     - Batch only within same Cycle Type
    ''     - Prioritize higher cycle types first (e.g., 180 OR > 150 VT > 102 VT > 65 VT > others)
    ''     - Greedy packing: prefer larger occupancies while keeping <= 1.0
    ''     - Only accept batch if total in [0.8, 1.0]
    ''
    ''   Phase 2 (Mix leftovers only; restricted rules):
    ''     - 150 VT batches may be topped up with 102 VT
    ''       Label batch as 150 VT
    ''     - 102 VT batches may be topped up with 65 VT
    ''       Label batch as 102 VT
    ''     - Priority: complete 150 VT mixed batches first (consume 102 VT there first)
    ''     - Only accept batch if total in [0.8, 1.0]
    ''
    '' NOTE:
    ''   - Volume occupancy is read from Operation Number = 300 (per your latest standard).
    ''   - Orders always have op=300 and occupancy present (per your clarification).
    ''===========================================================
    'Public Function AddFiringWeekAndBatchColumns(routingTable As DataTable,
    '                                            minOccupancy As Double,
    '                                            maxOccupancy As Double) As DataTable

    '    '-------------------------------------------------------
    '    ' Column names (edit here ONLY if your CSV headers differ)
    '    '-------------------------------------------------------
    '    Dim expectedDateCol As String = "ExpectedFiringStartDate"
    '    Dim kilnTypeCol As String = "Klin Type"
    '    Dim cycleTypeCol As String = "Cycle Type"

    '    Dim orderNoCol As String = "Order No"

    '    ' IMPORTANT: As per your instruction going forward
    '    Dim operationNoCol As String = "Operation Number"

    '    Dim volOccCol As String = "Volume Occupancy"

    '    ' Output columns
    '    Dim firingWeekCol As String = "firing week"
    '    Dim weekStartCol As String = "week start"
    '    Dim cycleBatchCol As String = "cycle+batch"
    '    Dim batchOccupancyCol As String = "batch+occupancy"
    '    Dim batchOrdersCol As String = "batch+orders"

    '    '-------------------------------------------------------
    '    ' Validate occupancy bounds (strict)
    '    '-------------------------------------------------------
    '    If routingTable Is Nothing Then Throw New ArgumentNullException(NameOf(routingTable))

    '    If minOccupancy <= 0 OrElse maxOccupancy <= 0 OrElse minOccupancy > maxOccupancy Then
    '        Throw New Exception("Invalid occupancy limits. Ensure: 0 < minOccupancy <= maxOccupancy")
    '    End If

    '    '-------------------------------------------------------
    '    ' Required column validation
    '    '-------------------------------------------------------
    '    If Not routingTable.Columns.Contains(expectedDateCol) Then
    '        Throw New Exception("Required column missing: " & expectedDateCol & ". Run AddExpectedFiringStartDate() first.")
    '    End If
    '    If Not routingTable.Columns.Contains(kilnTypeCol) Then Throw New Exception("Required column missing: " & kilnTypeCol)
    '    If Not routingTable.Columns.Contains(cycleTypeCol) Then Throw New Exception("Required column missing: " & cycleTypeCol)

    '    If Not routingTable.Columns.Contains(orderNoCol) Then Throw New Exception("Required column missing: " & orderNoCol)
    '    If Not routingTable.Columns.Contains(operationNoCol) Then Throw New Exception("Required column missing: " & operationNoCol)
    '    If Not routingTable.Columns.Contains(volOccCol) Then Throw New Exception("Required column missing: " & volOccCol)

    '    '-------------------------------------------------------
    '    ' Add output columns if needed
    '    '-------------------------------------------------------
    '    If Not routingTable.Columns.Contains(firingWeekCol) Then routingTable.Columns.Add(firingWeekCol, GetType(Integer))
    '    If Not routingTable.Columns.Contains(weekStartCol) Then routingTable.Columns.Add(weekStartCol, GetType(String))
    '    If Not routingTable.Columns.Contains(cycleBatchCol) Then routingTable.Columns.Add(cycleBatchCol, GetType(String))

    '    If Not routingTable.Columns.Contains(batchOccupancyCol) Then routingTable.Columns.Add(batchOccupancyCol, GetType(Double))
    '    If Not routingTable.Columns.Contains(batchOrdersCol) Then routingTable.Columns.Add(batchOrdersCol, GetType(String))

    '    '-------------------------------------------------------
    '    ' Date parsing config (dd-MM-yyyy)
    '    '-------------------------------------------------------
    '    Dim dateFormat As String = "dd-MM-yyyy"
    '    Dim culture As CultureInfo = CultureInfo.InvariantCulture

    '    Dim weekRule As CalendarWeekRule = CalendarWeekRule.FirstFourDayWeek
    '    Dim firstDayOfWeek As DayOfWeek = DayOfWeek.Monday

    '    '=======================================================
    '    ' STEP A: Populate week start + firing week (row-level)
    '    '=======================================================
    '    For Each row As DataRow In routingTable.Rows
    '        row(weekStartCol) = String.Empty
    '        row(cycleBatchCol) = String.Empty
    '        row(firingWeekCol) = DBNull.Value

    '        row(batchOccupancyCol) = DBNull.Value
    '        row(batchOrdersCol) = String.Empty

    '        Dim expectedDateStr As String = row(expectedDateCol).ToString().Trim()
    '        If String.IsNullOrWhiteSpace(expectedDateStr) Then Continue For

    '        Dim expectedDate As DateTime
    '        If Not DateTime.TryParseExact(expectedDateStr, dateFormat, culture, DateTimeStyles.None, expectedDate) Then
    '            Continue For
    '        End If

    '        Dim delta As Integer = (CInt(expectedDate.DayOfWeek) - CInt(DayOfWeek.Monday) + 7) Mod 7
    '        Dim weekStartDate As DateTime = expectedDate.AddDays(-delta)
    '        row(weekStartCol) = weekStartDate.ToString(dateFormat)

    '        Dim weekNo As Integer = culture.Calendar.GetWeekOfYear(expectedDate, weekRule, firstDayOfWeek)
    '        row(firingWeekCol) = weekNo
    '    Next

    '    '=======================================================
    '    ' STEP B: Build Order -> Occupancy map from Operation Number = 300
    '    '=======================================================
    '    Dim orderToOcc As New Dictionary(Of String, Double)(StringComparer.OrdinalIgnoreCase)

    '    For Each row As DataRow In routingTable.Rows
    '        Dim orderNo As String = row(orderNoCol).ToString().Trim()
    '        If String.IsNullOrWhiteSpace(orderNo) Then Continue For

    '        Dim opStr As String = row(operationNoCol).ToString().Trim()
    '        If Not opStr.Equals("300", StringComparison.OrdinalIgnoreCase) Then Continue For

    '        Dim occStr As String = row(volOccCol).ToString().Trim()
    '        Dim occ As Double

    '        If Not Double.TryParse(occStr, NumberStyles.Any, CultureInfo.InvariantCulture, occ) Then
    '            Throw New Exception("Invalid occupancy value for Order No " & orderNo & " at operation 300.")
    '        End If

    '        If occ > 1.0 Then
    '            Throw New Exception("Data issue: Occupancy > 1 for Order No " & orderNo & ".")
    '        End If

    '        orderToOcc(orderNo) = occ
    '    Next

    '    '=======================================================
    '    ' STEP C: Build one order-level record per Order No
    '    '=======================================================
    '    Dim orders As New List(Of OrderInfo)()
    '    Dim seenOrders As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

    '    For Each row As DataRow In routingTable.Rows
    '        Dim orderNo As String = row(orderNoCol).ToString().Trim()
    '        If String.IsNullOrWhiteSpace(orderNo) Then Continue For
    '        If seenOrders.Contains(orderNo) Then Continue For

    '        Dim weekStartStr As String = row(weekStartCol).ToString().Trim()
    '        If String.IsNullOrWhiteSpace(weekStartStr) Then Continue For

    '        Dim kilnType As String = row(kilnTypeCol).ToString().Trim()
    '        Dim cycleRaw As String = row(cycleTypeCol).ToString().Trim()
    '        If String.IsNullOrWhiteSpace(kilnType) OrElse String.IsNullOrWhiteSpace(cycleRaw) Then Continue For

    '        Dim cycleNorm As String = NormalizeCycleName(cycleRaw)

    '        Dim expectedDateStr As String = row(expectedDateCol).ToString().Trim()
    '        Dim expectedDate As DateTime
    '        If Not DateTime.TryParseExact(expectedDateStr, dateFormat, culture, DateTimeStyles.None, expectedDate) Then Continue For

    '        If Not orderToOcc.ContainsKey(orderNo) Then
    '            Throw New Exception("Missing occupancy at operation 300 for Order No " & orderNo & ".")
    '        End If

    '        orders.Add(New OrderInfo With {
    '            .OrderNo = orderNo,
    '            .WeekStart = weekStartStr,
    '            .ExpectedDate = expectedDate,
    '            .KilnType = kilnType,
    '            .CycleTypeRaw = cycleRaw,
    '            .CycleTypeNorm = cycleNorm,
    '            .Occupancy = orderToOcc(orderNo)
    '        })

    '        seenOrders.Add(orderNo)
    '    Next

    '    '=======================================================
    '    ' STEP D: Batching engine per (WeekStart + Kiln Type)
    '    '=======================================================

    '    ' Outputs at order level
    '    Dim orderToCycleBatch As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
    '    Dim orderToBatchInstanceKey As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

    '    ' Batch metadata by unique batch instance key
    '    Dim batchToTotalOcc As New Dictionary(Of String, Double)(StringComparer.OrdinalIgnoreCase)
    '    Dim batchToOrders As New Dictionary(Of String, List(Of String))(StringComparer.OrdinalIgnoreCase)

    '    ' Group by weekStart then kilnType
    '    Dim wkGroups = orders.
    '        GroupBy(Function(o) o.WeekStart).
    '        OrderBy(Function(g) DateTime.ParseExact(g.Key, dateFormat, culture))

    '    For Each wkGroup In wkGroups

    '        Dim kilnGroups = wkGroup.GroupBy(Function(o) o.KilnType)

    '        For Each kilnGroup In kilnGroups

    '            ' Batch counter resets per (WeekStart + Kiln Type)
    '            Dim batchNoCounter As Integer = 0

    '            ' Track unbatched orders in this bucket
    '            Dim unbatched As New List(Of OrderInfo)(kilnGroup)

    '            '-----------------------------
    '            ' PHASE 1: Same-cycle only
    '            '-----------------------------
    '            ' Determine cycle priority order for this bucket:
    '            '   - Known priorities: 180 OR, 150 VT, 102 VT, 65 VT
    '            '   - Others: higher numeric first if a number exists, else later.
    '            Dim cycleKeys As List(Of String) = unbatched.
    '                Select(Function(o) o.CycleTypeNorm).
    '                Distinct(StringComparer.OrdinalIgnoreCase).
    '                ToList()

    '            cycleKeys.Sort(Function(a, b) CompareCyclePriority(a, b))

    '            For Each cycleKey In cycleKeys

    '                ' Same-cycle candidate list
    '                Dim sameCycle As List(Of OrderInfo) = unbatched.
    '                    Where(Function(o) o.CycleTypeNorm.Equals(cycleKey, StringComparison.OrdinalIgnoreCase)).
    '                    OrderBy(Function(o) o.ExpectedDate).
    '                    ThenByDescending(Function(o) o.Occupancy).
    '                    ToList()

    '                ' Greedy build batches for this cycle type
    '                ' Only accept if total occupancy in [min, max]
    '                While sameCycle.Count > 0

    '                    Dim batch As List(Of OrderInfo) = BuildGreedyBatch(sameCycle, maxOccupancy)

    '                    ' Compute total
    '                    Dim totalOcc As Double = batch.Sum(Function(x) x.Occupancy)

    '                    If totalOcc >= minOccupancy AndAlso totalOcc <= maxOccupancy Then
    '                        ' Accept batch
    '                        batchNoCounter += 1

    '                        Dim primaryCycleLabel As String = cycleKey ' visible cycle name
    '                        Dim visibleCycleBatch As String = primaryCycleLabel & "_" & batchNoCounter.ToString()

    '                        Dim batchInstanceKey As String =
    '                            wkGroup.Key & "|" & kilnGroup.Key & "|" & primaryCycleLabel & "|" & batchNoCounter.ToString()

    '                        ' Store metadata
    '                        batchToTotalOcc(batchInstanceKey) = totalOcc
    '                        batchToOrders(batchInstanceKey) = batch.Select(Function(x) x.OrderNo).ToList()

    '                        ' Assign to each order
    '                        For Each item In batch
    '                            orderToCycleBatch(item.OrderNo) = visibleCycleBatch
    '                            orderToBatchInstanceKey(item.OrderNo) = batchInstanceKey
    '                        Next

    '                        ' Remove assigned orders from unbatched and sameCycle
    '                        For Each item In batch
    '                            unbatched.RemoveAll(Function(o) o.OrderNo.Equals(item.OrderNo, StringComparison.OrdinalIgnoreCase))
    '                            sameCycle.RemoveAll(Function(o) o.OrderNo.Equals(item.OrderNo, StringComparison.OrdinalIgnoreCase))
    '                        Next

    '                    Else
    '                        ' Cannot form a valid batch starting with the earliest seed under this cycle.
    '                        ' Because min is REQUIRED, we stop attempting further batches for this cycle.
    '                        Exit While
    '                    End If

    '                End While
    '            Next

    '            '-----------------------------
    '            ' PHASE 2: Mixing leftovers only (restricted)
    '            '   150 VT + 102 VT (label 150 VT)
    '            '   102 VT + 65 VT  (label 102 VT)
    '            ' Priority: do 150VT-mix first to consume 102VT for that before 102VT-mix
    '            '-----------------------------

    '            ' Helper local function to get remaining list by cycle
    '            Dim rem150 = unbatched.Where(Function(o) o.CycleTypeNorm.Equals("150 VT", StringComparison.OrdinalIgnoreCase)).ToList()
    '            Dim rem102 = unbatched.Where(Function(o) o.CycleTypeNorm.Equals("102 VT", StringComparison.OrdinalIgnoreCase)).ToList()
    '            Dim rem65 = unbatched.Where(Function(o) o.CycleTypeNorm.Equals("65 VT", StringComparison.OrdinalIgnoreCase)).ToList()

    '            ' 180 OR is never mixed; others beyond these are also not mixed (per your rule)
    '            ' Attempt mixed 150 VT batches first
    '            DoMixedBatches(primaryCycle:="150 VT",
    '                           primaryList:=rem150,
    '                           secondaryCycle:="102 VT",
    '                           secondaryList:=rem102,
    '                           wkKey:=wkGroup.Key,
    '                           kilnKey:=kilnGroup.Key,
    '                           minOcc:=minOccupancy,
    '                           maxOcc:=maxOccupancy,
    '                           batchNoCounter:=batchNoCounter,
    '                           orderToCycleBatch:=orderToCycleBatch,
    '                           orderToBatchInstanceKey:=orderToBatchInstanceKey,
    '                           batchToTotalOcc:=batchToTotalOcc,
    '                           batchToOrders:=batchToOrders,
    '                           unbatched:=unbatched)

    '            ' Refresh remaining after consuming some 102 in 150-mix
    '            rem102 = unbatched.Where(Function(o) o.CycleTypeNorm.Equals("102 VT", StringComparison.OrdinalIgnoreCase)).ToList()
    '            rem65 = unbatched.Where(Function(o) o.CycleTypeNorm.Equals("65 VT", StringComparison.OrdinalIgnoreCase)).ToList()

    '            ' Attempt mixed 102 VT batches next
    '            DoMixedBatches(primaryCycle:="102 VT",
    '                           primaryList:=rem102,
    '                           secondaryCycle:="65 VT",
    '                           secondaryList:=rem65,
    '                           wkKey:=wkGroup.Key,
    '                           kilnKey:=kilnGroup.Key,
    '                           minOcc:=minOccupancy,
    '                           maxOcc:=maxOccupancy,
    '                           batchNoCounter:=batchNoCounter,
    '                           orderToCycleBatch:=orderToCycleBatch,
    '                           orderToBatchInstanceKey:=orderToBatchInstanceKey,
    '                           batchToTotalOcc:=batchToTotalOcc,
    '                           batchToOrders:=batchToOrders,
    '                           unbatched:=unbatched)

    '            ' Any remaining orders in "unbatched" are intentionally left without cycle+batch (per requirement).
    '        Next
    '    Next

    '    '=======================================================
    '    ' STEP E: Write batch results back to ALL rows per order
    '    '=======================================================
    '    For Each row As DataRow In routingTable.Rows

    '        Dim orderNo As String = row(orderNoCol).ToString().Trim()
    '        If String.IsNullOrWhiteSpace(orderNo) Then Continue For

    '        If Not orderToBatchInstanceKey.ContainsKey(orderNo) Then
    '            ' Leave blank/unassigned
    '            Continue For
    '        End If

    '        Dim batchInstanceKey As String = orderToBatchInstanceKey(orderNo)

    '        ' Visible label
    '        row(cycleBatchCol) = orderToCycleBatch(orderNo)

    '        ' Batch occupancy
    '        row(batchOccupancyCol) = batchToTotalOcc(batchInstanceKey)

    '        ' Batch orders concat
    '        row(batchOrdersCol) = String.Join("_", batchToOrders(batchInstanceKey))

    '    Next

    '    Return routingTable

    'End Function


    ''===========================================================
    '' Helper: Normalize cycle name (trim + collapse spaces)
    '' Examples:
    ''   "150  VT" -> "150 VT"
    ''   " 180 OR " -> "180 OR"
    ''===========================================================
    'Private Function NormalizeCycleName(raw As String) As String
    '    If raw Is Nothing Then Return String.Empty
    '    Dim trimmed As String = raw.Trim()
    '    ' Replace multiple whitespace with single space
    '    Return Regex.Replace(trimmed, "\s+", " ")
    'End Function

    ''===========================================================
    '' Helper: Extract first integer from a cycle string (for "other" priorities)
    '' Returns -1 if none.
    ''===========================================================
    'Private Function ExtractCycleNumber(cycleNorm As String) As Integer
    '    Dim m As Match = Regex.Match(cycleNorm, "(\d+)")
    '    If Not m.Success Then Return -1
    '    Dim val As Integer
    '    If Integer.TryParse(m.Groups(1).Value, val) Then Return val
    '    Return -1
    'End Function



    ''===========================================================
    '' Helper: Compare cycle priorities (descending)
    '' Known priority: 180 OR > 150 VT > 102 VT > 65 VT
    '' Others: higher numeric first if number exists, else later.
    ''===========================================================
    'Private Function CompareCyclePriority(a As String, b As String) As Integer

    '    Dim pa As Integer = GetExplicitPriority(a)
    '    Dim pb As Integer = GetExplicitPriority(b)

    '    ' If explicit priority exists for both, compare
    '    If pa <> pb Then Return pb.CompareTo(pa) ' higher first

    '    ' If both are "others" (same explicit priority), use numeric descending
    '    Dim na As Integer = ExtractCycleNumber(a)
    '    Dim nb As Integer = ExtractCycleNumber(b)

    '    If na <> nb Then Return nb.CompareTo(na) ' higher number first

    '    ' Finally, stable string compare
    '    Return StringComparer.OrdinalIgnoreCase.Compare(a, b)
    'End Function

    '' Explicit priorities (bigger = higher priority)
    'Private Function GetExplicitPriority(cycleNorm As String) As Integer
    '    If cycleNorm.Equals("180 OR", StringComparison.OrdinalIgnoreCase) Then Return 1000
    '    If cycleNorm.Equals("150 VT", StringComparison.OrdinalIgnoreCase) Then Return 900
    '    If cycleNorm.Equals("102 VT", StringComparison.OrdinalIgnoreCase) Then Return 800
    '    If cycleNorm.Equals("65 VT", StringComparison.OrdinalIgnoreCase) Then Return 700
    '    ' others
    '    Return 0
    'End Function

    ''===========================================================
    '' Helper: Build a greedy batch from a candidate list
    '' Strategy:
    ''   - Seed with earliest expected date item (already sorted before calling)
    ''   - Then add largest occupancy items that fit under maxOcc
    '' Returns selected batch items.
    ''===========================================================
    'Private Function BuildGreedyBatch(candidatesSorted As List(Of OrderInfo),
    '                                 maxOcc As Double) As List(Of OrderInfo)

    '    Dim batch As New List(Of OrderInfo)()
    '    If candidatesSorted Is Nothing OrElse candidatesSorted.Count = 0 Then Return batch

    '    ' Seed: earliest remaining (index 0)
    '    Dim seed As OrderInfo = candidatesSorted(0)
    '    batch.Add(seed)

    '    Dim total As Double = seed.Occupancy

    '    ' Try to fill with remaining candidates by descending occupancy
    '    Dim rest = candidatesSorted.Skip(1).OrderByDescending(Function(o) o.Occupancy).ToList()

    '    For Each c In rest
    '        If total >= maxOcc Then Exit For
    '        If total + c.Occupancy <= maxOcc Then
    '            batch.Add(c)
    '            total += c.Occupancy
    '        End If
    '    Next

    '    Return batch
    'End Function

    ''===========================================================
    '' Helper: Create mixed batches for a (primary + secondary) pairing
    ''   - Fill with primary orders first, then top up with secondary
    ''   - Only accept if total in [minOcc, maxOcc]
    ''   - Label all orders as primaryCycle_BatchNo
    ''===========================================================
    'Private Sub DoMixedBatches(primaryCycle As String,
    '                           primaryList As List(Of OrderInfo),
    '                           secondaryCycle As String,
    '                           secondaryList As List(Of OrderInfo),
    '                           wkKey As String,
    '                           kilnKey As String,
    '                           minOcc As Double,
    '                           maxOcc As Double,
    '                           ByRef batchNoCounter As Integer,
    '                           orderToCycleBatch As Dictionary(Of String, String),
    '                           orderToBatchInstanceKey As Dictionary(Of String, String),
    '                           batchToTotalOcc As Dictionary(Of String, Double),
    '                           batchToOrders As Dictionary(Of String, List(Of String)),
    '                           unbatched As List(Of OrderInfo))

    '    ' Sort primary by earliest date then higher occupancy
    '    Dim prim As List(Of OrderInfo) = primaryList.
    '        OrderBy(Function(o) o.ExpectedDate).
    '        ThenByDescending(Function(o) o.Occupancy).
    '        ToList()

    '    ' Sort secondary by higher occupancy (to top up efficiently)
    '    Dim sec As List(Of OrderInfo) = secondaryList.
    '        OrderByDescending(Function(o) o.Occupancy).
    '        ThenBy(Function(o) o.ExpectedDate).
    '        ToList()

    '    While prim.Count > 0

    '        ' Start a batch with primary-first rule
    '        Dim batch As New List(Of OrderInfo)()
    '        Dim total As Double = 0.0

    '        ' Seed with earliest primary order
    '        Dim seed As OrderInfo = prim(0)
    '        batch.Add(seed)
    '        total += seed.Occupancy
    '        prim.RemoveAt(0)

    '        ' Add more primary orders (prefer large that fit) until close to max
    '        Dim primFill = prim.OrderByDescending(Function(o) o.Occupancy).ToList()
    '        For Each p In primFill
    '            If total >= maxOcc Then Exit For
    '            If total + p.Occupancy <= maxOcc Then
    '                batch.Add(p)
    '                total += p.Occupancy
    '            End If
    '        Next
    '        ' Remove picked primaries
    '        For Each picked In batch.Where(Function(x) x.CycleTypeNorm.Equals(primaryCycle, StringComparison.OrdinalIgnoreCase)).ToList()
    '            prim.RemoveAll(Function(o) o.OrderNo.Equals(picked.OrderNo, StringComparison.OrdinalIgnoreCase))
    '        Next

    '        ' Top up with secondary if needed
    '        If total < minOcc AndAlso sec.Count > 0 Then
    '            For Each s In sec.ToList()
    '                If total >= minOcc Then Exit For
    '                If total + s.Occupancy <= maxOcc Then
    '                    batch.Add(s)
    '                    total += s.Occupancy
    '                    sec.RemoveAll(Function(o) o.OrderNo.Equals(s.OrderNo, StringComparison.OrdinalIgnoreCase))
    '                End If
    '            Next
    '        End If

    '        ' Accept only if within [min, max]
    '        If total >= minOcc AndAlso total <= maxOcc Then

    '            batchNoCounter += 1

    '            Dim visibleLabel As String = primaryCycle & "_" & batchNoCounter.ToString()
    '            Dim batchInstanceKey As String = wkKey & "|" & kilnKey & "|" & primaryCycle & "|" & batchNoCounter.ToString()

    '            batchToTotalOcc(batchInstanceKey) = total
    '            batchToOrders(batchInstanceKey) = batch.Select(Function(x) x.OrderNo).ToList()

    '            For Each item In batch
    '                orderToCycleBatch(item.OrderNo) = visibleLabel
    '                orderToBatchInstanceKey(item.OrderNo) = batchInstanceKey
    '                unbatched.RemoveAll(Function(o) o.OrderNo.Equals(item.OrderNo, StringComparison.OrdinalIgnoreCase))
    '            Next

    '        Else
    '            ' Could not form a valid mixed batch; stop trying further for this primary cycle
    '            Exit While
    '        End If

    '    End While

    'End Sub

    'Public Sub writ()
    '    Trace.WriteLine("AlgoSeq started")
    '    Debug.WriteLine("Current opRec = ")
    '    Debug.WriteLine("Resource count = ")

    'End Sub


    '' Internal helper class
    'Private Class OrderInfo
    '    Public Property OrderNo As String
    '    Public Property WeekStart As String
    '    Public Property ExpectedDate As DateTime
    '    Public Property KilnType As String
    '    Public Property CycleTypeRaw As String
    '    Public Property CycleTypeNorm As String
    '    Public Property Occupancy As Double
    'End Class


    ' NOTE:
    ' - This is a NEW function reflecting the changed requirements.
    ' - It does NOT modify your previous AddFiringWeekAndBatchColumns().
    ' - It filters Klin Type = "Batch" first and batches only within those orders.
    Public Function AddFiringWeekAndBatchColumns_V2_ByExpectedDateAndCycle(ByVal routingTable As DataTable,
                                                                      Optional ByVal baseMaxOccupancy As Double = 1.0,
                                                                      Optional ByVal maxOccBufferPct As Double = 0.2) As DataTable
        '=========================================================
        ' PURPOSE (V2):
        '   - Filter to Klin Type = "Batch"
        '   - Use Operation Number = 300 rows as authoritative per Order No
        '   - Compute "firing week" from ExpectedFiringStartDate
        '   - Batch ONLY same cycle type, FCFS by ExpectedFiringStartDate
        '   - Within the FCFS date bucket, prefer lower occupancy first (helper function)
        '   - Max occupancy per batch = 1.0 + buffer (default 20% => 1.2)
        '   - week start field = earliest ExpectedFiringStartDate among orders in that batch (NOT Monday)
        '   - Write batch results back to ALL rows for that order (like earlier Step E)
        '=========================================================

        '-------------------------------------------------------
        ' Column names (edit here ONLY if your CSV headers differ)
        '-------------------------------------------------------
        Dim expectedDateCol As String = "ExpectedFiringStartDate"
        Dim kilnTypeCol As String = "Klin Type"
        Dim cycleTypeCol As String = "Cycle Type"

        Dim orderNoCol As String = "Order No"

        ' IMPORTANT: As per your instruction going forward
        Dim operationNoCol As String = "Operation Number"

        Dim volOccCol As String = "Volume Occupancy"

        ' Output columns
        Dim firingWeekCol As String = "firing week"
        Dim weekStartCol As String = "week start"             ' V2 meaning: earliest expected date in batch
        Dim cycleBatchCol As String = "cycle+batch"
        Dim batchOccupancyCol As String = "batch+occupancy"
        Dim batchOrdersCol As String = "batch+orders"

        '-------------------------------------------------------
        ' Validate inputs
        '-------------------------------------------------------
        If routingTable Is Nothing Then Throw New ArgumentNullException(NameOf(routingTable))
        If baseMaxOccupancy <= 0 Then Throw New Exception("baseMaxOccupancy must be > 0.")
        If maxOccBufferPct < 0 Then Throw New Exception("maxOccBufferPct must be >= 0.")

        Dim maxAllowed As Double = baseMaxOccupancy * (1.0 + maxOccBufferPct)

        '-------------------------------------------------------
        ' Required column validation
        '-------------------------------------------------------
        If Not routingTable.Columns.Contains(expectedDateCol) Then
            Throw New Exception("Required column missing: " & expectedDateCol & ". Run AddExpectedFiringStartDate() first.")
        End If
        If Not routingTable.Columns.Contains(kilnTypeCol) Then Throw New Exception("Required column missing: " & kilnTypeCol)
        If Not routingTable.Columns.Contains(cycleTypeCol) Then Throw New Exception("Required column missing: " & cycleTypeCol)

        If Not routingTable.Columns.Contains(orderNoCol) Then Throw New Exception("Required column missing: " & orderNoCol)
        If Not routingTable.Columns.Contains(operationNoCol) Then Throw New Exception("Required column missing: " & operationNoCol)
        If Not routingTable.Columns.Contains(volOccCol) Then Throw New Exception("Required column missing: " & volOccCol)

        '-------------------------------------------------------
        ' Add output columns if needed
        '-------------------------------------------------------
        If Not routingTable.Columns.Contains(firingWeekCol) Then routingTable.Columns.Add(firingWeekCol, GetType(Integer))
        If Not routingTable.Columns.Contains(weekStartCol) Then routingTable.Columns.Add(weekStartCol, GetType(String))
        If Not routingTable.Columns.Contains(cycleBatchCol) Then routingTable.Columns.Add(cycleBatchCol, GetType(String))
        If Not routingTable.Columns.Contains(batchOccupancyCol) Then routingTable.Columns.Add(batchOccupancyCol, GetType(Double))
        If Not routingTable.Columns.Contains(batchOrdersCol) Then routingTable.Columns.Add(batchOrdersCol, GetType(String))

        '-------------------------------------------------------
        ' Date parsing config (dd-MM-yyyy)
        '-------------------------------------------------------
        Dim dateFormat As String = "dd-MM-yyyy"
        Dim culture As CultureInfo = CultureInfo.InvariantCulture

        ' Week-of-year config (same as earlier function for consistency)
        Dim weekRule As CalendarWeekRule = CalendarWeekRule.FirstFourDayWeek
        Dim firstDayOfWeek As DayOfWeek = DayOfWeek.Monday

        '=======================================================
        ' STEP 0: Clear derived fields (safe re-run behavior)
        '=======================================================
        For Each row As DataRow In routingTable.Rows
            row(cycleBatchCol) = String.Empty
            row(weekStartCol) = String.Empty
            row(batchOrdersCol) = String.Empty
            row(batchOccupancyCol) = DBNull.Value
            row(firingWeekCol) = DBNull.Value
        Next

        '=======================================================
        ' STEP 1: Build one order-level record per Order No from:
        '   - Klin Type = "Batch"
        '   - Operation Number = 300
        '=======================================================
        Dim orders As New List(Of OrderInfo)()
        Dim seenOrders As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

        For Each row As DataRow In routingTable.Rows

            ' Filter to only Batch kiln type orders (confirmed)
            Dim kilnType As String = row(kilnTypeCol).ToString().Trim()
            If Not kilnType.Equals("Batch", StringComparison.OrdinalIgnoreCase) Then Continue For

            ' Use only operation 300 as authoritative (confirmed)
            Dim opStr As String = row(operationNoCol).ToString().Trim()
            If Not opStr.Equals("300", StringComparison.OrdinalIgnoreCase) Then Continue For

            Dim orderNo As String = row(orderNoCol).ToString().Trim()
            If String.IsNullOrWhiteSpace(orderNo) Then Continue For
            If seenOrders.Contains(orderNo) Then Continue For

            ' Parse expected firing start date
            Dim expectedDateStr As String = row(expectedDateCol).ToString().Trim()
            If String.IsNullOrWhiteSpace(expectedDateStr) Then Continue For

            Dim expectedDate As DateTime
            If Not DateTime.TryParseExact(expectedDateStr, dateFormat, culture, DateTimeStyles.None, expectedDate) Then
                ' Skip invalid date (or throw if you want strictness)
                Continue For
            End If

            ' Parse occupancy
            Dim occStr As String = row(volOccCol).ToString().Trim()
            Dim occ As Double
            If Not Double.TryParse(occStr, NumberStyles.Any, CultureInfo.InvariantCulture, occ) Then
                Throw New Exception("Invalid occupancy value for Order No " & orderNo & " at operation 300.")
            End If

            If occ <= 0 Then
                Throw New Exception("Invalid occupancy (<=0) for Order No " & orderNo & " at operation 300.")
            End If

            ' Confirmed rule: occupancy must be <= maxAllowed
            If occ > maxAllowed Then
                Throw New Exception("Order No " & orderNo &
                                " has Volume Occupancy=" & occ.ToString(CultureInfo.InvariantCulture) &
                                " which exceeds maxAllowed=" & maxAllowed.ToString(CultureInfo.InvariantCulture) &
                                ". It cannot be batched.")
            End If

            ' Cycle type normalization
            Dim cycleRaw As String = row(cycleTypeCol).ToString().Trim()
            If String.IsNullOrWhiteSpace(cycleRaw) Then Continue For
            Dim cycleNorm As String = NormalizeCycleName(cycleRaw)

            orders.Add(New OrderInfo With {
            .OrderNo = orderNo,
            .ExpectedDate = expectedDate,
            .KilnType = kilnType,
            .CycleTypeRaw = cycleRaw,
            .CycleTypeNorm = cycleNorm,
            .Occupancy = occ
        })

            seenOrders.Add(orderNo)
        Next

        '=======================================================
        ' STEP 2: Compute firing week per order (still required)
        '=======================================================
        Dim orderToFiringWeek As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
        For Each o In orders
            Dim weekNo As Integer = culture.Calendar.GetWeekOfYear(o.ExpectedDate, weekRule, firstDayOfWeek)
            orderToFiringWeek(o.OrderNo) = weekNo
        Next

        '=======================================================
        ' STEP 3: Batching per CycleTypeNorm
        '   - FCFS by ExpectedDate (earliest date gates the flow)
        '   - Within same date bucket, prefer lower occupancy first (helper)
        '   - No min occupancy requirement
        '=======================================================
        Dim orderToCycleBatch As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
        Dim orderToBatchInstanceKey As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

        Dim batchToTotalOcc As New Dictionary(Of String, Double)(StringComparer.OrdinalIgnoreCase)
        Dim batchToOrders As New Dictionary(Of String, List(Of String))(StringComparer.OrdinalIgnoreCase)
        Dim batchToEarliestExpectedDate As New Dictionary(Of String, DateTime)(StringComparer.OrdinalIgnoreCase)

        Dim cycleGroups = orders.
        GroupBy(Function(x) x.CycleTypeNorm, StringComparer.OrdinalIgnoreCase).
        OrderBy(Function(g) g.Key, StringComparer.OrdinalIgnoreCase)

        For Each cycleGroup In cycleGroups

            Dim cycleKey As String = cycleGroup.Key
            Dim batchNoCounter As Integer = 0

            Dim remaining As New List(Of OrderInfo)(cycleGroup)

            While remaining.Count > 0

                batchNoCounter += 1

                Dim batchOrders As New List(Of OrderInfo)()
                Dim batchTotal As Double = 0.0
                Dim batchEarliest As Nullable(Of DateTime) = Nothing

                While True

                    If remaining.Count = 0 Then Exit While

                    ' FCFS gate: earliest expected date among remaining defines "what can be considered next"
                    Dim dMin As DateTime = remaining.Min(Function(r) r.ExpectedDate)

                    ' Only orders on the current earliest date are eligible at this moment (FCFS)
                    Dim sameDateCandidates As List(Of OrderInfo) =
                    remaining.Where(Function(r) r.ExpectedDate.Date = dMin.Date).ToList()

                    ' Select orders from this earliest date bucket, prioritizing smaller occupancy first.
                    Dim picked As List(Of OrderInfo) =
                    SelectOrdersWithinSameDate_FCFS_LowOccupancyFirst(sameDateCandidates, maxAllowed - batchTotal)

                    If picked IsNot Nothing AndAlso picked.Count > 0 Then

                        For Each p In picked
                            batchOrders.Add(p)
                            batchTotal += p.Occupancy

                            If Not batchEarliest.HasValue OrElse p.ExpectedDate < batchEarliest.Value Then
                                batchEarliest = p.ExpectedDate
                            End If

                            remaining.RemoveAll(Function(x) x.OrderNo.Equals(p.OrderNo, StringComparison.OrdinalIgnoreCase))
                        Next

                        If batchTotal >= maxAllowed Then Exit While

                        ' Continue: if earliest date still has remaining, we try again;
                        ' otherwise dMin moves forward naturally (still FCFS).
                        Continue While

                    Else
                        ' We cannot pick any order from the earliest date bucket:
                        ' - Either no capacity left, or remaining orders are too large for remaining capacity.
                        ' Under FCFS, we cannot skip to later dates, so close this batch.
                        Exit While
                    End If

                End While

                If batchOrders.Count = 0 Then
                    Throw New Exception("Batching failed for cycle '" & cycleKey & "'. No order could be placed into a new batch. Check occupancy values.")
                End If

                ' Create labels/keys
                Dim visibleLabel As String = cycleKey & "_" & batchNoCounter.ToString(CultureInfo.InvariantCulture)
                Dim batchInstanceKey As String = "Batch" & "|" & cycleKey & "|" & batchNoCounter.ToString(CultureInfo.InvariantCulture)

                ' Persist batch metadata
                batchToTotalOcc(batchInstanceKey) = batchTotal
                batchToOrders(batchInstanceKey) = batchOrders.Select(Function(x) x.OrderNo).ToList()
                batchToEarliestExpectedDate(batchInstanceKey) = batchEarliest.Value

                ' Assign batch to orders
                For Each o In batchOrders
                    orderToCycleBatch(o.OrderNo) = visibleLabel
                    orderToBatchInstanceKey(o.OrderNo) = batchInstanceKey
                Next

            End While

        Next

        '=======================================================
        ' STEP 4: Write back to ALL rows for each order (like earlier Step E)
        '=======================================================
        For Each row As DataRow In routingTable.Rows

            Dim orderNo As String = row(orderNoCol).ToString().Trim()
            If String.IsNullOrWhiteSpace(orderNo) Then Continue For

            ' firing week for V2-relevant orders
            If orderToFiringWeek.ContainsKey(orderNo) Then
                row(firingWeekCol) = orderToFiringWeek(orderNo)
            End If

            ' Batch fields only for orders that were batched (Batch kiln type + op300)
            If Not orderToBatchInstanceKey.ContainsKey(orderNo) Then Continue For

            Dim batchInstanceKey As String = orderToBatchInstanceKey(orderNo)

            row(cycleBatchCol) = orderToCycleBatch(orderNo)
            row(batchOccupancyCol) = batchToTotalOcc(batchInstanceKey)
            row(batchOrdersCol) = String.Join("_", batchToOrders(batchInstanceKey))
            row(weekStartCol) = batchToEarliestExpectedDate(batchInstanceKey).ToString(dateFormat, culture)

        Next

        Return routingTable

    End Function

    '===========================================================
    ' Helper (pluggable): FCFS-within-same-date selection rule
    '   - "Options available" means: multiple orders share the SAME earliest ExpectedDate.
    '   - We prefer smaller occupancy first to pack as many orders as possible.
    '
    ' IMPORTANT:
    ' - This function is intentionally isolated so you can later swap it
    '   (e.g., to a different greedy heuristic) without changing batching flow.
    '===========================================================
    Private Function SelectOrdersWithinSameDate_FCFS_LowOccupancyFirst(ByVal sameDateCandidates As List(Of OrderInfo),
                                                                  ByVal remainingCapacity As Double) As List(Of OrderInfo)

        Dim selected As New List(Of OrderInfo)()

        If sameDateCandidates Is Nothing OrElse sameDateCandidates.Count = 0 Then Return selected
        If remainingCapacity <= 0 Then Return selected

        ' Sort by smallest occupancy first to maximize count of orders in the batch
        Dim pool As List(Of OrderInfo) = sameDateCandidates.
        OrderBy(Function(o) o.Occupancy).
        ThenBy(Function(o) o.OrderNo, StringComparer.OrdinalIgnoreCase).
        ToList()

        Dim cap As Double = remainingCapacity

        For Each c In pool
            If c.Occupancy <= cap Then
                selected.Add(c)
                cap -= c.Occupancy
            End If
        Next

        Return selected

    End Function

    '===========================================================
    ' Helper: Normalize cycle name (trim + collapse spaces)
    '===========================================================
    Private Function NormalizeCycleName(raw As String) As String
        If raw Is Nothing Then Return String.Empty
        Dim trimmed As String = raw.Trim()
        Return Regex.Replace(trimmed, "\s+", " ")
    End Function

    ' Internal helper class
    Private Class OrderInfo
        Public Property OrderNo As String
        Public Property ExpectedDate As DateTime
        Public Property KilnType As String
        Public Property CycleTypeRaw As String
        Public Property CycleTypeNorm As String
        Public Property Occupancy As Double
    End Class


    '===========================================================
    ' FUNCTION 4 : AddPressingFields
    '
    ' PURPOSE:
    '   Enrich an existing DataTable (output of AddFiringWeekAndBatchColumns)
    '   with pressing planning fields:
    '     1) expectedpressingstart = [week start] - [pressing buffer]
    '     2) pressing week         = week number of expectedpressingstart
    '     3) pressing week start   = Monday date of pressing week
    '
    ' INPUT:
    '   routingTable              : DataTable already containing "week start"
    '   pressingBufferColumnName  : column holding pressing buffer (days)
    '                               Default = "pressing buffer"
    '
    ' OUTPUT:
    '   Updated DataTable with new columns populated
    '
    ' NOTES:
    '   - Date format assumed "dd-MM-yyyy"
    '   - Week is computed using Monday as first day and FirstFourDayWeek rule
    '   - Pressing buffer assumed integer number of DAYS
    '===========================================================
    Public Function AddPressingFields(routingTable As DataTable,
                                     Optional pressingBufferColumnName As String = "pressing buffer") As DataTable

        '----------------------------
        ' Required input columns
        '----------------------------
        Dim weekStartCol As String = "week start"

        '----------------------------
        ' Output columns to add
        '----------------------------
        Dim expectedPressingStartCol As String = "expectedpressingstart"
        Dim pressingWeekCol As String = "pressing week"
        Dim pressingWeekStartCol As String = "pressing week start"

        If routingTable Is Nothing Then Throw New ArgumentNullException(NameOf(routingTable))

        ' Validate presence of required columns
        If Not routingTable.Columns.Contains(weekStartCol) Then
            Throw New Exception("Required column missing: " & weekStartCol & ". Run AddFiringWeekAndBatchColumns() first.")
        End If

        If Not routingTable.Columns.Contains(pressingBufferColumnName) Then
            Throw New Exception("Required column missing: " & pressingBufferColumnName)
        End If

        ' Add output columns if they don't exist
        If Not routingTable.Columns.Contains(expectedPressingStartCol) Then
            routingTable.Columns.Add(expectedPressingStartCol, GetType(String))
        End If

        If Not routingTable.Columns.Contains(pressingWeekCol) Then
            routingTable.Columns.Add(pressingWeekCol, GetType(Integer))
        End If

        If Not routingTable.Columns.Contains(pressingWeekStartCol) Then
            routingTable.Columns.Add(pressingWeekStartCol, GetType(String))
        End If

        ' Date parsing configuration
        Dim dateFormat As String = "dd-MM-yyyy"
        Dim culture As CultureInfo = CultureInfo.InvariantCulture

        ' Week calculation configuration (consistent with earlier function) 
        Dim weekRule As CalendarWeekRule = CalendarWeekRule.FirstFourDayWeek
        Dim firstDayOfWeek As DayOfWeek = DayOfWeek.Monday

        ' Iterate rows and compute new fields
        For Each row As DataRow In routingTable.Rows

            ' Default outputs
            row(expectedPressingStartCol) = String.Empty
            row(pressingWeekStartCol) = String.Empty
            row(pressingWeekCol) = DBNull.Value

            ' Read and parse "week start"
            Dim weekStartStr As String = row(weekStartCol).ToString().Trim()
            If String.IsNullOrWhiteSpace(weekStartStr) Then
                Continue For
            End If

            Dim weekStartDate As DateTime
            If Not DateTime.TryParseExact(weekStartStr, dateFormat, culture, DateTimeStyles.None, weekStartDate) Then
                Continue For
            End If

            ' Read and parse pressing buffer (assumed days)
            Dim bufStr As String = row(pressingBufferColumnName).ToString().Trim()
            If String.IsNullOrWhiteSpace(bufStr) Then
                Continue For
            End If

            Dim pressingBufferDays As Integer
            If Not Integer.TryParse(bufStr, pressingBufferDays) Then
                Continue For
            End If

            ' 1) expectedpressingstart = week start - pressing buffer days
            Dim expectedPressingStart As DateTime = weekStartDate.AddDays(-pressingBufferDays)
            row(expectedPressingStartCol) = expectedPressingStart.ToString(dateFormat)

            ' 2) pressing week = week number for expectedpressingstart
            Dim pWeekNo As Integer = culture.Calendar.GetWeekOfYear(expectedPressingStart, weekRule, firstDayOfWeek)
            row(pressingWeekCol) = pWeekNo

            ' 3) pressing week start = Monday of expectedpressingstart's week
            Dim delta As Integer = (CInt(expectedPressingStart.DayOfWeek) - CInt(DayOfWeek.Monday) + 7) Mod 7
            Dim pressingWeekStart As DateTime = expectedPressingStart.AddDays(-delta)
            row(pressingWeekStartCol) = pressingWeekStart.ToString(dateFormat)

        Next

        Return routingTable

    End Function

    Public Function CreatePressingBatches_PseudoSchedule(ByVal routingTable As DataTable,
                                                     ByVal maxTonnagePerResourceGroupPerDay As Decimal,
                                                     Optional ByVal lookaheadDays As Integer = 2,
                                                     Optional ByVal cooldownDays As Integer = 2,
                                                     Optional ByVal attributeOperationNumber As String = "300",
                                                     Optional ByVal kilnTypeBatchValue As String = "Batch",
                                                     Optional ByVal firingThresholdDateField As String = "ExpectedFiringStartDate") As DataTable
        '=========================================================
        ' PURPOSE:
        '   Day-by-day pseudo-scheduling to create PRESSING batches
        '   (for Klin Type = Batch) with changeover reduction keys:
        '       Resource Group + Wheel Dia + Wheel thickness
        '
        ' KEY RULES:
        '   - Use expectedpressingstart as the "press-by" signal (with lookahead)
        '   - HARD CONSTRAINT: never press after firing threshold date
        '       scheduledDay <= thresholdDate  (date-only comparison)
        '   - Capacity constraint per Resource Group per day: max tonnage
        '   - Cooldown per Resource Group:
        '       after running (Dia,Thk) on day D, cannot repeat D+1 and D+2
        '       next allowed is D + (cooldownDays + 1)
        '   - Prefer "campaign" behavior: fill chosen combo as much as possible,
        '     but allow additional combos the same day if capacity remains.
        '
        ' OUTPUT COLUMNS (written back to ALL rows per order):
        '   - pressing+batch        : "<WheelDia>_<globalNumber>"
        '   - pressing batch+orders : "order1_order2_..."
        '   - pressing batch+date   : "dd-MM-yyyy" (target pressing day)
        '=========================================================

        '-------------------------------------------------------
        ' Validate inputs
        '-------------------------------------------------------
        If routingTable Is Nothing Then Throw New ArgumentNullException(NameOf(routingTable))
        If maxTonnagePerResourceGroupPerDay <= 0D Then Throw New Exception("maxTonnagePerResourceGroupPerDay must be > 0.")
        If lookaheadDays < 0 Then Throw New Exception("lookaheadDays must be >= 0.")
        If cooldownDays < 0 Then Throw New Exception("cooldownDays must be >= 0.")
        If String.IsNullOrWhiteSpace(attributeOperationNumber) Then Throw New Exception("attributeOperationNumber is required.")
        If String.IsNullOrWhiteSpace(kilnTypeBatchValue) Then Throw New Exception("kilnTypeBatchValue is required.")
        If String.IsNullOrWhiteSpace(firingThresholdDateField) Then Throw New Exception("firingThresholdDateField is required.")

        '-------------------------------------------------------
        ' Column names (edit ONLY if headers differ)
        '-------------------------------------------------------
        Dim orderNoCol As String = "Order No"
        Dim operationNoCol As String = "Operation Number"
        Dim kilnTypeCol As String = "Klin Type"

        Dim resourceGroupCol As String = "Resource Group"
        Dim wheelDiaCol As String = "Wheel Dia"
        Dim wheelThkCol As String = "Wheel thickness"

        Dim expectedPressingStartCol As String = "expectedpressingstart"
        Dim tonnageCol As String = "Tonnage"

        ' Firing threshold is parameterized (default ExpectedFiringStartDate; can switch to "week start" later)
        Dim firingThresholdCol As String = firingThresholdDateField

        ' Output columns
        Dim outBatchCol As String = "pressing+batch"
        Dim outBatchOrdersCol As String = "pressing batch+orders"
        Dim outBatchDateCol As String = "pressing batch+date"

        '-------------------------------------------------------
        ' Required column validation
        '-------------------------------------------------------
        Dim requiredCols As String() = {
        orderNoCol, operationNoCol, kilnTypeCol,
        resourceGroupCol, wheelDiaCol, wheelThkCol,
        expectedPressingStartCol, tonnageCol, firingThresholdCol
    }

        For Each c In requiredCols
            If Not routingTable.Columns.Contains(c) Then
                Throw New Exception("Required column missing: " & c)
            End If
        Next

        '-------------------------------------------------------
        ' Add output columns if missing
        '-------------------------------------------------------
        If Not routingTable.Columns.Contains(outBatchCol) Then routingTable.Columns.Add(outBatchCol, GetType(String))
        If Not routingTable.Columns.Contains(outBatchOrdersCol) Then routingTable.Columns.Add(outBatchOrdersCol, GetType(String))
        If Not routingTable.Columns.Contains(outBatchDateCol) Then routingTable.Columns.Add(outBatchDateCol, GetType(String))

        '-------------------------------------------------------
        ' Date parsing (dd-MM-yyyy) consistent with your project
        '-------------------------------------------------------
        Dim dateFormat As String = "dd-MM-yyyy"
        Dim culture As CultureInfo = CultureInfo.InvariantCulture

        '=======================================================
        ' STEP 0: Clear existing pressing outputs (safe re-run)
        '=======================================================
        For Each row As DataRow In routingTable.Rows
            row(outBatchCol) = String.Empty
            row(outBatchOrdersCol) = String.Empty
            row(outBatchDateCol) = String.Empty
        Next

        '=======================================================
        ' STEP 1: Build order-level records from:
        '   - Klin Type = Batch
        '   - Operation Number = attributeOperationNumber (default 300)
        '=======================================================
        Dim orders As New List(Of PressOrder)()
        Dim seenOrders As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

        For Each row As DataRow In routingTable.Rows

            Dim kilnType As String = row(kilnTypeCol).ToString().Trim()
            If Not kilnType.Equals(kilnTypeBatchValue, StringComparison.OrdinalIgnoreCase) Then Continue For

            Dim opStr As String = row(operationNoCol).ToString().Trim()
            If Not opStr.Equals(attributeOperationNumber, StringComparison.OrdinalIgnoreCase) Then Continue For

            Dim orderNo As String = row(orderNoCol).ToString().Trim()
            If String.IsNullOrWhiteSpace(orderNo) Then Continue For
            If seenOrders.Contains(orderNo) Then Continue For

            Dim rg As String = row(resourceGroupCol).ToString().Trim()
            If String.IsNullOrWhiteSpace(rg) Then Continue For

            ' Parse expected pressing start date
            Dim epsStr As String = row(expectedPressingStartCol).ToString().Trim()
            Dim eps As DateTime
            If String.IsNullOrWhiteSpace(epsStr) Then Continue For
            If Not DateTime.TryParseExact(epsStr, dateFormat, culture, DateTimeStyles.None, eps) Then Continue For

            ' Parse firing threshold date (hard constraint day <= threshold)
            Dim thrStr As String = row(firingThresholdCol).ToString().Trim()
            Dim thr As DateTime
            If String.IsNullOrWhiteSpace(thrStr) Then Continue For
            If Not DateTime.TryParseExact(thrStr, dateFormat, culture, DateTimeStyles.None, thr) Then Continue For

            ' Parse numeric Dia/Thickness (exact match; normalized for keys/labels)
            Dim dia As Decimal
            If Not Decimal.TryParse(row(wheelDiaCol).ToString().Trim(), NumberStyles.Any, CultureInfo.InvariantCulture, dia) Then Continue For

            Dim thk As Decimal
            If Not Decimal.TryParse(row(wheelThkCol).ToString().Trim(), NumberStyles.Any, CultureInfo.InvariantCulture, thk) Then Continue For

            ' Parse tonnage (Decimal to avoid floating-point drift)
            Dim ton As Decimal
            If Not Decimal.TryParse(row(tonnageCol).ToString().Trim(), NumberStyles.Any, CultureInfo.InvariantCulture, ton) Then
                Throw New Exception("Invalid tonnage for Order No " & orderNo & " at operation " & attributeOperationNumber & ".")
            End If
            If ton <= 0D Then Throw New Exception("Invalid tonnage (<=0) for Order No " & orderNo & ".")

            ' If an order tonnage exceeds daily capacity, it can never be scheduled under this model.
            ' We fail fast because it indicates data/config mismatch.
            If ton > maxTonnagePerResourceGroupPerDay Then
                Throw New Exception("Order No " & orderNo & " tonnage=" & ton.ToString(CultureInfo.InvariantCulture) &
                                " exceeds daily maxTonnage=" & maxTonnagePerResourceGroupPerDay.ToString(CultureInfo.InvariantCulture) &
                                " for Resource Group " & rg & ". Cannot schedule.")
            End If

            orders.Add(New PressOrder With {
            .OrderNo = orderNo,
            .ResourceGroup = rg,
            .ExpectedPressingStart = eps.Date,   ' date-only semantics
            .FiringThreshold = thr.Date,         ' date-only semantics
            .WheelDia = dia,
            .WheelThickness = thk,
            .Tonnage = ton
        })

            seenOrders.Add(orderNo)
        Next

        If orders.Count = 0 Then
            ' Nothing to do; return table with cleared outputs.
            Return routingTable
        End If

        '=======================================================
        ' STEP 2: Planning horizon
        '   Start = min(expectedpressingstart)
        '   End   = max(firingThreshold)
        '=======================================================
        Dim planStart As DateTime = orders.Min(Function(o) o.ExpectedPressingStart)
        Dim planEnd As DateTime = orders.Max(Function(o) o.FiringThreshold)

        '=======================================================
        ' STEP 3: Pseudo-schedule per Resource Group, day by day
        '=======================================================

        ' Global batch counter across the whole run (your requirement)
        Dim globalBatchCounter As Integer = 0

        ' Track assigned batch per order
        Dim orderToBatchId As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
        Dim orderToBatchDate As New Dictionary(Of String, DateTime)(StringComparer.OrdinalIgnoreCase)
        Dim batchIdToOrders As New Dictionary(Of String, List(Of String))(StringComparer.OrdinalIgnoreCase)

        ' Group orders by RG for independent planning
        Dim rgGroups = orders.GroupBy(Function(o) o.ResourceGroup, StringComparer.OrdinalIgnoreCase)

        For Each rgGroup In rgGroups

            Dim rg As String = rgGroup.Key
            Dim unscheduled As New List(Of PressOrder)(rgGroup)

            ' Cooldown: comboKey -> nextAllowedDate (inclusive)
            ' If day < nextAllowedDate, the combo is blocked.
            Dim comboNextAllowed As New Dictionary(Of String, DateTime)(StringComparer.OrdinalIgnoreCase)

            Dim day As DateTime = planStart
            While day <= planEnd AndAlso unscheduled.Count > 0

                Dim remainingCapacity As Decimal = maxTonnagePerResourceGroupPerDay

                ' Try to create one or more batches today for this RG (campaign-first, then additional combos if capacity remains)
                While remainingCapacity > 0D AndAlso unscheduled.Count > 0

                    ' Build eligible window end for pull-ahead:
                    ' eligible if expectedpressingstart <= day + lookaheadDays
                    Dim windowEnd As DateTime = day.AddDays(lookaheadDays)

                    ' Eligible orders:
                    '  - Not yet scheduled
                    '  - Within pull-ahead window
                    '  - HARD: day <= firingThreshold (never press after threshold)
                    Dim eligible As List(Of PressOrder) = unscheduled.
                    Where(Function(o) o.ExpectedPressingStart <= windowEnd AndAlso day <= o.FiringThreshold).
                    ToList()

                    If eligible.Count = 0 Then Exit While

                    ' Group eligible by combo (Dia+Thk). Cooldown is per RG.
                    Dim comboGroups = eligible.
                    GroupBy(Function(o) MakeComboKey(o.WheelDia, o.WheelThickness), StringComparer.OrdinalIgnoreCase).
                    ToList()

                    ' Filter out combos blocked by cooldown
                    Dim availableCombos = comboGroups.
                    Where(Function(g)
                              Dim key As String = g.Key
                              If Not comboNextAllowed.ContainsKey(key) Then Return True
                              ' combo is allowed if day >= nextAllowedDate
                              Return day >= comboNextAllowed(key)
                          End Function).
                    ToList()

                    If availableCombos.Count = 0 Then
                        ' Nothing can run today due to cooldown blocks
                        Exit While
                    End If

                    '---------------------------------------------------------
                    ' Select the next combo to run today.
                    '
                    ' We prioritize protecting firing dates:
                    '   1) earliest firing threshold among orders in the combo
                    '   2) then earliest expected pressing start among orders in the combo
                    '
                    ' This keeps us aligned with "do not miss firing".
                    '---------------------------------------------------------
                    Dim chosen = availableCombos.
                    OrderBy(Function(g) g.Min(Function(o) o.FiringThreshold)).
                    ThenBy(Function(g) g.Min(Function(o) o.ExpectedPressingStart)).
                    ThenByDescending(Function(g) g.Sum(Function(o) o.Tonnage)). ' tie-break: more fill potential
                    First()

                    Dim chosenComboKey As String = chosen.Key

                    '---------------------------------------------------------
                    ' Build today's batch for this chosen combo:
                    ' Fill as many orders as possible until capacity is hit,
                    ' respecting earliest expectedpressingstart first (FCFS).
                    '---------------------------------------------------------
                    Dim chosenOrders As List(Of PressOrder) =
                    chosen.OrderBy(Function(o) o.ExpectedPressingStart).
                           ThenBy(Function(o) o.FiringThreshold).
                           ThenBy(Function(o) o.Tonnage). ' small tonnage earlier can pack more
                           ToList()

                    Dim batch As New List(Of PressOrder)()
                    Dim batchTonnage As Decimal = 0D

                    For Each o In chosenOrders
                        If batchTonnage + o.Tonnage <= remainingCapacity Then
                            batch.Add(o)
                            batchTonnage += o.Tonnage
                        End If
                    Next

                    If batch.Count = 0 Then
                        ' No order from chosen combo fits remaining capacity.
                        ' Since we allow multiple combos/day, try another combo that might fit.
                        ' We remove this chosen combo from consideration and continue.
                        availableCombos.RemoveAll(Function(g) g.Key.Equals(chosenComboKey, StringComparison.OrdinalIgnoreCase))
                        If availableCombos.Count = 0 Then Exit While

                        ' Recompute chosen from remaining combos
                        Dim alt = availableCombos.
                        OrderBy(Function(g) g.Min(Function(o) o.FiringThreshold)).
                        ThenBy(Function(g) g.Min(Function(o) o.ExpectedPressingStart)).
                        ThenByDescending(Function(g) g.Sum(Function(o) o.Tonnage)).
                        First()

                        chosen = alt
                        chosenComboKey = chosen.Key

                        ' Build batch again for alt chosen
                        chosenOrders = chosen.OrderBy(Function(o) o.ExpectedPressingStart).
                                        ThenBy(Function(o) o.FiringThreshold).
                                        ThenBy(Function(o) o.Tonnage).
                                        ToList()

                        batch.Clear()
                        batchTonnage = 0D

                        For Each o In chosenOrders
                            If batchTonnage + o.Tonnage <= remainingCapacity Then
                                batch.Add(o)
                                batchTonnage += o.Tonnage
                            End If
                        Next

                        If batch.Count = 0 Then
                            ' Still nothing fits; stop for today for this RG.
                            Exit While
                        End If
                    End If

                    '---------------------------------------------------------
                    ' Assign a new global batch id: "<WheelDia>_<globalNumber>"
                    ' WheelDia formatting: treat as integer-like (500 not 500.0)
                    '---------------------------------------------------------
                    globalBatchCounter += 1
                    Dim diaLabel As String = FormatDiaForBatchId(batch(0).WheelDia)
                    Dim batchId As String = diaLabel & "_" & globalBatchCounter.ToString(CultureInfo.InvariantCulture)

                    ' Persist batch order list (for writing batch+orders)
                    Dim batchOrderNos As List(Of String) = batch.Select(Function(x) x.OrderNo).ToList()
                    batchIdToOrders(batchId) = batchOrderNos

                    ' Assign each order to this batch and date
                    For Each o In batch
                        orderToBatchId(o.OrderNo) = batchId
                        orderToBatchDate(o.OrderNo) = day
                    Next

                    ' Remove batched orders from unscheduled
                    For Each o In batch
                        unscheduled.RemoveAll(Function(x) x.OrderNo.Equals(o.OrderNo, StringComparison.OrdinalIgnoreCase))
                    Next

                    ' Apply cooldown for this RG & combo:
                    ' If cooldownDays=2 and day=D, next allowed is D+3 (D + cooldownDays + 1)
                    comboNextAllowed(chosenComboKey) = day.AddDays(cooldownDays + 1)

                    ' Consume capacity
                    remainingCapacity -= batchTonnage

                    ' Continue to see if we can schedule another combo today with remaining capacity
                End While

                day = day.AddDays(1)
            End While

            ' Any remaining unscheduled orders in this RG are left blank (acceptable per your spec).
        Next

        '=======================================================
        ' STEP 4: Write results back to ALL rows per order
        '=======================================================
        For Each row As DataRow In routingTable.Rows

            Dim orderNo As String = row(orderNoCol).ToString().Trim()
            If String.IsNullOrWhiteSpace(orderNo) Then Continue For

            If Not orderToBatchId.ContainsKey(orderNo) Then
                ' Leave blank/unassigned for this order
                Continue For
            End If

            Dim batchId As String = orderToBatchId(orderNo)
            row(outBatchCol) = batchId

            ' Batch orders list as order1_order2...
            row(outBatchOrdersCol) = String.Join("_", batchIdToOrders(batchId))

            ' Target pressing date
            row(outBatchDateCol) = orderToBatchDate(orderNo).ToString(dateFormat, culture)

        Next

        Return routingTable

    End Function

    '===========================================================
    ' Helper: combo key for cooldown and grouping
    '===========================================================
    Private Function MakeComboKey(ByVal wheelDia As Decimal, ByVal wheelThk As Decimal) As String
        ' Use invariant formatting and normalize integer-like numbers
        Return NormalizeDecimalKey(wheelDia) & "|" & NormalizeDecimalKey(wheelThk)
    End Function

    '===========================================================
    ' Helper: normalize decimals into stable keys (e.g., 500.0 -> "500")
    '===========================================================
    Private Function NormalizeDecimalKey(ByVal value As Decimal) As String
        ' "G29" keeps precision without scientific notation for typical business decimals
        Dim s As String = value.ToString("G29", CultureInfo.InvariantCulture)

        ' Strip trailing ".0" if present (just in case)
        If s.Contains("."c) Then
            s = s.TrimEnd("0"c).TrimEnd("."c)
        End If

        Return s
    End Function

    '===========================================================
    ' Helper: format WheelDia for batch id: "500" not "500.0"
    '===========================================================
    Private Function FormatDiaForBatchId(ByVal wheelDia As Decimal) As String
        Return NormalizeDecimalKey(wheelDia)
    End Function

    '===========================================================
    ' Internal order model (kept minimal and explicit)
    '===========================================================
    Private Class PressOrder
        Public Property OrderNo As String
        Public Property ResourceGroup As String
        Public Property ExpectedPressingStart As DateTime  ' date-only
        Public Property FiringThreshold As DateTime        ' date-only (hard upper bound)
        Public Property WheelDia As Decimal
        Public Property WheelThickness As Decimal
        Public Property Tonnage As Decimal
    End Class


End Class





