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


    ' NOTE:
    ' - This is a NEW function reflecting the changed requirements.
    ' - It does NOT modify your previous AddFiringWeekAndBatchColumns().
    ' - It filters Kiln Type = "Batch" first and batches only within those orders.
    Public Function AddFiringWeekAndBatchColumns_V2_ByExpectedDateAndCycle(ByVal routingTable As DataTable,
                                                                      Optional ByVal baseMaxOccupancy As Double = 1.0,
                                                                      Optional ByVal maxOccBufferPct As Double = 0.2) As DataTable
        '=========================================================
        ' PURPOSE (V2):
        '   - Filter to Kiln Type = "Batch"
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
        Dim kilnTypeCol As String = "Kiln Type"
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
        '   - Kiln Type = "Batch"
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
        '   (for Kiln Type = Batch) with changeover reduction keys:
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
        Dim kilnTypeCol As String = "Kiln Type"

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
        '   - Kiln Type = Batch
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





