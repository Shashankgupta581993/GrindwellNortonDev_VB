Public Class pressingOptimizer_vf
    ' Pressing has two separate decisions:
    '   1. Rank eligible orders by the established WIP/date rules (this class).
    '   2. Choose and preserve die-ring campaigns on actual presses (the scheduler).
    ' Keeping them separate avoids confusing a queue sort with a machine plan.
    ' Legacy public callers still receive parent records; runPress uses the actual
    ' operation record plus these attributes so each placement can be validated.

    Friend Class Candidate
        Public Property RecordNo As Integer
        Public Property OrderNo As String
        Public Property OperationNumber As Integer
        Public Property ParentRecord As Integer          ' queue key to return
        Public Property ResourceGroup As String          ' resource group key
        Public Property Earliest As DateTime             ' Pressing earliest start (date-only)
        Public Property Due As DateTime                  ' Pressing Due date (date-only)
        Public Property Tier As Integer                  ' 0=approaching, 1=late, 2=other
        Public Property WheelDia As String
        Public Property WheelPin As String               ' Bore/pin; legacy table column is called Wheel thickness
        Public Property TypeKey As String                ' Legacy final rank tie-break: ring|pin, NOT the campaign key
        Public Property CycleRank As Integer
        Public Property MissingEarliest As Boolean
        Public Property MissingDue As Boolean

        Public Property PrevOpIsScheduled As Boolean     ' flag from prev_op_is_scheduled
        Public Property PrevOpPriority As Integer        ' 0 if PrevOpIsScheduled, 1 otherwise
        Public Property WipScore As Integer
        Public Property WipReadyTime As DateTime
        Public Property WipStatus As String
        Public Property WipRejectReason As String
    End Class

    ' ----------------------
    ' Pressing 200 queue builder
    ' ----------------------
    Public Function BuildPressing200Queue(dt As DataTable,
                                      currentDate As DateTime,
                                      Optional approachingDays As Integer = 2,
                                          Optional prioritizePrevOpFirst As Boolean = False,
                                          Optional debug As SchedulerDebugCollector = Nothing) As List(Of Integer)

        ' Retain the original public entrypoint for callers that need parent records.
        ' Actual pressing now uses the ranked operation candidates below; campaign
        ' continuity is decided on the chosen machine, not by shuffling this queue.
        Return BuildRankedCandidates(dt, currentDate, approachingDays,
                                     prioritizePrevOpFirst, debug).
            Select(Function(c) c.ParentRecord).Distinct().ToList()
    End Function

    Friend Function BuildRankedCandidates(dt As DataTable,
                                         currentDate As DateTime,
                                         Optional approachingDays As Integer = 2,
                                         Optional prioritizePrevOpFirst As Boolean = False,
                                         Optional debug As SchedulerDebugCollector = Nothing) As List(Of Candidate)

        If dt Is Nothing Then Throw New ArgumentNullException(NameOf(dt))

        ' ---- Required columns (these must exist in dt) ----
        ' NOTE: Keep these names EXACTLY matching your DataTable column names.
        SharedHelpers.RequireColumn(dt, "parent_record")
        SharedHelpers.RequireColumn(dt, "is_scheduled")
        SharedHelpers.RequireColumn(dt, "prev_op_is_scheduled")
        SharedHelpers.RequireColumn(dt, "wip_score")
        SharedHelpers.RequireColumn(dt, "wip_ready_time")
        SharedHelpers.RequireColumn(dt, "wip_status")
        SharedHelpers.RequireColumn(dt, "wip_reject_reason")
        SharedHelpers.RequireColumn(dt, "wip_prev_op_rec")
        SharedHelpers.RequireColumn(dt, "wip_prev_op_released")
        SharedHelpers.RequireColumn(dt, "Resource Group")
        SharedHelpers.RequireColumn(dt, "Operation Number")
        SharedHelpers.RequireColumn(dt, "Pressing earliest start")
        SharedHelpers.RequireColumn(dt, "Pressing Due date")
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
            If opNo <> 200 Then Continue For

            ' 2) Filter: STRICT UNSCHEDULED only
            Dim isScheduled As Boolean = SharedHelpers.SafeBool(r("is_scheduled"))
            If isScheduled Then Continue For
            ' A completed/actualized operation must never re-enter the campaign
            ' scheduler merely because it has no current planning-board allocation.
            If SharedHelpers.IsCompletedOrActualizedRow(r) Then Continue For

            ' 3) Queue key: parent_record must be valid
            Dim parentRec As Integer = SharedHelpers.SafeInt(r("parent_record"))
            If parentRec <= 0 Then Continue For

            ' 4) Read Resource Group
            Dim resourceGroup As String = SharedHelpers.SafeStr(r("Resource Group")).Trim()
            ' If you ever want to filter by specific groups, you could do it here.

            ' 5) Read pressing dates (date-only logic)
            Dim earliest As DateTime = SharedHelpers.SafeDate(r("Pressing earliest start")).Date
            Dim due As DateTime = SharedHelpers.SafeDate(r("Pressing Due date")).Date

            ' Treat missing/parse-failed dates as MinValue (consistent with SafeDate behavior)
            Dim missingEarliest As Boolean = (earliest = DateTime.MinValue)
            Dim missingDue As Boolean = (due = DateTime.MinValue)

            ' 6) Tiering logic:
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

            ' 7) Read batching/type attributes
            Dim wheelDia As String = SharedHelpers.SafeStr(r("Wheel Dia")).Trim()
            Dim wheelPin As String = SharedHelpers.SafeStr(r("Wheel thickness")).Trim()
            Dim cycleType As String = SharedHelpers.SafeStr(r("Cycle Type")).Trim()

            ' 8) Cycle ranking
            Dim cycleRank As Integer = GetCycleRank(cycleType)

            ' 9) Previous operation scheduled? (WIP reduction signal)
            Dim wipStatus As String = SharedHelpers.SafeStr(r("wip_status")).Trim()
            If Not wipStatus.Equals("Candidate", StringComparison.OrdinalIgnoreCase) Then Continue For

            Dim prevOpIsScheduled As Boolean = SharedHelpers.SafeBool(r("wip_prev_op_scheduled"))
            Dim prevOpPriority As Integer = If(prevOpIsScheduled, 0, 1)

            Dim wipScore As Integer = SharedHelpers.SafeInt(r("wip_score"))
            Dim wipReadyTime As DateTime = SharedHelpers.SafeDate(r("wip_ready_time"))
            Dim wipRejectReason As String = SharedHelpers.SafeStr(r("wip_reject_reason"))
            ' The common snapshot marks some rows Candidate even when their
            ' predecessor is not released (other stages use that broader pool).
            ' Pressing must independently enforce readiness. A first routing
            ' operation has no predecessor; otherwise a valid release is required.
            If SharedHelpers.SafeInt(r("wip_prev_op_rec")) > 0 AndAlso
               (Not SharedHelpers.SafeBool(r("wip_prev_op_released")) OrElse wipReadyTime = DateTime.MinValue) Then
                If debug IsNot Nothing AndAlso debug.Enabled Then
                    debug.TraceCandidateStep(New OptimizerCandidateTraceRow With {
                        .OptimizerName = "pressingOptimizer_vf", .Stage = "Pressing",
                        .StepName = "Readiness", .RecordNo = SharedHelpers.SafeInt(r("OrdersID")),
                        .OrderNo = SharedHelpers.SafeStr(r("Order No")), .OperationNumber = 200,
                        .Included = False, .ReasonCode = "PRESSING_PREV_OP_NOT_READY",
                        .ReasonDetail = "Predecessor is unreleased or its release time is missing."})
                End If
                Continue For
            End If
            ' 0 = “prev op is already scheduled” → better
            ' 1 = “prev op not scheduled / unknown” → slightly worse

            candidates.Add(New Candidate With {
            .RecordNo = SharedHelpers.SafeInt(r("OrdersID")),
            .OrderNo = SharedHelpers.SafeStr(r("Order No")).Trim(),
            .OperationNumber = opNo,
            .ParentRecord = parentRec,
            .ResourceGroup = resourceGroup,
            .Earliest = earliest,
            .Due = due,
            .Tier = tier,
            .WheelDia = wheelDia,
            .WheelPin = wheelPin,
            .TypeKey = wheelDia & "|" & wheelPin,
            .CycleRank = cycleRank,
            .MissingEarliest = missingEarliest,
            .MissingDue = missingDue,
            .PrevOpIsScheduled = prevOpIsScheduled,
            .PrevOpPriority = prevOpPriority,
            .WipScore = wipScore,
            .WipReadyTime = wipReadyTime,
            .WipStatus = wipStatus,
            .WipRejectReason = wipRejectReason
        })
        Next

        ' ---- Sorting: the exact queue priority order (within Resource Groups) ----
        ' The campaign scheduler uses this rank to seed new campaigns and protect urgent work.
        ' Order:
        ' 1) Tier asc (0,1,2)
        ' 2) Due asc (missing due goes last)
        ' 3) PrevOpPriority asc (0=prev op scheduled first)
        ' 4) Earliest asc (missing earliest goes last)
        ' 5) CycleRank desc (higher rank first)
        ' 6) TypeKey asc (clusters same Dia|Thickness)
        ' 7) ParentRecord asc (stable tie-breaker)
        Dim sorted As List(Of Candidate)

        If prioritizePrevOpFirst Then
            ' WIP-first strategy
            sorted = candidates.OrderByDescending(Function(c) c.WipScore) _
                   .ThenBy(Function(c) c.Tier) _
                   .ThenBy(Function(c) If(c.MissingDue, DateTime.MaxValue, c.Due)) _
                   .ThenBy(Function(c) If(c.MissingEarliest, DateTime.MaxValue, c.Earliest)) _
                   .ThenByDescending(Function(c) c.CycleRank) _
                   .ThenBy(Function(c) c.TypeKey) _
                   .ThenBy(Function(c) c.ParentRecord) _
                   .ToList()
        Else
            ' Due-date-first strategy (what we designed earlier)
            sorted = candidates.OrderBy(Function(c) c.Tier) _
                           .ThenBy(Function(c) If(c.MissingDue, DateTime.MaxValue, c.Due)) _
                           .ThenBy(Function(c) c.PrevOpPriority) _
                           .ThenBy(Function(c) If(c.MissingEarliest, DateTime.MaxValue, c.Earliest)) _
                           .ThenByDescending(Function(c) c.CycleRank) _
                           .ThenBy(Function(c) c.TypeKey) _
                           .ThenBy(Function(c) c.ParentRecord) _
                           .ToList()
        End If


        ' Preserve WIP/date priority as the seed ranking. Do NOT pull forward an
        ' exact ring+pin combination here: that old shortcut fragmented ring runs
        ' across due dates and could not enforce machine-level campaign continuity.
        Dim batched As List(Of Candidate) = sorted

        If debug IsNot Nothing AndAlso debug.Enabled Then
            Dim beforeCount As Integer = dt.AsEnumerable().Count(Function(r) SharedHelpers.SafeInt(r("Operation Number")) = 200)
            For i As Integer = 0 To batched.Count - 1
                Dim c As Candidate = batched(i)
                debug.TraceCandidateStep(New OptimizerCandidateTraceRow With {
                    .OptimizerName = "pressingOptimizer_vf",
                    .Stage = "Pressing",
                    .StepName = "FinalRankedQueue",
                    .OrderNo = c.OrderNo,
                    .ParentRecordNo = c.ParentRecord,
                    .RecordNo = c.RecordNo,
                    .OperationNumber = c.OperationNumber,
                    .BeforeCount = beforeCount,
                    .AfterCount = batched.Count,
                    .Included = True,
                    .ReasonCode = SchedulerDebugReasonCodes.OK_INCLUDED,
                    .ReasonDetail = "Included in ranked pressing candidates.",
                    .RankScore = c.WipScore,
                    .RankBreakdown = "Tier=" & c.Tier.ToString() & ";CycleRank=" & c.CycleRank.ToString() & ";Type=" & c.TypeKey
                })
            Next
        End If

        Return batched
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
