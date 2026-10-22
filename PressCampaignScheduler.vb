Option Strict On
Option Explicit On

Imports System
Imports System.Collections.Generic
Imports System.Linq

' These internal types deliberately contain no Opcenter objects. The same rules can
' therefore be exercised against a small test calendar before running on a live board.
Friend Class PressCampaignOrder
    Public Property RecordNo As Integer
    Public Property Rank As Integer
    Public Property Ring As String
    Public Property Pin As String
    Public Property Earliest As DateTime
    Public Property Ready As DateTime
    Public Property Due As DateTime
    Public Property Resources As New List(Of Integer)
End Class

Friend Class PressCampaignSlot
    Public Property RecordNo As Integer
    Public Property Resource As Integer
    Public Property Ring As String
    Public Property Pin As String
    ' Start includes changeover, should the placement adapter later supply it.
    Public Property Start As DateTime
    Public Property Finish As DateTime
    ' Filled before Commit is called, providing a clear integration point for
    ' the user's later changeover logic. These are decisions, not added durations.
    Public Property PreviousRing As String
    Public Property PreviousPin As String
    Public Property Transition As String
End Class

Friend Class PressCampaignDecision
    Public Property RecordNo As Integer
    Public Property Resource As Integer
    Public Property Reason As String
    Public Property Detail As String
    Public Property PreviousRing As String
    Public Property PreviousPin As String
    Public Property NextRing As String
    Public Property NextPin As String
    Public Property Transition As String
    Public Property CooldownUntil As DateTime?
    Public Property Start As DateTime?
    Public Property Finish As DateTime?
End Class

Friend Class PressCampaignScheduler
    Private ReadOnly boundary As DateTime
    Private ReadOnly timeline As New Dictionary(Of Integer, List(Of PressCampaignSlot))
    Private ReadOnly cursor As New Dictionary(Of Integer, DateTime)
    Private ReadOnly query As Func(Of PressCampaignOrder, Integer, DateTime, PressCampaignSlot)
    Private ReadOnly place As Func(Of PressCampaignSlot, PressCampaignSlot)
    Private ReadOnly trace As Action(Of PressCampaignDecision)
    Private ReadOnly historyCache As New Dictionary(Of Integer, HistoryState)

    Private Class HistoryState
        Public Slots As List(Of PressCampaignSlot)
        Public Violations As Dictionary(Of Integer, DateTime)
        Public FinishByRecord As Dictionary(Of Integer, DateTime)
        Public Blocked As New Dictionary(Of String, DateTime)(StringComparer.OrdinalIgnoreCase)
        Public Active As String = ""
        Public LatestFinish As DateTime
    End Class

    Private Function State(resource As Integer) As HistoryState
        Dim result As HistoryState = Nothing
        If historyCache.TryGetValue(resource, result) Then Return result
        Dim slots = History(resource).OrderBy(Function(s) s.Start).ThenBy(Function(s) s.RecordNo).ToList()
        result = New HistoryState With {.Slots = slots, .Violations = Violations(slots),
            .FinishByRecord = slots.ToDictionary(Function(s) s.RecordNo, Function(s) s.Finish)}
        For Each slot In slots
            If slot.Finish > result.LatestFinish Then result.LatestFinish = slot.Finish
            If slot.Ring = result.Active Then Continue For
            If result.Active <> "" Then result.Blocked(result.Active) = ReturnDate(slot.Start)
            result.Active = slot.Ring
        Next
        historyCache.Add(resource, result)
        Return result
    End Function

    Friend Sub New(scheduleBoundary As DateTime,
                   history As IEnumerable(Of PressCampaignSlot),
                   queryPlacement As Func(Of PressCampaignOrder, Integer, DateTime, PressCampaignSlot),
                   commitPlacement As Func(Of PressCampaignSlot, PressCampaignSlot),
                   decisionReporter As Action(Of PressCampaignDecision))
        boundary = scheduleBoundary
        query = queryPlacement
        place = commitPlacement
        trace = decisionReporter
        For Each item In history.OrderBy(Function(s) s.Start).ThenBy(Function(s) s.RecordNo)
            If item.Resource <= 0 OrElse item.Start = DateTime.MinValue OrElse item.Finish < item.Start Then
                Throw New InvalidOperationException("Invalid existing pressing allocation for record " & item.RecordNo.ToString())
            End If
            item.Ring = Key(item.Ring)
            item.Pin = Key(item.Pin)
            If item.Ring = "" Then Report(item.RecordNo, item.Resource, "PRESSING_HISTORY_RING_MISSING",
                "Existing allocation has unknown tooling; known previous ring is treated as removed at this allocation.")
            If Not timeline.ContainsKey(item.Resource) Then timeline(item.Resource) = New List(Of PressCampaignSlot)
            timeline(item.Resource).Add(item)
        Next
        ' Existing manual violations are reported, never repaired or hidden. Later
        ' insertions may not introduce a new violation or lengthen an existing one.
        If trace IsNot Nothing Then
            For Each resource In timeline.Keys
                For Each violation In Violations(timeline(resource))
                    Report(violation.Key, resource, "PRESSING_EXISTING_COOLDOWN_VIOLATION",
                           "Existing placement predates ring return eligibility.", violation.Value)
                Next
            Next
        End If
    End Sub

    Friend Shared Function Key(value As String) As String
        Return If(value, "").Trim().ToUpperInvariant()
    End Function

    Friend Shared Function ReturnDate(switchTime As DateTime) As DateTime
        ' Calendar days, not elapsed hours: any switch on the 7th permits the 10th.
        Return switchTime.Date.AddDays(3)
    End Function

    Private Sub Report(record As Integer, resource As Integer, reason As String,
                       detail As String, Optional untilDate As DateTime? = Nothing)
        If trace IsNot Nothing Then trace(New PressCampaignDecision With {
            .RecordNo = record, .Resource = resource, .Reason = reason,
            .Detail = detail, .CooldownUntil = untilDate})
    End Sub

    Private Function History(resource As Integer) As List(Of PressCampaignSlot)
        If Not timeline.ContainsKey(resource) Then timeline(resource) = New List(Of PressCampaignSlot)
        Return timeline(resource)
    End Function

    Private Function Previous(slot As PressCampaignSlot) As PressCampaignSlot
        ' Sort once per history revision; then use binary search even when the
        ' board has thousands of historical/fixed future pressing allocations.
        Dim slots = State(slot.Resource).Slots
        Dim low = 0
        Dim high = slots.Count - 1
        While low <= high
            Dim middle = low + (high - low) \ 2
            If slots(middle).Start <= slot.Start Then
                low = middle + 1
            Else
                high = middle - 1
            End If
        End While
        Return If(high < 0, Nothing, slots(high))
    End Function

    Private Class SearchOption
        Public Order As PressCampaignOrder
        Public Resource As Integer
        Public Start As DateTime
        Public Continuation As Boolean
        Public UnchangedPin As Boolean
    End Class

    Private Function SearchKey(order As PressCampaignOrder, resource As Integer, start As DateTime) As SearchOption
        Dim prior = Previous(New PressCampaignSlot With {.Resource = resource, .Start = start})
        Dim continuation = prior IsNot Nothing AndAlso prior.Ring = order.Ring
        Return New SearchOption With {.Order = order, .Resource = resource, .Start = start,
            .Continuation = continuation, .UnchangedPin = continuation AndAlso prior.Pin = order.Pin}
    End Function

    Private Shared Function CompareOptions(left As SearchOption, right As SearchOption) As Integer
        Dim comparison = left.Start.CompareTo(right.Start)
        If comparison <> 0 Then Return comparison
        comparison = right.Continuation.CompareTo(left.Continuation)
        If comparison <> 0 Then Return comparison
        comparison = right.UnchangedPin.CompareTo(left.UnchangedPin)
        If comparison <> 0 Then Return comparison
        comparison = left.Order.Rank.CompareTo(right.Order.Rank)
        If comparison <> 0 Then Return comparison
        Return left.Resource.CompareTo(right.Resource)
    End Function

    Private Function LowerBound(order As PressCampaignOrder, resource As Integer) As DateTime
        Dim result = boundary
        If order.Earliest > result Then result = order.Earliest
        If order.Ready > result Then result = order.Ready
        Dim resourceCursor As DateTime
        If cursor.TryGetValue(resource, resourceCursor) AndAlso resourceCursor > result Then result = resourceCursor
        Return result
    End Function

    Private Shared Function PairKey(record As Integer, resource As Integer) As Long
        Return (CLng(record) << 32) Or CLng(CUInt(resource))
    End Function

    Private Function GetOption(order As PressCampaignOrder, resource As Integer,
                               evaluated As Dictionary(Of Long, PressCampaignSlot)) As PressCampaignSlot
        Dim key = PairKey(order.RecordNo, resource)
        Dim result As PressCampaignSlot = Nothing
        If Not evaluated.TryGetValue(key, result) Then
            result = FindSlot(order, resource, boundary)
            evaluated.Add(key, result)
        End If
        Return result
    End Function

    ' Replay ring transitions, not individual order completions. Same-ring jobs and
    ' idle gaps keep the campaign alive; pin changes do not start a ring cooldown.
    Private Shared Function Violations(slots As IEnumerable(Of PressCampaignSlot)) As Dictionary(Of Integer, DateTime)
        Dim blocked As New Dictionary(Of String, DateTime)(StringComparer.OrdinalIgnoreCase)
        Dim active As String = ""
        Dim result As New Dictionary(Of Integer, DateTime)
        For Each slot In slots.OrderBy(Function(s) s.Start).ThenBy(Function(s) s.RecordNo)
            Dim ring = Key(slot.Ring)
            If ring = active Then Continue For
            If active <> "" Then blocked(active) = ReturnDate(slot.Start)
            Dim allowed As DateTime
            If ring <> "" AndAlso blocked.TryGetValue(ring, allowed) AndAlso slot.Start < allowed Then
                result(slot.RecordNo) = allowed
            End If
            active = ring
        Next
        Return result
    End Function

    ' Validate the COMPLETE affected press timeline, including fixed future work.
    ' Checking only the last order would allow a gap insertion to break a later
    ' pre-existing campaign. A rejected insertion is retried after the blocking
    ' future operation, or on the candidate ring's own next eligible date.
    Private Function RetryAfter(slot As PressCampaignSlot) As DateTime
        Dim snapshot = State(slot.Resource)
        Dim existing = snapshot.Slots
        ' The common path appends after all existing work. Historical violations
        ' cannot change, and just one ring-return lookup proves cooldown safety.
        If existing.Count = 0 Then Return DateTime.MinValue
        If slot.Start >= snapshot.LatestFinish AndAlso slot.Start > existing(existing.Count - 1).Start Then
            Dim allowed As DateTime
            If slot.Ring <> snapshot.Active AndAlso snapshot.Blocked.TryGetValue(slot.Ring, allowed) AndAlso slot.Start < allowed Then Return allowed
            Return DateTime.MinValue
        End If
        For Each other In existing
            If slot.Start < other.Finish AndAlso other.Start < slot.Finish Then Return other.Finish
        Next
        Dim baseline = snapshot.Violations
        Dim proposed = Violations(existing.Concat({slot}))
        Dim retry As DateTime = DateTime.MinValue
        For Each issue In proposed
            Dim oldLimit As DateTime
            If issue.Key = slot.RecordNo Then
                retry = If(issue.Value > retry, issue.Value, retry)
            ElseIf Not baseline.TryGetValue(issue.Key, oldLimit) OrElse issue.Value > oldLimit Then
                Dim protectedFinish = snapshot.FinishByRecord(issue.Key)
                If protectedFinish > retry Then retry = protectedFinish
            End If
        Next
        Return retry
    End Function

    Private Function FindSlot(order As PressCampaignOrder, resource As Integer,
                              requested As DateTime, Optional emitTrace As Boolean = True,
                              Optional hypothetical As PressCampaignSlot = Nothing) As PressCampaignSlot
        Dim atTime = {boundary, order.Earliest, order.Ready, requested}.Max()
        If cursor.ContainsKey(resource) AndAlso cursor(resource) > atTime Then atTime = cursor(resource)
        ' The hypothetical slot is used only by urgency checks. It never reaches
        ' the board; replaying it here models the ring switch/cooldown it would cause.
        Dim savedState As HistoryState = Nothing
        If hypothetical IsNot Nothing Then
            savedState = State(resource)
            History(resource).Add(hypothetical)
            historyCache.Remove(resource)
        End If
        Try
            Do
                Dim slot = query(order, resource, atTime)
                If slot Is Nothing Then Return Nothing
                slot.RecordNo = order.RecordNo
                slot.Resource = resource
                slot.Ring = order.Ring
                slot.Pin = order.Pin
                If slot.Start < atTime OrElse slot.Finish <= slot.Start Then
                    If emitTrace Then Report(order.RecordNo, resource, "PRESSING_INVALID_QUERY_RESULT", "Query returned an invalid interval or ignored the minimum start.")
                    Return Nothing
                End If
                Dim retry = RetryAfter(slot)
                If retry = DateTime.MinValue Then Return slot
                If emitTrace Then Report(order.RecordNo, resource, "PRESSING_COOLDOWN_BLOCKED",
                    "Retrying after ring cooldown or a protected existing placement.", retry)
                ' Every retry must move forward. Query supplies the finite board
                ' horizon; do not invent a small retry cap that drops feasible work.
                If retry <= atTime Then Return Nothing
                atTime = retry
            Loop
        Finally
            If hypothetical IsNot Nothing Then
                History(resource).Remove(hypothetical)
                ' Restore the unchanged real history cache, not the trial state.
                historyCache(resource) = savedState
            End If
        End Try
    End Function

    Private Shared Function Deadline(order As PressCampaignOrder) As DateTime
        If order.Due = DateTime.MinValue OrElse order.Due.Date = DateTime.MaxValue.Date Then Return DateTime.MaxValue
        ' Pressing due is date-only. Completing at midnight immediately after that
        ' day is allowed; dates must not accidentally become midnight-at-day-start.
        Return order.Due.Date.AddDays(1)
    End Function

    Private Function Continues(slot As PressCampaignSlot) As Boolean
        Dim prior = Previous(slot)
        Return prior IsNot Nothing AndAlso Key(prior.Ring) = slot.Ring
    End Function

    Private Function SamePin(slot As PressCampaignSlot) As Boolean
        Dim prior = Previous(slot)
        Return prior IsNot Nothing AndAlso Key(prior.Pin) = slot.Pin
    End Function

    Private Function ThreatenedBy(selected As PressCampaignSlot, selectedOrder As PressCampaignOrder,
                                  orders As IEnumerable(Of PressCampaignOrder),
                                  failed As Dictionary(Of Integer, HashSet(Of Integer)),
                                  evaluated As Dictionary(Of Long, PressCampaignSlot)) As PressCampaignSlot
        For Each order In orders
            If order.Rank >= selectedOrder.Rank Then Exit For
            If Not order.Resources.Contains(selected.Resource) OrElse failed(order.RecordNo).Contains(selected.Resource) Then Continue For
            If LowerBound(order, selected.Resource) > selected.Start Then Continue For
            Dim due = Deadline(order)
            If due = DateTime.MaxValue Then Continue For
            ' Only these higher-ranked orders can interrupt. Query them on demand;
            ' reuse any result from this decision, never from a previous placement.
            Dim urgent = GetOption(order, selected.Resource, evaluated)
            If urgent Is Nothing Then Continue For
            ' Only protect work already eligible at this opportunity. Future work
            ' remains governed by its release/earliest-start date.
            If urgent.Start > selected.Start Then Continue For
            If urgent.Finish > due Then Return urgent
            Dim delayed = FindSlot(order, selected.Resource, selected.Finish, False, selected)
            If delayed Is Nothing OrElse delayed.Finish > due Then Return urgent
        Next
        Return Nothing
    End Function

    Friend Sub Run(input As IEnumerable(Of PressCampaignOrder))
        Dim pending As New Dictionary(Of Integer, PressCampaignOrder)
        Dim failedResources As New Dictionary(Of Integer, HashSet(Of Integer))
        For Each order In input.OrderBy(Function(o) o.Rank)
            order.Ring = Key(order.Ring)
            order.Pin = Key(order.Pin)
            If order.Ring = "" Then
                Report(order.RecordNo, 0, "PRESSING_MISSING_RING", "Manual review required: no die-ring identity.")
            ElseIf order.Earliest = DateTime.MinValue Then
                Report(order.RecordNo, 0, "PRESSING_MISSING_EARLIEST", "Manual review required: earliest pressing date is missing.")
            Else
                order.Resources = order.Resources.Distinct().OrderBy(Function(r) r).ToList()
                pending(order.RecordNo) = order
                failedResources(order.RecordNo) = New HashSet(Of Integer)
            End If
        Next

        ' Stable seed order is built once; removing a placed order does not change
        ' anyone else's WIP/date rank. Native results are scoped to ONE decision.
        Dim ranked = pending.Values.OrderBy(Function(o) o.Rank).ToList()
        While pending.Count > 0
            Dim remaining = ranked.Where(Function(o) pending.ContainsKey(o.RecordNo)).ToList()
            Dim evaluated As New Dictionary(Of Long, PressCampaignSlot)
            Dim search As New List(Of SearchOption)
            For Each order In remaining
                For Each resource In order.Resources
                    If failedResources(order.RecordNo).Contains(resource) Then Continue For
                    search.Add(SearchKey(order, resource, LowerBound(order, resource)))
                Next
            Next
            search.Sort(AddressOf CompareOptions)

            ' Exact branch-and-bound search, not a heuristic shortlist:
            ' no native result can start before its boundary/release/press cursor.
            ' If it starts AT that lower bound, its tooling preference is already
            ' known from history; if later, the time key alone makes it worse.
            ' Therefore stop only when even the next optimistic key cannot beat
            ' our verified choice. Future fixed work and native constraints are
            ' still queried whenever they could affect the winner.
            Dim selected As PressCampaignSlot = Nothing
            Dim bestKey As SearchOption = Nothing
            For Each candidate In search
                If bestKey IsNot Nothing AndAlso CompareOptions(bestKey, candidate) <= 0 Then Exit For
                Dim slot = GetOption(candidate.Order, candidate.Resource, evaluated)
                If slot Is Nothing Then Continue For
                Dim actualKey = SearchKey(candidate.Order, candidate.Resource, slot.Start)
                If bestKey Is Nothing OrElse CompareOptions(actualKey, bestKey) < 0 Then
                    selected = slot
                    bestKey = actualKey
                End If
            Next
            If selected Is Nothing Then
                For Each order In remaining
                    Report(order.RecordNo, 0, "PRESSING_NO_FEASIBLE_PLACEMENT", "No eligible resource/time remains within the planning-board horizon.")
                Next
                Exit While
            End If

            Dim interrupted As Boolean = False
            ' Replacing a continuation with more urgent work can itself delay an
            ' even higher-ranked order. Rank strictly decreases, so this terminates.
            Do
                Dim urgent = ThreatenedBy(selected, pending(selected.RecordNo), remaining, failedResources, evaluated)
                If urgent Is Nothing Then Exit Do
                selected = urgent
                interrupted = True
            Loop

            Dim prior = Previous(selected)
            ' Tooling context is required by placement even with diagnostics off.
            ' A trace object/string payload is only built when a reporter exists.
            Dim continuing = prior IsNot Nothing AndAlso prior.Ring = selected.Ring
            selected.PreviousRing = If(prior Is Nothing, "", prior.Ring)
            selected.PreviousPin = If(prior Is Nothing, "", prior.Pin)
            selected.Transition = If(prior Is Nothing OrElse prior.Ring = "", "INITIAL_SETUP",
                If(Not continuing, "RING_CHANGE", If(prior.Pin = selected.Pin, "UNCHANGED_TOOLING", "PIN_ONLY_CHANGE")))

            ' The adapter returns the actual committed interval. Do not update
            ' tooling history for failed placements. A surprising actual result is
            ' a hard stop: silently continuing would compound a board inconsistency.
            Dim actual = place(selected)
            If actual Is Nothing Then
                Report(selected.RecordNo, selected.Resource, "PRESSING_PLACEMENT_FAILED", "Operation was not placed; continuing with other orders.")
                ' A failed commit must not spin forever on the same resource, but
                ' it must not prevent an otherwise valid alternate press either.
                failedResources(selected.RecordNo).Add(selected.Resource)
            Else
                actual.Ring = selected.Ring
                actual.Pin = selected.Pin
                If actual.RecordNo <> selected.RecordNo OrElse actual.Resource <> selected.Resource OrElse
                   actual.Start <> selected.Start OrElse actual.Finish <> selected.Finish OrElse RetryAfter(actual) <> DateTime.MinValue Then
                    Report(selected.RecordNo, selected.Resource, "PRESSING_ACTUAL_PLACEMENT_INVALID", "Actual board placement differs from the validated campaign constraints; stopped for manual review.")
                    Throw New InvalidOperationException("Pressing actual placement violates the validated campaign interval. Inspect schedule diagnostics before continuing.")
                End If
                History(actual.Resource).Add(actual)
                historyCache.Remove(actual.Resource)
                cursor(actual.Resource) = actual.Finish
                If trace IsNot Nothing Then
                    trace(New PressCampaignDecision With {
                        .RecordNo = selected.RecordNo, .Resource = selected.Resource,
                        .PreviousRing = selected.PreviousRing, .PreviousPin = selected.PreviousPin,
                        .NextRing = selected.Ring, .NextPin = selected.Pin,
                        .Transition = selected.Transition, .Start = actual.Start, .Finish = actual.Finish,
                        .CooldownUntil = If(selected.Transition = "RING_CHANGE", CType(ReturnDate(actual.Start), DateTime?), Nothing),
                        .Reason = If(interrupted, "PRESSING_URGENCY_INTERRUPTION",
                            If(continuing, "PRESSING_CAMPAIGN_CONTINUED", "PRESSING_CAMPAIGN_STARTED")),
                        .Detail = "Setup durations excluded in this phase; original WIP/date rank retained within campaign decisions."})
                End If
                pending.Remove(selected.RecordNo)
            End If
        End While
    End Sub
End Class
