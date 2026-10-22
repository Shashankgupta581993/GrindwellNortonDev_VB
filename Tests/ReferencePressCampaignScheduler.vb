Option Strict On
Imports System
Imports System.Collections.Generic
Imports System.Linq

' Frozen correctness oracle: the original exhaustive campaign search. Test project only.
Friend Class ReferencePressCampaignScheduler
    Private ReadOnly boundary As DateTime
    Private ReadOnly timeline As New Dictionary(Of Integer, List(Of PressCampaignSlot))
    Private ReadOnly cursor As New Dictionary(Of Integer, DateTime)
    Private ReadOnly query As Func(Of PressCampaignOrder, Integer, DateTime, PressCampaignSlot)
    Private ReadOnly place As Func(Of PressCampaignSlot, PressCampaignSlot)
    Private ReadOnly trace As Action(Of PressCampaignDecision)

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
        For Each resource In timeline.Keys
            For Each violation In Violations(timeline(resource))
                Report(violation.Key, resource, "PRESSING_EXISTING_COOLDOWN_VIOLATION",
                       "Existing placement predates ring return eligibility.", violation.Value)
            Next
        Next
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
        Return History(slot.Resource).Where(Function(s) s.Start <= slot.Start).
            OrderByDescending(Function(s) s.Start).ThenByDescending(Function(s) s.RecordNo).FirstOrDefault()
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
        Dim existing = History(slot.Resource)
        For Each other In existing
            If slot.Start < other.Finish AndAlso other.Start < slot.Finish Then Return other.Finish
        Next
        Dim baseline = Violations(existing)
        Dim proposed = Violations(existing.Concat({slot}))
        Dim retry As DateTime = DateTime.MinValue
        For Each issue In proposed
            Dim oldLimit As DateTime
            If issue.Key = slot.RecordNo Then
                retry = If(issue.Value > retry, issue.Value, retry)
            ElseIf Not baseline.TryGetValue(issue.Key, oldLimit) OrElse issue.Value > oldLimit Then
                Dim protectedSlot = existing.First(Function(s) s.RecordNo = issue.Key)
                If protectedSlot.Finish > retry Then retry = protectedSlot.Finish
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
        If hypothetical IsNot Nothing Then History(resource).Add(hypothetical)
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
            If hypothetical IsNot Nothing Then History(resource).Remove(hypothetical)
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
                                  options As List(Of PressCampaignSlot),
                                  orders As Dictionary(Of Integer, PressCampaignOrder)) As PressCampaignSlot
        For Each urgent In options.Where(Function(s) s.Resource = selected.Resource AndAlso
                                             orders(s.RecordNo).Rank < selectedOrder.Rank).
                                             OrderBy(Function(s) orders(s.RecordNo).Rank)
            Dim order = orders(urgent.RecordNo)
            Dim due = Deadline(order)
            ' Only protect work already eligible at this opportunity. Future work
            ' remains governed by its release/earliest-start date.
            If urgent.Start > selected.Start Then Continue For
            If due = DateTime.MaxValue Then Continue For
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
                pending(order.RecordNo) = order
                failedResources(order.RecordNo) = New HashSet(Of Integer)
            End If
        Next

        While pending.Count > 0
            ' Requery after every successful placement. A resource calendar and the
            ' campaign state have changed, so yesterday's best resource is not a
            ' reliable choice for the next order. Resource IDs provide stable ties.
            Dim options As New List(Of PressCampaignSlot)
            For Each order In pending.Values.OrderBy(Function(o) o.Rank)
                For Each resource In order.Resources.Distinct().OrderBy(Function(r) r)
                    If failedResources(order.RecordNo).Contains(resource) Then Continue For
                    Dim slot = FindSlot(order, resource, boundary)
                    If slot IsNot Nothing Then options.Add(slot)
                Next
            Next
            If options.Count = 0 Then
                For Each order In pending.Values
                    Report(order.RecordNo, 0, "PRESSING_NO_FEASIBLE_PLACEMENT", "No eligible resource/time remains within the planning-board horizon.")
                Next
                Exit While
            End If

            ' Work at the next available scheduling opportunity. This prevents a
            ' preference for a ring that is ready tomorrow from idling a press today.
            Dim nextStart = options.Min(Function(s) s.Start)
            Dim ready = options.Where(Function(s) s.Start = nextStart).ToList()
            Dim selected = ready.OrderByDescending(Function(s) Continues(s)).
                ThenByDescending(Function(s) Continues(s) AndAlso SamePin(s)).
                ThenBy(Function(s) pending(s.RecordNo).Rank).
                ThenBy(Function(s) s.Resource).First()
            Dim interrupted As Boolean = False
            ' Replacing a continuation with more urgent work can itself delay an
            ' even higher-ranked order. Rank strictly decreases, so this terminates.
            Do
                Dim urgent = ThreatenedBy(selected, pending(selected.RecordNo), options, pending)
                If urgent Is Nothing Then Exit Do
                selected = urgent
                interrupted = True
            Loop

            Dim prior = Previous(selected)
            Dim decision As New PressCampaignDecision With {
                .RecordNo = selected.RecordNo, .Resource = selected.Resource,
                .PreviousRing = If(prior Is Nothing, "", prior.Ring),
                .PreviousPin = If(prior Is Nothing, "", prior.Pin),
                .NextRing = selected.Ring, .NextPin = selected.Pin,
                .Start = selected.Start, .Finish = selected.Finish}
            decision.Transition = If(prior Is Nothing OrElse prior.Ring = "", "INITIAL_SETUP",
                If(Not Continues(selected), "RING_CHANGE", If(SamePin(selected), "UNCHANGED_TOOLING", "PIN_ONLY_CHANGE")))
            If decision.Transition = "RING_CHANGE" Then decision.CooldownUntil = ReturnDate(selected.Start)
            decision.Reason = If(interrupted, "PRESSING_URGENCY_INTERRUPTION",
                If(Continues(selected), "PRESSING_CAMPAIGN_CONTINUED", "PRESSING_CAMPAIGN_STARTED"))
            decision.Detail = "Setup durations excluded in this phase; original WIP/date rank retained within campaign decisions."
            selected.PreviousRing = decision.PreviousRing
            selected.PreviousPin = decision.PreviousPin
            selected.Transition = decision.Transition

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
                cursor(actual.Resource) = actual.Finish
                decision.Start = actual.Start
                decision.Finish = actual.Finish
                If decision.Transition = "RING_CHANGE" Then decision.CooldownUntil = ReturnDate(actual.Start)
                If trace IsNot Nothing Then trace(decision)
                pending.Remove(selected.RecordNo)
            End If
        End While
    End Sub
End Class
