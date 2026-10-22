Option Strict On
Imports System
Imports System.Collections.Generic
Imports System.Linq

' A deliberately small finite-capacity board substitute. The production campaign
' engine is linked into this executable unchanged; no Opcenter installation or
' live allocation is modified by these tests.
Module PressCampaignHarness
    Private ReadOnly day As New DateTime(2026, 10, 7, 8, 0, 0)
    Private checks As Integer

    Private Class Board
        Public Slots As New List(Of PressCampaignSlot)
        Public Decisions As New List(Of PressCampaignDecision)
        Public Durations As New Dictionary(Of Integer, Double)
        Public Failed As New HashSet(Of Integer)
        Public FailedPairs As New HashSet(Of String)
        Public Commits As New List(Of PressCampaignSlot)
        Public QueryCount As Integer
        Public MoveActual As Boolean
        Public Horizon As DateTime = day.AddDays(20)
        Public UseReference As Boolean
        Public TraceEnabled As Boolean = True

        Public Function Query(order As PressCampaignOrder, resource As Integer, earliest As DateTime) As PressCampaignSlot
            QueryCount += 1
            Dim duration = If(Durations.ContainsKey(order.RecordNo), Durations(order.RecordNo), 1.0)
            Dim start = earliest
            For Each item In Slots.Where(Function(s) s.Resource = resource).OrderBy(Function(s) s.Start)
                If start < item.Finish AndAlso item.Start < start.AddHours(duration) Then start = item.Finish
            Next
            If start.AddHours(duration) > Horizon Then Return Nothing
            Return Slot(order.RecordNo, resource, order.Ring, order.Pin, start, duration)
        End Function

        Public Function Commit(proposed As PressCampaignSlot) As PressCampaignSlot
            If Failed.Contains(proposed.RecordNo) OrElse FailedPairs.Contains(proposed.RecordNo.ToString() & ":" & proposed.Resource.ToString()) Then Return Nothing
            Dim actual = Slot(proposed.RecordNo, proposed.Resource, proposed.Ring, proposed.Pin,
                              proposed.Start, (proposed.Finish - proposed.Start).TotalHours)
            If MoveActual Then actual.Finish = actual.Finish.AddMinutes(1)
            Slots.Add(actual)
            Commits.Add(actual)
            Return actual
        End Function

        Public Sub Run(ParamArray orders As PressCampaignOrder())
            Dim reporter As Action(Of PressCampaignDecision) = Nothing
            If TraceEnabled Then reporter = AddressOf Decisions.Add
            If UseReference Then
                Dim baseline As New ReferencePressCampaignScheduler(day, Slots.ToList(), AddressOf Query,
                    AddressOf Commit, reporter)
                baseline.Run(orders)
                Return
            End If
            Dim engine As New PressCampaignScheduler(day, Slots.ToList(), AddressOf Query,
                                                     AddressOf Commit, reporter)
            engine.Run(orders)
        End Sub
    End Class

    Private Function Slot(id As Integer, resource As Integer, ring As String, pin As String,
                          start As DateTime, hours As Double) As PressCampaignSlot
        Return New PressCampaignSlot With {.RecordNo = id, .Resource = resource, .Ring = ring,
            .Pin = pin, .Start = start, .Finish = start.AddHours(hours)}
    End Function

    Private Function Order(id As Integer, rank As Integer, ring As String, Optional pin As String = "P",
                           Optional resource As Integer = 1) As PressCampaignOrder
        Return New PressCampaignOrder With {.RecordNo = id, .Rank = rank, .Ring = ring, .Pin = pin,
            .Earliest = day.Date, .Ready = day, .Due = day.AddDays(15).Date,
            .Resources = New List(Of Integer) From {resource}}
    End Function

    Private Sub Check(ok As Boolean, message As String)
        If Not ok Then Throw New Exception("FAIL: " & message)
        checks += 1
    End Sub

    Sub Main()
        CampaignAndPins()
        CooldownCalendar()
        UrgentInterruption()
        IdleAndReadiness()
        Assignments()
        FutureProtection()
        FailuresAndMissingData()
        RerunAndExistingViolations()
        InvalidActual()
        ControlledComparison()
        DifferentialComparison()
        PerformanceComparison()
        Console.WriteLine("PASS: " & checks.ToString() & " campaign assertions; production rule engine executed without a live board.")
    End Sub

    Private Sub CampaignAndPins()
        Dim b As New Board
        Dim a = Order(1, 0, "510-G", "X")
        Dim other = Order(2, 1, "600-F")
        Dim pinChange = Order(3, 2, "510-G", "Y")
        Dim samePin = Order(4, 3, " 510-g ", "x")
        samePin.Due = day.AddDays(10)
        b.Run(a, other, pinChange, samePin)
        Check(String.Join(",", b.Commits.Select(Function(s) s.RecordNo)) = "1,4,3,2", "ring grouping spans different due dates; same pin preferred")
        Check(b.Decisions.Any(Function(d) d.RecordNo = 3 AndAlso d.Transition = "PIN_ONLY_CHANGE" AndAlso Not d.CooldownUntil.HasValue), "pin changes do not trigger cooldown")
        Check(b.Decisions.Any(Function(d) d.RecordNo = 4 AndAlso d.Transition = "UNCHANGED_TOOLING"), "trim/case normalization")
        Check(PressCampaignScheduler.Key("510-F") <> PressCampaignScheduler.Key("510-G"), "full suffix retained")

        Dim cross As New Board
        cross.Slots.Add(Slot(90, 1, "A", "P", day.AddHours(-12), 10))
        Dim nextDay = Order(1, 0, "A")
        nextDay.Earliest = day.AddDays(1).Date
        cross.Run(nextDay)
        Check(cross.Commits(0).Start = day.AddDays(1).Date, "idle overnight does not end campaign")
        Check(cross.Decisions.Last().Transition = "UNCHANGED_TOOLING", "campaign spans calendar days")
    End Sub

    Private Sub CooldownCalendar()
        Dim b As New Board
        b.Slots.Add(Slot(90, 1, "A", "P", day.AddDays(-1), 1))
        b.Slots.Add(Slot(91, 1, "B", "P", day, 1))
        b.Run(Order(1, 0, "A"))
        Check(b.Commits(0).Start = New DateTime(2026, 10, 10), "switch on 7th blocks 7th/8th/9th")
        Check(PressCampaignScheduler.ReturnDate(New DateTime(2026, 10, 7, 23, 59, 0)) = New DateTime(2026, 10, 10), "late switch still uses calendar date")
        Check(b.Decisions.Any(Function(d) d.Reason = "PRESSING_COOLDOWN_BLOCKED"), "cooldown has a diagnostic")
    End Sub

    Private Sub UrgentInterruption()
        Dim b As New Board
        b.Slots.Add(Slot(90, 1, "A", "P", day.AddHours(-2), 1))
        Dim urgent = Order(1, 0, "B")
        urgent.Due = day.Date
        Dim extension = Order(2, 1, "A")
        b.Durations(2) = 20
        b.Run(urgent, extension)
        Check(b.Commits(0).RecordNo = 1, "continuation cannot make urgent work late")
        Check(b.Commits(1).Start = New DateTime(2026, 10, 10), "interrupted A campaign must wait")
        Check(b.Decisions.Any(Function(d) d.Reason = "PRESSING_URGENCY_INTERRUPTION"), "urgency interruption explained")

        Dim blocked As New Board
        blocked.Slots.Add(Slot(90, 1, "B", "P", day.AddDays(-1), 1))
        blocked.Slots.Add(Slot(91, 1, "A", "P", day.AddHours(-1), 1))
        urgent = Order(1, 0, "B")
        urgent.Due = day.Date
        blocked.Run(urgent, Order(2, 1, "A"))
        Check(blocked.Commits(0).RecordNo = 2, "urgent work cannot override cooldown")
        Check(blocked.Commits(1).Start.Date = New DateTime(2026, 10, 10), "urgent ring waits for eligibility")

        Dim late As New Board
        late.Slots.Add(Slot(90, 1, "A", "P", day.AddHours(-2), 1))
        urgent = Order(1, 0, "B")
        urgent.Due = day.AddDays(-1).Date
        late.Run(urgent, Order(2, 1, "A"))
        Check(late.Commits(0).RecordNo = 1, "already-late higher-ranked work interrupts")
    End Sub

    Private Sub IdleAndReadiness()
        Dim b As New Board
        b.Slots.Add(Slot(90, 1, "A", "P", day.AddHours(-2), 1))
        Dim futureA = Order(1, 0, "A")
        futureA.Earliest = day.AddDays(1)
        b.Run(futureA, Order(2, 1, "B"))
        Check(b.Commits(0).RecordNo = 2 AndAlso b.Commits(0).Start = day, "do not idle for an unreleased same-ring job")
        Check(b.Commits(1).Start.Date = New DateTime(2026, 10, 10), "switch while waiting starts cooldown")
        Dim release As New Board
        Dim o = Order(1, 0, "A")
        o.Ready = day.AddDays(2)
        o.Earliest = day.AddDays(1)
        release.Run(o)
        Check(release.Commits(0).Start = o.Ready, "predecessor readiness is a hard lower bound")
    End Sub

    Private Sub Assignments()
        Dim b As New Board
        b.Slots.Add(Slot(90, 1, "A", "P", day.AddHours(-2), 1))
        b.Slots.Add(Slot(91, 2, "B", "P", day.AddHours(-2), 1))
        Dim flexible = Order(1, 0, "B")
        flexible.Resources.Add(2)
        b.Run(flexible, Order(2, 1, "C", "P", 1))
        Check(b.Commits.First(Function(s) s.RecordNo = 1).Resource = 2, "flexible order prefers campaign continuation")
        Check(b.Commits.First(Function(s) s.RecordNo = 2).Resource = 1, "fixed assignment retained")
        Dim independent As New Board
        independent.Slots.Add(Slot(90, 1, "A", "P", day.AddDays(-1), 1))
        independent.Slots.Add(Slot(91, 1, "B", "P", day, 1))
        independent.Run(Order(1, 0, "A", "P", 2))
        Check(independent.Commits(0).Start = day, "cooldown belongs to one press, not all presses")
    End Sub

    Private Sub FutureProtection()
        Dim b As New Board
        b.Slots.Add(Slot(90, 1, "A", "P", day.AddHours(-2), 1))
        b.Slots.Add(Slot(91, 1, "A", "P", day.AddDays(1), 1))
        b.Run(Order(1, 0, "B"))
        Check(b.Commits(0).Start >= day.AddDays(1).AddHours(1), "insertion cannot invalidate a fixed future A campaign")
        Check(b.Slots.First(Function(s) s.RecordNo = 91).Start = day.AddDays(1), "existing allocation unchanged")
    End Sub

    Private Sub FailuresAndMissingData()
        Dim b As New Board
        b.Failed.Add(1)
        b.Run(Order(1, 0, "A"), Order(2, 1, "B"), Order(3, 2, " "))
        Check(b.Commits.Count = 1 AndAlso b.Commits(0).RecordNo = 2, "failed placement does not prevent another order")
        Check(b.Decisions.Any(Function(d) d.RecordNo = 2 AndAlso d.Transition = "INITIAL_SETUP"), "failed placement does not change tooling state")
        Check(b.Decisions.Any(Function(d) d.Reason = "PRESSING_MISSING_RING"), "missing ring held for review")
        Dim missing As New Board
        Dim o = Order(1, 0, "A")
        o.Earliest = DateTime.MinValue
        missing.Run(o)
        Check(missing.Commits.Count = 0, "missing earliest date is not silently treated as ready")
        Dim none As New Board
        o = Order(1, 0, "A")
        o.Resources.Clear()
        none.Run(o)
        Check(none.Decisions.Any(Function(d) d.Reason = "PRESSING_NO_FEASIBLE_PLACEMENT"), "no-resource candidate reported")
        Dim alternate As New Board
        o = Order(1, 0, "A")
        o.Resources.Add(2)
        alternate.FailedPairs.Add("1:1")
        alternate.Run(o)
        Check(alternate.Commits.Count = 1 AndAlso alternate.Commits(0).Resource = 2, "failed commit still tries a valid alternate press")
    End Sub

    Private Sub RerunAndExistingViolations()
        Dim b As New Board
        b.Run(Order(1, 0, "A"), Order(2, 1, "B"))
        b.Run(Order(3, 0, "A"))
        Check(b.Commits.Last().Start.Date = New DateTime(2026, 10, 10), "fresh invocation rebuilds cooldown from board")
        Check(b.Commits.Select(Function(s) s.RecordNo).Distinct().Count() = 3, "no duplicate records")
        Dim existing As New Board
        existing.Slots.Add(Slot(90, 1, "A", "P", day.AddHours(-3), 1))
        existing.Slots.Add(Slot(91, 1, "B", "P", day.AddHours(-2), 1))
        existing.Slots.Add(Slot(92, 1, "A", "P", day.AddHours(-1), 1))
        existing.Run(Order(1, 0, "A"))
        Check(existing.Decisions.Any(Function(d) d.Reason = "PRESSING_EXISTING_COOLDOWN_VIOLATION"), "manual historical violation is visible")
        Check(existing.Commits.Count = 1, "unchanged historical violation does not prevent safe continuation")
    End Sub

    Private Sub InvalidActual()
        Dim b As New Board With {.MoveActual = True}
        Dim threw As Boolean = False
        Try
            b.Run(Order(1, 0, "A"), Order(2, 1, "B"))
        Catch ex As InvalidOperationException
            threw = True
        End Try
        Check(threw AndAlso b.Commits.Count = 1, "unexpected committed interval stops further placement")
    End Sub

    Private Sub ControlledComparison()
        ' Same input/calendar for both policies. The baseline mimics the former
        ' ranked per-order earliest-placement path; the campaign path uses exactly
        ' the same ranking and durations. This is a synthetic regression fixture,
        ' not a claim about the customer's production schedule or speed.
        Dim orders As New List(Of PressCampaignOrder)
        Dim oldBoard As New Board
        Dim campaignBoard As New Board
        For i As Integer = 0 To 239
            orders.Add(Order(i + 1, i, "R" & ((i \ 4) Mod 6).ToString(), "P", (i Mod 4) + 1))
            oldBoard.Durations(i + 1) = 0.25
            campaignBoard.Durations(i + 1) = 0.25
        Next
        For Each o In orders
            oldBoard.Commit(oldBoard.Query(o, o.Resources(0), day))
        Next
        Dim watch = Diagnostics.Stopwatch.StartNew()
        campaignBoard.Run(orders.ToArray())
        watch.Stop()
        Dim oldChanges = CountRingChanges(oldBoard.Slots)
        Dim newChanges = CountRingChanges(campaignBoard.Slots)
        Check(oldChanges = 236 AndAlso newChanges = 20, "controlled fixture reduces ring changes")
        Check(campaignBoard.Commits.Count = orders.Count, "controlled fixture schedules every order")
        Check(campaignBoard.Commits.Max(Function(s) s.Finish) = oldBoard.Commits.Max(Function(s) s.Finish), "controlled fixture retains makespan")
        Check(campaignBoard.Commits.All(Function(s) s.Finish <= orders(s.RecordNo - 1).Due.Date.AddDays(1)), "controlled fixture has no late orders")
        Console.WriteLine("Synthetic comparison: 240 orders / 4 presses; ring changes " & oldChanges.ToString() & " -> " & newChanges.ToString() &
                          "; same finish time; 0 late; " & watch.ElapsedMilliseconds.ToString() & " ms (fake board).")
    End Sub

    Private Function CountRingChanges(slots As IEnumerable(Of PressCampaignSlot)) As Integer
        Dim count As Integer = 0
        For Each group In slots.GroupBy(Function(s) s.Resource)
            Dim prior As String = Nothing
            For Each s In group.OrderBy(Function(x) x.Start)
                If prior IsNot Nothing AndAlso prior <> s.Ring Then count += 1
                prior = s.Ring
            Next
        Next
        Return count
    End Function

    Private Sub DifferentialComparison()
        Dim baselineCalls As Long = 0
        Dim optimizedCalls As Long = 0
        Dim baselineMs As Long = 0
        Dim optimizedMs As Long = 0
        ' Compare the complete ordered placement sequence with the frozen original,
        ' not just total changes. Random calendars include fixed future work,
        ' variable release/due dates, alternate presses and unsuccessful commits.
        For seed As Integer = 0 To 79
            Dim random As New Random(seed)
            Dim baseline As New Board With {.UseReference = True, .TraceEnabled = False}
            Dim optimized As New Board With {.TraceEnabled = False}
            For resource As Integer = 1 To 3
                For j As Integer = 0 To 3
                    Dim atTime = If(j = 0, day.AddHours(-2), day.AddDays(j).AddHours(random.Next(0, 5)))
                    Dim ring = "R" & random.Next(4).ToString()
                    Dim id = 1000 + resource * 10 + j
                    baseline.Slots.Add(Slot(id, resource, ring, "P", atTime, 1))
                    optimized.Slots.Add(Slot(id, resource, ring, "P", atTime, 1))
                Next
            Next
            Dim a As New List(Of PressCampaignOrder)
            Dim b As New List(Of PressCampaignOrder)
            For id As Integer = 1 To 35
                Dim resource = random.Next(1, 4)
                Dim ring = "R" & random.Next(4).ToString()
                Dim pin = "P" & random.Next(2).ToString()
                Dim x = Order(id, id, ring, pin, resource)
                x.Earliest = day.AddHours(random.Next(0, 50))
                x.Ready = day.AddHours(random.Next(0, 50))
                x.Due = day.AddDays(random.Next(-1, 10)).Date
                If random.Next(2) = 1 Then x.Resources = New List(Of Integer) From {1, 2, 3}
                Dim y = Order(id, id, ring, pin, resource)
                y.Earliest = x.Earliest : y.Ready = x.Ready : y.Due = x.Due
                y.Resources = x.Resources.ToList()
                Dim duration = 0.25 + random.Next(1, 12) * 0.25
                baseline.Durations(id) = duration : optimized.Durations(id) = duration
                If random.Next(25) = 0 Then
                    baseline.Failed.Add(id) : optimized.Failed.Add(id)
                End If
                a.Add(x) : b.Add(y)
            Next
            Dim watch = Diagnostics.Stopwatch.StartNew()
            baseline.Run(a.ToArray())
            baselineMs += watch.ElapsedMilliseconds
            watch.Restart()
            optimized.Run(b.ToArray())
            optimizedMs += watch.ElapsedMilliseconds
            Check(Signature(baseline) = Signature(optimized), "exact reference parity, seed " & seed.ToString())
            Check(optimized.Decisions.Count = 0, "disabled tracing produces no decision records")
            baselineCalls += baseline.QueryCount
            optimizedCalls += optimized.QueryCount
        Next
        Console.WriteLine("80 differential calendars: exact placement parity. Queries " & baselineCalls.ToString() & " -> " & optimizedCalls.ToString() &
                          "; fake-board runtime " & baselineMs.ToString() & " -> " & optimizedMs.ToString() & " ms.")
    End Sub

    Private Function Signature(board As Board) As String
        Return String.Join(";", board.Commits.Select(Function(s) s.RecordNo.ToString() & ":" & s.Resource.ToString() & ":" & s.Start.Ticks.ToString() & ":" & s.Finish.Ticks.ToString()))
    End Function

    Private Sub PerformanceComparison()
        Dim baseline As New Board With {.UseReference = True, .TraceEnabled = False}
        Dim optimized As New Board With {.TraceEnabled = False}
        Dim orders As New List(Of PressCampaignOrder)
        For i As Integer = 0 To 999
            orders.Add(Order(i + 1, i, "R" & ((i \ 4) Mod 10).ToString(), "P", (i Mod 4) + 1))
            baseline.Durations(i + 1) = 0.25 : optimized.Durations(i + 1) = 0.25
        Next
        Dim watch = Diagnostics.Stopwatch.StartNew()
        baseline.Run(orders.ToArray())
        Dim oldMs = watch.ElapsedMilliseconds
        watch.Restart()
        optimized.Run(orders.ToArray())
        Dim newMs = watch.ElapsedMilliseconds
        Check(Signature(baseline) = Signature(optimized), "1000-order exact placement parity")
        Check(optimized.QueryCount < baseline.QueryCount, "native query count reduced on representative campaign workload")
        Console.WriteLine("1000-order performance fixture: queries " & baseline.QueryCount.ToString() & " -> " & optimized.QueryCount.ToString() &
            "; fake-board runtime " & oldMs.ToString() & " -> " & newMs.ToString() & " ms.")
    End Sub
End Module
