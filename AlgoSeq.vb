Option Strict On
Option Explicit On

Imports System
Imports System.Runtime.InteropServices
Imports Preactor
Imports Preactor.Interop.PreactorObject

<ComVisible(True)> _
<Microsoft.VisualBasic.ComClass("6e4ab73c-5de3-4108-ae88-4c4675df4992", "eec811c8-27f6-43d3-8962-28472d9f325c")> _
Public Class AlgoSeq
    Public Function Run(ByRef preactorComObject As PreactorObj, ByRef pespComObject As Object) As Integer

        Dim preactor As IPreactor = PreactorFactory.CreatePreactorObject(preactorComObject)
        Dim planningboard As IPlanningBoard = preactor.PlanningBoard

        Dim ordersTable As Integer
        Dim opRec As Integer
        Dim ResRec As Integer
        Dim ResRecs As IEnumerable(Of Integer)
        Dim opTimes As Nullable(Of Preactor.OperationTimes)

        ordersTable = preactor.FindFirstClassificationString("LAUNCH TIME").Value.FormatNumber
        opRec = 0
        CreateRankedParentQueue(preactor, planningboard, ordersTable, "JobsQueue")

        ' --- DEBUG: snapshot the JobsQueue contents (operation record numbers) ---
        Dim jobsQueueSnapshot As List(Of Integer) = GetQueueSnapshot(planningboard, "JobsQueue")

        '------------------------------------------------------------
        ' Select the resource that gives the earliest feasible
        ' ChangeStart time for the operation.
        '
        ' Key idea:
        '   - Test the operation on ALL valid alternate resources
        '   - Choose the candidate with the minimum ChangeStart
        '   - Then load the operation on that chosen resource
        '------------------------------------------------------------

        While (planningboard.GetOperationInQueue("JobsQueue", 1, opRec))

            ' Take the next operation out of the ranked queue so we can decide where to load it.
            planningboard.RemoveOperationFromQueue("JobsQueue", opRec)

            ' Inner loop: schedule this operation and then walk to subsequent operations
            ' (your "family" / routing chain) using GetNextOperation.
            While (opRec > 0)

                ' Find all valid alternate resources for this operation.
                ResRecs = planningboard.FindResources(opRec)

                ' Track the best (earliest) feasible candidate we find.
                Dim bestResRec As Integer = 0
                Dim bestOpTimes As Nullable(Of Preactor.OperationTimes) = Nothing

                ' Loop through *all* alternate resources and test feasibility on each.
                For Each ResRec In ResRecs

                    ' Test if the operation can be placed on this resource, and get the timing result.
                    ' TerminatorTime is the boundary between schedule history and schedule future;
                    ' using it here aligns with "schedule as soon as possible" in the future horizon. :contentReference[oaicite:3]{index=3}
                    opTimes = planningboard.TestOperationOnResource(opRec, ResRec, planningboard.TerminatorTime)

                    If opTimes.HasValue Then
                        ' This resource is feasible. Compare it to the current best candidate.
                        ' We want the earliest possible start time (ChangeStart).
                        If (Not bestOpTimes.HasValue) Then
                            ' First feasible candidate becomes the best by default.
                            bestResRec = ResRec
                            bestOpTimes = opTimes
                        Else
                            ' Replace best candidate if this one starts earlier.
                            If opTimes.Value.ChangeStart < bestOpTimes.Value.ChangeStart Then
                                bestResRec = ResRec
                                bestOpTimes = opTimes
                            End If
                        End If
                    End If

                Next ' evaluate next alternate resource

                ' After scanning all alternates:
                If bestOpTimes.HasValue AndAlso bestResRec > 0 Then
                    ' Load the operation onto the resource that gives the earliest feasible start.
                    planningboard.PutOperationOnResource(opRec, bestResRec, bestOpTimes.Value.ChangeStart)
                Else
                    ' No feasible resource was found.
                    ' Practical meaning:
                    '   - This operation cannot be scheduled on any alternate resource at/after the terminator boundary
                    '     under current constraints (calendars, setups, secondary constraints, etc.).
                    ' Leave it unscheduled (or handle with a custom queue / reason code if your design requires).
                End If

                ' Move to the next operation in the routing chain.
                opRec = planningboard.GetNextOperation(opRec, 1) ' API-supported routing traversal:contentReference[oaicite:4]{index=4}

            End While ' next operation in chain

        End While ' next op in JobsQueue



        Return 0
    End Function

    Private Function CreateRankedParentQueue(ByRef preactor As IPreactor, ByVal planningboard As IPlanningBoard,
                                             ByVal ordersTable As Integer, ByVal QName As String) As Integer

        Dim ordersParent As Preactor.FormatFieldPair
        Dim dueDateField As Nullable(Of Preactor.FormatFieldPair)
        Dim priorityField As Nullable(Of Preactor.FormatFieldPair)
        Dim parentRecord As Integer
        Dim SequenceMode As Preactor.SequenceMode
        Dim familyFields As IEnumerable(Of Preactor.FormatFieldPair)
        Dim nextrec As Integer
        ordersParent = New FormatFieldPair()
        familyFields = preactor.FindClassificationString("FAMILY")

        For Each familyField In familyFields
            If (familyField.FormatNumber = ordersTable) Then
                ordersParent = familyField
            End If
        Next
        'My code starts
        Dim ordersOpNoField As Preactor.FormatFieldPair
        ordersOpNoField = New FormatFieldPair()
        Dim opNoFields As IEnumerable(Of Preactor.FormatFieldPair)
        opNoFields = preactor.FindClassificationString("OP NO")

        For Each opNofield In opNoFields
            If (opNofield.FormatNumber = ordersTable) Then
                ordersOpNoField = opNofield
            End If
        Next

        'end
        dueDateField = preactor.FindFirstClassificationString("DUE DATE")
        priorityField = preactor.FindFirstClassificationString("PRIORITY")
        planningboard.CreateQueue(QName)
        parentRecord = preactor.FindMatchingRecord(ordersParent, parentRecord, -1)
        While (parentRecord > 0)
            If (planningboard.GetOperationLocateState(parentRecord)) Then
                If (planningboard.IsOperationScheduled(parentRecord)) Then
                    nextrec = parentRecord
                    While (nextrec > 0)
                        If (Not planningboard.IsOperationScheduled(nextrec)) Then
                            planningboard.AddOperationToQueue(QName, nextrec, QueuePosition.End)
                            nextrec = 0
                        Else
                            nextrec = planningboard.GetNextOperation(nextrec, 1)
                        End If
                    End While
                End If
                parentRecord = preactor.FindMatchingRecord(ordersParent, parentRecord, -1)
            End If ' if this order was highlighted
        End While

        SequenceMode = planningboard.SequenceMode
        Select Case SequenceMode.Priority

            Case SequencePriority.DueDate
                If (dueDateField.HasValue) Then
                    planningboard.RankQueueByFieldName(QName, preactor.GetFieldName(dueDateField.Value), QueueRanking.Ascending)
                End If
            Case SequencePriority.Priority
                If (priorityField.HasValue) Then
                    planningboard.RankQueueByFieldName(QName, preactor.GetFieldName(priorityField.Value), QueueRanking.Ascending)
                End If
            Case SequencePriority.ReversePriority
                If (priorityField.HasValue) Then
                    planningboard.RankQueueByFieldName(QName, preactor.GetFieldName(priorityField.Value), QueueRanking.Descending)
                End If

            Case Else
        End Select
        Return 0
    End Function

    ' Returns the current contents of a PlanningBoard queue as a list of operation record numbers,
    ' in queue order (position 1, 2, 3, ...). This does NOT modify the queue.
    Private Function GetQueueSnapshot(ByVal planningboard As IPlanningBoard, ByVal queueName As String) As List(Of Integer)

        Dim snapshot As New List(Of Integer)()

        Dim pos As Integer = 1
        Dim opRec As Integer = 0

        ' GetOperationInQueue(queueName, position, opRec) returns True if an item exists at that position
        ' and sets opRec to the record number. When there are no more items, it returns False.
        While planningboard.GetOperationInQueue(queueName, pos, opRec)
            snapshot.Add(opRec)
            pos += 1
        End While

        Return snapshot
    End Function

End Class
