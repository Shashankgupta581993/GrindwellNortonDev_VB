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

        While (planningboard.GetOperationInQueue("JobsQueue", 1, opRec))

            planningboard.RemoveOperationFromQueue("JobsQueue", opRec)

            While (opRec > 0) ' inner loop for operations of the same family

                ResRecs = planningboard.FindResources(opRec)
                For Each ResRec In ResRecs

                    opTimes = planningboard.TestOperationOnResource(opRec, ResRec, planningboard.TerminatorTime)

                    If (opTimes.HasValue) Then
                        planningboard.PutOperationOnResource(opRec, ResRec, opTimes.Value.ChangeStart)
                        ' if the operation times had a value
                    End If
                    Exit For ' only do this for the first resource in this simple example
                Next ' for each resource record

                opRec = planningboard.GetNextOperation(opRec, 1)

            End While ' whilst there is another operation
        End While ' whilst there is another operation in the queue

        'TODO : Your code goes here

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

        ordersParent = New FormatFieldPair()
        familyFields = preactor.FindClassificationString("FAMILY")

        For Each familyField In familyFields
            If (familyField.FormatNumber = ordersTable) Then
                ordersParent = familyField
            End If
        Next

        dueDateField = preactor.FindFirstClassificationString("DUE DATE")
        priorityField = preactor.FindFirstClassificationString("PRIORITY")
        planningboard.CreateQueue(QName)
        parentRecord = preactor.FindMatchingRecord(ordersParent, parentRecord, -1)
        While (parentRecord > 0)
            If (planningboard.GetOperationLocateState(parentRecord)) Then
                planningboard.AddOperationToQueue(QName, parentRecord, QueuePosition.End)
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

End Class
