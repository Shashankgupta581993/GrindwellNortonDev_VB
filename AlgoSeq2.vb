Option Strict On
Option Explicit On

Imports System
Imports System.Runtime.InteropServices
Imports Preactor
Imports Preactor.Interop.PreactorObject

<ComVisible(True)>
<Microsoft.VisualBasic.ComClass("f4873a5e-e4a1-4aae-a637-9ccb743205e3", "e03a0e50-cb57-4f0b-a78b-f374e04365cc")>
Public Class AlgoSeq2
    Public Function Run(ByRef preactorComObject As PreactorObj, ByRef pespComObject As Object) As Integer

        Dim preactor As IPreactor = PreactorFactory.CreatePreactorObject(preactorComObject)
        Dim planningboard As IPlanningBoard = preactor.PlanningBoard

        Dim ordersTable As Integer = preactor.GetFormatNumber("Orders")
        Dim opRec As Integer = 0
        Dim ResRec As Integer
        Dim ResRecs As IEnumerable(Of Integer)
        Dim opTimes As Nullable(Of Preactor.OperationTimes)
        Dim dueDateField As Nullable(Of Preactor.FormatFieldPair)

        ' build initial queue
        CreateRankedParentQueue(preactor, planningboard, ordersTable, "JobsQueue")

        ' find due date field
        dueDateField = preactor.FindFirstClassificationString("DUE DATE")

        While (planningboard.GetOperationInQueue("JobsQueue", 1, opRec))

            planningboard.RemoveOperationFromQueue("JobsQueue", opRec)

            Do While (opRec > 0)

                ' get due date of the current operation
                Dim dueDate As DateTime = planningboard.TerminatorTime ' default if no due date
                If dueDateField.HasValue Then
                    Dim ddVal As Object = preactor.ReadFieldDateTime(opRec, dueDateField.Value)
                    If ddVal IsNot Nothing AndAlso IsDate(ddVal) Then
                        dueDate = CDate(ddVal)
                    End If
                End If

                ' test resources backwards from due date
                Dim scheduled As Boolean = False
                ResRecs = planningboard.FindResources(opRec)

                For Each ResRec In ResRecs
                    ' backward test
                    opTimes = planningboard.TestOperationOnResource(opRec, ResRec, dueDate)

                    If opTimes.HasValue Then
                        planningboard.PutOperationOnResource(opRec, ResRec, opTimes.Value.ChangeStart)
                        scheduled = True
                        Exit For
                    End If
                Next

                ' if backward scheduling failed, try forward
                If Not scheduled Then
                    For Each ResRec In ResRecs
                        opTimes = planningboard.TestOperationOnResource(opRec, ResRec, planningboard.time, ScheduleDirection.Forward)

                        If opTimes.HasValue Then
                            planningboard.PutOperationOnResource(opRec, ResRec, opTimes.Value.ChangeStart)
                            Exit For
                        End If
                    Next
                End If

                ' move to next sibling operation
                opRec = planningboard.GetNextOperation(opRec, 1)
            Loop
        End While

        Return 0
    End Function

    Private Function CreateRankedParentQueue(ByRef preactor As IPreactor, ByVal planningboard As IPlanningBoard,
                                             ByVal ordersTable As Integer, ByVal QName As String) As Integer

        Dim ordersParent As Preactor.FormatFieldPair
        Dim dueDateField As Nullable(Of Preactor.FormatFieldPair)
        Dim priorityField As Nullable(Of Preactor.FormatFieldPair)
        Dim parentRecord As Integer

        ordersParent = New FormatFieldPair()
        dueDateField = preactor.FindFirstClassificationString("DUE DATE")
        priorityField = preactor.FindFirstClassificationString("PRIORITY")

        planningboard.CreateQueue(QName)

        parentRecord = preactor.FindMatchingRecord(ordersParent, 0, -1)
        While (parentRecord > 0)
            If (planningboard.GetOperationLocateState(parentRecord)) Then
                planningboard.AddOperationToQueue(QName, parentRecord, QueuePosition.End)
            End If
            parentRecord = preactor.FindMatchingRecord(ordersParent, parentRecord, -1)
        End While

        ' Rank by due date ascending so earliest deadlines are scheduled first
        If (dueDateField.HasValue) Then
            planningboard.RankQueueByFieldName(QName, preactor.GetFieldName(dueDateField.Value), QueueRanking.Ascending)
        ElseIf (priorityField.HasValue) Then
            planningboard.RankQueueByFieldName(QName, preactor.GetFieldName(priorityField.Value), QueueRanking.Ascending)
        End If

        Return 0
    End Function
End Class