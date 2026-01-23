Option Strict Off
Option Explicit Off

Imports System
Imports System.Runtime.InteropServices
Imports Preactor
Imports Preactor.Interop.PreactorObject

<ComVisible(True)> _
<Microsoft.VisualBasic.ComClass("8e37ffa4-7db4-49ff-a5fb-470ef5bfd962", "e14e8cbc-14c4-42bd-877e-e8719abdca7b")> _
Public Class EventSeq
    Public Function Run(ByRef preactorComObject As PreactorObj, ByRef pespComObject As Object) As Integer
        Dim preactor As IPreactor = PreactorFactory.CreatePreactorObject(preactorComObject)
        Dim planningboard As IPlanningBoard = preactor.PlanningBoard

        Dim EventParameters = planningboard.NextEvent()
        Dim ResRec As Integer
        Dim Qname As String
        Dim ResIndex As Integer
        Dim QNumber As Integer

        While EventParameters.HasValue

            Select Case EventParameters.Value.EventType

                Case EventTypes.OperationFinished

                    ' Event Parameter 1 is the Operation record that finished
                    ' Event Parameter 2 is the Resource record that became available
                    ' check all resources for this event because secondary constraints may have
                    ' changed

                    For ResRec = 1 To preactor.RecordCount("Resources")

                        Qname = planningboard.GetResourceQueueName(ResRec)
                        ScheduleOperations(preactor, Qname, ResRec,
                        EventParameters.Value.EventTime)

                    Next ResRec
                Case EventTypes.QueueChange
                    ' Event Parameter 1 is the number of the queue that changed
                    ' check all resources which use this queue
                    ResIndex = 1
                    ResRec = 0
                    Qname = planningboard.GetQueueName(EventParameters.Value.Parameter1)
                    While (planningboard.GetQueuesResource(QName, ResIndex, ResRec))

                        ScheduleOperations(preactor, QName, ResRec, EventParameters.Value.EventTime)
                        ResIndex = ResIndex + 1

                    End While ' whilst there is another resource for this queue
                Case EventTypes.ShiftChange

                    ' Event Parameter 2 is the Resource record that had a shift change
                    ' check the resource that had the shift change
                    QNumber = planningboard.GetResourceQueue(EventParameters.Value.Parameter2)
                    QName = planningboard.GetQueueName(QNumber)
                    Q = planningboard.GetResourceQueue(ResRec)

                    ScheduleOperations(preactor, QName, EventParameters.Value.Parameter2,
                                       EventParameters.Value.EventTime)
                Case EventTypes.UserEvent

                Case Else
            End Select

            EventParameters = planningboard.NextEvent()
        End While ' whilst there is another event
        Return 0
    End Function

    ' Make method Public and fix missing parameters
    Public Function ScheduleOperations(preactor As IPreactor, QName As String, ResRec As Integer, TestEventTime As DateTime) As Integer
        ' Dim preactor As IPreactor = PreactorFactory.CreatePreactorObject(preactorComObject)
        Dim planningboard As IPlanningBoard = preactor.PlanningBoard

        Dim CurrentRank As Integer = 1
        Dim OpRecord As Integer = 0
        Dim ResourceFree As Boolean = planningboard.IsResourceFree(ResRec, TestEventTime.AddDays(planningboard.SchedulingAccuracy))
        Dim value = planningboard.GetOperationInQueue(QName, CurrentRank, OpRecord)
        While planningboard.GetOperationInQueue(QName, CurrentRank, OpRecord) AndAlso ResourceFree

            Dim TestOpResults = planningboard.TestOperationOnResource(OpRecord, ResRec, TestEventTime)

            If Not TestOpResults.HasValue Then
                CurrentRank += 1
                Continue While
            End If

            If TestOpResults.Value.ChangeStart <= TestEventTime.AddDays(planningboard.SchedulingAccuracy) Then
                planningboard.PutOperationOnResource(OpRecord, ResRec, TestOpResults.Value.ChangeStart)

            Else
                ' if the operation cannot start now, check the next job in the queue
                CurrentRank += 1
            End If

            ' check if the resource is still free at this time
            ResourceFree = planningboard.IsResourceFree(ResRec, TestEventTime.AddDays(planningboard.SchedulingAccuracy))
        End While

        Return 0
    End Function
End Class
