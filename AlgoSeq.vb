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

        ' Put a breakpoint here and inspect:
        '   jobsQueueSnapshot
        ' It will contain the opRec values currently in JobsQueue, in queue order.


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




'' ======================================================================================
'' AlgoSeq
'' --------------------------------------------------------------------------------------
'' Purpose:
''   Example of an algorithmic sequencing (scheduling) rule for Siemens Opcenter AS
''   (Preactor). It:
''     1) Builds a queue of "parent" operations (top-level / order-level records)
''        that are currently "locatable" on the planning board.
''     2) Ranks that queue based on Planning Board Sequence Mode (due date / priority).
''     3) Iterates through the queue and attempts to schedule operations onto resources
''        using TestOperationOnResource + PutOperationOnResource.
''
'' Notes on Opcenter/Preactor objects used:
''   - IPreactor: main API object to read fields/tables/classifications and general model
''   - IPlanningBoard: API for queues, operations, resources, and scheduling calls
''
'' Important:
''   This sample uses "parentRecord" / "parent operation" patterns often used in Opcenter
''   where an "order" (parent) links to one or more operations (children).
'' ======================================================================================
'Public Class AlgoSeq

'    ' ----------------------------------------------------------------------------------
'    ' Run: Entry point called by Opcenter AS when executing the rule.
'    '
'    ' Parameters:
'    '   preactorComObject: COM object from Opcenter which we wrap into IPreactor
'    '   pespComObject    : Another COM object (often used in integrations); unused here
'    '
'    ' Returns:
'    '   Integer status code (0 = success in this sample)
'    ' ----------------------------------------------------------------------------------
'    Public Function Run(ByRef preactorComObject As PreactorObj, ByRef pespComObject As Object) As Integer

'        ' Wrap COM object into the strongly-typed .NET IPreactor interface
'        Dim preactor As IPreactor = PreactorFactory.CreatePreactorObject(preactorComObject)

'        ' PlanningBoard provides scheduling-specific API: queues, operations, resources, etc.
'        Dim planningboard As IPlanningBoard = preactor.PlanningBoard

'        ' ----------------------------------------------------------------------------------
'        ' Local variables used for table/record navigation and scheduling
'        ' ----------------------------------------------------------------------------------
'        Dim ordersTable As Integer         ' Format (table) number for a chosen classification (used as "Orders table")
'        Dim opRec As Integer               ' Operation record pointer (also used for parent record in some calls)
'        Dim ResRec As Integer              ' Single resource record id
'        Dim ResRecs As IEnumerable(Of Integer) ' Collection of eligible resources for an operation
'        Dim opTimes As Nullable(Of Preactor.OperationTimes) ' Holds feasible timing results from TestOperationOnResource

'        ' ----------------------------------------------------------------------------------
'        ' Resolve which "format/table" we are working with.
'        '
'        ' FindFirstClassificationString("LAUNCH TIME") returns a field reference.
'        ' Value.FormatNumber gives the table/format number that field belongs to.
'        ' In many models, "LAUNCH TIME" is on the Orders (parent) table.
'        ' ----------------------------------------------------------------------------------
'        ordersTable = preactor.FindFirstClassificationString("LAUNCH TIME").Value.FormatNumber

'        ' Initialize operation record pointer
'        opRec = 0

'        ' Build and rank a queue of parent operations (orders) into a planning board queue named "JobsQueue"
'        CreateRankedParentQueue(preactor, planningboard, ordersTable, "JobsQueue")

'        ' ==================================================================================
'        ' OUTER LOOP:
'        '   Fetch the first operation in queue "JobsQueue".
'        '   Signature: GetOperationInQueue(queueName, position, ByRef operationRecord)
'        '     - position=1 usually means "first position"
'        '     - opRec is set (ByRef) to the operation/record id found at that position
'        '
'        ' Loop continues while there is an operation at position 1 of JobsQueue.
'        ' ==================================================================================
'        While (planningboard.GetOperationInQueue("JobsQueue", 1, opRec))

'            ' Remove that operation from the queue so it doesn't get picked again
'            planningboard.RemoveOperationFromQueue("JobsQueue", opRec)

'            ' ------------------------------------------------------------------------------
'            ' INNER LOOP:
'            '   Intended comment in your code says:
'            '     "inner loop for operations of the same family"
'            '
'            '   But as written it is: While (opRec > 0)
'            '   That means it will run only if opRec is a positive record id.
'            '   In practice, opRec will be > 0 here (since it was retrieved from queue),
'            '   so the inner loop will run.
'            '
'            '   However, note the comment "operations of the same family" does not
'            '   exist in logic here; there is no explicit family comparison/filter.
'            '   The loop progresses using GetNextOperation(opRec, 1), which walks the
'            '   routing/operation chain depending on your model's linkage.
'            ' ------------------------------------------------------------------------------
'            While (opRec > 0)

'                ' Find all resources that can potentially run this operation record
'                ' Returns an enumerable list of resource record ids.
'                ResRecs = planningboard.FindResources(opRec)

'                ' Try to schedule the operation on a resource.
'                ' In this simple example, we only consider the FIRST resource found.
'                For Each ResRec In ResRecs

'                    ' TestOperationOnResource checks feasibility and returns candidate times:
'                    '   - Inputs:
'                    '       opRec: operation record
'                    '       ResRec: resource record
'                    '       planningboard.TerminatorTime: typically the horizon end time
'                    '   - Output:
'                    '       Nullable(OperationTimes) where .HasValue means feasible.
'                    '       OperationTimes contains start/end/changeover details.
'                    opTimes = planningboard.TestOperationOnResource(opRec, ResRec, planningboard.TerminatorTime)

'                    ' If feasible timing is returned, place the operation on the resource.
'                    If (opTimes.HasValue) Then

'                        ' PutOperationOnResource actually schedules the operation.
'                        ' Using opTimes.Value.ChangeStart means you are starting at the
'                        ' "ChangeStart" time (often includes changeover/setup start).
'                        planningboard.PutOperationOnResource(opRec, ResRec, opTimes.Value.ChangeStart)

'                        ' Note: if the operation times had a value, it was scheduled.
'                    End If

'                    ' IMPORTANT:
'                    ' Exit For means we stop after the first resource in ResRecs.
'                    ' So even if the first resource is infeasible (HasValue=False),
'                    ' the code will still Exit For and will NOT try the next resource.
'                    Exit For

'                Next ' each resource

'                ' Move to the next operation linked to this one.
'                ' GetNextOperation(opRec, 1) typically means:
'                '   - from current opRec, get the next operation in the chain
'                '   - "1" is usually the direction / relationship index depending on API
'                '
'                ' When there is no next operation, this typically returns 0,
'                ' which will end the inner While(opRec > 0).
'                opRec = planningboard.GetNextOperation(opRec, 1)

'            End While ' end inner loop across linked operations

'        End While ' end outer loop over queue items

'        ' TODO : Your code goes here
'        ' You could add logging, stats, exception handling, etc.

'        Return 0
'    End Function

'    ' ======================================================================================
'    ' CreateRankedParentQueue
'    ' --------------------------------------------------------------------------------------
'    ' Purpose:
'    '   1) Identify the "FAMILY" classification field that belongs to the Orders table
'    '      (ordersTable format number).
'    '   2) Iterate through matching records (parent orders) and add those that are
'    '      "locatable" to the queue.
'    '   3) Rank the queue based on current PlanningBoard SequenceMode.Priority setting:
'    '      - DueDate -> rank ascending by DUE DATE
'    '      - Priority -> rank ascending by PRIORITY
'    '      - ReversePriority -> rank descending by PRIORITY
'    '
'    ' Parameters:
'    '   preactor     : IPreactor object for classification lookup and record searching
'    '   planningboard: IPlanningBoard object for queue operations/ranking
'    '   ordersTable  : format/table number that represents Orders table
'    '   QName        : planningboard queue name to create/populate
'    '
'    ' Returns:
'    '   Integer status (always 0 in this sample)
'    ' ======================================================================================
'    Private Function CreateRankedParentQueue(ByRef preactor As IPreactor,
'                                             ByVal planningboard As IPlanningBoard,
'                                             ByVal ordersTable As Integer,
'                                             ByVal QName As String) As Integer

'        ' Field reference that will represent the "FAMILY" field on the Orders table
'        Dim ordersParent As Preactor.FormatFieldPair

'        ' Optional field refs for sorting
'        Dim dueDateField As Nullable(Of Preactor.FormatFieldPair)
'        Dim priorityField As Nullable(Of Preactor.FormatFieldPair)

'        ' Record pointer used to iterate through records
'        Dim parentRecord As Integer

'        ' Captures the planning board sequencing mode (contains .Priority among other settings)
'        Dim SequenceMode As Preactor.SequenceMode

'        ' FAMILY can exist in multiple tables; FindClassificationString returns all matches
'        Dim familyFields As IEnumerable(Of Preactor.FormatFieldPair)

'        ' Initialize the ordersParent field pair (placeholder)
'        ordersParent = New FormatFieldPair()

'        ' Find all fields classified as "FAMILY" (across formats/tables)
'        familyFields = preactor.FindClassificationString("FAMILY")

'        ' Choose the FAMILY field that belongs to the ordersTable format number
'        For Each familyField In familyFields
'            If (familyField.FormatNumber = ordersTable) Then
'                ordersParent = familyField
'            End If
'        Next

'        ' Lookup common fields for ranking
'        dueDateField = preactor.FindFirstClassificationString("DUE DATE")
'        priorityField = preactor.FindFirstClassificationString("PRIORITY")

'        ' Create (or recreate) the queue in the planning board
'        planningboard.CreateQueue(QName)

'        ' ------------------------------------------------------------------------------
'        ' FindMatchingRecord is used to iterate over records that match a field criteria.
'        '
'        ' Here it is called as:
'        '   parentRecord = preactor.FindMatchingRecord(ordersParent, parentRecord, -1)
'        '
'        ' Two important nuances:
'        '   1) parentRecord is not explicitly initialized before first call; in VB.NET
'        '      locals default to 0. So this first call is effectively:
'        '        FindMatchingRecord(ordersParent, 0, -1)
'        '      which typically means "start search from the beginning".
'        '   2) The meaning of "-1" depends on the API; commonly it's used as a match value
'        '      or "any" placeholder in sample code. In some models this is intended to
'        '      retrieve parent records linked in a certain way.
'        ' ------------------------------------------------------------------------------
'        parentRecord = preactor.FindMatchingRecord(ordersParent, parentRecord, -1)

'        ' Iterate through all returned records (>0 means a valid record found)
'        While (parentRecord > 0)

'            ' GetOperationLocateState checks if an operation/record is currently "locatable"
'            ' (i.e., available/eligible to be scheduled/handled in the current context)
'            If (planningboard.GetOperationLocateState(parentRecord)) Then

'                ' Add to end of queue
'                planningboard.AddOperationToQueue(QName, parentRecord, QueuePosition.End)

'                ' Get next matching record
'                parentRecord = preactor.FindMatchingRecord(ordersParent, parentRecord, -1)
'            End If

'            ' NOTE (as written):
'            '   If GetOperationLocateState(parentRecord) is FALSE, parentRecord is NOT advanced,
'            '   so this While loop would become infinite (stuck on the same record).
'            '
'            '   Usually sample code advances parentRecord regardless, e.g.:
'            '     parentRecord = FindMatchingRecord(...)
'            '
'            '   I am NOT changing logic, just highlighting this behavior for your review.
'        End While

'        ' Read current planning board sequencing configuration
'        SequenceMode = planningboard.SequenceMode

'        ' Rank queue based on Priority mode in SequenceMode
'        Select Case SequenceMode.Priority

'            Case SequencePriority.DueDate
'                ' If DUE DATE field exists, rank ascending (earliest due date first)
'                If (dueDateField.HasValue) Then
'                    planningboard.RankQueueByFieldName(QName,
'                                                       preactor.GetFieldName(dueDateField.Value),
'                                                       QueueRanking.Ascending)
'                End If

'            Case SequencePriority.Priority
'                ' If PRIORITY field exists, rank ascending (depends on your priority convention)
'                If (priorityField.HasValue) Then
'                    planningboard.RankQueueByFieldName(QName,
'                                                       preactor.GetFieldName(priorityField.Value),
'                                                       QueueRanking.Ascending)
'                End If

'            Case SequencePriority.ReversePriority
'                ' If PRIORITY field exists, rank descending
'                If (priorityField.HasValue) Then
'                    planningboard.RankQueueByFieldName(QName,
'                                                       preactor.GetFieldName(priorityField.Value),
'                                                       QueueRanking.Descending)
'                End If

'            Case Else
'                ' No ranking applied
'        End Select

'        Return 0
'    End Function

'End Class
