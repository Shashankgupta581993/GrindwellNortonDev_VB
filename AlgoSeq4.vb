Option Strict On
Option Explicit On

Imports System
Imports System.Globalization
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Threading
Imports Microsoft.VisualBasic.FileIO
Imports Preactor
Imports Preactor.Interop.PreactorObject
Imports System.Windows.Forms
Imports System.Linq

<ComVisible(True)>
<Microsoft.VisualBasic.ComClass("4196dd4d-4e89-45a5-9ca5-4fc6dcf10308", "ef5b2382-ab81-47a5-9c8d-0826dcc85a0a")>
Public Class AlgoSeq4
    Private _schedulerDebug As SchedulerDebugCollector

    Public Function runFiring(ByRef preactorComObject As PreactorObj, ByRef pespComObject As Object) As Integer
        ' Batch firing logic
        Dim preactor As IPreactor = PreactorFactory.CreatePreactorObject(preactorComObject)
        BeginSchedulerDebug(preactor)
        Dim planningboard As IPlanningBoard = preactor.PlanningBoard
        Dim ordersTable As Integer = preactor.FindFirstClassificationString("LAUNCH TIME").Value.FormatNumber

        Dim routingDt As DataTable = readOrderTable(preactor)
        Dim currentDate As DateTime = planningboard.TerminatorTime

        ' 1. Optimization: Cache Kiln Resources in a Dictionary for fast O(1) lookups
        Dim kilnResources As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase) From {
        {"AKLN", planningboard.GetResourceNumber("AKLN")},
        {"BKLN", planningboard.GetResourceNumber("BKLN")},
        {"CKLN", planningboard.GetResourceNumber("CKLN")},
        {"DKLN", planningboard.GetResourceNumber("DKLN")},
        {"RKLN", planningboard.GetResourceNumber("RKLN")},
        {"NKLN", planningboard.GetResourceNumber("NKLN")}
    }

        Dim LOADBICK As Integer = planningboard.GetResourceNumber("LOADBICK")
        Dim ULDBICK As Integer = planningboard.GetResourceNumber("ULDBICK")
        Dim PREINSPC As Integer = planningboard.GetResourceNumber("PREINSPC")
        Dim KILNACK As Integer = planningboard.GetResourceNumber("KILNACK")

        ' Fetch Optimizer Settings
        Dim GNOptimizerSettings As Integer = preactor.GetFormatNumber("GN Optimizer Settings")
        Dim GNOptimizerSettings_Numeric As Integer = preactor.GetFieldNumber(GNOptimizerSettings, "Numeric Value")
        Dim GNOptimizerSettings_Boolean As Integer = preactor.GetFieldNumber(GNOptimizerSettings, "Toggle Value")

        Dim maxOcc As Double = preactor.ReadFieldDouble(GNOptimizerSettings, GNOptimizerSettings_Numeric, 2)
        Dim minOcc As Double = preactor.ReadFieldDouble(GNOptimizerSettings, GNOptimizerSettings_Numeric, 1)
        Dim allowUnderfilledTail As Boolean = preactor.ReadFieldInt(GNOptimizerSettings, GNOptimizerSettings_Boolean, 3) = 1
        Dim batchStartDelayMins As Integer = preactor.ReadFieldInt(GNOptimizerSettings, GNOptimizerSettings_Numeric, 4)

        Dim configDir As String = preactor.ParseShellString("{PATH}")
        Dim debugFolder As String = configDir & "\Debug\Firing_" & DateTime.Now.ToString("yyyyMMdd_HHmmss")

        Dim enableDebugExport As Boolean = False
        Dim firingObj As New firingOptimizer_vf()

        If enableDebugExport Then
            firingObj.ExportFiringCandidateDebug(routingDt, debugFolder)
        End If

        Dim batchKilnNames As New List(Of String) From {
        "AKLN", "BKLN", "CKLN", "DKLN", "RKLN", "NKLN"
        }

        Dim batchInitialAvailability As Dictionary(Of String, DateTime) =
        SharedHelpers.BuildEffectiveStartByResourceFromGnKilnAvailability(preactor,
                                                                      planningboard,
                                                                      batchKilnNames)
        ' Build firing plan using firing optimizer
        'Dim plan = firingObj.BuildBatchKilnPlan(routingDt, configDir & "\kilndata.csv", currentDate, minOcc, maxOcc,
        '                                    allowUnderfilledTail:=True,
        '                                    batchStartDelayMins:=60,
        '                                    maxBatchesPerDayGlobal:=2)
        Dim plan = firingObj.BuildBatchKilnPlan(routingDt,
                                        configDir & "\kilndata.csv",
                                        currentDate,
                                        minOcc,
                                        maxOcc,
                                        allowUnderfilledTail:=True,
                                        batchStartDelayMins:=60,
                                        maxBatchesPerDayGlobal:=2,
                                        initialKilnAvailability:=batchInitialAvailability,
                                        debug:=_schedulerDebug)

        ' Debugger
        If enableDebugExport Then
            firingObj.ExportPlanToCsv(plan, debugFolder)
        End If

        ' 2. Optimization: Declare loop variables outside to reduce stack allocations
        Dim Times As OperationTimes?
        Dim batchNo As Integer
        Dim batchStart As DateTime
        Dim kilnName As String
        Dim kilnResId As Integer

        ' Iterate firing queue (these are op 300 record numbers)
        For Each firingOpRec As Integer In plan.QueueFiringOpRecs

            ' Get batch metadata
            batchNo = plan.BatchNoByFiringOpRec(firingOpRec)
            batchStart = plan.BatchStartByBatchNo(batchNo)
            kilnName = plan.KilnByBatchNo(batchNo)

            ' 3. Optimization: Replace slow Select Case with instant Dictionary lookup
            If kilnResources.TryGetValue(kilnName, kilnResId) Then
                If planningboard.IsOperationScheduled(firingOpRec) Then Continue For
                PutOperationWithTrace(preactor, planningboard, "BatchFiring", firingOpRec, kilnResId, batchStart, "Forward")
            End If

            ' Handle Previous Operation
            Dim PREVIOUSOP As Integer = planningboard.GetPreviousOperation(firingOpRec, 1)
            If PREVIOUSOP > 0 AndAlso
                Not SharedHelpers.IsCompletedOrActualizedOp(routingDt, PREVIOUSOP) Then
                Try
                    Times = planningboard.BackTestOpOnResource(PREVIOUSOP, LOADBICK, batchStart)
                    If Times.HasValue Then
                        PutOperationWithTrace(preactor, planningboard, "BatchLoading", PREVIOUSOP, LOADBICK, Times.Value.ProcessStart, "Backward")
                    End If
                Catch ex As Exception
                    ' Silently handled per original logic
                End Try
            End If

            ' Handle Next Operations
            Dim NEXTOP As Integer = planningboard.GetNextOperation(firingOpRec, 1)
            If NEXTOP > 0 AndAlso
   Not SharedHelpers.IsCompletedOrActualizedOp(routingDt, NEXTOP) Then
                ' LAZY EVALUATION: Maintained from original logic
                Dim batchEnd As DateTime = plan.BatchEndByBatchNo(batchNo)

                Try
                    Times = planningboard.TestOperationOnResource(NEXTOP, ULDBICK, batchEnd)
                    If Times.HasValue Then
                        PutOperationWithTrace(preactor, planningboard, "BatchUnloading", NEXTOP, ULDBICK, Times.Value.ProcessStart, "Forward")
                    End If
                Catch ex As Exception
                End Try

                ' Re-evaluate for the subsequent operation in the sequence
                NEXTOP = planningboard.GetNextOperation(NEXTOP, 1)
                If NEXTOP > 0 AndAlso
   Not SharedHelpers.IsCompletedOrActualizedOp(routingDt, NEXTOP) Then
                    Try
                        Times = planningboard.TestOperationOnResource(NEXTOP, PREINSPC, batchEnd)
                        If Times.HasValue Then
                            PutOperationWithTrace(preactor, planningboard, "PostFiring", NEXTOP, PREINSPC, Times.Value.ProcessStart.AddDays(2), "Forward") '2 days
                        End If
                    Catch ex As Exception
                    End Try

                    NEXTOP = planningboard.GetNextOperation(NEXTOP, 1)
                    If NEXTOP > 0 AndAlso
   Not SharedHelpers.IsCompletedOrActualizedOp(routingDt, NEXTOP) Then
                        Try
                            Times = planningboard.TestOperationOnResource(NEXTOP, KILNACK, batchEnd)
                            If Times.HasValue Then
                                PutOperationWithTrace(preactor, planningboard, "PostFiring", NEXTOP, KILNACK, Times.Value.ProcessStart, "Forward")
                            End If
                        Catch ex As Exception
                        End Try
                    End If
                End If
            End If
        Next

        FinishSchedulerDebug(preactor)
        preactor.DestroyStatus()
        Return 0
    End Function
    'Public Function runFiring(ByRef preactorComObject As PreactorObj, ByRef pespComObject As Object) As Integer
    '    ' Batch firing logic

    '    Dim preactor As IPreactor = PreactorFactory.CreatePreactorObject(preactorComObject)
    '    Dim planningboard As IPlanningBoard = preactor.PlanningBoard
    '    Dim ordersTable As Integer = preactor.FindFirstClassificationString("LAUNCH TIME").Value.FormatNumber

    '    ' Example: import a CSV, build pressing queue, create ranked queue and schedule
    '    'Dim filePath As String = "D:\Documents\Opcenter\Cases\Grindwell Norton\Opcenter SC - Dev\Files\Templates\Routing.csv"

    '    'Dim routingDt As DataTable = ImportRoutingCsvToDataTable(filePath)
    '    Dim routingDt As DataTable = readOrderTable(preactor)

    '    'Dim currentDate As New System.DateTime(2025, 8, 1, 0, 0, 0)
    '    Dim currentDate As DateTime = planningboard.TerminatorTime

    '    'Append schedule times from board for a few operation numbers (example)

    '    Dim AKLN As Integer = planningboard.GetResourceNumber("AKLN")
    '    Dim BKLN As Integer = planningboard.GetResourceNumber("BKLN")
    '    Dim CKLN As Integer = planningboard.GetResourceNumber("CKLN")
    '    Dim DKLN As Integer = planningboard.GetResourceNumber("DKLN")
    '    Dim RKLN As Integer = planningboard.GetResourceNumber("RKLN")
    '    Dim NKLN As Integer = planningboard.GetResourceNumber("NKLN")
    '    Dim LOADBICK As Integer = planningboard.GetResourceNumber("LOADBICK")
    '    Dim ULDBICK As Integer = planningboard.GetResourceNumber("ULDBICK")
    '    Dim PREINSPC As Integer = planningboard.GetResourceNumber("PREINSPC")
    '    Dim KILNACK As Integer = planningboard.GetResourceNumber("KILNACK")
    '    Dim GNOptimizerSettings As Integer = preactor.GetFormatNumber("GN Optimizer Settings")
    '    Dim GNOptimizerSettings_Numeric As Integer = preactor.GetFieldNumber(GNOptimizerSettings, "Numeric Value")
    '    Dim GNOptimizerSettings_Boolean As Integer = preactor.GetFieldNumber(GNOptimizerSettings, "Toggle Value")


    '    ' Build firing plan using firing optimizer (external class)
    '    ' Dim opSettings As FormatFieldPair
    '    Dim maxOcc As Double = preactor.ReadFieldDouble(GNOptimizerSettings, GNOptimizerSettings_Numeric, 2)
    '    Dim minOcc As Double = preactor.ReadFieldDouble(GNOptimizerSettings, GNOptimizerSettings_Numeric, 1)
    '    Dim allowUnderfilledTail As Boolean = preactor.ReadFieldInt(GNOptimizerSettings, GNOptimizerSettings_Boolean, 3) = 1
    '    Dim batchStartDelayMins As Integer = preactor.ReadFieldInt(GNOptimizerSettings, GNOptimizerSettings_Numeric, 4)
    '    'Dim maxBatchesPerDayGlobal As Integer = preactor.ReadFieldInt(GNOptimizerSettings, GNOptimizerSettings_Numeric, 5)

    '    Dim configDir As String = preactor.ParseShellString("{PATH}")
    '    Dim debugFolder As String = configDir & "\Debug\Firing_" & DateTime.Now.ToString("yyyyMMdd_HHmmss")

    '    ' Adding a boolean flag to control debug export, so you can easily turn it on/off without commenting code
    '    Dim enableDebugExport As Boolean = False
    '    Dim firingObj As New firingOptimizer_vf()

    '    If enableDebugExport Then
    '        firingObj.ExportFiringCandidateDebug(routingDt, debugFolder)
    '    End If
    '    Dim ConfigPath As String = preactor.ParseShellString("{PATH}")

    '    Dim plan = firingObj.BuildBatchKilnPlan(routingDt, ConfigPath & "\kilndata.csv", currentDate, minOcc, maxOcc,
    '                                            allowUnderfilledTail:=True,
    '                                            batchStartDelayMins:=60,
    '                                            maxBatchesPerDayGlobal:=2)


    '    ' Debugger
    '    If enableDebugExport Then
    '        firingObj.ExportPlanToCsv(plan, debugFolder)
    '    End If

    '    ' 1) iterate firing queue (these are op 300 record numbers)
    '    For Each firingOpRec As Integer In plan.QueueFiringOpRecs

    '        ' 2) get batch metadata
    '        Dim batchNo As Integer = plan.BatchNoByFiringOpRec(firingOpRec)
    '        Dim batchStart As DateTime = plan.BatchStartByBatchNo(batchNo)
    '        'Dim batchEnd As DateTime = plan.BatchEndByBatchNo(batchNo)
    '        Dim kilnName As String = plan.KilnByBatchNo(batchNo)
    '        Dim batchKind As String = plan.BatchKindByBatchNo(batchNo)
    '        Dim Times As OperationTimes?



    '        Select Case (kilnName)
    '            Case "AKLN"
    '                planningboard.PutOperationOnResource(firingOpRec, AKLN, batchStart)
    '            Case "BKLN"
    '                planningboard.PutOperationOnResource(firingOpRec, BKLN, batchStart)
    '            Case "CKLN"
    '                planningboard.PutOperationOnResource(firingOpRec, CKLN, batchStart)
    '            Case "DKLN"
    '                planningboard.PutOperationOnResource(firingOpRec, DKLN, batchStart)
    '            Case "RKLN"
    '                planningboard.PutOperationOnResource(firingOpRec, RKLN, batchStart)
    '            Case "NKLN"
    '                planningboard.PutOperationOnResource(firingOpRec, NKLN, batchStart)
    '        End Select

    '        ' 3) Handle Previous Operation
    '        Dim PREVIOUSOP As Integer = planningboard.GetPreviousOperation(firingOpRec, 1)
    '        If PREVIOUSOP > 0 Then
    '            Try
    '                Times = planningboard.BackTestOpOnResource(PREVIOUSOP, LOADBICK, batchStart)
    '                If Times.HasValue Then
    '                    planningboard.PutOperationOnResource(PREVIOUSOP, LOADBICK, Times.Value.ProcessStart)
    '                End If
    '            Catch ex As Exception

    '            End Try
    '        End If

    '        ' 4) Handle Next Operations
    '        Dim NEXTOP As Integer = planningboard.GetNextOperation(firingOpRec, 1)
    '        If NEXTOP > 0 Then
    '            ' LAZY EVALUATION: Only query batchEnd if a NEXTOP actually exists.
    '            ' This saves execution time by avoiding an unnecessary lookup.
    '            Dim batchEnd As DateTime = plan.BatchEndByBatchNo(batchNo)

    '            Try
    '                Times = planningboard.TestOperationOnResource(NEXTOP, ULDBICK, batchEnd)
    '                If Times.HasValue Then
    '                    planningboard.PutOperationOnResource(NEXTOP, ULDBICK, Times.Value.ProcessStart)
    '                End If
    '            Catch ex As Exception

    '            End Try

    '            ' Re-evaluate for the subsequent operation in the sequence
    '            NEXTOP = planningboard.GetNextOperation(NEXTOP, 1)
    '            If NEXTOP > 0 Then

    '                Try
    '                    Times = planningboard.TestOperationOnResource(NEXTOP, PREINSPC, batchEnd)
    '                    If Times.HasValue Then
    '                        planningboard.PutOperationOnResource(NEXTOP, PREINSPC, Times.Value.ProcessStart.AddDays(2)) '2 days
    '                    End If
    '                Catch ex As Exception

    '                End Try

    '                NEXTOP = planningboard.GetNextOperation(NEXTOP, 1)
    '                If NEXTOP > 0 Then
    '                    Try
    '                        Times = planningboard.TestOperationOnResource(NEXTOP, KILNACK, batchEnd)
    '                        If Times.HasValue Then
    '                            planningboard.PutOperationOnResource(NEXTOP, KILNACK, Times.Value.ProcessStart)
    '                        End If
    '                    Catch ex As Exception

    '                    End Try
    '                End If
    '            End If
    '        End If

    '    Next

    '    preactor.DestroyStatus()
    '    Return 0
    'End Function

    Public Function runSWKFiring(ByRef preactorComObject As PreactorObj,
                             ByRef pespComObject As Object) As Integer

        Dim preactor As IPreactor = PreactorFactory.CreatePreactorObject(preactorComObject)
        BeginSchedulerDebug(preactor)
        Dim planningboard As IPlanningBoard = preactor.PlanningBoard

        Dim ordersTable As Integer = preactor.FindFirstClassificationString("LAUNCH TIME").Value.FormatNumber


        Dim routingDt As DataTable = readOrderTable(preactor)
        Dim currentDate As DateTime = planningboard.TerminatorTime

        'Dim swkMetadataDate As DateTime =
        'SharedHelpers.ReadOptimizerSettingDate(preactor,
        '                                   "SWBKILN Available From",
        'DateTime.MinValue)

        Dim swkStart As DateTime =
        SharedHelpers.GetEffectiveStartFromGnKilnAvailability(preactor,
                                                          planningboard,
                                                          "SWBKILN")        ' ------------------------------------------------------------
        ' SWK verified parameters
        ' ------------------------------------------------------------
        Dim swkMinTonnage As Double = 0.8
        Dim swkMaxTonnage As Double = 1.0
        Dim swkDailyBatchLimit As Integer = 2
        Dim swkBatchStartDelayMins As Integer = 60
        Dim swkAllowUnderfilledTail As Boolean = True

        ' Later replace above constants with GN Optimizer Settings reads:
        '   SWK Min Tonnage
        '   SWK Max Tonnage
        '   SWK Daily Batch Limit
        '   SWK Batch Start Delay Mins
        '   SWK Allow Underfilled Tail

        Dim SWBKILN As Integer = planningboard.GetResourceNumber("SWBKILN")
        Dim LOADSW As Integer = planningboard.GetResourceNumber("LOADSW")
        Dim ULDSW As Integer = planningboard.GetResourceNumber("ULDSW")
        Dim PREINSPC As Integer = planningboard.GetResourceNumber("PREINSPC")
        Dim KILNACK As Integer = planningboard.GetResourceNumber("KILNACK")

        If SWBKILN <= 0 Then Throw New Exception("SWK resource not found: SWBKILN")
        If LOADSW <= 0 Then Throw New Exception("SWK loading resource not found: LOADSW")
        If ULDSW <= 0 Then Throw New Exception("SWK unloading resource not found: ULDSW")
        If PREINSPC <= 0 Then Throw New Exception("Resource not found: PREINSPC")
        If KILNACK <= 0 Then Throw New Exception("Resource not found: KILNACK")

        Dim swkObj As New swkOptimizer_vf()

        Dim plan As swkOptimizer_vf.SwkBatchPlan =
        swkObj.BuildSwkPlan(routingDt,
                            swkStart,
                            swkMinTonnage,
                            swkMaxTonnage,
                            dailyBatchLimit:=swkDailyBatchLimit,
                            batchStartDelayMins:=swkBatchStartDelayMins,
                            allowUnderfilledTail:=swkAllowUnderfilledTail,
                            swkResourceName:="SWBKILN",
                            debug:=_schedulerDebug)

        Dim configDir As String = preactor.ParseShellString("{PATH}")
        Dim debugFolder As String = configDir & "\Debug\SWK_" & DateTime.Now.ToString("yyyyMMdd_HHmmss")
        Dim enableDebugExport As Boolean = False

        If enableDebugExport Then
            swkObj.ExportSwkPlanToCsv(plan, debugFolder)
        End If

        For Each firingOpRec As Integer In plan.QueueFiringOpRecs

            If Not plan.BatchNoByFiringOpRec.ContainsKey(firingOpRec) Then Continue For

            Dim batchNo As Integer = plan.BatchNoByFiringOpRec(firingOpRec)
            Dim batchStart As DateTime = plan.BatchStartByBatchNo(batchNo)
            Dim batchEnd As DateTime = plan.BatchEndByBatchNo(batchNo)

            ' --------------------------------------------------------
            ' 1. Schedule firing op 300 on SWBKILN
            ' --------------------------------------------------------
            Dim firingTimes As OperationTimes? =
            planningboard.TestOperationOnResource(firingOpRec, SWBKILN, batchStart)

            If planningboard.IsOperationScheduled(firingOpRec) Then Continue For
            'If firingTimes.HasValue Then
            PutOperationWithTrace(preactor, planningboard, "SWK", firingOpRec, SWBKILN, batchStart, "Forward")
            'Else
            'System.Diagnostics.Debug.WriteLine("SWK: Cannot place firing op " &
            'firingOpRec &
            '                                   " on SWBKILN at " &
            'batchStart.ToString("yyyy-MM-dd HH:mm:ss"))
            '    Continue For
            'End If

            ' --------------------------------------------------------
            ' 2. Schedule previous loading operation on LOADSW
            ' --------------------------------------------------------
            Dim previousOp As Integer = planningboard.GetPreviousOperation(firingOpRec, 1)

            If previousOp > 0 Then
                Dim loadTimes As OperationTimes? =
                planningboard.BackTestOpOnResource(previousOp, LOADSW, batchStart)

                If loadTimes.HasValue Then
                    PutOperationWithTrace(preactor, planningboard, "SWKLoading", previousOp, LOADSW, loadTimes.Value.ProcessStart, "Backward")
                Else
                    System.Diagnostics.Debug.WriteLine("SWK: Cannot back-schedule LOADSW for previous op " &
                                                   previousOp &
                                                   ", firing op " &
                                                   firingOpRec)
                End If
            End If

            ' --------------------------------------------------------
            ' 3. Schedule next unloading operation on ULDSW
            ' --------------------------------------------------------
            Dim nextOp As Integer = planningboard.GetNextOperation(firingOpRec, 1)

            If nextOp > 0 Then

                Dim unloadTimes As OperationTimes? =
                planningboard.TestOperationOnResource(nextOp, ULDSW, batchEnd)

                If unloadTimes.HasValue Then
                    PutOperationWithTrace(preactor, planningboard, "SWKUnloading", nextOp, ULDSW, unloadTimes.Value.ProcessStart, "Forward")
                Else
                    System.Diagnostics.Debug.WriteLine("SWK: Cannot schedule ULDSW for next op " &
                                                   nextOp &
                                                   ", firing op " &
                                                   firingOpRec)
                    Continue For
                End If

                ' ----------------------------------------------------
                ' 4. PREINSPC
                ' ----------------------------------------------------
                nextOp = planningboard.GetNextOperation(nextOp, 1)

                If nextOp > 0 Then

                    Dim preInspTimes As OperationTimes? =
                    planningboard.TestOperationOnResource(nextOp, PREINSPC, batchEnd)

                    If preInspTimes.HasValue Then
                        PutOperationWithTrace(preactor, planningboard, "PostFiring", nextOp, PREINSPC, preInspTimes.Value.ProcessStart.AddDays(1), "Forward")
                    Else
                        System.Diagnostics.Debug.WriteLine("SWK: Cannot schedule PREINSPC for op " & nextOp)
                        Continue For
                    End If

                    ' ------------------------------------------------
                    ' 5. KILNACK
                    ' ------------------------------------------------
                    nextOp = planningboard.GetNextOperation(nextOp, 1)

                    If nextOp > 0 Then
                        Dim ackTimes As OperationTimes? =
                        planningboard.TestOperationOnResource(nextOp, KILNACK, batchEnd)

                        If ackTimes.HasValue Then
                            PutOperationWithTrace(preactor, planningboard, "PostFiring", nextOp, KILNACK, ackTimes.Value.ProcessStart, "Forward")
                        Else
                            System.Diagnostics.Debug.WriteLine("SWK: Cannot schedule KILNACK for op " & nextOp)
                        End If
                    End If
                End If
            End If

        Next

        FinishSchedulerDebug(preactor)
        preactor.DestroyStatus()
        Return 0

    End Function

    Public Function runFiring2(ByRef preactorComObject As PreactorObj, ByRef pespComObject As Object) As Integer
        Dim preactor As IPreactor = PreactorFactory.CreatePreactorObject(preactorComObject)
        BeginSchedulerDebug(preactor)
        Dim planningboard As IPlanningBoard = preactor.PlanningBoard

        Dim ordersTable As Integer = preactor.FindFirstClassificationString("LAUNCH TIME").Value.FormatNumber
        Dim routingDt As DataTable = readOrderTable(preactor)
        Dim currentDate As DateTime = planningboard.TerminatorTime

        Dim GNOptimizerSettings As Integer = preactor.GetFormatNumber("GN Optimizer Settings")
        Dim GNOptimizerSettings_Numeric As Integer = preactor.GetFieldNumber(GNOptimizerSettings, "Numeric Value")
        Dim maxOcc As Double = preactor.ReadFieldDouble(GNOptimizerSettings, GNOptimizerSettings_Numeric, 9)
        Dim minOccPreferred As Double = preactor.ReadFieldDouble(GNOptimizerSettings, GNOptimizerSettings_Numeric, 8)

        ' Parameters
        Dim totalCartsAvailable As Integer = preactor.ReadFieldInt(GNOptimizerSettings, GNOptimizerSettings_Numeric, 7)
        Dim cartsPerDay As Double = preactor.ReadFieldDouble(GNOptimizerSettings, GNOptimizerSettings_Numeric, 5)
        Dim dryingToFiringBufferHours As Double = preactor.ReadFieldDouble(GNOptimizerSettings, GNOptimizerSettings_Numeric, 8) / 60

        Dim TCBK As Integer = planningboard.GetResourceNumber("TCBK")
        Dim LOADPTK As Integer = planningboard.GetResourceNumber("LOADPTK")
        Dim ULDPTK As Integer = planningboard.GetResourceNumber("ULDPTK")
        Dim BATCHTIME As Integer = preactor.GetFieldNumber(ordersTable, "Batch Time")
        Dim PREINSPC As Integer = planningboard.GetResourceNumber("PREINSPC")
        Dim KILNACK As Integer = planningboard.GetResourceNumber("KILNACK")
        Dim FTDSD20 As Integer = planningboard.GetResourceNumber("FTDSD20")

        'Dim tunnelMetadataDate As DateTime =
        'SharedHelpers.ReadOptimizerSettingDate(preactor,
        '                                  "TCBK Available From",
        'DateTime.MinValue)

        Dim tunnelStart As DateTime =
        SharedHelpers.GetEffectiveStartFromGnKilnAvailability(preactor,
                                                          planningboard,
                                                          "TCBK")

        Dim tunnelObj As New tunnelOptimizer_vf()

        Dim plan = tunnelObj.BuildTunnelPlan(
        routingDt,
        startTime:=tunnelStart,
        cartsPerDay:=cartsPerDay,
        totalCarts:=totalCartsAvailable,
        minOccPreferred:=minOccPreferred,
        maxOcc:=maxOcc,
        dryingToFiringBufferHours:=dryingToFiringBufferHours,
        debug:=_schedulerDebug
    )

        ' 1. Optimization: Pre-calculate constants outside the loop
        Dim batchTimeValue As Double = totalCartsAvailable / cartsPerDay

        ' 2. Optimization: Declare loop variables outside to prevent repeated stack allocations
        Dim cartNo As Integer
        Dim batchstart As DateTime
        Dim batchStartOffset As DateTime
        Dim PREVIOUSOP As Integer
        Dim NEXTOP As Integer
        Dim NEXTOPNO As Integer
        Dim opTimes As OperationTimes? ' Used to test .HasValue safely

        For Each firingOpRec As Integer In plan.CartNoByFiringOpRec.Keys

            cartNo = plan.CartNoByFiringOpRec(firingOpRec)
            batchstart = plan.StartByFiringOpRec(firingOpRec)

            ' 3. Optimization: Calculate the 2-hour offset once per iteration
            batchStartOffset = batchstart.AddHours(2)

            If planningboard.IsOperationScheduled(firingOpRec) Then Continue For
            preactor.WriteField(ordersTable, BATCHTIME, firingOpRec, batchTimeValue)
            PutOperationWithTrace(preactor, planningboard, "TunnelFiring", firingOpRec, TCBK, batchStartOffset, "Forward")

            ' Handle Previous Operation
            PREVIOUSOP = planningboard.GetPreviousOperation(firingOpRec, 1)
            If PREVIOUSOP > 0 Then
                opTimes = planningboard.BackTestOpOnResource(PREVIOUSOP, LOADPTK, batchStartOffset)
                If opTimes.HasValue Then
                    PutOperationWithTrace(preactor, planningboard, "TunnelLoading", PREVIOUSOP, LOADPTK, opTimes.Value.ProcessStart, "Backward")
                End If
            End If

            ' Handle Next Operations
            NEXTOP = planningboard.GetNextOperation(firingOpRec, 1)
            If NEXTOP > 0 Then
                opTimes = planningboard.TestOperationOnResource(NEXTOP, ULDPTK, batchStartOffset)
                If opTimes.HasValue Then
                    PutOperationWithTrace(preactor, planningboard, "TunnelUnloading", NEXTOP, ULDPTK, opTimes.Value.ProcessStart, "Forward")
                End If

                ' Re-evaluate for the subsequent operation in the sequence
                NEXTOP = planningboard.GetNextOperation(NEXTOP, 1)

                If NEXTOP > 0 Then
                    NEXTOPNO = preactor.ReadFieldInt(ordersTable, "Op. No.", NEXTOP)

                    If NEXTOPNO = 320 Then
                        opTimes = planningboard.TestOperationOnResource(NEXTOP, FTDSD20, batchStartOffset)
                        If opTimes.HasValue Then
                            PutOperationWithTrace(preactor, planningboard, "PostFiring", NEXTOP, FTDSD20, opTimes.Value.ProcessStart, "Forward")
                        End If
                        NEXTOP = planningboard.GetNextOperation(NEXTOP, 1)
                    End If

                    If NEXTOP > 0 Then
                        ' 4. Optimization: Removed double-calling of TestOperationOnResource
                        opTimes = planningboard.TestOperationOnResource(NEXTOP, PREINSPC, batchStartOffset)
                        If opTimes.HasValue Then
                            PutOperationWithTrace(preactor, planningboard, "PostFiring", NEXTOP, PREINSPC, opTimes.Value.ProcessStart.AddDays(1), "Forward") '1 day offset
                        End If

                        NEXTOP = planningboard.GetNextOperation(NEXTOP, 1)
                        If NEXTOP > 0 Then
                            opTimes = planningboard.TestOperationOnResource(NEXTOP, KILNACK, batchStartOffset)
                            If opTimes.HasValue Then
                                PutOperationWithTrace(preactor, planningboard, "PostFiring", NEXTOP, KILNACK, opTimes.Value.ProcessStart, "Forward")
                            End If
                        End If
                    End If
                End If
            End If
        Next

        FinishSchedulerDebug(preactor)
        preactor.DestroyStatus()
        Return 0
    End Function
    'Public Function runFiring2(ByRef preactorComObject As PreactorObj, ByRef pespComObject As Object) As Integer
    '    Dim preactor As IPreactor = PreactorFactory.CreatePreactorObject(preactorComObject)
    '    Dim planningboard As IPlanningBoard = preactor.PlanningBoard

    '    Dim ordersTable As Integer = preactor.FindFirstClassificationString("LAUNCH TIME").Value.FormatNumber
    '    Dim routingDt As DataTable = readOrderTable(preactor)

    '    'Dim currentDate As New System.DateTime(2025, 8, 1, 0, 0, 0)
    '    Dim currentDate As DateTime = planningboard.TerminatorTime

    '    Dim GNOptimizerSettings As Integer = preactor.GetFormatNumber("GN Optimizer Settings")
    '    Dim GNOptimizerSettings_Numeric As Integer = preactor.GetFieldNumber(GNOptimizerSettings, "Numeric Value")
    '    Dim maxOcc As Double = preactor.ReadFieldDouble(GNOptimizerSettings, GNOptimizerSettings_Numeric, 9)
    '    Dim minOccPreferred As Double = preactor.ReadFieldDouble(GNOptimizerSettings, GNOptimizerSettings_Numeric, 8)

    '    ' Parameters you will provide
    '    Dim totalCartsAvailable As Integer = preactor.ReadFieldInt(GNOptimizerSettings, GNOptimizerSettings_Numeric, 7)
    '    Dim cartsPerDay As Double = preactor.ReadFieldDouble(GNOptimizerSettings, GNOptimizerSettings_Numeric, 5)
    '    Dim dryingToFiringBufferHours As Double = preactor.ReadFieldDouble(GNOptimizerSettings, GNOptimizerSettings_Numeric, 8) / 60
    '    Dim TCBK As Integer = planningboard.GetResourceNumber("TCBK")
    '    Dim LOADPTK As Integer = planningboard.GetResourceNumber("LOADPTK")
    '    Dim ULDPTK As Integer = planningboard.GetResourceNumber("ULDPTK")
    '    Dim BATCHTIME As Integer = preactor.GetFieldNumber(ordersTable, "Batch Time")
    '    Dim PREINSPC As Integer = planningboard.GetResourceNumber("PREINSPC")
    '    Dim KILNACK As Integer = planningboard.GetResourceNumber("KILNACK")
    '    Dim FTDSD20 As Integer = planningboard.GetResourceNumber("FTDSD20")

    '    Dim PREVIOUSOP As Integer
    '    'Dim NextOprec As Integer
    '    'Dim PrevOpRecStart As DateTime
    '    'Dim NextOpRecStart As DateTime
    '    Dim NEXTOPNO As Integer


    '    Dim tunnelObj As New tunnelOptimizer_vf

    '    ' currentDate = your scheduling anchor (same as you used earlier)
    '    ' This is the start cursor for cart generation. Your logic can choose:
    '    ' - currentDate, or
    '    ' - earliest ReadyTime in the dataset
    '    Dim plan = tunnelObj.BuildTunnelPlan(
    '    routingDt,
    '    startTime:=currentDate,
    '    cartsPerDay:=cartsPerDay,
    '    totalCarts:=totalCartsAvailable,
    '    minOccPreferred:=minOccPreferred,
    '    maxOcc:=maxOcc,
    '    dryingToFiringBufferHours:=dryingToFiringBufferHours
    '    )
    '    Dim cartNo As Integer
    '    Dim batchstart As DateTime

    '    For Each firingOpRec As Integer In plan.CartNoByFiringOpRec.Keys

    '        cartNo = plan.CartNoByFiringOpRec(firingOpRec)
    '        batchstart = plan.StartByFiringOpRec(firingOpRec)
    '        preactor.WriteField(ordersTable, BATCHTIME, firingOpRec, totalCartsAvailable / cartsPerDay)
    '        'planningboard.PutOperationOnResource(firingOpRec, TCBK, batchstart.AddDays(1))
    '        planningboard.PutOperationOnResource(firingOpRec, TCBK, batchstart.AddHours(2))

    '        ' 3) Handle Previous Operation
    '        PREVIOUSOP = planningboard.GetPreviousOperation(firingOpRec, 1)
    '        If PREVIOUSOP > 0 Then
    '            planningboard.PutOperationOnResource(PREVIOUSOP, LOADPTK, planningboard.BackTestOpOnResource(PREVIOUSOP, LOADPTK, batchstart.AddHours(2)).Value.ProcessStart)
    '        End If

    '        ' 4) Handle Next Operations
    '        Dim NEXTOP As Integer = planningboard.GetNextOperation(firingOpRec, 1)
    '        If NEXTOP > 0 Then
    '            ' LAZY EVALUATION: Only query batchEnd if a NEXTOP actually exists.
    '            ' This saves execution time by avoiding an unnecessary lookup.
    '            planningboard.PutOperationOnResource(NEXTOP, ULDPTK, planningboard.TestOperationOnResource(NEXTOP, ULDPTK, batchstart.AddHours(2)).Value.ProcessStart)

    '            ' Re-evaluate for the subsequent operation in the sequence
    '            NEXTOP = planningboard.GetNextOperation(NEXTOP, 1)
    '            NEXTOPNO = preactor.ReadFieldInt(ordersTable, "Op. No.", NEXTOP)
    '            If NEXTOPNO = 320 Then
    '                planningboard.PutOperationOnResource(NEXTOP, FTDSD20, planningboard.TestOperationOnResource(NEXTOP, FTDSD20, batchstart.AddHours(2)).Value.ProcessStart)
    '                NEXTOP = planningboard.GetNextOperation(NEXTOP, 1)
    '            End If
    '            If NEXTOP > 0 Then
    '                Dim testop As OperationTimes? = planningboard.TestOperationOnResource(NEXTOP, PREINSPC, batchstart.AddHours(2))
    '                planningboard.PutOperationOnResource(NEXTOP, PREINSPC, planningboard.TestOperationOnResource(NEXTOP, PREINSPC, batchstart.AddHours(2)).Value.ProcessStart.AddDays(1)) '2 days
    '                NEXTOP = planningboard.GetNextOperation(NEXTOP, 1)
    '                If NEXTOP > 0 Then
    '                    Dim testop2 As OperationTimes? = planningboard.TestOperationOnResource(NEXTOP, KILNACK, batchstart.AddHours(2))
    '                    planningboard.PutOperationOnResource(NEXTOP, KILNACK, planningboard.TestOperationOnResource(NEXTOP, KILNACK, batchstart.AddHours(2)).Value.ProcessStart)
    '                End If

    '            End If
    '        End If
    '    Next

    '    preactor.DestroyStatus()

    '    Return 0
    'End Function

    Public Function runFix(ByRef preactorComObject As PreactorObj, ByRef pespComObject As Object) As Integer
        ' Initialize Preactor and Planning Board objects
        Dim preactor As IPreactor = PreactorFactory.CreatePreactorObject(preactorComObject)
        Dim planningboard As IPlanningBoard = preactor.PlanningBoard

        ' Get table and field references
        Dim ordersTable As Integer = preactor.FindFirstClassificationString("LAUNCH TIME").Value.FormatNumber
        Dim opNoField As Integer = preactor.GetFieldNumber(ordersTable, "Op. No.")
        Dim currentDate As DateTime = planningboard.TerminatorTime
        Dim DRYER As Integer = planningboard.GetResourceNumber("DRYER")

        Dim reccount As Integer = preactor.RecordCount(ordersTable)

        ' OPTIMIZATION: Pre-declare variables outside the loop to reduce stack overhead
        Dim recOpNo As Integer
        Dim nxtrecOpNo As Integer
        Dim nextOpIndex As Integer
        Dim oprec As Integer
        Dim resrecs As IEnumerable(Of Integer)
        Dim resrec As Integer
        Dim targetTime As DateTime
        Dim nextOpStartTime As DateTime
        Dim times2 As DateTime
        Dim times As OperationTimes? ' This remains valid as it was in your original code

        ' ==========================================
        ' OPTIMIZATION: COMBINED LOOP
        ' We process Phase 1 and Phase 2 simultaneously, eliminating the need to loop through 'reccount' twice.
        ' ==========================================
        For i As Integer = 1 To reccount
            Try
                ' Read the Operation Number exactly once per record
                recOpNo = preactor.ReadFieldInt(ordersTable, opNoField, i)

                ' -----------------------------------------------------
                ' PHASE 1: DRYER Fix (Op 260 -> Op 290)
                ' -----------------------------------------------------
                If recOpNo = 260 Then
                    nextOpIndex = planningboard.GetNextOperation(i, 1)

                    If nextOpIndex > 0 Then
                        nxtrecOpNo = preactor.ReadFieldInt(ordersTable, opNoField, nextOpIndex)

                        If nxtrecOpNo = 290 Then
                            ' Restored your exact original syntax here
                            nextOpStartTime = planningboard.GetOperationTimes(i + 1).Value.OperationTimes.ProcessStart

                            If nextOpStartTime >= currentDate Then
                                times = planningboard.BackTestOpOnResource(i, DRYER, nextOpStartTime)

                                ' Safe check using .HasValue on the standard OperationTimes? object
                                If times.HasValue Then
                                    planningboard.PutOperationOnResource(i, DRYER, times.Value.ProcessStart)
                                End If
                            End If
                        End If
                    End If

                    ' -----------------------------------------------------
                    ' PHASE 2: Previous Operations Adjustment (Op 200)
                    ' -----------------------------------------------------
                ElseIf recOpNo = 200 Then
                    ' Restored your exact original syntax here
                    times2 = planningboard.GetOperationTimes(i).Value.OperationTimes.ProcessStart

                    If times2 >= currentDate Then
                        ' Calculate target time once outside the While loop
                        targetTime = times2.AddDays(-1)
                        oprec = planningboard.GetPreviousOperation(i, 1)

                        ' Chain backward through all preceding operations
                        While (oprec > 0)
                            resrecs = planningboard.FindResources(oprec)

                            If resrecs IsNot Nothing Then
                                resrec = resrecs.FirstOrDefault()

                                ' Ensure a valid resource was found before placing
                                If resrec > 0 Then
                                    planningboard.PutOperationOnResource(oprec, resrec, targetTime)
                                End If
                            End If

                            ' Move backward to the next previous operation
                            oprec = planningboard.GetPreviousOperation(oprec, 1)
                        End While
                    End If
                End If

            Catch ex As Exception
                ' The Try...Catch block safely handles scenarios where GetOperationTimes returns null
                Debug.WriteLine("Failed scheduling op " & i & ": " & ex.Message)
            End Try
        Next

        preactor.DestroyStatus()
        Return 0
    End Function
    'Public Function runFix(ByRef preactorComObject As PreactorObj, ByRef pespComObject As Object) As Integer
    '    ' Initialize Preactor and Planning Board objects
    '    Dim preactor As IPreactor = PreactorFactory.CreatePreactorObject(preactorComObject)
    '    Dim planningboard As IPlanningBoard = preactor.PlanningBoard

    '    ' Get table and field references
    '    Dim ordersTable As Integer = preactor.FindFirstClassificationString("LAUNCH TIME").Value.FormatNumber
    '    Dim opNoField As Integer = preactor.GetFieldNumber(ordersTable, "Op. No.")

    '    ' OPTIMIZATION: Commented out 'routingDt' as it is declared and populated but never used. 
    '    ' Bypassing 'readOrderTable(preactor)' saves I/O overhead and memory.
    '    ' Dim routingDt As DataTable = readOrderTable(preactor)

    '    Dim currentDate As DateTime = planningboard.TerminatorTime
    '    Dim DRYER As Integer = planningboard.GetResourceNumber("DRYER")
    '    Dim times As OperationTimes?

    '    Dim reccount As Integer = preactor.RecordCount(ordersTable)
    '    Dim recOpNo As Integer
    '    Dim nxtrecOpNo As Integer

    '    ' ==========================================
    '    ' PHASE 1: DRYER Fix (Op 260 -> Op 290)
    '    ' ==========================================
    '    For i As Integer = 1 To reccount
    '        recOpNo = preactor.ReadFieldInt(ordersTable, opNoField, i)

    '        ' Filter for operations that are explicitly Op. No. 260
    '        If recOpNo <> 260 Then Continue For

    '        ' Find the logical next operation and check if it is Op. No. 290
    '        Dim nextOpIndex As Integer = planningboard.GetNextOperation(i, 1)
    '        nxtrecOpNo = preactor.ReadFieldInt(ordersTable, opNoField, nextOpIndex)
    '        If nxtrecOpNo <> 290 Then Continue For

    '        Try
    '            ' OPTIMIZATION: Cache the start time of the next operation to avoid 
    '            ' querying the 'planningboard' COM object multiple times.
    '            ' (Note: Preserving original logic which explicitly targets index 'i + 1')
    '            Dim nextOpStartTime As DateTime = planningboard.GetOperationTimes(i + 1).Value.OperationTimes.ProcessStart

    '            ' Skip if the next operation starts before the current terminator time
    '            If nextOpStartTime < currentDate Then Continue For

    '            ' Back-test Op 260 on the DRYER resource from the start time of Op 290
    '            times = planningboard.BackTestOpOnResource(i, DRYER, nextOpStartTime)

    '            ' Place Op 260 on the DRYER at the newly calculated start time
    '            planningboard.PutOperationOnResource(i, DRYER, times.Value.ProcessStart)
    '        Catch ex As Exception
    '            Debug.WriteLine("Failed scheduling op " & i & ": " & ex.Message)
    '        End Try
    '    Next

    '    Dim resrecs As IEnumerable(Of Integer)
    '    Dim resrec As Integer
    '    Dim times2 As DateTime
    '    Dim oprec As Integer

    '    ' ==========================================
    '    ' PHASE 2: Previous Operations Adjustment
    '    ' ==========================================
    '    For i As Integer = 1 To reccount
    '        Try
    '            ' Filter for operations that are explicitly Op. No. 200
    '            If preactor.ReadFieldInt(ordersTable, opNoField, i) <> 200 Then Continue For

    '            ' Fetch and evaluate the current operation's start time
    '            times2 = planningboard.GetOperationTimes(i).Value.OperationTimes.ProcessStart
    '            If times2 < currentDate Then Continue For

    '            ' Get the immediately preceding operation
    '            oprec = planningboard.GetPreviousOperation(i, 1)

    '            ' OPTIMIZATION: Pre-calculate the target date (-1 day) outside of the While loop.
    '            ' This prevents 'AddDays' from being unnecessarily recalculated in every iteration.
    '            Dim targetTime As DateTime = times2.AddDays(-1)

    '            ' Chain backward through all preceding operations
    '            While (oprec > 0)
    '                ' Find eligible resources and pick the first one
    '                resrecs = planningboard.FindResources(oprec)
    '                resrec = resrecs.FirstOrDefault()

    '                ' Place the previous operation on the resource exactly 1 day prior to Op 200
    '                planningboard.PutOperationOnResource(oprec, resrec, targetTime)

    '                ' Move backward to the next previous operation
    '                oprec = planningboard.GetPreviousOperation(oprec, 1)
    '            End While
    '        Catch ex As Exception
    '            Debug.WriteLine("Failed scheduling op " & oprec & ": " & ex.Message)
    '        End Try
    '    Next

    '    preactor.DestroyStatus()
    '    Return 0
    'End Function


    Public Function afterFiring(ByRef preactorComObject As PreactorObj,
                            ByRef pespComObject As Object) As Integer

        Dim preactor As IPreactor = PreactorFactory.CreatePreactorObject(preactorComObject)
        BeginSchedulerDebug(preactor)
        Dim planningboard As IPlanningBoard = preactor.PlanningBoard

        Dim postFiring As New PostFiringScheduler()

        Dim routingDt As DataTable = readOrderTable(preactor)

        Dim queue As List(Of PostFiringScheduler.QueueItem) =
        postFiring.BuildQueue(preactor, planningboard, routingDt, "KILNACK", _schedulerDebug)

        If queue.Count > 0 Then
            postFiring.ScheduleQueue(preactor, planningboard, queue, _schedulerDebug)
        End If

        FinishSchedulerDebug(preactor)
        preactor.DestroyStatus()
        Return 0

    End Function

    Public Function showMeta(ByRef preactorComObject As PreactorObj, ByRef pespComObject As Object) As Integer
        Dim preactor As IPreactor = PreactorFactory.CreatePreactorObject(preactorComObject)
        Dim planningboard As IPlanningBoard = preactor.PlanningBoard

        Debug.WriteLine(preactor.GetFieldNumber(143, "Numeric Value"))
        Debug.WriteLine(preactor.ReadFieldDouble(143, 11, 1))

        Return 0
    End Function
    Public Function testFunction(ByRef preactorComObject As PreactorObj, ByRef pespComObject As Object) As Integer
        Dim preactor As IPreactor = PreactorFactory.CreatePreactorObject(preactorComObject)
        Dim planningboard As IPlanningBoard = preactor.PlanningBoard

        Return 0
    End Function

    Public Function untilPress(ByRef preactorComObject As PreactorObj, ByRef pespComObject As Object) As Integer
        Dim preactor As IPreactor = PreactorFactory.CreatePreactorObject(preactorComObject)
        Dim planningboard As IPlanningBoard = preactor.PlanningBoard
        Dim opRec As Integer
        Dim routingdt As DataTable
        Dim ResRec As Integer
        Dim ResRecs As IEnumerable(Of Integer)
        Dim opTimes As Nullable(Of Preactor.OperationTimes)
        Dim ordersTable As Integer = preactor.FindFirstClassificationString("LAUNCH TIME").Value.FormatNumber
        Dim opNoField As Integer = preactor.GetFieldNumber(ordersTable, "Op. No.")
        'Dim currentDate As New System.DateTime(2025, 8, 1, 0, 0, 0)
        Dim currentDate As DateTime = planningboard.TerminatorTime
        routingdt = readOrderTable(preactor)
        Dim pressingQueue As List(Of Integer) = BuildPressing200Queue(routingdt, currentDate)
        CreateRankedOperationQueue(preactor, planningboard, ordersTable, "JobsQueue", pressingQueue)

        ' Snapshot for debugging
        Dim jobsQueueSnapshot As List(Of Integer) = GetQueueSnapshot(planningboard, "JobsQueue")

        While (planningboard.GetOperationInQueue("JobsQueue", 1, opRec))
            ' Take the next operation out of the ranked queue so we can decide where to load it.
            planningboard.RemoveOperationFromQueue("JobsQueue", opRec)

            ' Inner loop: schedule this operation and then walk to subsequent operations
            ' (your "family" / routing chain) using GetNextOperation.
            While (opRec > 0)
                If preactor.ReadFieldInt(ordersTable, opNoField, opRec) >= 200 Then Exit While
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
                    ' planningboard.PutOperationOnResource(opRec, bestResRec, bestOpTimes.Value.ChangeStart)
                    planningboard.PutOperationOnResource(opRec, bestResRec, bestOpTimes.Value.ProcessStart.AddDays(-1))
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


    Public Function runPress(ByRef preactorComObject As PreactorObj, ByRef pespComObject As Object) As Integer
        Dim preactor As IPreactor = PreactorFactory.CreatePreactorObject(preactorComObject)
        BeginSchedulerDebug(preactor)
        Dim planningboard As IPlanningBoard = preactor.PlanningBoard
        Dim opRec As Integer
        Dim routingdt As DataTable
        Dim ResRec As Integer
        Dim ResRecs As IEnumerable(Of Integer)
        Dim opTimes As Nullable(Of Preactor.OperationTimes)
        Dim ordersTable As Integer = preactor.FindFirstClassificationString("LAUNCH TIME").Value.FormatNumber
        routingdt = readOrderTable(preactor)
        Dim currentDate As DateTime = planningboard.TerminatorTime
        Dim pressingObj As New pressingOptimizer_vf()
        Dim pressingQueue = pressingObj.BuildPressing200Queue(routingdt, currentDate, prioritizePrevOpFirst:=True, debug:=_schedulerDebug)
        CreateRankedOperationQueue(preactor, planningboard, ordersTable, "JobsQueue", pressingQueue)
        Dim opNoField As Integer = preactor.GetFieldNumber(ordersTable, "Op. No.")

        ' Snapshot for debugging
        Dim jobsQueueSnapshot As List(Of Integer) = GetQueueSnapshot(planningboard, "JobsQueue")

        While (planningboard.GetOperationInQueue("JobsQueue", 1, opRec))

            ' Take the next operation out of the ranked queue so we can decide where to load it.
            planningboard.RemoveOperationFromQueue("JobsQueue", opRec)

            ' Inner loop: schedule this operation and then walk to subsequent operations
            ' (your "family" / routing chain) using GetNextOperation.
            While (opRec > 0) ' this condition is wrong
                If preactor.ReadFieldInt(ordersTable, opNoField, opRec) > 200 Then Exit While
                If preactor.ReadFieldInt(ordersTable, opNoField, opRec) = 200 Then
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
                        opTimes = planningboard.QueryOperationOnResource(opRec, ResRec, planningboard.TerminatorTime)

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
                        If planningboard.IsOperationScheduled(opRec) Then Exit While
                        ' Load the operation onto the resource that gives the earliest feasible start.
                        PutOperationWithTrace(preactor, planningboard, "Pressing", opRec, bestResRec, bestOpTimes.Value.ChangeStart, "Forward")
                        Try
                            'planningboard.PutOperationOnResource(planningboard.GetPreviousOperation(opRec, 1), bestResRec, bestOpTimes.Value.ChangeStart.AddDays(-1))
                            'planningboard.PutOperationOnResource(planningboard.GetPreviousOperation(planningboard.GetPreviousOperation(opRec, 1), 1), bestResRec, bestOpTimes.Value.ChangeStart.AddDays(-1))
                        Catch ex As Exception

                        End Try

                    Else
                        ' No feasible resource was found.
                        ' Practical meaning:
                        '   - This operation cannot be scheduled on any alternate resource at/after the terminator boundary
                        '     under current constraints (calendars, setups, secondary constraints, etc.).
                        ' Leave it unscheduled (or handle with a custom queue / reason code if your design requires).
                    End If
                End If
                ' Move to the next operation in the routing chain.
                opRec = planningboard.GetNextOperation(opRec, 1) ' API-supported routing traversal:contentReference[oaicite:4]{index=4}

            End While ' next operation in chain

        End While ' next op in JobsQueue
        FinishSchedulerDebug(preactor)
        Return 0
    End Function

    'Public Function runPressToFiring(ByRef preactorComObject As PreactorObj, ByRef pespComObject As Object) As Integer
    '    Dim preactor As IPreactor = PreactorFactory.CreatePreactorObject(preactorComObject)
    '    Dim planningboard As IPlanningBoard = preactor.PlanningBoard

    '    Dim ordersTable As Integer = preactor.FindFirstClassificationString("LAUNCH TIME").Value.FormatNumber
    '    Dim opNofield As Integer = preactor.GetFieldNumber(ordersTable, "Op. No.")

    '    Dim currentDate As DateTime = planningboard.TerminatorTime
    '    Dim routingdt As DataTable = readOrderTable(preactor)

    '    Dim pressingQueue As List(Of Integer) = BuildPressing200Queue(routingdt, currentDate)
    '    CreateRankedOperationQueue(preactor, planningboard, ordersTable, "JobsQueue", pressingQueue)

    '    Dim reccount As Integer = preactor.RecordCount(ordersTable)
    '    Dim AIRDRY As Integer = planningboard.GetResourceNumber("AIRDRY")
    '    Dim DRYER As Integer = planningboard.GetResourceNumber("DRYER") ' Kept for reference, though logically merged below

    '    Dim terminatorBoundary As DateTime = planningboard.TerminatorTime

    '    ' OPTIMIZATION: Pre-declare all variables outside the loop to prevent repeated stack allocations
    '    Dim opNo As Integer
    '    Dim resRecs As IEnumerable(Of Integer)
    '    Dim bestResRec As Integer
    '    Dim opTimes As OperationTimes?
    '    Dim bestOpTimes As OperationTimes?
    '    Dim finalChangeStart As DateTime

    '    ' OPTIMIZATION: Use a 'For' loop instead of 'While' for slightly faster and cleaner index iteration
    '    For opRec As Integer = 1 To reccount
    '        ' Read the field once per record
    '        opNo = preactor.ReadFieldInt(ordersTable, opNofield, opRec)

    '        ' Short-circuit evaluation: Only proceed if opNo is exactly within the range
    '        If opNo > 200 AndAlso opNo < 290 Then
    '            resRecs = planningboard.FindResources(opRec)

    '            ' SAFETY: Ensure FindResources didn't return a null collection before iterating
    '            If resRecs IsNot Nothing Then
    '                bestResRec = 0
    '                bestOpTimes = Nothing

    '                For Each resRec As Integer In resRecs
    '                    opTimes = planningboard.TestOperationOnResource(opRec, resRec, terminatorBoundary)

    '                    ' SAFE CHECK: Only compare if TestOperationOnResource returned a valid object
    '                    If opTimes.HasValue Then 
    '                        ' Compare to find the earliest start time
    '                        If Not bestOpTimes.HasValue OrElse opTimes.Value.ChangeStart < bestOpTimes.Value.ChangeStart Then
    '                            bestResRec = resRec
    '                            bestOpTimes = opTimes
    '                        End If
    '                    End If
    '                Next

    '                ' FINAL ASSIGNMENT: Ensure we found a valid time and resource
    '                If bestOpTimes.HasValue AndAlso bestResRec > 0 Then
    '                    finalChangeStart = bestOpTimes.Value.ChangeStart

    '                    ' OPTIMIZATION: Cleaned up redundant ElseIf block
    '                    If bestResRec = AIRDRY Then
    '                        planningboard.PutOperationOnResource(opRec, bestResRec, finalChangeStart.Date.AddDays(1))
    '                    Else
    '                        ' Handles DRYER and any other resource exactly the same way (as in your original logic)
    '                        planningboard.PutOperationOnResource(opRec, bestResRec, finalChangeStart)
    '                    End If
    '                End If
    '            End If
    '        End If
    '    Next

    '    preactor.DestroyStatus()
    '    Return 0
    'End Function

    Public Function runPressToFiring(ByRef preactorComObject As PreactorObj,
                                 ByRef pespComObject As Object) As Integer

        Dim preactor As IPreactor = PreactorFactory.CreatePreactorObject(preactorComObject)
        BeginSchedulerDebug(preactor)
        Dim planningboard As IPlanningBoard = preactor.PlanningBoard

        Dim ordersTable As Integer =
        preactor.FindFirstClassificationString("LAUNCH TIME").Value.FormatNumber

        Dim opNoField As Integer =
        preactor.GetFieldNumber(ordersTable, "Op. No.")

        Dim airDryResourceRec As Integer =
        planningboard.GetResourceNumber("AIRDRY")

        Dim currentDate As DateTime = planningboard.TerminatorTime
        Dim routingDt As DataTable = readOrderTable(preactor)


        ' Build first eligible WIP operations.
        Dim queue As List(Of Integer) =
        BuildPressToFiringQueue(preactor,
                                planningboard,
                                routingDt,
                                currentDate,
                                ordersTable,
                                opNoField)

        Dim queueName As String =
        "JobsQueue_PressToFiring_" & DateTime.Now.ToString("yyyyMMddHHmmssfff",
                                                           CultureInfo.InvariantCulture)

        CreateRankedOperationQueue(preactor, planningboard, ordersTable, queueName, queue)

        Dim opRec As Integer = 0

        While planningboard.GetOperationInQueue(queueName, 1, opRec)

            planningboard.RemoveOperationFromQueue(queueName, opRec)

            ' Schedule this operation and all following unscheduled operations
            ' until firing/loading starts.
            SchedulePressToFiringChain(preactor,
                       planningboard,
                       ordersTable,
                       opNoField,
                       airDryResourceRec,
                       routingDt,
                       opRec)

        End While

        FinishSchedulerDebug(preactor)
        preactor.DestroyStatus()
        Return 0

    End Function
    Private Sub SchedulePressToFiringChain(preactor As IPreactor,
                                   planningboard As IPlanningBoard,
                                   ordersTable As Integer,
                                   opNoField As Integer,
                                   airDryResourceRec As Integer,
                                   routingDt As DataTable,
                                   startOpRec As Integer)

        Dim opRec As Integer = startOpRec

        While opRec > 0

            Dim opNo As Integer =
            preactor.ReadFieldInt(ordersTable, opNoField, opRec)

            ' Stop when firing/loading starts.
            If opNo >= 290 Then Exit While

            ' Skip pressing and earlier operations.
            If opNo > 200 AndAlso opNo < 290 Then

                ' Never reschedule an already scheduled operation.
                If Not IsScheduledLive(planningboard, opRec) Then

                    'Dim readyTime As DateTime? =
                    'GetReadyTimeFromScheduledPredecessors(planningboard, opRec)
                    Dim readyTime As DateTime?

                    If opRec = startOpRec Then

                        Dim wipReadyTime As DateTime =
        GetWipReadyTimeFromRouting(routingDt, opRec)

                        If wipReadyTime = DateTime.MinValue Then
                            readyTime = Nothing
                        Else
                            readyTime = wipReadyTime
                        End If

                    Else

                        readyTime = GetReadyTimeFromScheduledPredecessors(planningboard, opRec)

                    End If
                    If Not readyTime.HasValue Then
                        Exit While
                    End If

                    Dim placement As Placement? =
                    FindEarliestFeasiblePlacement(preactor,
                                                  planningboard,
                                                  ordersTable,
                                                  opRec,
                                                  readyTime.Value)

                    If placement.HasValue Then

                        If Not IsScheduledLive(planningboard, opRec) Then
                            Dim operationStart As DateTime = placement.Value.StartTime

                            ' Air drying starts on the next day. Drying and all
                            ' other operations retain their normal earliest start.
                            If placement.Value.ResourceRec = airDryResourceRec Then
                                operationStart = operationStart.Date.AddDays(1)
                            End If

                            PutOperationWithTrace(preactor, planningboard, "Drying", opRec, placement.Value.ResourceRec, operationStart, "Forward")
                        End If

                    Else
                        Exit While
                    End If

                End If

            End If

            ' Continue through the routing so DRYING follows AIRDRY and the
            ' remaining operations are scheduled in sequence.
            opRec = planningboard.GetNextOperation(opRec, 1)

        End While

    End Sub

    Private Class PressToFiringCandidate
        Public Property OpRec As Integer
        Public Property OrderNo As String
        Public Property OpNo As Integer
        Public Property IsWip As Boolean
        Public Property ReadyTime As DateTime
        Public Property DueTime As DateTime
        Public Property CycleRank As Integer

        Public Property WipScore As Integer
    End Class

    Private Structure Placement
        Public ResourceRec As Integer
        Public StartTime As DateTime
    End Structure

    Private Function IsScheduledLive(planningboard As IPlanningBoard,
                                 opRec As Integer) As Boolean

        If opRec <= 0 Then Return False

        Dim times As Nullable(Of Preactor.OperationResourceTimes) =
        planningboard.GetOperationTimes(opRec)

        Return times.HasValue

    End Function
    Private Function GetWipReadyTimeFromRouting(routingDt As DataTable,
                                            opRec As Integer) As DateTime

        If routingDt Is Nothing OrElse opRec <= 0 Then Return DateTime.MinValue

        For Each r As DataRow In routingDt.Rows

            If SharedHelpers.SafeInt(r("OrdersID")) = opRec Then
                Return SharedHelpers.SafeDate(r("wip_ready_time"))
            End If

        Next

        Return DateTime.MinValue

    End Function
    Private Function GetReadyTimeFromScheduledPredecessors(planningboard As IPlanningBoard,
                                                       opRec As Integer) As DateTime?

        Dim idx As Integer = 1
        Dim prevRec As Integer
        Dim foundAny As Boolean = False
        Dim maxEnd As DateTime = DateTime.MinValue

        Do
            prevRec = planningboard.GetPreviousOperation(opRec, idx)

            If prevRec <= 0 Then Exit Do

            Dim prevTimes As Nullable(Of Preactor.OperationResourceTimes) =
            planningboard.GetOperationTimes(prevRec)

            ' If predecessor exists but is not scheduled, this operation is not WIP-ready.
            If Not prevTimes.HasValue Then Return Nothing

            foundAny = True

            If prevTimes.Value.OperationTimes.ProcessEnd > maxEnd Then
                maxEnd = prevTimes.Value.OperationTimes.ProcessEnd
            End If

            idx += 1
        Loop

        If Not foundAny Then Return Nothing

        Return maxEnd

    End Function

    Private Function FindEarliestFeasiblePlacement(preactor As IPreactor,
                                               planningboard As IPlanningBoard,
                                               ordersTable As Integer,
                                               opRec As Integer,
                                               readyTime As DateTime) As Placement?

        Dim testFrom As DateTime = readyTime

        If testFrom < planningboard.TerminatorTime Then
            testFrom = planningboard.TerminatorTime
        End If

        Dim bestPlacement As Placement? = Nothing
        Dim bestTimes As Nullable(Of Preactor.OperationTimes) = Nothing

        For Each resRec As Integer In planningboard.FindResources(opRec)

            Dim testTimes As Nullable(Of Preactor.OperationTimes) =
            planningboard.TestOperationOnResource(opRec, resRec, testFrom)

            If Not testTimes.HasValue Then Continue For

            If Not bestTimes.HasValue OrElse
           testTimes.Value.ChangeStart < bestTimes.Value.ChangeStart Then

                bestTimes = testTimes

                bestPlacement = New Placement With {
                .ResourceRec = resRec,
                .StartTime = testTimes.Value.ChangeStart
            }

            End If

        Next

        Return bestPlacement

    End Function

    Private Function BuildPressToFiringQueue(preactor As IPreactor,
                                         planningboard As IPlanningBoard,
                                         routingDt As DataTable,
                                         currentDate As DateTime,
                                         ordersTable As Integer,
                                         opNoField As Integer) As List(Of Integer)

        If routingDt Is Nothing Then Throw New ArgumentNullException(NameOf(routingDt))

        SharedHelpers.RequireColumn(routingDt, "OrdersID")
        SharedHelpers.RequireColumn(routingDt, "Operation Number")
        SharedHelpers.RequireColumn(routingDt, "Order No")
        SharedHelpers.RequireColumn(routingDt, "is_scheduled")
        SharedHelpers.RequireColumn(routingDt, "prev_op_is_scheduled")
        SharedHelpers.RequireColumn(routingDt, "firing due date")
        SharedHelpers.RequireColumn(routingDt, "Cycle Type")
        SharedHelpers.RequireColumn(routingDt, "wip_status")
        SharedHelpers.RequireColumn(routingDt, "wip_ready_time")
        SharedHelpers.RequireColumn(routingDt, "wip_score")

        Dim candidates As New List(Of PressToFiringCandidate)()
        Dim seen As New HashSet(Of Integer)()

        For Each r As DataRow In routingDt.Rows

            Dim opRec As Integer = SharedHelpers.SafeInt(r("OrdersID"))

            If opRec <= 0 Then Continue For
            If seen.Contains(opRec) Then Continue For

            Dim opNo As Integer = SharedHelpers.SafeInt(r("Operation Number"))

            ' Press-to-firing window only.
            If opNo <= 200 OrElse opNo >= 290 Then Continue For

            ' Snapshot guard.
            If SharedHelpers.SafeBool(r("is_scheduled")) Then Continue For

            ' Live planning board guard.
            If IsScheduledLive(planningboard, opRec) Then Continue For

            'Dim readyTime As DateTime? =
            'GetReadyTimeFromScheduledPredecessors(planningboard, opRec)

            'Dim isWip As Boolean =
            'SharedHelpers.SafeBool(r("prev_op_is_scheduled")) OrElse readyTime.HasValue

            '' WIP-first requirement:
            '' for this run, only operations whose predecessor is scheduled are eligible.
            'If Not isWip Then Continue For
            'If Not readyTime.HasValue Then Continue For
            Dim wipStatus As String =
    SharedHelpers.SafeStr(r("wip_status")).Trim()

            If Not wipStatus.Equals("Candidate", StringComparison.OrdinalIgnoreCase) Then
                Continue For
            End If

            Dim readyTime As DateTime =
    SharedHelpers.SafeDate(r("wip_ready_time"))

            If readyTime = DateTime.MinValue Then Continue For

            Dim wipScore As Integer =
    SharedHelpers.SafeInt(r("wip_score"))

            Dim isWip As Boolean = wipScore > 0

            Dim dueTime As DateTime =
            SharedHelpers.ParseDueAsEndOfDay(r("firing due date"))

            If dueTime = DateTime.MinValue Then dueTime = DateTime.MaxValue

            Dim cycleRank As Integer =
            GetCycleRank(SharedHelpers.SafeStr(r("Cycle Type")))

            candidates.Add(New PressToFiringCandidate With {
            .OpRec = opRec,
            .OrderNo = SharedHelpers.SafeStr(r("Order No")).Trim(),
            .OpNo = opNo,
            .IsWip = isWip,
            .ReadyTime = readyTime,
            .WipScore = wipScore,
            .DueTime = dueTime,
            .CycleRank = cycleRank
        })

            seen.Add(opRec)

        Next

        Return candidates.
        OrderByDescending(Function(c) c.WipScore).
        ThenBy(Function(c) c.ReadyTime).
        ThenBy(Function(c) c.DueTime).
        ThenByDescending(Function(c) c.CycleRank).
        ThenBy(Function(c) c.OpNo).
        ThenBy(Function(c) c.OpRec).
        Select(Function(c) c.OpRec).
        ToList()

    End Function
    Private Function CreateRankedQueue(ByRef preactor As IPreactor, ByVal planningboard As IPlanningBoard,
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

    Private Sub ShowExplorerUI()
        ' If Run() is not on an STA thread, create one for WinForms
        Dim t As New Thread(Sub()
                                Application.EnableVisualStyles()
                                Application.SetCompatibleTextRenderingDefault(False)

                                Using frm As New DataViewExplorerForm()
                                    frm.StartPosition = FormStartPosition.CenterScreen
                                    frm.ShowDialog() ' blocks until closed
                                End Using
                            End Sub)

        t.SetApartmentState(ApartmentState.STA)
        t.IsBackground = True
        t.Start()
        t.Join() ' Wait until user closes form (remove Join() if you do NOT want to block)
    End Sub


    Private Function CreateRankedOperationQueue(ByRef preactor As IPreactor, ByVal planningboard As IPlanningBoard,
                                                ByVal ordersTable As Integer, ByVal QName As String, ByVal queue As List(Of Integer)) As Integer
        planningboard.CreateQueue(QName)
        For Each q In queue
            planningboard.AddOperationToQueue(QName, q, QueuePosition.End)
        Next
        Return 0
    End Function

    Public Function freezeWindow(ByRef preactorComObject As PreactorObj, ByRef pespComObject As Object, Optional days As Integer = 2) As Integer
        Dim preactor As IPreactor = PreactorFactory.CreatePreactorObject(preactorComObject)
        Dim planningboard As IPlanningBoard = preactor.PlanningBoard
        Dim currentDate As DateTime = planningboard.TerminatorTime

        Dim resrecs As Integer = planningboard.ResourceCount
        For i As Integer = 1 To resrecs
            planningboard.LockResource(i, currentDate, currentDate.AddDays(days), OperationReferencePoint.AnyPart, True)
        Next

        preactor.DestroyStatus()
        Return 0
    End Function


    ' ----------------------
    ' CSV import helper
    ' ----------------------
    Public Function ImportRoutingCsvToDataTable(csvPath As String) As DataTable
        If String.IsNullOrWhiteSpace(csvPath) Then Throw New ArgumentException("csvPath is empty.")
        If Not IO.File.Exists(csvPath) Then Throw New IO.FileNotFoundException("CSV not found.", csvPath)

        Dim dt As New DataTable("Routing")
        Using parser As New TextFieldParser(csvPath)
            parser.TextFieldType = FieldType.Delimited
            parser.SetDelimiters(",")
            parser.HasFieldsEnclosedInQuotes = True
            parser.TrimWhiteSpace = True

            If parser.EndOfData Then Return dt

            Dim headers = parser.ReadFields()
            If headers Is Nothing OrElse headers.Length = 0 Then Return dt

            For Each h In headers
                Dim colName = h.Trim()
                Dim colType As Type = GetType(String)
                Select Case colName
                    Case "OrdersID", "Order No", "Part Number", "Sales Order", "Quantity", "Operation Number"
                        colType = GetType(Integer)
                    Case "Setup Time", "Time Per Item", "Batch Time", "Tonnage", "Volume Occupancy", "Firing buffer", "MTS/MTO", "MTS/MTO priority", "Que Time", "Pressing buffer", "Wheel thickness", "Week start"
                        colType = GetType(Double)
                    Case "Due Date", "Pressing earliest start", "pressing due date"
                        colType = GetType(DateTime)
                    Case Else
                        colType = GetType(String)
                End Select
                dt.Columns.Add(New DataColumn(colName, colType))
            Next

            While Not parser.EndOfData
                Dim fields = parser.ReadFields()
                If fields Is Nothing Then Continue While
                Dim row = dt.NewRow()
                For i As Integer = 0 To dt.Columns.Count - 1
                    Dim raw As String = If(i < fields.Length, fields(i), "")
                    raw = If(raw, "").Trim()
                    Dim col = dt.Columns(i)
                    If String.IsNullOrEmpty(raw) Then
                        row(i) = GetDefaultValue(col.DataType)
                        Continue For
                    End If
                    Try
                        If col.DataType Is GetType(Integer) Then
                            row(i) = Integer.Parse(raw, CultureInfo.InvariantCulture)
                        ElseIf col.DataType Is GetType(Double) Then
                            row(i) = Double.Parse(raw, CultureInfo.InvariantCulture)
                        ElseIf col.DataType Is GetType(DateTime) Then
                            row(i) = SharedHelpers.ParseDateDdMmYyyy(raw)
                        Else
                            row(i) = raw
                        End If
                    Catch
                        row(i) = GetDefaultValue(col.DataType)
                    End Try
                Next
                dt.Rows.Add(row)
            End While
        End Using
        Return dt
    End Function

    Private Function GetDefaultValue(t As Type) As Object
        If t Is GetType(String) Then Return ""
        If t Is GetType(Integer) Then Return 0
        If t Is GetType(Double) Then Return 0.0R
        If t Is GetType(DateTime) Then Return DateTime.MinValue
        Return Nothing
    End Function

    ' ----------------------
    ' Append operation times from planning board into CSV datatable
    ' ----------------------
    Public Function AppendOperationTimesFromBoard(dt As DataTable,
                                                 preactor As IPreactor,
                                                 planningboard As IPlanningBoard,
                                                 opNumber As Integer) As DataTable
        If dt Is Nothing Then Throw New ArgumentNullException(NameOf(dt))
        If preactor Is Nothing Then Throw New ArgumentNullException(NameOf(preactor))
        If planningboard Is Nothing Then Throw New ArgumentNullException(NameOf(planningboard))

        SharedHelpers.RequireColumn(dt, "OrdersID")
        SharedHelpers.RequireColumn(dt, "Operation Number")

        Dim startColName As String = "scheduled_start_time"
        Dim endColName As String = "scheduled_end_time"
        Dim schColName As String = "is_scheduled"

        If Not dt.Columns.Contains(startColName) Then dt.Columns.Add(startColName, GetType(DateTime))
        If Not dt.Columns.Contains(endColName) Then dt.Columns.Add(endColName, GetType(DateTime))
        If Not dt.Columns.Contains(schColName) Then dt.Columns.Add(schColName, GetType(Boolean))

        Dim ordersTable As Integer = getformatfieldpair(preactor, format:="ORDERS").FormatNumber
        Dim opNoFields = preactor.FindClassificationString("OP NO")
        Dim ordersOpNoField As Preactor.FormatFieldPair = Nothing
        For Each f In opNoFields
            If f.FormatNumber = ordersTable Then ordersOpNoField = f
        Next
        'If ordersOpNoField Is Nothing Then Return dt

        Dim opCount As Integer = preactor.RecordCount(ordersTable)

        Dim idColumn As DataColumn = dt.Columns("OrdersID")
        dt.PrimaryKey = New DataColumn() {idColumn}

        For opRec As Integer = 1 To opCount
            Dim oNo As Integer
            Try
                oNo = preactor.ReadFieldInt(ordersOpNoField.FormatNumber, ordersOpNoField.FieldNumber, opRec)
            Catch
                Continue For
            End Try
            If oNo <> opNumber Then Continue For

            ' find corresponding CSV row(s) by OrdersID (order record assumed stored in OrdersID column)
            Dim rowtoupdate As DataRow = Nothing
            Try
                rowtoupdate = dt.Rows.Find(opRec)
            Catch
            End Try

            If rowtoupdate IsNot Nothing Then
                Dim ot = planningboard.GetOperationTimes(opRec)
                If ot.HasValue Then
                    rowtoupdate(startColName) = ot.Value.OperationTimes.ProcessStart
                    rowtoupdate(endColName) = ot.Value.OperationTimes.ProcessEnd
                    rowtoupdate(schColName) = True
                End If
            End If
        Next

        Return dt
    End Function


    ' New pressing optimizer
    ' ----------------------
    ' Pressing 200 queue builder (STRICT UNSCHEDULED) returning PARENT RECORD
    ' Minimal changes from your original:
    '   1) Adds is_scheduled + parent_record columns
    '   2) Filters only op=PRESS_OP_NUMBER AND is_scheduled=False
    '   3) Uses ParentRecord as the queue item instead of OrdersID
    '   4) Cycle ranking remains via GetCycleRank(), but you can later replace it with a map/dictionary
    ' ----------------------

    ' Candidate structure used for sorting/batching
    Private Class Candidate
        Public Property ParentRecord As Integer          ' <-- NEW: queue key to return
        Public Property Earliest As DateTime
        Public Property Due As DateTime
        Public Property Tier As Integer
        Public Property WheelDia As String
        Public Property WheelPin As String
        Public Property TypeKey As String                ' WheelDia|WheelThickness
        Public Property CycleRank As Integer
        Public Property MissingEarliest As Boolean
        Public Property MissingDue As Boolean
    End Class

    ' ----------------------
    ' Pressing 200 queue builder
    ' ----------------------
    Public Function BuildPressing200Queue(dt As DataTable,
                                      currentDate As DateTime,
                                      Optional approachingDays As Integer = 2) As List(Of Integer)

        If dt Is Nothing Then Throw New ArgumentNullException(NameOf(dt))

        ' ---- Required columns (these must exist in dt) ----
        ' NOTE: Keep these names EXACTLY matching your DataTable column names.
        SharedHelpers.RequireColumn(dt, "parent_record")               ' <-- NEW
        SharedHelpers.RequireColumn(dt, "is_scheduled")                ' <-- NEW

        SharedHelpers.RequireColumn(dt, "Operation Number")
        SharedHelpers.RequireColumn(dt, "Pressing earliest start")
        SharedHelpers.RequireColumn(dt, "Pressing Due date")           ' (fixes earlier mismatch)
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
            ' is_scheduled is expected to be reliably boolean-like based on your confirmation.
            Dim isScheduled As Boolean = SharedHelpers.SafeBool(r("is_scheduled"))
            If isScheduled Then Continue For

            ' 3) Queue key: parent_record must be valid
            Dim parentRec As Integer = SharedHelpers.SafeInt(r("parent_record"))
            If parentRec <= 0 Then Continue For

            ' 4) Read pressing dates (date-only logic)
            Dim earliest As DateTime = SharedHelpers.SafeDate(r("Pressing earliest start")).Date
            Dim due As DateTime = SharedHelpers.SafeDate(r("Pressing Due date")).Date

            ' Treat missing/parse-failed dates as MinValue (consistent with your existing SafeDate behavior)
            Dim missingEarliest As Boolean = (earliest = DateTime.MinValue)
            Dim missingDue As Boolean = (due = DateTime.MinValue)

            ' 5) Tiering logic:
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

            ' 6) Read batching/type attributes
            Dim wheelDia As String = SharedHelpers.SafeStr(r("Wheel Dia")).Trim()
            Dim wheelPin As String = SharedHelpers.SafeStr(r("Wheel thickness")).Trim()
            Dim cycleType As String = SharedHelpers.SafeStr(r("Cycle Type")).Trim()

            ' CycleRank currently uses GetCycleRank() (hardcoded list).
            ' Later you can replace this call with a dictionary/map lookup.
            Dim cycleRank As Integer = GetCycleRank(cycleType)


            candidates.Add(New Candidate With {
            .ParentRecord = parentRec,
            .Earliest = earliest,
            .Due = due,
            .Tier = tier,
            .WheelDia = wheelDia,
            .WheelPin = wheelPin,
            .TypeKey = wheelDia & "|" & wheelPin,
            .CycleRank = cycleRank,
            .MissingEarliest = missingEarliest,
            .MissingDue = missingDue
        })
        Next

        ' ---- Sorting: the exact queue priority order ----
        ' 1) Tier asc (0,1,2)
        ' 2) Due asc (missing due goes last)
        ' 3) Earliest asc (missing earliest goes last)
        ' 4) CycleRank desc (higher rank first)
        ' 5) TypeKey asc (clusters same Dia|Thickness)
        ' 6) ParentRecord asc (stable tie-breaker)
        Dim sorted As List(Of Candidate) =
        candidates.OrderBy(Function(c) c.Tier) _
                  .ThenBy(Function(c) If(c.MissingDue, DateTime.MaxValue, c.Due)) _
                  .ThenBy(Function(c) If(c.MissingEarliest, DateTime.MaxValue, c.Earliest)) _
                  .ThenByDescending(Function(c) c.CycleRank) _
                  .ThenBy(Function(c) c.TypeKey) _
                  .ThenBy(Function(c) c.ParentRecord) _
                  .ToList()

        ' ---- Soft batching: greedy pull-forward of same Tier + TypeKey + Due ----
        Dim batched As List(Of Candidate) = GreedyTypeBatchingWithinTier(sorted, lookahead:=50)

        ' ---- Output: return parent_record list (distinct to be safe) ----
        Return batched.Select(Function(c) c.ParentRecord).Distinct().ToList()
    End Function

    ' ----------------------
    ' Greedy batching (unchanged logic, but now works on ParentRecord candidates)
    ' Pull-forward rules:
    '   - same Tier
    '   - same TypeKey (WheelDia|WheelThickness)
    '   - same Due date
    ' This is a "soft clustering" step; it does NOT change tier ordering.
    ' ----------------------
    Private Function GreedyTypeBatchingWithinTier(sorted As List(Of Candidate),
                                             Optional lookahead As Integer = 50) As List(Of Candidate)

        If sorted Is Nothing OrElse sorted.Count <= 2 Then Return If(sorted, New List(Of Candidate)())

        Dim work As New List(Of Candidate)(sorted)
        Dim result As New List(Of Candidate)(work.Count)

        Dim i As Integer = 0
        While i < work.Count
            Dim cur As Candidate = work(i)
            result.Add(cur)

            ' Only attempt pull-forward if TypeKey is meaningful
            If Not String.IsNullOrEmpty(cur.TypeKey) Then
                Dim pulled As Integer = 0
                Dim j As Integer = i + 1

                ' Scan forward but stop after lookahead pulls
                While j < work.Count AndAlso pulled < lookahead

                    ' Match only within same tier, same type, same due date
                    If work(j).Tier = cur.Tier AndAlso
                   work(j).TypeKey = cur.TypeKey AndAlso
                   work(j).Due = cur.Due Then

                        result.Add(work(j))
                        work.RemoveAt(j)     ' remove from work list so it doesn’t appear again later
                        pulled += 1
                        Continue While       ' keep scanning from same j index (since list shifted)
                    End If

                    j += 1
                End While
            End If

            i += 1
        End While

        Return result
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

    Private Sub BeginSchedulerDebug(preactor As IPreactor)
        Try
            _schedulerDebug = New SchedulerDebugCollector()
            Dim snapshot As List(Of OperationSnapshot) =
                SchedulerStageDiagnostics.BuildOrderOperationSnapshot(preactor, _schedulerDebug)
            SchedulerStageDiagnostics.DiagnoseWip(snapshot, _schedulerDebug)
            SchedulerStageDiagnostics.DiagnoseAll(snapshot, _schedulerDebug)
        Catch ex As Exception
            Debug.WriteLine("Scheduler diagnostics initialization failed: " & ex.Message)
            _schedulerDebug = Nothing
        End Try
    End Sub

    Private Sub FinishSchedulerDebug(preactor As IPreactor)
        If _schedulerDebug Is Nothing OrElse Not _schedulerDebug.Enabled Then Return
        Try
            Dim snapshot As List(Of OperationSnapshot) =
                SchedulerStageDiagnostics.BuildOrderOperationSnapshot(preactor, _schedulerDebug)
            SchedulerStageDiagnostics.DiagnoseWip(snapshot, _schedulerDebug)
            SchedulerStageDiagnostics.DiagnoseAll(snapshot, _schedulerDebug)
            _schedulerDebug.ExportAll(preactor)
        Catch ex As Exception
            Debug.WriteLine("Scheduler diagnostics export failed: " & ex.Message)
        End Try
    End Sub

    Private Sub PutOperationWithTrace(preactor As IPreactor,
                                      planningboard As IPlanningBoard,
                                      stage As String,
                                      opRec As Integer,
                                      resourceRec As Integer,
                                      startTime As DateTime,
                                      direction As String)
        Dim trace As ScheduleAttemptTraceRow = Nothing
        If _schedulerDebug IsNot Nothing AndAlso _schedulerDebug.Enabled Then
            Try
                Dim snapshot As OperationSnapshot =
                    _schedulerDebug.OperationSnapshots.FirstOrDefault(Function(x) x.RecordNo = opRec)
                trace = New ScheduleAttemptTraceRow With {
                    .Stage = stage,
                    .OrderNo = If(snapshot Is Nothing, "", snapshot.OrderNo),
                    .ParentRecordNo = If(snapshot Is Nothing, 0, snapshot.ParentRecordNo),
                    .RecordNo = opRec,
                    .OperationNumber = If(snapshot Is Nothing, 0, snapshot.OperationNumber),
                    .RequestedResource = resourceRec.ToString(CultureInfo.InvariantCulture),
                    .RequestedStartTime = startTime,
                    .SchedulingDirection = direction,
                    .WasAttempted = True
                }
                _schedulerDebug.TraceScheduleAttempt(trace)
            Catch ex As Exception
                Debug.WriteLine("Schedule-attempt trace initialization failed: " & ex.Message)
                trace = Nothing
            End Try
        End If

        Try
            planningboard.PutOperationOnResource(opRec, resourceRec, startTime)
        Catch ex As Exception
            If trace IsNot Nothing Then
                trace.ExceptionType = ex.GetType().FullName
                trace.ExceptionMessage = ex.Message
                trace.FailureReasonCode = SchedulerDebugReasonCodes.SCHEDULE_EXCEPTION_THROWN
                trace.FailureReasonDetail = ex.Message
            End If
            Throw
        End Try

        If trace Is Nothing Then Return
        Try
            trace.ScheduledAfterAttempt = planningboard.IsOperationScheduled(opRec)
            trace.PlanningBoardResultCode = If(trace.ScheduledAfterAttempt, 0, -1)
            trace.PlanningBoardResultMeaning = If(trace.ScheduledAfterAttempt, "Scheduled", "Operation remained unscheduled")
            trace.FailureReasonCode = If(trace.ScheduledAfterAttempt,
                                         SchedulerDebugReasonCodes.OK_SCHEDULED,
                                         SchedulerDebugReasonCodes.SCHEDULE_RESULT_NOT_SCHEDULED)
            Dim times As Nullable(Of OperationResourceTimes) = planningboard.GetOperationTimes(opRec)
            If times.HasValue Then
                trace.ActualStartTime = times.Value.OperationTimes.ProcessStart
                trace.ActualEndTime = times.Value.OperationTimes.ProcessEnd
            End If
            trace.ActualResource = resourceRec.ToString(CultureInfo.InvariantCulture)
        Catch ex As Exception
            Debug.WriteLine("Schedule-attempt trace completion failed: " & ex.Message)
        End Try
    End Sub


End Class
