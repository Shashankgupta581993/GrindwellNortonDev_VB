Option Strict On
Option Explicit On

Imports System

Public Class DebugFieldMapRow
    Public Property RunId As String
    Public Property LogicalFieldName As String
    Public Property RequestedFieldName As String
    Public Property ResolvedFieldName As String
    Public Property FieldNumber As Integer
    Public Property Exists As Boolean
    Public Property UsedFallback As Boolean
    Public Property Detail As String
End Class

Public Class OperationSnapshot
    Public Property RunId As String
    Public Property ExportedAt As DateTime
    Public Property RecordNo As Integer
    Public Property ParentRecordNo As Integer
    Public Property OrderNo As String
    Public Property OperationNumber As Integer
    Public Property OperationName As String
    Public Property ResourceGroup As String
    Public Property RequiredResource As String
    Public Property KilnType As String
    Public Property CycleType As String
    Public Property VolumeOccupancy As Double
    Public Property Quantity As Double
    Public Property DueDate As DateTime?
    Public Property EarliestStart As DateTime?
    Public Property ScheduledStartTime As DateTime?
    Public Property ScheduledEndTime As DateTime?
    Public Property IsScheduled As Boolean
    Public Property PrevOperationRecordNo As Integer
    Public Property PrevOperationNumber As Integer
    Public Property PrevOperationIsScheduled As Boolean
    Public Property PrevOperationEndTime As DateTime?
    Public Property NextOperationRecordNo As Integer
    Public Property NextOperationNumber As Integer
    Public Property IsDisabled As Boolean
    Public Property IsComplete As Boolean
    Public Property Show As Boolean
    Public Property TableAttribute1 As String
    Public Property TableAttribute2 As String
    Public Property TableAttribute3 As String
    Public Property WheelDia As String
    Public Property WheelThickness As String
End Class

Public Class WipDiagnosticRow
    Public Property RunId As String
    Public Property OrderNo As String
    Public Property ParentRecordNo As Integer
    Public Property RecordNo As Integer
    Public Property OperationNumber As Integer
    Public Property OperationName As String
    Public Property Status As String
    Public Property ReasonCode As String
    Public Property ReasonDetail As String
    Public Property PreviousOperationRecordNo As Integer
    Public Property PreviousOperationNumber As Integer
    Public Property PreviousOperationScheduled As Boolean
    Public Property PreviousOperationEndTime As DateTime?
    Public Property NextOperationRecordNo As Integer
    Public Property NextOperationNumber As Integer
End Class

Public Class StageEligibilityRow
    Public Property RunId As String
    Public Property Stage As String
    Public Property OrderNo As String
    Public Property ParentRecordNo As Integer
    Public Property RecordNo As Integer
    Public Property OperationNumber As Integer
    Public Property IsCandidate As Boolean
    Public Property CandidateRank As Integer
    Public Property IncludedInOptimizer As Boolean
    Public Property ExcludedReasonCode As String
    Public Property ExcludedReasonDetail As String
    Public Property RequiredResource As String
    Public Property ResourceGroup As String
    Public Property KilnType As String
    Public Property CycleType As String
    Public Property VolumeOccupancy As Double
    Public Property DueDate As DateTime?
    Public Property EarliestAllowedStart As DateTime?
    Public Property PreviousOperationEndTime As DateTime?
    Public Property WipStatus As String
End Class

Public Class OptimizerCandidateTraceRow
    Public Property RunId As String
    Public Property OptimizerName As String
    Public Property Stage As String
    Public Property StepName As String
    Public Property OrderNo As String
    Public Property ParentRecordNo As Integer
    Public Property RecordNo As Integer
    Public Property OperationNumber As Integer
    Public Property BeforeCount As Integer
    Public Property AfterCount As Integer
    Public Property Included As Boolean
    Public Property ReasonCode As String
    Public Property ReasonDetail As String
    Public Property RankScore As Double
    Public Property RankBreakdown As String
End Class

Public Class BatchTunnelSwkPlanTraceRow
    Public Property RunId As String
    Public Property OptimizerName As String
    Public Property Stage As String
    Public Property StepName As String
    Public Property OrderNo As String
    Public Property ParentRecordNo As Integer
    Public Property RecordNo As Integer
    Public Property OperationNumber As Integer
    Public Property PlanGroup As String
    Public Property PlannedResource As String
    Public Property PlannedStart As DateTime?
    Public Property PlannedEnd As DateTime?
    Public Property Included As Boolean
    Public Property ReasonCode As String
    Public Property ReasonDetail As String
End Class

Public Class ScheduleAttemptTraceRow
    Public Property RunId As String
    Public Property AttemptNo As Integer
    Public Property Stage As String
    Public Property OrderNo As String
    Public Property ParentRecordNo As Integer
    Public Property RecordNo As Integer
    Public Property OperationNumber As Integer
    Public Property RequestedResource As String
    Public Property RequestedStartTime As DateTime?
    Public Property RequestedEndTime As DateTime?
    Public Property SchedulingDirection As String
    Public Property WasAttempted As Boolean
    Public Property PlanningBoardResultCode As Integer
    Public Property PlanningBoardResultMeaning As String
    Public Property ExceptionType As String
    Public Property ExceptionMessage As String
    Public Property ScheduledAfterAttempt As Boolean
    Public Property ActualStartTime As DateTime?
    Public Property ActualEndTime As DateTime?
    Public Property ActualResource As String
    Public Property FailureReasonCode As String
    Public Property FailureReasonDetail As String
End Class

Public Class ResourceValidationRow
    Public Property RunId As String
    Public Property Stage As String
    Public Property RecordNo As Integer
    Public Property OperationNumber As Integer
    Public Property ResourceName As String
    Public Property ResourceGroup As String
    Public Property ResourceRecordNo As Integer
    Public Property Exists As Boolean
    Public Property IsValidForOperation As Boolean
    Public Property ReasonCode As String
    Public Property ReasonDetail As String
End Class

Public Class ReasonCodeSummaryRow
    Public Property ReasonCode As String
    Public Property Count As Integer
    Public Property Sources As String
End Class

Public Class DebugConfigSnapshotRow
    Public Property RunId As String
    Public Property Name As String
    Public Property Value As String
End Class

Public Class SchedulerActionMetricsRow
    Public Property RunId As String
    Public Property ActionName As String
    Public Property StartedAt As DateTime
    Public Property FinishedAt As DateTime
    Public Property ElapsedMilliseconds As Long
    Public Property DebugInitializationMilliseconds As Long
    Public Property RoutingSnapshotMilliseconds As Long
    Public Property QueueBuildMilliseconds As Long
    Public Property SchedulingMilliseconds As Long
    Public Property RecordsScanned As Integer
    Public Property RoutingRowsCreated As Integer
    Public Property CandidateCount As Integer
    Public Property BoundaryCount As Integer
    Public Property QueueCount As Integer
    Public Property FormatLookupCalls As Long
    Public Property FieldLookupCalls As Long
    Public Property ResourceLookupCalls As Long
    Public Property ReadOperationNumberCalls As Long
    Public Property IsOperationScheduledCalls As Long
    Public Property GetOperationTimesCalls As Long
    Public Property GetPreviousOperationCalls As Long
    Public Property GetNextOperationCalls As Long
    Public Property FindResourcesCalls As Long
    Public Property FeasibilityTestCalls As Long
    Public Property PlacementAttempts As Long
    Public Property PlacementSuccesses As Long
    Public Property AlreadyScheduledSkips As Long
    Public Property CompletedSkips As Long
    Public Property HandledExceptions As Long
    Public Property FormatLookupCacheHits As Long
    Public Property FieldLookupCacheHits As Long
    Public Property ResourceLookupCacheHits As Long
    Public Property OperationNumberCacheHits As Long
    Public Property ScheduledStateCacheHits As Long
    Public Property OperationTimesCacheHits As Long
    Public Property PreviousOperationCacheHits As Long
    Public Property NextOperationCacheHits As Long
    Public Property FirstResourceCacheHits As Long
End Class
