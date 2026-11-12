# Code-based debugger settings

Edit the `BEGIN EDITABLE DEBUGGER SETTINGS` section in
`Debugging/SchedulerDebugCollector.vb`. This is the single configuration source
for this assembly's diagnostics. Values are literal VB.NET `True` or `False`.

```vb
Friend ReadOnly MasterEnabled As Boolean = False
Friend ReadOnly ActionMetricsEnabled As Boolean = True
```

The effective state of each channel is `MasterEnabled AndAlso ChannelEnabled`.
The shipped defaults are master **False**, every individual setting **True**.
Turning only the master to True therefore enables every channel. To enable a
subset, set the other individual values to False before rebuilding.

These fields are read-only after the assembly loads. Rebuild and reload the DLL
in the controlled Opcenter host after editing them. Changing source while an old
DLL is already loaded does not change that session. No COM registration or
deployment is performed by the local verification scripts.

## Individual channels

| Boolean field | Controlled diagnostics |
|---|---|
| `ActionMetricsEnabled` | Action counters/timings and `12_ActionMetrics.csv` |
| `FieldMapEnabled` | Field-resolution rows and `02_FieldMap.csv` |
| `OperationSnapshotsEnabled` | `03_OrderOperationSnapshot.csv` and `04_UnscheduledOperations.csv` |
| `WipDiagnosticsEnabled` | WIP analysis and `05_WipChainDiagnostics.csv` |
| `StageEligibilityEnabled` | Eligibility analysis and `06_StageEligibilityDiagnostics.csv` |
| `OptimizerCandidateTraceEnabled` | Candidate decisions, pressing campaign reporter, and `07_OptimizerCandidateTrace.csv` |
| `BatchPlanTraceEnabled` | Existing batch/tunnel/SWK trace collection API and `08_BatchTunnelSwkPlanTrace.csv` |
| `ScheduleAttemptTraceEnabled` | Attempt traces, diagnostic result reads, and `09_ScheduleAttemptTrace.csv` |
| `ResourceValidationEnabled` | Existing resource-validation collection API and `10_ResourceValidation.csv` |
| `ReasonSummaryEnabled` | `11_ReasonCodeSummary.csv`, using enabled producers only |
| `ProblemStatementExportEnabled` | Diagnostic Markdown/JSON problem statements |
| `BatchDebugCsvEnabled` | Standalone firing candidate/plan debug exports |
| `SwkDebugCsvEnabled` | Standalone SWK summary/queue debug exports |
| `AlgoSeq3DecisionLogEnabled` | Legacy `DecisionLog_*.csv` |
| `AlgoSeq3ScheduleLogEnabled` | Legacy diagnostic `PlannedSchedule_*.csv` |
| `DebugOutputEnabled` | Existing diagnostic `Debug.WriteLine`/`Trace` calls and debug-only verification reads |
| `DiagnosticPopupsEnabled` | Informational debugging popups |

Channels preserve their existing producers; enabling a collection API does not
invent new trace producers where the original scheduler never recorded them.

## Profiles

- **All off:** `MasterEnabled=False`. Individual values may remain True.
- **All on:** master and every individual field True.
- **Metrics only:** master and `ActionMetricsEnabled` True; every other channel False.
- **Candidate/attempt investigation:** master, `OptimizerCandidateTraceEnabled`,
  `ScheduleAttemptTraceEnabled`, and `ReasonSummaryEnabled` True; other channels
  False unless their output is also wanted.

Candidate and attempt tracing, WIP analysis, and stage eligibility can need an
internal operation snapshot even when `OperationSnapshotsEnabled=False`. Such
prerequisites do not write the disabled snapshot CSVs. Field-map-only mode resolves
fields without reading every operation. Metrics-only mode avoids the detailed
snapshot. Reason-summary-only mode contains no reasons unless a producer is enabled.

Aggregate runs write `00_RunSummary.md`; detailed runs also write
`01_ConfigSnapshot.csv`. These are package metadata, not additional channels.
Problem-statement manifests include artifacts generated during that export,
not unrelated files left in a reused directory.

## Compatibility and boundaries

The public collector properties and configuration-query methods remain present.
They cannot enable a channel disabled by these code settings. Detailed collection
can still be suppressed through `collector.Enabled=False`; metrics have their
own independent compatibility property.

The previous diagnostic app-setting keys and environment-variable overrides are
no longer read. Unrelated application settings remain in place. Existing
`Debug` calls retain their original Debug/Release conditional compilation; the
code switch cannot restore calls omitted by the Release compiler.

Required prompts, operational errors, business CSV exports, the calendar-state
viewer, database operations, and Opcenter's own unrelated diagnostics are retained.
These settings control this project's debugging, not Visual Studio or Opcenter
globally. Scheduling feasibility checks, campaign rules, and MES stamping continue
to execute with all debug channels off.

## Verification

Run `Tests/DebuggerSettingsMatrix.ps1` with Windows PowerShell. It builds 21
isolated configurations: master-off with all child flags on, all-on, all channels
off with master on, attempts plus reasons, and each of the 17 individual channels.
Production source settings are never rewritten by the matrix.

`Tests/SchedulerDiagnosticsHarness.ps1 -AssemblyPath <dll>` checks an all-off
build by default. For a custom enabled build, pass its comma-separated effective
channel names through `-EnabledChannels`. Those arguments describe expectations;
they do not change the compiled settings.
