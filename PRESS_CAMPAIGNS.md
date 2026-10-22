# Pressing campaigns: implementation and validation

`AlgoSeq4.runPress` now passes eligible operation-200 records to a per-press
campaign scheduler. Its public COM signature is unchanged. The earlier mixing,
drying, firing and finishing actions are not replaced.

## Business rules

- A campaign uses the complete die-ring identity. Only surrounding whitespace
  and letter case are normalized; suffixes and annotations are preserved.
- The existing WIP/date/cycle ranking selects campaign seeds. On a press with an
  active campaign, eligible orders using that ring are preferred, with the same
  pin first. Exact equality of due dates is no longer required for grouping.
- Readiness, pressing earliest start, and the schedule boundary are hard minimum
  start times. Unreleased predecessors are not admitted. Completed, actualized,
  already scheduled and WIP-rejected operations remain excluded.
- A campaign can cross midnight and idle periods. When another ring starts on
  that press, the removed ring cannot return until midnight on switch date + 3.
  Switching at any time on 7 October permits return on 10 October. A pin-only
  change does not start this restriction.
- Higher-ranked ready work interrupts a continuation if it is already late or
  would become late because of the proposed placement. Due dates are inclusive
  calendar days. This check also considers the ring cooldown that a trial
  placement would cause. Urgent work cannot bypass cooldown.
- Choose work at the next feasible start opportunity rather than deliberately
  idling for an unreleased ring. At equal opportunities, prefer continuation,
  then unchanged pin, then original rank, then resource number.
- Explicit press assignments are intersected with Opcenter's valid resources.
  “Select from Group” permits available eligible alternatives. Ring inventory is
  unlimited in this phase, but Opcenter's own feasibility constraints still apply.

## Existing and failed placements

The engine reconstructs tooling transitions from actual scheduled pressing
allocations before and after the schedule boundary. It does not move existing
orders. Proposed gap insertions are checked against the complete affected press
timeline so they cannot introduce or worsen a future cooldown violation.

New placements are chronological per press. Existing manual violations are
reported rather than repaired. Missing ring or earliest-start data holds a new
order for review. Existing allocations with unknown ring identity are reported;
their unknown tooling cannot establish a named campaign, so production data
should be corrected before relying on its tooling history.

A failed placement does not change campaign state. Other eligible presses may
be tried once; other orders continue. An actual placement that differs from the
validated resource/interval stops the action for review. The already placed
operation is not silently deleted or moved. Diagnostics are finalized even on
that error path.

## Setup-time integration, deliberately deferred

The workbook's `SET UP  DATA` tab is the agreed future source. This phase neither
imports its durations nor resolves “Old” versus “SEMED” methods, aggregate setup
times, NULLs or nomenclature aliases.

Before commit, each selected slot carries previous ring/pin, next ring/pin, and
`INITIAL_SETUP`, `UNCHANGED_TOOLING`, `PIN_ONLY_CHANGE`, or `RING_CHANGE`.
`QueryPressCampaign` and `CommitPressCampaign` are the adapter seams. Future setup
time must be included in the queried interval and placed exactly once. Appending
it only after feasibility testing would invalidate capacity and urgency results.

No automatic ad hoc or cooldown overrides are provided. Those remain planner
decisions. Capacity and lateness results are provisional until setup is modeled.

## Diagnostics

Existing scheduler diagnostics must be enabled through the existing configuration
when a production comparison is performed. No configuration toggle is changed by
this implementation. With diagnostics enabled, `07_OptimizerCandidateTrace.csv`
contains `CampaignDecision` entries with resource, previous/next tooling,
transition classification, actual interval and cooldown expiry in `ReasonDetail`.
Readiness exclusions and placement attempts remain in the existing trace files.

Reasons distinguish campaign continuation/start, urgency interruption, cooldown
or fixed-future-placement deferral, missing data, failed resource queries,
failed commits, existing violations and unexpected actual placements. Repeated
deferral entries represent separate evaluations, not separate orders.

## Local verification

Use Visual Studio MSBuild for .NET Framework 4.8. Build without COM registration
into separate output, intermediate and base-intermediate folders; an old normal-build registration
cache can otherwise try to unregister the installed DLL even with registration
disabled. The existing `bin/Debug` dependencies can be supplied as ReferencePath.

```powershell
$msbuild = 'C:/Program Files/Microsoft Visual Studio/2022/Community/MSBuild/Current/Bin/MSBuild.exe'
& $msbuild GrindwellNortonDev_VB.vbproj /t:Build /p:Configuration=Debug /p:Platform=x64 /p:RegisterForComInterop=false /p:ReferencePath="$pwd/bin/Debug" /p:BaseIntermediateOutputPath=obj/PressCampaignDebug/ /p:IntermediateOutputPath=obj/PressCampaignDebug/ /p:OutputPath=bin/PressCampaignDebug/ /v:minimal
& $msbuild Tests/PressCampaignHarness.vbproj /t:Build /v:minimal
& ./Tests/bin/PressCampaignHarness/PressCampaignHarness.exe
powershell.exe -NoProfile -File Tests/PressingCandidateHarness.ps1 -AssemblyPath "$pwd/bin/PressCampaignDebug/GrindwellNortonDev_VB.dll"
powershell.exe -NoProfile -File Tests/SharedHelpersReadinessHarness.ps1 -AssemblyPath "$pwd/bin/PressCampaignDebug/GrindwellNortonDev_VB.dll"
```

The pure-engine harness executes calendar, urgency, assignments, fixed future
work, failure, rerun and transition scenarios. Its controlled comparison uses
240 orders and four presses on the same synthetic calendar: the former ranked
earliest-placement policy makes 236 ring changes, versus 20 with campaigns,
with identical finish time and no late orders. This is not a production forecast.

The candidate harness loads the compiled production assembly and checks ranking,
completion/WIP exclusions, predecessor readiness and the legacy public queue.
The existing readiness harness checks shared-helper regressions.

Before client rollout, run old and new versions on copies of the same Opcenter
input dataset with identical calendars, assignments and schedule boundary. Compare
ring/pin changes, campaign lengths, scheduled/unscheduled counts, utilisation,
lateness and action runtime. Check all new placements for readiness, overlap,
assignment and cooldown violations. Export actual schedule results and decision
traces. Excel outputs alone do not contain the complete input/readiness/calendar
state required to reproduce this live comparison.

## Performance optimisation

The original campaign engine queried every pending order/resource after every
placement. The optimised engine orders candidates by their provable earliest
start and tooling preference, querying only candidates whose best possible
result could beat the current verified winner. Urgency checks still query any
higher-ranked ready order that could require interruption. This is an exact
search, not a fixed-size shortlist or a change to campaign priorities.

Native query results are shared only within a single decision. They are discarded
after each placement, so shared native constraints cannot leave stale results on
other presses. There is no assumed independence between machines. Eligibility
checks at commit still guard against an operation being scheduled externally.

Press history is sorted and indexed once per revision. Binary predecessor lookup
and an append cooldown fast path avoid repeated full timeline replay. Insertion
before fixed future work retains the full validation. Hypothetical urgency checks
restore the original history cache when their trial is removed.

The test-only `ReferencePressCampaignScheduler` is the frozen original exhaustive
engine. Differential tests compare every committed order, resource, start and
finish on 80 varied calendars, plus a 1,000-order workload. These tests also
measure native-query counts; fake-board runtime does not predict live COM cost.

Measured on this development machine with tracing disabled: the 1,000-order
fixture made 554,500 feasibility calls before optimisation versus 109,000 after
(80.3% fewer). Fake-board runtime was 41.8 seconds versus 3.9 seconds. The complete
placement sequence was identical. All 200 campaign assertions, the candidate
and shared-readiness harnesses, and the diagnostics kill-switch tests passed.
These figures are local regression measurements, not a promised live speedup.

## Turning diagnostics off

The easiest explicit override is the Windows user environment variable
`GN_SCHEDULER_DIAGNOSTICS_OFF` set to `1`. It disables this project's detailed
scheduler diagnostics **and** metrics, even if the host configuration enables
them. It does not disable Opcenter's own unrelated logging.

1. Close Opcenter.
2. In Windows, open **Edit environment variables for your account**.
3. Under user variables, create `GN_SCHEDULER_DIAGNOSTICS_OFF` with value `1`.
4. Relaunch Opcenter from a fresh launcher so it inherits the variable. If a
   long-running launcher retains the old environment, restart it too.

Alternatively, run this once in PowerShell, then restart Opcenter:

```powershell
[Environment]::SetEnvironmentVariable('GN_SCHEDULER_DIAGNOSTICS_OFF', '1', 'User')
```

Remove the override to restore the normal debug/metrics settings:

```powershell
[Environment]::SetEnvironmentVariable('GN_SCHEDULER_DIAGNOSTICS_OFF', $null, 'User')
```

Without this override, the existing `SchedulerDebugEnabled` and
`SchedulerMetricsEnabled` app settings still control diagnostics. Both are
`false` in the repository's `app.config`. For a hosted .NET DLL,
`ConfigurationManager.AppSettings` reads the **host executable's configuration**;
editing only `GrindwellNortonDev_VB.dll.config` may have no effect. An absent app
setting falls back to `GN_SCHEDULER_DEBUG` or `GN_SCHEDULER_METRICS`, respectively;
an explicit app setting takes precedence over those two fallback variables.

When diagnostics are off, no campaign reporter is installed, decision payloads
and strings are not constructed, and full before/after snapshots and CSV exports
are skipped. Functional validation still runs. Tests change only their own
process environment and do not change the user's saved Windows settings.
