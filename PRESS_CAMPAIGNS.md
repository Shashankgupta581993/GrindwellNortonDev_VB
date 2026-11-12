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

## Press tooling setup

`Opcenter/PressSetupRates.csv` contains the 34 workstation rows from
`Press Setup Data.xlsx`, sheet `Setup for Press`. Maintain all rates in the
workbook. Publish a comma-separated UTF-8 CSV containing the workstations in the
target Opcenter database. The `GrindwellNorton1` deployment has 21 matching finite
resources. Blank rates are unavailable. Zero is an explicit rate. PL12001A and
PL12001B use SEMED.

- During **Pressing**, contiguous operation-200 placements compare the ring and bore/pin
  independently. Both known identities unchanged gives zero. If the press is
  empty or setup starts after any idle time, both components are due. A gap
  across nonworking calendar time also counts. Two timestamps that round to
  the same saved Orders timestamp count as contiguous.
- Add the applicable rates that exist and write their sum to the existing
  Orders **Setup Time** field for operation 200. An unknown identity cannot
  establish a change for that component. When no applicable rate exists, use
  the site's original zero-minute base setup. This clears a prior calculated
  setup if a rate disappears. If the site's imports introduce nonzero base
  setups, the code needs a stable source for those values before deployment.
- `runPress` uses the selected duration in the native feasibility query and
  persists it only on a verified placement. A query that moves setup into an
  idle gap is repeated with the new duration. The scheduler protects existing
  allocations and includes the tested interval in campaign urgency decisions.
- **Repair Press Setup** keeps future movable pressing operations on the same
  press and in the same order. It computes the final Setup Time values from
  the allocated sequence before writing any fields. Setup is charged only for tooling values that
  differ and are known on both the current and preceding order, regardless
  of any gap. The first order on an empty press uses the available rates for
  its known tooling. Opcenter recalculates operation times once per affected
  press using the existing allocations. Repair verifies the allocation set,
  resource sequence, Setup Time, and Total Setup Time before accepting the result.
  It restores the original fields and schedule if an allocation changes
  unexpectedly or timing violates readiness, calendar, or dependency checks.
  Use this action after changing press rates, tooling mappings, or a Gantt sequence.

### Opcenter 2210 deployment

1. Verify the existing Orders **Setup Time** field has the `SEQ SETUP`
   classification. Opcenter can combine this duration with other setup
   classifications. The audited `GrindwellNorton1` pressing orders had zero
   base setup, zero positive resource-specific setup values, and no configured
   changeover groups or setup-efficiency multipliers on the resources.
2. Deploy the Branch2 assembly and its `.dll.config` at the registered x64
   COM path. Set `PressSetupCsvPath` in the DLL configuration to the published
   CSV. A relative path resolves from Opcenter `{PATH}`. The press action
   explicitly reads the sidecar file. Existing host values take precedence by
   key, including a blank path that disables the feature. The default mappings
   are `PressRingField=String Attribute 5` and
   `PressBorePinField=String Attribute 4`. Restart Opcenter to load a new DLL.
3. Create a PESP script named **Repair Press Setup** with one Custom Action
   calling `GrindwellNortonDev_VB.AlgoSeq4.RepairPressSetup`. Bind it to an
   available Sequencer user-tool slot through **Tool Configuration**. Use the
   same custom-action registration path as `runPress`. The native Repair
   Schedule command does not run this script.
4. In a named schedule copy, run the custom action and verify the recalculated
   setup and process times, missing-rate fallback, protected operations, and
   downstream dependencies. Confirm that the same records remain on each press
   in the same sequence. Run the action again and confirm zero field changes.
   Reopen the saved copy and check that the allocations and setup values persisted
   before rolling out the build.

Use the installed Siemens Opcenter APS 2210 configuration help for
`SEQ SETUP` and PESP Custom Actions. The 2210 SDK API
help covers the planning-board query, placement and schedule-state methods.
To remove the feature from an enabled dataset, restore the site's base
**Setup Time** values and rebuild the affected schedule on a copy first.
A blank or absent `PressSetupCsvPath` in the effective configuration disables
the Branch2 code path.

## Diagnostics

Scheduler diagnostics are controlled by the code-defined master and channel switches
in `Debugging/SchedulerDebugCollector.vb`; see `DEBUGGER_SETTINGS.md`.
For a production comparison, enable the same channels in both test builds.
With candidate diagnostics enabled, `07_OptimizerCandidateTrace.csv`
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

Set `SchedulerDebugSettings.MasterEnabled = False` in
`Debugging/SchedulerDebugCollector.vb`, rebuild, and reload the DLL in Opcenter.
Each diagnostic channel also has its own Boolean setting. The master overrides
all channels; environment variables and host app settings no longer enable or
disable this project's diagnostics. See `DEBUGGER_SETTINGS.md` for profiles.

With candidate diagnostics off, no campaign reporter is installed. Existing
campaign feasibility and safety validation still runs. Operational errors,
required prompts, business exports, and Opcenter's own logging remain separate.
The diagnostics harness now tests isolated builds with code-defined settings.
