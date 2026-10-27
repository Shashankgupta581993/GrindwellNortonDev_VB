# MES Plan Effective Date — branch2

The custom rules prompt for a date before the first scheduling action. Each
successful `PutOperationOnResource` result stamps that operation. Identical
placements and moves back to the original allocation qualify. Failed placements
do not. The date is reused across separate rule buttons and separate COM class
instances in the same Opcenter session and dataset.

There are no baseline scans, SQL calls, or extra schedule reads for this feature.
Format/field/type resolution happens once per action; the first successful
placement per operation per cycle adds one `WriteField` call. Further placements
of that operation in the cycle skip the write.

## Required Opcenter configuration

These configuration and SQL view sources are **not in this repository**. Complete
the following in the GN configuration before enabling the updated rules in UAT:

1. Add or map an **operation-level**, persisted Orders **DateTime** field named
   `MES Plan Effective Date`. It must allow an empty value for utouched rows and
   be included in then normal dataset save/load mapping. Do not configure it as an
   inherited order-wide value or an input to scheduling calculations. Existing
   imports must preserve this APS-owned field unless deliberately resetting it.
2. If the existing field has another name, set `MesPlanEffectiveDateField` in the
   effective application configuration to that exact name. The default is
   `MES Plan Effective Date`. A library's `.dll.config` alone does not change
   the host application's settings. Missing or non-DateTime fields block the
   rule before placement.
3. Register the updated assembly using the site's existing deployment procedure.
   Existing rule COM GUIDs and method signatures are unchanged. Add a custom
   action pointing to `GrindwellNortonDev_VB.MesSchedulingActions.ChangeDate`,
   labelled **New Scheduling Cycle / Change Date**. Its signature follows the
   existing two-parameter custom actions and returns `0` on success, `-1` on
   cancellation or initialization failure. Respect that result in calling scripts.
4. Add `GrindwellNortonDev_VB.MesSchedulingActions.InvalidateCycle` to the
   dataset load/reload, import, undo/restore and close workflows, before another
   rule can be executed. It returns `0` and clears only temporary cycle state,
   never the saved operation dates. **This wiring is mandatory**, including a
   reload of the same dataset: `CurrentFileName` cannot identify that event.
   If automatic hooks are unavailable, the dataset workflow must invoke this
   action explicitly and require a new cycle before enabling scheduling.
5. Keep the existing Save workflow. Expose the persisted operation field in the
   existing MES view with its agreed column name/date format. Preserve existing
   row selection criteria. Verify that the view reads the intended saved dataset
   and that the Orders field's physical column is populated after Save. No view
   definition or physical column name is assumed by this change.

Session separation uses the native Opcenter object's identity plus the
configuration path and Orders `CurrentFileName`. Restarting the application
discards all temporary context. A date does not expire automatically at midnight.
Cancel at the first prompt prevents the rule running. Cancel in Change Date
preserves the previous cycle and date. A confirmed Change Date starts a fresh
cycle, even when the planner chooses the same date again. Changing dates does
not modify operations until they are placed again.

## Failure handling and boundaries

A failed date write raises `MesSchedulingException`, bypassing the existing
per-record/per-order catch-and-continue handlers. Earlier successful placements
remain on the live board. The failed operation is retained for retry; the next
rule or Change Date action repairs it with the **original** date before allowing
more scheduling or selecting another date. If repair still fails, the action
does not start. Fix the field/save configuration and retry; do not export/save a
partially stamped run for MES. Normal Save is not intercepted by this MVP.

The configuration must invalidate the cycle after any reload, import or undo
because these can change operation record identities or restore older field
values. Imports, manual edits and standard rules are outside an active MVP cycle.
Operations moved indirectly by Opcenter are not stamped unless a custom rule
explicitly places them. Successful placement means the SDK returned `True`;
an API exception is not treated as a successful placement.

All compiled placement paths in `AlgoSeq`, `AlgoSeq3`, `AlgoSeq4`, `EventSeq` and
`PostFiringScheduler` are instrumented. `EventSeq.ScheduleOperations` and
`PostFiringScheduler.ScheduleQueue` require the caller's `IPreactor` wrapper to
have been prepared by their scheduling entry point; the deployed `Run` and
`afterFiring` paths do this. Confirm live button mappings against this list.

## Validation

Build x64 with COM registration disabled and isolated intermediate/output paths.
Run `Tests/MesSchedulingDatesHarness.ps1` with **Windows PowerShell** and the
built DLL path. It tests date normalization, untouched rows, deduplication,
new-cycle overwrite, pending-write recovery and the real shared placement helper
against SDK interface proxies, including false/throwing placements and move-back
behavior. Also run the existing readiness, candidate, diagnostics and press
campaign harnesses.

UAT must verify the real date dialog, cancellation, date reuse across rule
buttons, field validation, every deployed placement path, lifecycle invalidation
(especially reload of the same dataset and undo), save/reload persistence, and
the MES view. Measure added action time on the representative approximately
110,000-operation dataset with identical inputs and diagnostic settings. Local
tests cannot establish live Opcenter/SQL performance or save/export correctness.
