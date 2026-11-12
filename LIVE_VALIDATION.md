# Controlled Opcenter replay and release gate

Status on 2026-11-03: **not executed**. The user confirmed that Opcenter is not
running for this validation. Local process/window discovery alone did not establish
a usable test session, Dataset 15 access, two restored copies, or the exact replay
inputs. Existing model folders and `.bak` files were found, but their contents
were not restored or assumed to be identical current scheduling data.

No live schedule, SQL database, deployment, or COM registration was changed.

## Required test setup

1. Identify the disposable test model and database server. Restore two independent
   copies from the same authoritative snapshot. Confirm `Schedule`, Dataset ID 15,
   record count, resource calendars, optimizer settings, external CSVs, and initial
   schedule state match. A folder copy alone does not duplicate a SQL database.
2. Record the exact configured action sequence and MES effective date. Preserve
   the same existing allocations, planning boundary, application version, machine,
   and input state for both implementations. Do not assume the four action names
   listed below imply their intended production execution order.
3. Run the saved baseline assembly and optimized assembly in separate controlled
   host sessions. Confirm which DLL is loaded. Use matching diagnostic channels;
   see `DEBUGGER_SETTINGS.md` for optimized code settings. Baseline configuration
   must be enabled in the disposable baseline host using its original mechanism.
4. Do not replace the currently registered DLL during local verification. Controlled
   deployment/reload belongs to this live gate; retain an immediate rollback copy.

## Equivalence captures

Capture each of `runFiringFollowOn`, `runFix`, `Schedule400`, and `afterFiring`
independently from identical initial state, then the actual complete configured
action sequence. Include pressing, batch, tunnel, and SWK actions used by the model.

Compare ordered candidates, resource enumeration/ties, batching, requested start
times/directions, placement sequence, actual resources/start/end, return values,
handled failures, completion/WIP boundaries, and MES date stamps/retries. Preserve
full timestamp precision in an authoritative business-state export. The diagnostic
CSV formatter uses seconds and is not an oracle for subsecond differences or
business fields it does not export.

Use focused scenarios for completed/actualized operations, already scheduled work,
missing fields, exact ties, readiness/cooldown dates, failed queries/placements,
partial batches, and repeated actions. Restore initial state between runs.

The read-only helper compares the existing detailed diagnostic captures:

```powershell
powershell.exe -NoProfile -File Tests/Compare-OpcenterReplay.ps1 `
    -BaselineFolder 'D:/Validation/Baseline/runFiringFollowOn' `
    -OptimizedFolder 'D:/Validation/Optimized/runFiringFollowOn'
```

It requires all eight default capture files, preserves row/column order, parses
quoted multiline CSV, and ignores only `RunId` and `ExportedAt`. Missing evidence
or a differing header, value, or row count fails. It never launches Opcenter or
changes either capture. Metrics are deliberately not compared for equality.
An identical authoritative business-state CSV can be compared separately with
`-Files`; choose its fields to include full timestamps, actual resource, MES
stamps, and all effects relevant to the action. Passing diagnostic comparison alone
does not complete the business-equivalence gate.

## Performance protocol

- Measure diagnostics-off, metrics-only, and detailed-diagnostics modes separately.
- Warm each implementation, then run five alternating baseline/optimized pairs,
  restoring the same initial state before every measured run.
- Record total and phase duration, allocations, GC counts, peak process memory,
  record/candidate/queue counts, API-call counts, and actual successful placements.
- Use `12_ActionMetrics.csv` for available phase/API counters, supplemented with
  an external .NET/Opcenter profiler for allocations, GC, native/COM time, and peak
  memory. The existing counters do not cover every COM read; see `CODE_AUDIT.md`.
- Compare medians and retain raw captures, build hashes, hardware, application
  version, and profiling settings. Local harness duration is not scheduler time.

## Acceptance record

- [ ] Authoritative identical test copies, action sequence, and MES date recorded.
- [ ] Independent action and full-sequence business effects match.
- [ ] Required edge cases and failed/partial placements match.
- [ ] Diagnostics-off, metrics-only, and detailed mode outcomes match.
- [ ] Five paired production measurements and peak-memory/API evidence recorded.
- [ ] No repeatable speed or memory regression in the targeted workloads.
- [ ] Controlled deployment and rollback verification completed.

Until these items pass, production equivalence and performance remain pending;
the local implementation and regression results must not be described as a
completed live validation.
