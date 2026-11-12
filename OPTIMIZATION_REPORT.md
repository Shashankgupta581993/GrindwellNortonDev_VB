# branch2 optimization and final verification

Date: 2026-11-03. Baseline commit: `04f0daa`.

**Status: local implementation, documentation, and regression verification complete;
live Opcenter acceptance remains pending.** The user confirmed that Opcenter is
not running for validation. No live dataset, SQL database, COM registration,
deployment, public signature, COM identity, or production project structure was
changed. Existing `MES_SCHEDULING_DATE.md` work was preserved byte-for-byte.

## Optimization register

| Change | Why it improves the targeted cost | Behavior protection |
|---|---|---|
| Code master and 17 individual debugger switches | Skips disabled diagnostic producers, trace payloads, snapshot reads, directories, files, and standalone logging. Export-only SWK path resolution and firing folder formatting are guarded too. | Explicitly requested diagnostic change; operational messages, business exports, scheduling decisions, and MES effects retained. |
| Snapshot prerequisite calculation | Builds operation snapshots only for enabled consumers; field-map-only and metrics-only modes avoid per-operation snapshots. | Internal prerequisites do not enable another channel's output. |
| Compact firing parent-stage index | Replaces a complete snapshot filter/list per firing operation with a pre-sized per-call dictionary of loading/unloading bits. | Stable output order, missing/zero/negative parents, and existing eligibility rules compared with baseline. An initial two-index design was rejected because it increased allocations. |
| Single SWK order-count pass | Removes a full order-map scan for every exported batch. | Sorted batch order, queue order, counts, empty plans, and CSV bytes preserved. |
| Reusable SWK summary/queue field arrays | Removes one string-array allocation for each output row, offsetting the count index's storage cost. | Field evaluation order, format, escaping, encoding, and newlines retained. |
| Reusable diagnostic CSV field array | Removes a temporary List and its backing array per exported row. | Reflection property/getter order, exceptions, partial output, and CSV bytes tested. |
| Legacy CSV indexed reads/buffer reuse | Removes LINQ projections and permits avoiding ItemArray copies for framework primitive/sealed-value columns. | Object/custom columns retain ItemArray snapshots, including a ToString callback that mutates another cell. Legacy comma-to-space behavior remains. |
| Private reusable character arrays | Removes per-call separator-array allocations in numeric-date detection and diagnostic escaping. | Date parsing and escaping behavior retained. |
| Process mapping TryGetValue lookups | Avoids duplicate dictionary hashing at both operation and cycle levels. | Successful, missing, and empty lookup results/out values compared against baseline. |
| Source cleanup | Removes ten fully commented obsolete methods and clarifies diagnostic ownership/guards. | Active code and algorithms retained; no runtime gain claimed for deleting comments. |

Existing action-scoped caches, bulk loading, invalidation, WIP/release-boundary
logic, stage arrays, stable sorting, and campaign/packing algorithms remain.
`Schedule400` and `afterFiring` stay independent; `runFix.nextOpIndex` remains.

No blanket LINQ/For Each rewrite, eager-operator substitution, typed routing-schema
migration, cross-action cache, shared pool, background COM work, or database/UI
rewrite was introduced without evidence. See `CODE_AUDIT.md` for the complete
retained-opportunity and suspected-defect discussion.

## Measured local results

Release .NET Framework 4.8; warm-up followed by five alternating paired runs.
The diagnostic workload has 6,000 operation snapshots. The SWK export workload
has 500 batches and 5,000 queued operations. Values are medians from the final
run, not a projected production speedup.

| Synthetic workload | Baseline ms | Optimized ms | Baseline allocated bytes | Optimized allocated bytes |
|---|---:|---:|---:|---:|
| Stage eligibility diagnostics | 341.262 | 6.418 | 1,285,760 | 754,296 |
| SWK CSV export | 82.726 | 51.118 | 6,029,744 | 5,575,288 |

Allocations use AppDomain monitoring around the measured invocation, including
some harness overhead. Elapsed results vary with machine load. These measurements
show lower targeted runtime and allocations on the fixtures; they do not measure
live scheduler COM time, whole-process peak memory, or every optimizer workload.

The raw five-pair measurements are retained locally in
`Tests/obj/OptimizationEvidence/diagnostic-benchmark.json` and
`Tests/obj/OptimizationEvidence/swk-benchmark.json`. That generated directory is
ignored by Git; the median results above remain reviewable in this report.

## Final verification

| Check | Result |
|---|---|
| Final Debug x64 build, isolated output, COM registration disabled | Passed |
| Final Release x64 build, isolated output, COM registration disabled | Passed |
| Shared readiness/WIP/boundary/settings/schema harness | Passed against Debug and Release |
| Pressing candidate harness | Passed against Debug and Release |
| MES date/stamp/retry/placement harness | Passed against Debug and Release |
| Default code-settings diagnostic harness | Passed against Debug and Release |
| Isolated debugger matrix | All 21 configurations passed |
| Pressing campaign engine/reference harness | 200 assertions passed, including 80 differential calendars and 1,000-order exact-placement parity |
| Exported public types/member signatures and type attributes | Baseline and optimized all-on build match |
| CSV differential checks | Three cultures, quoted/newline fields, getter order, throwing getter/partial output, legacy custom-object snapshots passed |
| Enabled legacy decision/schedule logging | Byte parity passed with normalized fixed run identity |
| SWK summary/queue exports and empty-plan handling | Byte parity/edge checks passed |
| Process-mapping dictionary lookup behavior | Baseline parity passed |
| Eligibility output | 6,000-row baseline parity passed, including zero/negative parent grouping |
| Read-only replay comparator harness | Passed; differing values, headers, counts, and missing files rejected |
| Production diagnostic environment/config override scan | No old override reads or diagnostic app-setting keys remain |
| Whitespace/diff review | Passed; unrelated MES document preserved |

The campaign reference comparison predates these changes; its query reduction is
not claimed as a new optimization from this task. Local fixtures do not establish
universal equivalence for untested external inputs or a live PlanningBoard.

### Artifact fingerprints (SHA-256)

| Artifact | SHA-256 |
|---|---|
| Saved pre-change baseline | `C4FAA00414A2BA98A6622FD50F8FEA321EB25D32C59D82BF6754962D4101C40C` |
| Final Debug | `41678D6C7C7CEF84F8C2628997BE2EBC7B401AF870A52B14510DB4F030F8D9A8` |
| Final Release | `47F03188D7E7A20F660FA94C3ED4F9100ECB7742A70A8BCEF63BB3FE44ED9DC3` |

The normal `bin/Debug` and `bin/Release` deployment outputs were not replaced.
Compilation used `bin/OptimizationDebug` and `bin/OptimizationRelease` with
separate intermediate directories and `RegisterForComInterop=false`.

## Reproduce local checks

Use Windows PowerShell and Visual Studio MSBuild. Build Debug/Release into the
isolated directories documented in `PRESS_CAMPAIGNS.md`, substituting
`OptimizationDebug`/`OptimizationRelease` for the example output/intermediate paths.
Run the four existing PowerShell harnesses against each fresh assembly, then:

```powershell
powershell.exe -NoProfile -File Tests/DebuggerSettingsMatrix.ps1
powershell.exe -NoProfile -File Tests/OptimizationDifferentialHarness.ps1
powershell.exe -NoProfile -File Tests/ReplayComparisonHarness.ps1
& Tests/bin/PressCampaignHarness/PressCampaignHarness.exe
```

The differential harness requires the saved baseline DLL at its default path,
or `-BaselineAssembly` pointing to a genuine pre-change build. It uses the
matrix's all-on build as its optimized comparison so enabled diagnostics can be
compared. The production default remains master-off. Harnesses do not deploy or
register either assembly.

## Remaining external gate

`LIVE_VALIDATION.md` contains the exact setup, capture/comparison procedure,
profiling protocol, and unchecked release criteria. Still required are identical
restored Dataset 15 copies, action sequence/MES date, controlled host loading,
full-precision business-effect comparisons, and five paired production profiles
covering all three diagnostic modes, native/COM time, GC, and peak memory.

No live result or production improvement percentage is inferred from historical
approvals, synthetic fixtures, folder names, or an apparent process/window.
Deployment and production acceptance are not marked complete.
