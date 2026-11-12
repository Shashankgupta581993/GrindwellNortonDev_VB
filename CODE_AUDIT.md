# Unused-code recommendations and suspected defects

Audit date: 2026-11-03. Scope: compiled production VB.NET source on branch2.
Recommendations below are separate from the semantic-equivalence refactoring.
No public/COM class or method was deleted based on a repository-only search.

## Private members with no observed production callers

| Location | Candidate members | Recommendation |
|---|---|---|
| `AlgoSeq3.vb` | `DumpOpRec` | Optional diagnostic helper; remove if manual/reflection use is unnecessary. |
| `AlgoSeq4.vb` | `ScheduleTunnelReleasedContinuations`, `GetWipReadyTimeFromRouting`, `CreateRankedQueue`, `ShowExplorerUI` | Review retained alternate implementations and any external reflection callers before deleting. |
| `PostFiringScheduler.vb` | `IsKilnAckRow` | Review whether this compatibility predicate is still needed. |
| `PressCampaignScheduler.vb` | `Continues`, `SamePin` | Current production selection uses `SearchKey`/`CompareOptions`; these helpers appear unused. Keep similarly named helpers in the frozen reference engine: its tests still call them. |

Evidence: identifier/call searches over production `.vb` files found only these
declarations for the listed production members. This is a static reachability
indication, not proof against reflection. WinForms `Handles` event methods were
excluded from the deletion recommendations even when they have no textual callers.

Legacy public actions such as `AlgoSeq2`, `AlgoSeq3`, `EventSeq`,
`Preactor_Extensions`, `showMeta`, `testFunction`, and `freezeWindow` may be bound
from Opcenter configuration. Their absence from internal calls does not establish
that they are unused. Preserve their identities and inspect the actual action
configuration before considering retirement. No entire class/module is proven
unused by this audit.

Ten fully commented-out obsolete methods were removed from firing, tunnel, and
shared-helper source. Active replacements and explanatory comments remain. This
improves readability; commented code did not incur runtime cost.

## Suspected defects retained for separate review

1. **Placement-success metrics can overcount.** In `Schedule400`, `runFix`,
   `PutOperationWithTrace`, and post-firing placement paths, some success counters
   increment after `PutOperationOnResource` returns without consistently checking
   its Boolean result. The post-firing path also marks the cached operation as
   placed. Treat these counters as instrumentation, not an independent success
   oracle. Correcting this could affect traversal and is outside this refactoring.
2. **`EventSeq` shift-change branch uses retained resource state.** The branch
   computes `Q = planningboard.GetResourceQueue(ResRec)` after obtaining the event's
   resource through `Parameter2`. `Q` is implicit under `Option Explicit Off` and
   `ResRec` can contain earlier-loop state. Investigate with actual event fixtures;
   no typing or event-handling change was made.
3. **Operation-row index validation is incomplete for arbitrary table edits.**
   `BuildOperationRowIndex` checks row count and table ownership. Changing an
   `OrdersID` in place, or replacing rows without changing count, can leave cached
   keys stale. Current action-owned snapshot use was retained. A broader
   invalidation change requires consumer and mutation tests.

## Performance opportunities deliberately retained

- Keep existing run-scoped COM caches, invalidation, bulk loading, and row indexes.
  There is no new cross-action scheduling cache or background COM work.
- Keep stable LINQ ordering and current press/batch/tunnel/SWK selection algorithms.
  Equal-key behavior and mutable feasibility state make blanket rewrites unsafe.
- Keep the routing DataTable fully `Object`-typed and preserve `DBNull`. A typed
  schema or reduced field profile needs a production profile and column oracle.
- Keep existing public `Object`/ByRef COM parameters. `EventSeq`'s relaxed typing
  needs event-level fixtures before conversion behavior can be changed safely.
- Keep existing database/UI disposal, connection ownership, transactions, error
  messages, and event timing. Prior hardening already uses `Using` and
  `SqlConnectionStringBuilder`; no new evidence justified a database/UI rewrite.
- Retain short concatenations, efficient collection enumeration, and deliberate
  copies whose ownership/order guarantees matter. Existing parsers/builders
  already use `StringBuilder` where appropriate. Shared pools were not introduced.

These decisions prevent speculative performance edits from altering business
behavior. They are not claims that live profiling found no remaining bottlenecks.
