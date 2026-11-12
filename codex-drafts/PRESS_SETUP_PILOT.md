# Pressing setup pilot

**Review draft.** The expected results below describe the current code and project notes. Confirm with the business owner whether the difference between scheduling and repair is intended before proposing a rule change.

## Requirement and scope

Verify how operation 200 setup time is calculated when two orders with identical ring and bore/pin identities have an idle gap. Trace the configured PESP action into `AlgoSeq4.runPress` or `AlgoSeq4.RepairPressSetup` and follow the duration calculation in `PressSetupBoard.vb`. Use `PressCampaignScheduler.vb` if the result depends on campaign or candidate selection.

The pilot covers one press, one preceding order, one target order, and a protected operation as a control. Run the two actions from their appropriate starting states. For a subsequent change, compare baseline and candidate from identical starting state for each action.

## Small reproducible example

Use a disposable model and a press that can run both orders. Set ring setup to **6 minutes** and bore/pin setup to **9 minutes**, matching the synthetic rates in `Tests/PressSetupHarness.ps1`. Use `305-D` and `10` for the preceding operation. For the core case, finish it at 10:00 and arrange the target's feasible setup start at 10:15 on the same press. Use a working calendar and zero base setup. If the model has other setup contributions, record them separately and compare the explicit Orders `Setup Time` field first.

| Case | Target ring / bore-pin | Setup start relative to preceding finish | `runPress` expected minutes | `RepairPressSetup` expected minutes |
| --- | --- | --- | ---: | ---: |
| Pilot | `305-D` / `10` | 15-minute gap | 15 | 0 |
| Contiguous control | `305-D` / `10` | Contiguous | 0 | 0 |
| Ring only | `310-D` / `10` | Contiguous | 6 | 6 |
| Bore/pin only | `305-D` / `11` | Contiguous | 9 | 9 |
| Both changed | `310-D` / `11` | Contiguous | 15 | 15 |
| Bore/pin rate unavailable | `305-D` / `11` | Contiguous | 0 | 0 |

Each row is a separate case. A blank rate is unavailable. An explicit zero rate is valid. For the `runPress` comparison use the setup start returned by the native query. If a nominally contiguous case moves into a gap, record that as a separate outcome. For repair, compare the existing press sequence and capture any timing shifts from native recalculation.

## Replay steps

1. Record the source revision, DLL hash, exact APS build, disposable database and dataset, rates file hash, actual tooling mappings, diagnostic settings, and PESP actions. Verify Orders `Setup Time` uses `SEQ SETUP` and record other native setup settings. Use the installed Siemens help topics listed in [LAPTOP_SETUP.md](LAPTOP_SETUP.md).
2. Build locally and run `Tests/PressSetupHarness.ps1` against the resulting DLL. Use the harness for re-query, fallback, and simulated write/recalculation failures. Check the native outcomes below in Opcenter.
3. In the scheduling copy, start with the target unallocated, run Pressing, and capture its assigned press, setup start, process start, end, `Setup Time`, and `Total Setup Time`. Confirm the protected operation's state and the applicable calendar and predecessor constraints.
4. In the repair copy, start with the target already allocated behind the same preceding order with a stale 15-minute `Setup Time`. Run the PESP custom action **Repair Press Setup**. Check the target setup, actual timing, unchanged resource assignments and sequence, and the protected operation's state. Run repair a second time and check for zero additional field changes. Save and reopen the copy to verify persistence.
5. If the requirement is a runtime improvement, follow the paired measurement protocol in [LIVE_VALIDATION.md](../LIVE_VALIDATION.md) with restored inputs and identical diagnostic settings. Compare action time and available API counters alongside the scheduling outcomes.

## Evidence to record

| Item | Result |
| --- | --- |
| Approved expected behavior for idle gaps in each action | Pending business confirmation |
| Exact APS build and matching Siemens help topics | Pending |
| Revision, DLL hash, dataset, fixture and configuration | Pending |
| Harness command and outcome | Pending pilot run |
| Native pressing and repair before/after exports | Pending pilot run |
| Protected operation, repeat repair, and saved/reopened checks | Pending pilot run |
| Unexpected behavior or follow-up requirement | Pending pilot run |

Sources for this draft are [PRESS_CAMPAIGNS.md](../PRESS_CAMPAIGNS.md), [PressSetupBoard.vb](../PressSetupBoard.vb), and [PressSetupHarness.ps1](../Tests/PressSetupHarness.ps1).
