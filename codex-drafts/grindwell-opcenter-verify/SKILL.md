---
name: grindwell-opcenter-verify
description: Build and verify Grindwell Norton VB.NET Opcenter APS scheduling changes on Windows. Use for isolated x64 builds, targeted scheduler harnesses, or preparing a controlled Opcenter replay.
---

# Verify Grindwell Norton Opcenter changes

1. Find the repository root containing `GrindwellNortonDev_VB.vbproj`. Record the source revision and relevant pre-existing edits. Identify the configured PESP action, the changed rule, and the expected scheduling effects.
2. Read the project notes for that action. For press setup use `PRESS_CAMPAIGNS.md`, `AlgoSeq4.vb`, and `PressSetupBoard.vb`. Locate the installed Siemens configuration or SDK help for native behavior the change relies on. Record the product build and help topic.
3. Run `scripts/Build-Local.ps1` beside this skill in Windows PowerShell with `-RepositoryRoot <repo>`. Use `-SdkReferencePath <directory>` if the repository's `bin/Debug` lacks matching Opcenter references. The script returns the fresh Debug x64 DLL path. If reference resolution fails, report the missing dependency and locate the matching SDK. Preserve the project reference configuration during verification.
4. For setup calculations or repair, run `powershell.exe -NoProfile -File Tests/PressSetupHarness.ps1 -AssemblyPath <dll>` from the repository root. Add the pressing candidate harness for readiness/ranking changes and the campaign harness for campaign selection or resource-order changes. Use the commands and fixture notes in `PRESS_CAMPAIGNS.md`. For other actions, select the relevant harness in `Tests/` and check its parameters.
5. For host behavior, prepare an Opcenter replay from `LIVE_VALIDATION.md` and the action-specific notes. Record the disposable model and database, input snapshot, loaded DLL hash, configuration, action order, and expected differences. Run a mutating test only in a confirmed disposable host within the task's authorization. Keep any pending host step visible in the report.

Report the build path and hash, harness commands and outcomes, replay observations, and any unresolved native behavior. For performance changes include equivalent workload details, timing, and available API-call counts. The build helper leaves COM registration disabled. A host load, registration, or deployment needs its own explicit scope.
