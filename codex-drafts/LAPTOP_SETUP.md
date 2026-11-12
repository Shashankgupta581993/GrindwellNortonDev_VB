# Grindwell Norton laptop setup

The laptop and client run the same Opcenter version according to the project owner. The repo documents APS 2210, and the laptop has Visual Studio MSBuild and the .NET Framework 4.8 reference assemblies. Record the exact Opcenter build and SDK path for a reproducible run.

## One-time setup

1. **Source baseline.** Review `git status --short` and choose which pending work belongs in the next baseline, including any untracked press setup sources and tests. Checkpoint the selected work before creating a Codex worktree. Record the commit SHA and the DLL SHA-256 separately.
2. **Build references.** Identify the matching `Preactor.*.dll` directory from the installed SDK or a verified build output. A new worktree may lack the ignored `bin/Debug` directory. Pass the verified reference directory to the build helper using `-SdkReferencePath`. Keep local signing material outside the checkout and use [BUILD_SIGNING.md](../BUILD_SIGNING.md) for release builds.
3. **Siemens references.** Locate the installed configuration and SDK help for the exact APS build or obtain the matching release documentation from Siemens Support Center. Record the location and topic titles for `SEQ SETUP`, PESP Custom Actions, planning-board feasibility and placement, schedule state, and resource-time recalculation. For CHM or other help that Codex cannot search directly, provide approved excerpts with the source title, release, and topic. [Siemens Support Center](https://support.sw.siemens.com/)
4. **Local test model.** Bring over an approved snapshot and restore it under a distinct local database name. Point a separate Opcenter model at that database. Confirm the effective connection, dataset, PESP action configuration, and registered DLL before running a scheduling action. For baseline and candidate comparisons, restore identical starting state between runs. A model folder copy alone leaves the database connection to verify.
5. **Codex permissions.** Use read-only for discovery and workspace-write for code and harness work. Use explicit authorization for COM registration and a controlled test load of the DLL. Keep client credentials and signing keys outside the Codex workspace. [Codex approvals and security](https://learn.chatgpt.com/docs/agent-approvals-security)

## Client material for the first replay

| Bring over | Capture alongside it |
| --- | --- |
| Model and SQL backup or equivalent restorable test snapshot | Opcenter build, dataset, restore point, planning boundary, and schedule name |
| Effective model and DLL configuration | `PressSetupCsvPath`, `PressRingField`, `PressBorePinField`, Orders `Setup Time` classification, native setup configuration, and PESP bindings |
| Press rates CSV and source workbook version | File hash, tested workstation, and the corresponding ring and bore/pin rates |
| Calendar and initial state for the selected press | Shift gaps, preceding operation, tested operation, protected allocations, and dependencies |
| A baseline before/after capture | Order ID and operation, resource, sequence, setup start, process start, end, `Setup Time`, and `Total Setup Time` at available precision |

## Local build and pressing harness

Run from the repository root in Windows PowerShell. Supply a verified SDK directory if `bin/Debug` is absent or contains assemblies from another build.

```powershell
$dll = & .\codex-drafts\grindwell-opcenter-verify\scripts\Build-Local.ps1 -RepositoryRoot .
powershell.exe -NoProfile -File .\Tests\PressSetupHarness.ps1 -AssemblyPath $dll
Get-FileHash -LiteralPath $dll -Algorithm SHA256
```

The helper creates fresh folders under the ignored `bin/CodexLocal` and `obj/CodexLocal` paths. Use the returned DLL path for the harness. The helper performs a Debug build with COM registration disabled. A controlled Opcenter host load and replay are separate steps described in [PRESS_CAMPAIGNS.md](../PRESS_CAMPAIGNS.md) and [LIVE_VALIDATION.md](../LIVE_VALIDATION.md).
