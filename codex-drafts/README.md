# Grindwell Norton Codex setup drafts

These files are staged here for review. Use the pilot to confirm the workflow on the laptop before adopting the instructions and skill.

| Draft | Review or adoption step |
| --- | --- |
| [AGENTS.md.draft](AGENTS.md.draft) | Review the standing project rules, then copy the file to the repository root as `AGENTS.md`. |
| [LAPTOP_SETUP.md](LAPTOP_SETUP.md) | Record the actual Opcenter build, SDK path, Siemens help locations, and disposable database details. |
| [PRESS_SETUP_PILOT.md](PRESS_SETUP_PILOT.md) | Review the proposed business examples, run the first replay, and record the observed results. |
| [grindwell-opcenter-verify](grindwell-opcenter-verify/SKILL.md) | After the pilot, move the whole skill folder into `.agents/skills/` to make it discoverable in this repository. The folder includes the isolated build helper. |

Use ChatGPT to discuss the requirement and expected outcomes. Use a local Codex chat to read the repository and Siemens help, and a Codex worktree for implementation once the baseline is checkpointed. Use Plan mode when the requirement crosses actions or changes scheduling behavior with unclear consequences. For a larger change, an independent subagent can verify Siemens SDK behavior or review the diff while one agent owns edits.

## Prompt for the first Codex chat

> Read `codex-drafts/PRESS_SETUP_PILOT.md`, `PRESS_CAMPAIGNS.md`, and the relevant code and harness. Trace the same-tooling idle-gap case through `runPress` and `RepairPressSetup`. Identify the exact Siemens help topics to verify, any difference between the code and the notes, and what you need from the laptop model to reproduce it. This first pass is read-only. Give me the expected results and a short replay plan.

[Codex environments](https://learn.chatgpt.com/docs/environments/modes) · [Codex slash commands](https://learn.chatgpt.com/docs/reference/slash-commands) · [Build skills](https://learn.chatgpt.com/docs/build-skills)
