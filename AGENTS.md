# AGENTS.md

## File Metadata

- Last reviewed: 2026-03-31

## Purpose

- Repo-local notes for `C:\Scripts\MicVolumeGuard`, a standalone Windows PowerShell utility for keeping microphone volume fixed.
- Keep project-specific workflow, validation notes, and publication caveats here; keep broader cross-repo notes in `C:\Users\marcm\.codex\AGENTS.md`.

## Current Snapshot

- Current task: initial repo packaging and GitHub publication prep for `MicVolumeGuard`.
- Known caveat: `gh auth status` reported an invalid GitHub token on 2026-03-31, so remote repo creation and push are blocked until re-authenticated.
- Recent status: source was imported from `MicVolumeGuard.zip` with install/uninstall wrappers plus the main microphone volume guard script.
- Build status: no build step applies; validate changes with PowerShell syntax checks and manual Windows install/uninstall testing when behavior changes.
