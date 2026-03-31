# AGENTS.md

## File Metadata

- Last reviewed: 2026-03-31

## Purpose

- Repo-local notes for `C:\Scripts\MicVolumeGuard`, a standalone Windows PowerShell utility for keeping microphone volume fixed.
- Keep project-specific workflow and validation notes here; keep broader cross-repo notes in `C:\Users\marcm\.codex\AGENTS.md`.

## Current Snapshot

- Current task: maintenance and publication polish for `MicVolumeGuard`.
- Known caveat: behavior depends on Windows Scheduled Tasks and the selected default capture-device role, so functional validation is manual on a Windows desktop.
- Recent status: source was imported from `MicVolumeGuard.zip`, documented, committed, and published as the public GitHub repo `marcmy/MicVolumeGuard`.
- Build status: no build step applies; validate changes with PowerShell syntax checks and manual Windows install/uninstall testing when behavior changes.
