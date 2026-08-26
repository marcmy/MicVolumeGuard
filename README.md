# MicVolumeGuard

MicVolumeGuard is a small Windows PowerShell utility that keeps the current Windows default microphone volume pinned to a target percentage so automatic gain control or other apps cannot drift it away from your preferred level.

## What it does

- Watches both the current **Default Device** and **Default Communications Device** for Windows capture audio.
- Re-resolves both defaults on every poll, so changing the default microphone in Windows automatically moves the guard to the new device.
- Deduplicates the endpoints when both Windows defaults point to the same microphone.
- Restores microphone volume when it moves outside a tolerance threshold.
- Can optionally restore any change immediately and optionally restore mute state too.
- Logs endpoint changes so it is visible when Windows switches either default microphone.
- Installs as a Scheduled Task with optional auto-start at logon plus desktop shortcuts to start and stop the guard manually.
- Writes lightweight runtime logs to `%LocalAppData%\MicVolumeGuard\MicVolumeGuard.log`.

## Files

- `Install.cmd` runs the installer with elevation.
- `Install.ps1` prompts for the target percentage and whether the guard should auto-start at logon.
- `MicVolumeGuard.ps1` is the long-running guard process.
- `Uninstall.cmd` runs the uninstaller with elevation.
- `Uninstall.ps1` removes the scheduled task, shortcuts, running guard process, and the default log file.

## Requirements

- Windows
- Windows PowerShell 5.1 or newer
- Permission to create a scheduled task

## Install

1. Extract the folder anywhere you want to keep it.
2. Run `Install.cmd` as administrator.
3. Enter the target microphone percentage.
4. Choose whether the guard should auto-start at logon.

The installer creates a scheduled task named `MicVolumeGuard` and two desktop shortcuts:

- `Start Mic Volume Guard`
- `Stop Mic Volume Guard`

The guard follows both Windows capture defaults dynamically; there is no device or role selection to maintain after installation.

## Uninstall

Run `Uninstall.cmd` as administrator.

## Direct script usage

You can also run the guard directly:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\MicVolumeGuard.ps1 -TargetPercent 100
```

Useful parameters:

- `-TargetPercent <0-100>`
- `-PollMs <milliseconds>`
- `-TolerancePercent <0-100>`
- `-RestoreAnyChange`
- `-AlsoRestoreMute`
- `-ProcessPriority Normal|AboveNormal|High`
- `-LogPath <path>`

`-Role` is retained only for compatibility with older scheduled tasks and direct invocations. Current MicVolumeGuard versions always follow both the normal default capture device and the default communications capture device.

If `-TargetPercent` is omitted, the script uses the current level of the normal Windows default microphone as its baseline.
