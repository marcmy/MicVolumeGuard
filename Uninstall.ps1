#Requires -Version 5.1

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$taskName = 'MicVolumeGuard'

try {
    Stop-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
}
catch {
}

Get-CimInstance Win32_Process |
Where-Object {
    $_.Name -match '^(powershell|pwsh)\.exe$' -and
    $_.CommandLine -match 'MicVolumeGuard\.ps1'
} |
ForEach-Object {
    Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue
}

try {
    Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction SilentlyContinue
}
catch {
}

$desktop = [Environment]::GetFolderPath('Desktop')
Remove-Item (Join-Path $desktop 'Start Mic Volume Guard.lnk') -Force -ErrorAction SilentlyContinue
Remove-Item (Join-Path $desktop 'Stop Mic Volume Guard.lnk') -Force -ErrorAction SilentlyContinue