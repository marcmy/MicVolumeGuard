#Requires -Version 5.1

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Read-YesNo {
    param(
        [string]$Prompt,
        [bool]$Default = $true
    )

    $suffix = if ($Default) { 'Y/n' } else { 'y/N' }

    while ($true) {
        $inputValue = Read-Host "$Prompt [$suffix]"
        if ([string]::IsNullOrWhiteSpace($inputValue)) {
            return $Default
        }

        switch ($inputValue.Trim().ToLowerInvariant()) {
            'y' { return $true }
            'yes' { return $true }
            'n' { return $false }
            'no' { return $false }
        }

        Write-Host 'Please enter yes or no.' -ForegroundColor Yellow
    }
}

function Read-AudioRole {
    param(
        [string]$DefaultRole = 'Communications'
    )

    while ($true) {
        $inputValue = Read-Host "Choose audio role [$DefaultRole] (Communications/Console/Multimedia)"
        if ([string]::IsNullOrWhiteSpace($inputValue)) {
            return $DefaultRole
        }

        switch ($inputValue.Trim().ToLowerInvariant()) {
            'communications' { return 'Communications' }
            'comm' { return 'Communications' }
            'console' { return 'Console' }
            'con' { return 'Console' }
            'multimedia' { return 'Multimedia' }
            'multi' { return 'Multimedia' }
        }

        Write-Host 'Please enter Communications, Console, or Multimedia.' -ForegroundColor Yellow
    }
}

$base = Split-Path -Parent $MyInvocation.MyCommand.Path
$scriptPath = Join-Path $base 'MicVolumeGuard.ps1'
$iconPath = Join-Path $base 'microphone.ico'
$defaultLogDir = Join-Path $env:LocalAppData 'MicVolumeGuard'
$defaultLogPath = Join-Path $defaultLogDir 'MicVolumeGuard.log'

if (-not (Test-Path -LiteralPath $scriptPath)) {
    throw "MicVolumeGuard.ps1 not found in: $base"
}

if (-not ('MicVolInstallHelper' -as [type])) {
    Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

public static class MicVolInstallHelper
{
    public enum EDataFlow { eRender, eCapture, eAll, EDataFlow_enum_count }
    public enum ERole { eConsole, eMultimedia, eCommunications, ERole_enum_count }

    [Flags]
    public enum CLSCTX : uint
    {
        INPROC_SERVER = 0x1,
        INPROC_HANDLER = 0x2,
        LOCAL_SERVER = 0x4,
        REMOTE_SERVER = 0x10,
        ALL = INPROC_SERVER | INPROC_HANDLER | LOCAL_SERVER | REMOTE_SERVER
    }

    [ComImport]
    [Guid("BCDE0395-E52F-467C-8E3D-C4579291692E")]
    private class MMDeviceEnumeratorComObject
    {
    }

    [ComImport]
    [Guid("A95664D2-9614-4F35-A746-DE8DB63617E6")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IMMDeviceEnumerator
    {
        int NotImpl1();
        int GetDefaultAudioEndpoint(EDataFlow dataFlow, ERole role, out IMMDevice ppDevice);
    }

    [ComImport]
    [Guid("D666063F-1587-4E43-81F1-B948E807363F")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IMMDevice
    {
        int Activate(ref Guid iid, CLSCTX dwClsCtx, IntPtr pActivationParams, [MarshalAs(UnmanagedType.IUnknown)] out object ppInterface);
        int OpenPropertyStore(int stgmAccess, out object ppProperties);
        int GetId([MarshalAs(UnmanagedType.LPWStr)] out string ppstrId);
        int GetState(out int pdwState);
    }

    [ComImport]
    [Guid("5CDF2C82-841E-4546-9722-0CF74078229A")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IAudioEndpointVolume
    {
        int RegisterControlChangeNotify(IntPtr pNotify);
        int UnregisterControlChangeNotify(IntPtr pNotify);
        int GetChannelCount(out uint pnChannelCount);
        int SetMasterVolumeLevel(float fLevelDB, ref Guid pguidEventContext);
        int SetMasterVolumeLevelScalar(float fLevel, ref Guid pguidEventContext);
        int GetMasterVolumeLevel(out float pfLevelDB);
        int GetMasterVolumeLevelScalar(out float pfLevel);
        int SetChannelVolumeLevel(uint nChannel, float fLevelDB, ref Guid pguidEventContext);
        int SetChannelVolumeLevelScalar(uint nChannel, float fLevel, ref Guid pguidEventContext);
        int GetChannelVolumeLevel(uint nChannel, out float pfLevelDB);
        int GetChannelVolumeLevelScalar(uint nChannel, out float pfLevel);
        int SetMute([MarshalAs(UnmanagedType.Bool)] bool bMute, ref Guid pguidEventContext);
        int GetMute(out bool pbMute);
        int GetVolumeStepInfo(out uint pnStep, out uint pnStepCount);
        int VolumeStepUp(ref Guid pguidEventContext);
        int VolumeStepDown(ref Guid pguidEventContext);
        int QueryHardwareSupport(out uint pdwHardwareSupportMask);
        int GetVolumeRange(out float pflVolumeMindB, out float pflVolumeMaxdB, out float pflVolumeIncrementdB);
    }

    public static int GetMicPercent(ERole role)
    {
        var enumerator = (IMMDeviceEnumerator)(new MMDeviceEnumeratorComObject());
        IMMDevice device;
        Marshal.ThrowExceptionForHR(enumerator.GetDefaultAudioEndpoint(EDataFlow.eCapture, role, out device));

        object endpointObj;
        Guid iid = typeof(IAudioEndpointVolume).GUID;
        Marshal.ThrowExceptionForHR(device.Activate(ref iid, CLSCTX.ALL, IntPtr.Zero, out endpointObj));
        var endpoint = (IAudioEndpointVolume)endpointObj;

        float level;
        Marshal.ThrowExceptionForHR(endpoint.GetMasterVolumeLevelScalar(out level));

        try
        {
            return (int)Math.Round(level * 100.0f);
        }
        finally
        {
            if (endpoint != null) Marshal.ReleaseComObject(endpoint);
            if (device != null) Marshal.ReleaseComObject(device);
            if (enumerator != null) Marshal.ReleaseComObject(enumerator);
        }
    }
}
"@
}

$roleMap = @{
    Console        = [MicVolInstallHelper+ERole]::eConsole
    Multimedia     = [MicVolInstallHelper+ERole]::eMultimedia
    Communications = [MicVolInstallHelper+ERole]::eCommunications
}

$selectedRole = Read-AudioRole -DefaultRole 'Communications'
$defaultPercent = [MicVolInstallHelper]::GetMicPercent($roleMap[$selectedRole])
if ($defaultPercent -lt 0 -or $defaultPercent -gt 100) {
    $defaultPercent = 100
}

while ($true) {
    $inputValue = Read-Host "Enter target mic volume percent [$defaultPercent]"
    if ([string]::IsNullOrWhiteSpace($inputValue)) {
        $targetPercent = $defaultPercent
        break
    }

    $parsed = 0
    if ([int]::TryParse($inputValue, [ref]$parsed) -and $parsed -ge 0 -and $parsed -le 100) {
        $targetPercent = $parsed
        break
    }

    Write-Host 'Please enter a whole number from 0 to 100.' -ForegroundColor Yellow
}

$autoStartAtLogOn = Read-YesNo -Prompt 'Start automatically at logon?' -Default $true

$taskName = 'MicVolumeGuard'
$taskDescription = "Keeps $selectedRole mic volume fixed at $targetPercent% against AGC changes."
$taskArgs = "-NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$scriptPath`" -TargetPercent $targetPercent -PollMs 1000 -TolerancePercent 1 -Role $selectedRole -ProcessPriority High -LogPath `"$defaultLogPath`""

$action = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument $taskArgs
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable -MultipleInstances IgnoreNew
$principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType Interactive -RunLevel Highest
$trigger = $null
if ($autoStartAtLogOn) {
    $trigger = New-ScheduledTaskTrigger -AtLogOn -User $env:USERNAME
}

Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction SilentlyContinue
$registerTaskSplat = @{
    TaskName    = $taskName
    Action      = $action
    Settings    = $settings
    Principal   = $principal
    Description = $taskDescription
}
if ($null -ne $trigger) {
    $registerTaskSplat.Trigger = $trigger
}
Register-ScheduledTask @registerTaskSplat | Out-Null

New-Item -ItemType Directory -Path $defaultLogDir -Force | Out-Null

$desktop = [Environment]::GetFolderPath('Desktop')
$wsh = New-Object -ComObject WScript.Shell

$startShortcut = $wsh.CreateShortcut((Join-Path $desktop 'Start Mic Volume Guard.lnk'))
$startShortcut.TargetPath = "$env:SystemRoot\System32\schtasks.exe"
$startShortcut.Arguments = '/Run /TN "MicVolumeGuard"'
if (Test-Path -LiteralPath $iconPath) {
    $startShortcut.IconLocation = "$iconPath,0"
} else {
    $startShortcut.IconLocation = "$env:SystemRoot\System32\SndVol.exe,0"
}
$startShortcut.Save()

$stopShortcut = $wsh.CreateShortcut((Join-Path $desktop 'Stop Mic Volume Guard.lnk'))
$stopShortcut.TargetPath = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
$stopShortcut.Arguments = '-NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -Command "Stop-ScheduledTask -TaskName ''MicVolumeGuard'' -ErrorAction SilentlyContinue; Get-CimInstance Win32_Process | Where-Object { $_.Name -match ''^(powershell|pwsh)\.exe$'' -and $_.CommandLine -match ''MicVolumeGuard\.ps1'' } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force }"'
if (Test-Path -LiteralPath $iconPath) {
    $stopShortcut.IconLocation = "$iconPath,0"
} else {
    $stopShortcut.IconLocation = "$env:SystemRoot\System32\SndVol.exe,0"
}
$stopShortcut.Save()

Start-ScheduledTask -TaskName $taskName

$autoStartLabel = if ($autoStartAtLogOn) { 'Enabled at logon' } else { 'Manual start only' }

Write-Host ''
Write-Host "Installed successfully. Target mic volume: $targetPercent%" -ForegroundColor Green
Write-Host "Audio role: $selectedRole" -ForegroundColor Green
Write-Host "Auto-start: $autoStartLabel" -ForegroundColor Green
Write-Host "Log file: $defaultLogPath" -ForegroundColor Green
Write-Host 'Desktop shortcuts created:' -ForegroundColor Green
Write-Host '  Start Mic Volume Guard'
Write-Host '  Stop Mic Volume Guard'
Write-Host ''
