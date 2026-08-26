#Requires -Version 5.1

param(
    [Nullable[int]]$TargetPercent = 100,

    # Retained for compatibility with older scheduled tasks and direct invocations.
    # MicVolumeGuard now always follows both the normal default capture device
    # (Console) and the default communications capture device.
    [ValidateSet('Console', 'Multimedia', 'Communications')]
    [string]$Role = 'Console',

    [ValidateRange(50, 5000)]
    [int]$PollMs = 1000,

    [ValidateRange(0, 100)]
    [int]$TolerancePercent = 1,

    [switch]$RestoreAnyChange,
    [switch]$AlsoRestoreMute,

    [ValidateSet('Normal', 'AboveNormal', 'High')]
    [string]$ProcessPriority = 'High',

    [string]$LogPath = (Join-Path (Join-Path $env:LocalAppData 'MicVolumeGuard') 'MicVolumeGuard.log')
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$script:GuardLogPath = $LogPath

function Write-GuardLog {
    param(
        [string]$Message,
        [ValidateSet('INFO', 'WARN', 'ERROR')]
        [string]$Level = 'INFO'
    )

    if ([string]::IsNullOrWhiteSpace($script:GuardLogPath)) {
        return
    }

    try {
        $logDirectory = Split-Path -Parent $script:GuardLogPath
        if (-not [string]::IsNullOrWhiteSpace($logDirectory)) {
            New-Item -ItemType Directory -Path $logDirectory -Force | Out-Null
        }

        $timestamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
        Add-Content -LiteralPath $script:GuardLogPath -Value "[$timestamp] [$Level] $Message"
    }
    catch {
        return
    }
}

try {
    $currentProcess = Get-Process -Id $PID -ErrorAction Stop
    switch ($ProcessPriority) {
        'Normal'      { $currentProcess.PriorityClass = [System.Diagnostics.ProcessPriorityClass]::Normal }
        'AboveNormal' { $currentProcess.PriorityClass = [System.Diagnostics.ProcessPriorityClass]::AboveNormal }
        'High'        { $currentProcess.PriorityClass = [System.Diagnostics.ProcessPriorityClass]::High }
    }
}
catch {
    Write-GuardLog "Unable to set process priority to $ProcessPriority." 'WARN'
}

$mutexName = 'Local\MicVolumeGuard_Global'
$createdNew = $false
$script:Mutex = [System.Threading.Mutex]::new($true, $mutexName, [ref]$createdNew)

if (-not $createdNew) {
    exit
}

try {
    if (-not ('CoreAudioMicGuard' -as [type])) {
        Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

public static class CoreAudioMicGuard
{
    public enum EDataFlow
    {
        eRender,
        eCapture,
        eAll,
        EDataFlow_enum_count
    }

    public enum ERole
    {
        eConsole,
        eMultimedia,
        eCommunications,
        ERole_enum_count
    }

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
        int GetDevice([MarshalAs(UnmanagedType.LPWStr)] string pwstrId, out IMMDevice ppDevice);
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

    private static IMMDeviceEnumerator CreateEnumerator()
    {
        return (IMMDeviceEnumerator)(new MMDeviceEnumeratorComObject());
    }

    private static IAudioEndpointVolume GetEndpointVolume(IMMDevice device)
    {
        Guid iid = typeof(IAudioEndpointVolume).GUID;
        object endpointObj;
        Marshal.ThrowExceptionForHR(device.Activate(ref iid, CLSCTX.ALL, IntPtr.Zero, out endpointObj));
        return (IAudioEndpointVolume)endpointObj;
    }

    public static string GetDefaultCaptureDeviceId(ERole role)
    {
        IMMDeviceEnumerator enumerator = null;
        IMMDevice device = null;

        try
        {
            enumerator = CreateEnumerator();
            Marshal.ThrowExceptionForHR(enumerator.GetDefaultAudioEndpoint(EDataFlow.eCapture, role, out device));

            string id;
            Marshal.ThrowExceptionForHR(device.GetId(out id));
            return id;
        }
        finally
        {
            if (device != null) Marshal.ReleaseComObject(device);
            if (enumerator != null) Marshal.ReleaseComObject(enumerator);
        }
    }

    private static IMMDevice GetDeviceById(string deviceId, out IMMDeviceEnumerator enumerator)
    {
        enumerator = CreateEnumerator();
        IMMDevice device;
        Marshal.ThrowExceptionForHR(enumerator.GetDevice(deviceId, out device));
        return device;
    }

    public static float GetCaptureVolumeScalar(string deviceId)
    {
        IMMDeviceEnumerator enumerator = null;
        IMMDevice device = null;
        IAudioEndpointVolume endpoint = null;

        try
        {
            device = GetDeviceById(deviceId, out enumerator);
            endpoint = GetEndpointVolume(device);

            float level;
            Marshal.ThrowExceptionForHR(endpoint.GetMasterVolumeLevelScalar(out level));
            return level;
        }
        finally
        {
            if (endpoint != null) Marshal.ReleaseComObject(endpoint);
            if (device != null) Marshal.ReleaseComObject(device);
            if (enumerator != null) Marshal.ReleaseComObject(enumerator);
        }
    }

    public static void SetCaptureVolumeScalar(string deviceId, float scalar)
    {
        IMMDeviceEnumerator enumerator = null;
        IMMDevice device = null;
        IAudioEndpointVolume endpoint = null;

        try
        {
            device = GetDeviceById(deviceId, out enumerator);
            endpoint = GetEndpointVolume(device);

            float clamped = Math.Max(0.0f, Math.Min(1.0f, scalar));
            Guid context = Guid.Empty;
            Marshal.ThrowExceptionForHR(endpoint.SetMasterVolumeLevelScalar(clamped, ref context));
        }
        finally
        {
            if (endpoint != null) Marshal.ReleaseComObject(endpoint);
            if (device != null) Marshal.ReleaseComObject(device);
            if (enumerator != null) Marshal.ReleaseComObject(enumerator);
        }
    }

    public static bool GetCaptureMute(string deviceId)
    {
        IMMDeviceEnumerator enumerator = null;
        IMMDevice device = null;
        IAudioEndpointVolume endpoint = null;

        try
        {
            device = GetDeviceById(deviceId, out enumerator);
            endpoint = GetEndpointVolume(device);

            bool muted;
            Marshal.ThrowExceptionForHR(endpoint.GetMute(out muted));
            return muted;
        }
        finally
        {
            if (endpoint != null) Marshal.ReleaseComObject(endpoint);
            if (device != null) Marshal.ReleaseComObject(device);
            if (enumerator != null) Marshal.ReleaseComObject(enumerator);
        }
    }

    public static void SetCaptureMute(string deviceId, bool muted)
    {
        IMMDeviceEnumerator enumerator = null;
        IMMDevice device = null;
        IAudioEndpointVolume endpoint = null;

        try
        {
            device = GetDeviceById(deviceId, out enumerator);
            endpoint = GetEndpointVolume(device);

            Guid context = Guid.Empty;
            Marshal.ThrowExceptionForHR(endpoint.SetMute(muted, ref context));
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

    $consoleRole = [CoreAudioMicGuard+ERole]::eConsole
    $communicationsRole = [CoreAudioMicGuard+ERole]::eCommunications

    if ($null -eq $TargetPercent) {
        $defaultDeviceId = [CoreAudioMicGuard]::GetDefaultCaptureDeviceId($consoleRole)
        $targetScalar = [double][CoreAudioMicGuard]::GetCaptureVolumeScalar($defaultDeviceId)
        $targetPercentLabel = [int][math]::Round($targetScalar * 100.0)
    }
    else {
        if ($TargetPercent -lt 0 -or $TargetPercent -gt 100) {
            throw 'TargetPercent must be between 0 and 100.'
        }
        $targetScalar = [double]$TargetPercent / 100.0
        $targetPercentLabel = $TargetPercent
    }

    $toleranceScalar = [double]$TolerancePercent / 100.0
    $lastLoopErrorMessage = $null
    $lastLoopErrorAt = [datetime]::MinValue
    $lastConsoleDeviceId = $null
    $lastCommunicationsDeviceId = $null
    $targetMuteByDevice = @{}

    Write-GuardLog "Starting MicVolumeGuard. TargetPercent=$targetPercentLabel FollowDefault=True FollowCommunications=True LegacyRole=$Role PollMs=$PollMs TolerancePercent=$TolerancePercent RestoreAnyChange=$($RestoreAnyChange.IsPresent) AlsoRestoreMute=$($AlsoRestoreMute.IsPresent)" 'INFO'

    while ($true) {
        try {
            $consoleDeviceId = [CoreAudioMicGuard]::GetDefaultCaptureDeviceId($consoleRole)
            $communicationsDeviceId = [CoreAudioMicGuard]::GetDefaultCaptureDeviceId($communicationsRole)

            if ($consoleDeviceId -ne $lastConsoleDeviceId -or $communicationsDeviceId -ne $lastCommunicationsDeviceId) {
                $sameDevice = $consoleDeviceId -eq $communicationsDeviceId
                Write-GuardLog "Capture endpoints changed. Default=$consoleDeviceId Communications=$communicationsDeviceId SameDevice=$sameDevice" 'INFO'
                $lastConsoleDeviceId = $consoleDeviceId
                $lastCommunicationsDeviceId = $communicationsDeviceId
            }

            $deviceIds = @($consoleDeviceId, $communicationsDeviceId) | Select-Object -Unique

            foreach ($deviceId in $deviceIds) {
                $currentScalar = [double][CoreAudioMicGuard]::GetCaptureVolumeScalar($deviceId)
                $delta = [math]::Abs($currentScalar - $targetScalar)

                $needsRestore = if ($RestoreAnyChange) {
                    $delta -gt $toleranceScalar
                }
                else {
                    $currentScalar -lt ($targetScalar - $toleranceScalar)
                }

                if ($needsRestore) {
                    [CoreAudioMicGuard]::SetCaptureVolumeScalar($deviceId, [float]$targetScalar)
                }

                if ($AlsoRestoreMute) {
                    if (-not $targetMuteByDevice.ContainsKey($deviceId)) {
                        $targetMuteByDevice[$deviceId] = [CoreAudioMicGuard]::GetCaptureMute($deviceId)
                    }

                    $currentMute = [CoreAudioMicGuard]::GetCaptureMute($deviceId)
                    if ($currentMute -ne $targetMuteByDevice[$deviceId]) {
                        [CoreAudioMicGuard]::SetCaptureMute($deviceId, [bool]$targetMuteByDevice[$deviceId])
                    }
                }
            }

            if ($null -ne $lastLoopErrorMessage) {
                Write-GuardLog 'Capture device access recovered.' 'INFO'
                $lastLoopErrorMessage = $null
                $lastLoopErrorAt = [datetime]::MinValue
            }
        }
        catch {
            $loopErrorMessage = $_.Exception.Message
            if ([string]::IsNullOrWhiteSpace($loopErrorMessage)) {
                $loopErrorMessage = $_.Exception.GetType().FullName
            }

            $now = Get-Date
            if ($loopErrorMessage -ne $lastLoopErrorMessage -or ($now - $lastLoopErrorAt).TotalMinutes -ge 5) {
                Write-GuardLog "Loop error: $loopErrorMessage" 'WARN'
                $lastLoopErrorMessage = $loopErrorMessage
                $lastLoopErrorAt = $now
            }
        }

        Start-Sleep -Milliseconds $PollMs
    }
}
catch {
    $fatalMessage = $_.Exception.Message
    if ([string]::IsNullOrWhiteSpace($fatalMessage)) {
        $fatalMessage = $_.Exception.GetType().FullName
    }

    Write-GuardLog "Fatal error: $fatalMessage" 'ERROR'
    throw
}
finally {
    if ($null -ne $script:Mutex) {
        try { $script:Mutex.ReleaseMutex() } catch { $null = $_ }
        try { $script:Mutex.Dispose() } catch { $null = $_ }
    }
}
