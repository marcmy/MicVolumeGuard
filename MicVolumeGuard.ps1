#Requires -Version 5.1

param(
    [Nullable[int]]$TargetPercent = 100,

    [ValidateSet('Console', 'Multimedia', 'Communications')]
    [string]$Role = 'Communications',

    [ValidateRange(50, 5000)]
    [int]$PollMs = 1000,

    [ValidateRange(0, 100)]
    [int]$TolerancePercent = 1,

    [switch]$RestoreAnyChange,
    [switch]$AlsoRestoreMute,

    [ValidateSet('Normal', 'AboveNormal', 'High')]
    [string]$ProcessPriority = 'High'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

try {
    $currentProcess = Get-Process -Id $PID -ErrorAction Stop
    switch ($ProcessPriority) {
        'Normal'      { $currentProcess.PriorityClass = [System.Diagnostics.ProcessPriorityClass]::Normal }
        'AboveNormal' { $currentProcess.PriorityClass = [System.Diagnostics.ProcessPriorityClass]::AboveNormal }
        'High'        { $currentProcess.PriorityClass = [System.Diagnostics.ProcessPriorityClass]::High }
    }
} catch {}

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

    private static IMMDevice GetDefaultCaptureDevice(ERole role, out IMMDeviceEnumerator enumerator)
    {
        enumerator = (IMMDeviceEnumerator)(new MMDeviceEnumeratorComObject());
        IMMDevice device;
        Marshal.ThrowExceptionForHR(enumerator.GetDefaultAudioEndpoint(EDataFlow.eCapture, role, out device));
        return device;
    }

    private static IAudioEndpointVolume GetEndpointVolume(IMMDevice device)
    {
        Guid iid = typeof(IAudioEndpointVolume).GUID;
        object endpointObj;
        Marshal.ThrowExceptionForHR(device.Activate(ref iid, CLSCTX.ALL, IntPtr.Zero, out endpointObj));
        return (IAudioEndpointVolume)endpointObj;
    }

    public static float GetDefaultCaptureVolumeScalar(ERole role)
    {
        IMMDeviceEnumerator enumerator = null;
        IMMDevice device = null;
        IAudioEndpointVolume endpoint = null;

        try
        {
            device = GetDefaultCaptureDevice(role, out enumerator);
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

    public static void SetDefaultCaptureVolumeScalar(ERole role, float scalar)
    {
        IMMDeviceEnumerator enumerator = null;
        IMMDevice device = null;
        IAudioEndpointVolume endpoint = null;

        try
        {
            device = GetDefaultCaptureDevice(role, out enumerator);
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

    public static bool GetDefaultCaptureMute(ERole role)
    {
        IMMDeviceEnumerator enumerator = null;
        IMMDevice device = null;
        IAudioEndpointVolume endpoint = null;

        try
        {
            device = GetDefaultCaptureDevice(role, out enumerator);
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

    public static void SetDefaultCaptureMute(ERole role, bool muted)
    {
        IMMDeviceEnumerator enumerator = null;
        IMMDevice device = null;
        IAudioEndpointVolume endpoint = null;

        try
        {
            device = GetDefaultCaptureDevice(role, out enumerator);
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

    $roleMap = @{
        Console        = [CoreAudioMicGuard+ERole]::eConsole
        Multimedia     = [CoreAudioMicGuard+ERole]::eMultimedia
        Communications = [CoreAudioMicGuard+ERole]::eCommunications
    }

    $roleEnum = $roleMap[$Role]

    if ($null -eq $TargetPercent) {
        $targetScalar = [double][CoreAudioMicGuard]::GetDefaultCaptureVolumeScalar($roleEnum)
    } else {
        if ($TargetPercent -lt 0 -or $TargetPercent -gt 100) {
            throw 'TargetPercent must be between 0 and 100.'
        }
        $targetScalar = [double]$TargetPercent / 100.0
    }

    $targetMute = $false
    if ($AlsoRestoreMute) {
        $targetMute = [CoreAudioMicGuard]::GetDefaultCaptureMute($roleEnum)
    }

    $toleranceScalar = [double]$TolerancePercent / 100.0

    while ($true) {
        try {
            $currentScalar = [double][CoreAudioMicGuard]::GetDefaultCaptureVolumeScalar($roleEnum)
            $delta = [math]::Abs($currentScalar - $targetScalar)

            $needsRestore = $false

            if ($RestoreAnyChange) {
                if ($delta -gt $toleranceScalar) {
                    $needsRestore = $true
                }
            } else {
                if ($currentScalar -lt ($targetScalar - $toleranceScalar)) {
                    $needsRestore = $true
                }
            }

            if ($needsRestore) {
                [CoreAudioMicGuard]::SetDefaultCaptureVolumeScalar($roleEnum, [float]$targetScalar)
            }

            if ($AlsoRestoreMute) {
                $currentMute = [CoreAudioMicGuard]::GetDefaultCaptureMute($roleEnum)
                if ($currentMute -ne $targetMute) {
                    [CoreAudioMicGuard]::SetDefaultCaptureMute($roleEnum, $targetMute)
                }
            }
        } catch {}

        Start-Sleep -Milliseconds $PollMs
    }
}
finally {
    if ($null -ne $script:Mutex) {
        try { $script:Mutex.ReleaseMutex() } catch {}
        try { $script:Mutex.Dispose() } catch {}
    }
}
