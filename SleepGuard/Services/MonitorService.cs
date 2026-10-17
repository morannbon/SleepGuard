using System.Collections.Concurrent;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Timers;

namespace SleepGuard;

/// <summary>
/// プロセス監視とスリープ防止を担当するサービス。
///
/// 監視対象が実行中の間は Windows の電源要求を保持する。
/// 監視対象がすべて停止した時点で要求を解除し、Windows 標準のスリープ設定に委ねる。
/// </summary>
public class MonitorService : IDisposable
{
    // ── Win32 API ────────────────────────────────────────────────
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct ReasonContext
    {
        public uint Version;
        public uint Flags;
        public IntPtr SimpleReasonString;
    }

    private enum PowerRequestType
    {
        // POWER_REQUEST_TYPE (WinNT.h)
        // 0 = PowerRequestDisplayRequired
        // 1 = PowerRequestSystemRequired
        DisplayRequired = 0,
        SystemRequired = 1
    }

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern IntPtr PowerCreateRequest(ref ReasonContext context);

    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool PowerSetRequest(IntPtr powerRequest, PowerRequestType requestType);

    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool PowerClearRequest(IntPtr powerRequest, PowerRequestType requestType);

    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool CloseHandle(IntPtr handle);

    private const uint PowerRequestContextVersion = 0;
    private const uint PowerRequestContextSimpleString = 0x1;
    private static readonly IntPtr InvalidHandleValue = new(-1);

    // ── 電源要求ハンドル ──────────────────────────────────────────
    private IntPtr _powerRequestHandle = IntPtr.Zero;
    private bool _systemRequestHeld;
    private bool _displayRequestHeld;

    // ── 状態 ─────────────────────────────────────────────────────
    private AppSettings _settings = SettingsManager.Load();
    private readonly object _settingsLock = new();
    private readonly object _stateLock = new();

    private System.Timers.Timer? _timer;
    private bool _isMonitoring;
    private bool      _isSleepPrevented;
    private List<string> _lastRunningProcs = new();
    private bool _disposed;
    private int _tickCounter;
    private DateTime? _lastDiagnosticLogAt;
    private readonly object _matchLogLock = new();
    private readonly Dictionary<string, string> _lastMatchDiagnostics = new(StringComparer.OrdinalIgnoreCase);

    // ── イベント ─────────────────────────────────────────────────
    public event Action<bool>?         StatusChanged;
    public event Action<MonitorStatus>? StatusUpdated;

    // ── パブリック ────────────────────────────────────────────────
    public bool IsSleepPrevented => _isSleepPrevented;

    public MonitorService() { }

    public void Start()
    {
        lock (_stateLock)
        {
            ThrowIfDisposed();

            StopTimerCore();
            _isMonitoring = true;

            // 監視開始直後に正しい状態へ同期し、最初のタイマー待ち中の誤動作を防ぐ。
            TickCore(forceStatus: true);

            var intervalMs = GetIntervalMs();
            _timer = new System.Timers.Timer(intervalMs)
            {
                AutoReset = true
            };
            _timer.Elapsed += OnTimerElapsed;
            _timer.Start();

            SettingsManager.WriteLog($"SleepGuard 監視開始 version={System.Reflection.Assembly.GetExecutingAssembly().GetName().Version} intervalMs={intervalMs:0} watched={_settings.WatchedProcesses.Count} preventDisplay={_settings.PreventDisplaySleep}");
        }
    }

    public void Stop()
    {
        lock (_stateLock)
        {
            StopTimerCore();
            _isMonitoring = false;
            ApplySleepState(false, null, "monitor-stop");
            PublishCurrentStatus(force: true);
            SettingsManager.WriteLog("SleepGuard 監視停止");
        }
    }

    /// <summary>設定変更時に呼ぶ。タイマー間隔も即時反映する。</summary>
    public void ReloadSettings()
    {
        var settings = SettingsManager.Load();

        lock (_stateLock)
        {
            lock (_settingsLock) { _settings = settings; }

            if (_timer != null)
                _timer.Interval = GetIntervalMs();

            // 設定保存直後に監視対象と表示スリープ設定を実状態へ反映する。
            if (_isMonitoring)
                TickCore(forceStatus: true);
            else
                PublishCurrentStatus(force: true);

            SettingsManager.WriteLog($"設定再読込 intervalMs={GetIntervalMs():0} watched={settings.WatchedProcesses.Count} preventDisplay={settings.PreventDisplaySleep}");
        }
    }

    // ── タイマー ─────────────────────────────────────────────────
    private void OnTimerElapsed(object? sender, ElapsedEventArgs e)
    {
        lock (_stateLock)
        {
            if (!_isMonitoring || _disposed)
                return;

            TickCore(forceStatus: false);
        }
    }

    private void TickCore(bool forceStatus)
    {
        AppSettings settings;
        lock (_settingsLock) { settings = _settings; }

        var now = DateTime.Now;
        _tickCounter++;

        if (settings.WatchedProcesses.Count == 0)
        {
            ApplySleepState(false, null, "watched-empty");

            var empty = new List<string>();
            WriteDiagnosticSnapshotIfNeeded(now, settings, empty, "watched-empty");
            PublishStatusIfChanged(settings, empty, now, true);
            return;
        }

        var runningProcs = GetRunningWatchedProcesses(settings);
        var anyRunning = runningProcs.Count > 0;
        var procsChanged = !runningProcs.SequenceEqual(_lastRunningProcs);

        if (procsChanged)
        {
            SettingsManager.WriteLog($"実行中プロセス変化 running=[{string.Join(", ", runningProcs)}] previous=[{string.Join(", ", _lastRunningProcs)}]");
        }

        var shouldBlock = anyRunning;
        var blockStateChanged = ApplySleepState(
            shouldBlock,
            shouldBlock ? settings : null,
            shouldBlock ? "watched-running" : "watched-stopped",
            runningProcs);

        WriteDiagnosticSnapshotIfNeeded(now, settings, runningProcs, blockStateChanged ? "block-state-changed" : (procsChanged ? "running-changed" : null));
        PublishStatusIfChanged(settings, runningProcs, now, forceStatus || procsChanged || blockStateChanged);
    }

    // ── プロセス検索 ─────────────────────────────────────────────
    /// <summary>
    /// 監視対象のうち実行中のプロセス名一覧を返す。
    /// Process 配列は使用後即 Dispose してハンドルを解放する。
    /// </summary>
    public List<string> GetRunningWatchedProcesses(AppSettings? settings = null)
    {
        AppSettings s;
        if (settings != null)
            s = settings;
        else
            lock (_settingsLock) { s = _settings; }

        var running = new List<string>(s.WatchedProcesses.Count);
        foreach (var wp in s.WatchedProcesses)
        {
            Process[]? procs = null;
            try
            {
                procs = Process.GetProcessesByName(wp.ProcessName);
                if (TryMatchWatchedProcess(wp, procs, out var diagnostics))
                {
                    running.Add(wp.DisplayName);
                    LogMatchIfChanged(wp, diagnostics);
                }
                else
                {
                    lock (_matchLogLock)
                        _lastMatchDiagnostics.Remove(wp.Id);
                }
            }
            catch (Exception ex)
            {
                SettingsManager.WriteLog($"監視列挙失敗 display={wp.DisplayName} process={wp.ProcessName} error={ex.GetType().Name}:{ex.Message}");
            }
            finally
            {
                if (procs != null)
                    foreach (var p in procs) p.Dispose();
            }
        }
        return running;
    }

    private void LogMatchIfChanged(WatchedProcess wp, string diagnostics)
    {
        lock (_matchLogLock)
        {
            if (_lastMatchDiagnostics.TryGetValue(wp.Id, out var previous) &&
                string.Equals(previous, diagnostics, StringComparison.Ordinal))
                return;

            _lastMatchDiagnostics[wp.Id] = diagnostics;
        }

        SettingsManager.WriteLog($"監視一致 {diagnostics}");
    }

    private static bool TryMatchWatchedProcess(WatchedProcess wp, Process[] processes, out string diagnostics)
    {
        diagnostics = $"display={wp.DisplayName} process={wp.ProcessName} candidates={processes.Length}";

        if (processes.Length == 0)
            return false;

        var expectedPath = NormalizePath(wp.FilePath);
        if (string.IsNullOrWhiteSpace(expectedPath))
        {
            diagnostics += " match=process-name-only";
            return true;
        }

        foreach (var process in processes)
        {
            try
            {
                var actualPath = NormalizePath(process.MainModule?.FileName);
                if (string.Equals(actualPath, expectedPath, StringComparison.OrdinalIgnoreCase))
                {
                    diagnostics += $" pid={process.Id} path={actualPath} match=path";
                    return true;
                }
            }
            catch (Exception ex)
            {
                diagnostics += $" pid={process.Id} path-read-error={ex.GetType().Name}";
            }
        }

        diagnostics += $" expectedPath={expectedPath} match=none";
        return false;
    }

    private static string NormalizePath(string? path)
    {
        if (string.IsNullOrWhiteSpace(path))
            return string.Empty;

        try
        {
            return Path.GetFullPath(path.Trim().Trim('"'));
        }
        catch
        {
            return path.Trim().Trim('"');
        }
    }

    // ── スリープ制御 ─────────────────────────────────────────────
    /// <summary>
    /// Windows API の適用完了後だけ公開状態を更新する。
    /// 監視判定、停止、設定変更の全経路はこのメソッドを通す。
    /// </summary>
    private bool ApplySleepState(
        bool shouldPrevent,
        AppSettings? settings,
        string reason,
        List<string>? runningProcs = null)
    {
        if (shouldPrevent)
        {
            if (!EnsurePowerRequestHandle())
                return false;

            var changed = false;
            if (!_systemRequestHeld)
            {
                if (!PowerSetRequest(_powerRequestHandle, PowerRequestType.SystemRequired))
                {
                    SettingsManager.WriteLog($"スリープ制御失敗 action=hold-system error={Marshal.GetLastWin32Error()} reason={reason}");
                    return false;
                }

                _systemRequestHeld = true;
                changed = true;
                SettingsManager.WriteLog($"スリープ制御 action=hold-system requestType={(int)PowerRequestType.SystemRequired} handle=0x{_powerRequestHandle.ToInt64():X} success=True reason={reason} running=[{string.Join(", ", runningProcs ?? new List<string>())}]");
            }

            var shouldHoldDisplay = settings?.PreventDisplaySleep == true;
            if (shouldHoldDisplay && !_displayRequestHeld)
            {
                if (!PowerSetRequest(_powerRequestHandle, PowerRequestType.DisplayRequired))
                {
                    SettingsManager.WriteLog($"スリープ制御失敗 action=hold-display error={Marshal.GetLastWin32Error()} reason={reason}");
                    return false;
                }

                _displayRequestHeld = true;
                changed = true;
                SettingsManager.WriteLog($"スリープ制御 action=hold-display requestType={(int)PowerRequestType.DisplayRequired} handle=0x{_powerRequestHandle.ToInt64():X} success=True reason={reason}");
            }
            else if (!shouldHoldDisplay && _displayRequestHeld)
            {
                if (!PowerClearRequest(_powerRequestHandle, PowerRequestType.DisplayRequired))
                {
                    SettingsManager.WriteLog($"スリープ制御失敗 action=release-display error={Marshal.GetLastWin32Error()} reason={reason}");
                    return false;
                }

                _displayRequestHeld = false;
                changed = true;
                SettingsManager.WriteLog($"スリープ制御 action=release-display requestType={(int)PowerRequestType.DisplayRequired} handle=0x{_powerRequestHandle.ToInt64():X} success=True reason={reason}");
            }

            if (!_isSleepPrevented)
            {
                _isSleepPrevented = true;
                StatusChanged?.Invoke(true);
                changed = true;
            }

            return changed;
        }

        var released = ReleasePowerRequests(reason);
        if (_isSleepPrevented && released)
        {
            _isSleepPrevented = false;
            StatusChanged?.Invoke(false);
            return true;
        }

        return false;
    }

    private bool EnsurePowerRequestHandle()
    {
        if (_powerRequestHandle != IntPtr.Zero && _powerRequestHandle != InvalidHandleValue)
            return true;

        var reasonPtr = Marshal.StringToHGlobalUni("SleepGuard monitored process is running");
        try
        {
            var context = new ReasonContext
            {
                Version = PowerRequestContextVersion,
                Flags = PowerRequestContextSimpleString,
                SimpleReasonString = reasonPtr
            };

            _powerRequestHandle = PowerCreateRequest(ref context);
            if (_powerRequestHandle == IntPtr.Zero || _powerRequestHandle == InvalidHandleValue)
            {
                SettingsManager.WriteLog($"スリープ制御失敗 action=create-request error={Marshal.GetLastWin32Error()}");
                _powerRequestHandle = IntPtr.Zero;
                return false;
            }

            SettingsManager.WriteLog($"スリープ制御 action=create-request success=True handle=0x{_powerRequestHandle.ToInt64():X}");
            return true;
        }
        finally
        {
            Marshal.FreeHGlobal(reasonPtr);
        }
    }

    private bool ReleasePowerRequests(string reason)
    {
        var succeeded = true;

        if (_powerRequestHandle == IntPtr.Zero || _powerRequestHandle == InvalidHandleValue)
            return true;

        if (_displayRequestHeld)
        {
            if (PowerClearRequest(_powerRequestHandle, PowerRequestType.DisplayRequired))
            {
                _displayRequestHeld = false;
                SettingsManager.WriteLog($"スリープ制御 action=release-display requestType={(int)PowerRequestType.DisplayRequired} handle=0x{_powerRequestHandle.ToInt64():X} success=True reason={reason}");
            }
            else
            {
                succeeded = false;
                SettingsManager.WriteLog($"スリープ制御失敗 action=release-display error={Marshal.GetLastWin32Error()} reason={reason}");
            }
        }

        if (_systemRequestHeld)
        {
            if (PowerClearRequest(_powerRequestHandle, PowerRequestType.SystemRequired))
            {
                _systemRequestHeld = false;
                SettingsManager.WriteLog($"スリープ制御 action=release-system requestType={(int)PowerRequestType.SystemRequired} handle=0x{_powerRequestHandle.ToInt64():X} success=True reason={reason}");
            }
            else
            {
                succeeded = false;
                SettingsManager.WriteLog($"スリープ制御失敗 action=release-system error={Marshal.GetLastWin32Error()} reason={reason}");
            }
        }

        return succeeded;
    }

    private void ClosePowerRequestHandle()
    {
        if (_powerRequestHandle == IntPtr.Zero || _powerRequestHandle == InvalidHandleValue)
            return;

        var handle = _powerRequestHandle;
        var closed = CloseHandle(handle);
        SettingsManager.WriteLog($"スリープ制御 action=close-request handle=0x{handle.ToInt64():X} success={closed} error={(closed ? 0 : Marshal.GetLastWin32Error())}");
        _powerRequestHandle = IntPtr.Zero;
    }

    private void PublishStatusIfChanged(AppSettings settings, List<string> runningProcs, DateTime now, bool force)
    {
        if (!force)
            return;

        _lastRunningProcs = new List<string>(runningProcs);
        StatusUpdated?.Invoke(new MonitorStatus
        {
            IsSleepPrevented = _isSleepPrevented,
            RunningProcesses = new List<string>(runningProcs),
            WatchedCount         = settings.WatchedProcesses.Count,
            CheckIntervalSeconds = settings.CheckIntervalSeconds,
            CheckedAt            = now
        });
    }

    private void WriteDiagnosticSnapshotIfNeeded(DateTime now, AppSettings settings, List<string> runningProcs, string? reason)
    {
        var shouldWrite = reason != null
                          || _tickCounter <= 3
                          || (_tickCounter % 30 == 0)
                          || !_lastDiagnosticLogAt.HasValue
                          || (now - _lastDiagnosticLogAt.Value) >= TimeSpan.FromMinutes(5);
        if (!shouldWrite)
            return;

        _lastDiagnosticLogAt = now;
        SettingsManager.WriteLog(
            $"監視診断 reason={reason ?? "periodic"} tick={_tickCounter} watched={settings.WatchedProcesses.Count} running={runningProcs.Count} blocked={_isSleepPrevented} mode=power-request-hold watchedList=[{string.Join(" | ", settings.WatchedProcesses.Select(w => $"{w.DisplayName}:{w.ProcessName}:{NormalizePath(w.FilePath)}"))}] runningList=[{string.Join(", ", runningProcs)}]");
    }

    // ── 現在状態スナップショット ──────────────────────────────────
    /// <summary>ウィンドウ再表示時の UI 初期化に使用する。</summary>
    public MonitorStatus GetCurrentStatus()
    {
        AppSettings settings;
        lock (_settingsLock) { settings = _settings; }
        var runningProcs = GetRunningWatchedProcesses(settings);
        return new MonitorStatus
        {
            IsSleepPrevented = _isSleepPrevented,
            RunningProcesses = runningProcs,
            WatchedCount         = settings.WatchedProcesses.Count,
            CheckIntervalSeconds = settings.CheckIntervalSeconds,
            CheckedAt            = DateTime.Now
        };
    }

    // ── ヘルパー ─────────────────────────────────────────────────
    private double GetIntervalMs()
    {
        lock (_settingsLock)
            return Math.Clamp(_settings.CheckIntervalSeconds, 10, 60) * 1000.0;
    }

    private void StopTimerCore()
    {
        if (_timer == null)
            return;

        _timer.Stop();
        _timer.Elapsed -= OnTimerElapsed;
        _timer.Dispose();
        _timer = null;
    }

    private void PublishCurrentStatus(bool force)
    {
        AppSettings settings;
        lock (_settingsLock) { settings = _settings; }
        var running = _isMonitoring ? GetRunningWatchedProcesses(settings) : new List<string>();
        PublishStatusIfChanged(settings, running, DateTime.Now, force);
    }

    private void ThrowIfDisposed()
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
    }

    public void Dispose()
    {
        lock (_stateLock)
        {
            if (_disposed) return;
            StopTimerCore();
            _isMonitoring = false;
            ApplySleepState(false, null, "dispose");
            ClosePowerRequestHandle();
            _disposed = true;
        }

        GC.SuppressFinalize(this);
    }
}

/// <summary>監視状態のスナップショット</summary>
public record MonitorStatus
{
    public bool          IsSleepPrevented { get; init; }
    public List<string>  RunningProcesses { get; init; } = new();
    public int           WatchedCount         { get; init; }
    public int           CheckIntervalSeconds { get; init; }
    public DateTime      CheckedAt            { get; init; }
}
