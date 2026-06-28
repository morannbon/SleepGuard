using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Timers;

namespace SleepGuard;

/// <summary>
/// プロセス監視とスリープ防止を担当するサービス。
///
/// 10秒tickごとに監視対象プロセスを確認し、
/// 動作中であれば SendInput でマウス移動量0を送りWindowsのアイドルタイマーをリセットする。
/// プロセスが消えたら何もしないのでWindowsのスリープタイマーが自然に動き出す。
/// </summary>
public class MonitorService : IDisposable
{
    // ── Win32 API ────────────────────────────────────────────────
    [DllImport("user32.dll", SetLastError = true)]
    private static extern uint SendInput(uint nInputs, INPUT[] pInputs, int cbSize);

    [StructLayout(LayoutKind.Sequential)]
    private struct INPUT
    {
        public uint type;       // 0 = MOUSE
        public MOUSEINPUT mi;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct MOUSEINPUT
    {
        public int    dx;
        public int    dy;
        public uint   mouseData;
        public uint   dwFlags;  // 0x0001 = MOUSEEVENTF_MOVE
        public uint   time;
        public IntPtr dwExtraInfo;
    }

    // マウス移動量0の入力（staticで使い回し、アロケーションなし）
    private static readonly INPUT[] _mouseInput = new[]
    {
        new INPUT { type = 0, mi = new MOUSEINPUT { dwFlags = 0x0001 } }
    };
    private static readonly int _inputSize = Marshal.SizeOf<INPUT>();

    // ── 状態 ─────────────────────────────────────────────────────
    private AppSettings  _settings = SettingsManager.Load();
    private readonly object _settingsLock = new();
    private readonly object _timerLock   = new();

    private System.Timers.Timer? _timer;
    private bool         _isSleepPrevented;
    private List<string> _lastRunningProcs = new();
    private DateTime?    _lastDiagnosticLogAt;
    private bool         _disposed;

    // ── イベント ─────────────────────────────────────────────────
    public event Action<bool>?          StatusChanged;
    public event Action<MonitorStatus>? StatusUpdated;

    // ── パブリック ────────────────────────────────────────────────
    public bool IsSleepPrevented => _isSleepPrevented;

    public void Start()
    {
        _timer?.Stop();
        _timer?.Dispose();

        _timer = new System.Timers.Timer(10_000) { AutoReset = true };
        _timer.Elapsed += OnTimerElapsed;
        _timer.Start();
        SettingsManager.WriteLog(
            $"SleepGuard 監視開始 version={System.Reflection.Assembly.GetExecutingAssembly().GetName().Version}" +
            $" watched={_settings.WatchedProcesses.Count}");
    }

    public void Stop()
    {
        _timer?.Stop();
        _timer?.Dispose();
        _timer = null;

        if (_isSleepPrevented)
        {
            _isSleepPrevented = false;
            StatusChanged?.Invoke(false);
            SettingsManager.WriteLog("スリープ防止 終了 — 監視停止");
        }
        SettingsManager.WriteLog("SleepGuard 監視停止");
    }

    public void ReloadSettings()
    {
        var s = SettingsManager.Load();
        lock (_settingsLock) { _settings = s; }
    }

    // ── タイマー ─────────────────────────────────────────────────
    private void OnTimerElapsed(object? sender, ElapsedEventArgs e)
    {
        if (!Monitor.TryEnter(_timerLock)) return;
        try { Tick(); }
        finally { Monitor.Exit(_timerLock); }
    }

    private void Tick()
    {
        AppSettings settings;
        lock (_settingsLock) { settings = _settings; }

        var runningProcs = GetRunningWatchedProcesses(settings);
        var anyRunning   = runningProcs.Count > 0;
        var procsChanged = !runningProcs.SequenceEqual(_lastRunningProcs);

        if (anyRunning)
        {
            // プロセスあり → マウス移動量0でアイドルタイマーリセット
            SendInput(1, _mouseInput, _inputSize);

            if (!_isSleepPrevented)
            {
                _isSleepPrevented = true;
                StatusChanged?.Invoke(true);
                SettingsManager.WriteLog(
                    $"スリープ防止 開始 [{string.Join(", ", runningProcs)}]");
            }
        }
        else
        {
            // プロセスなし → 何もしない。Windowsタイマーが自然に進む。
            if (_isSleepPrevented)
            {
                _isSleepPrevented = false;
                StatusChanged?.Invoke(false);
                SettingsManager.WriteLog("スリープ防止 終了 — プロセスなし");
            }
        }

        // 診断ログ：状態変化時または5分に1回のみ
        var now = DateTime.Now;
        if (procsChanged || !_lastDiagnosticLogAt.HasValue ||
            (now - _lastDiagnosticLogAt.Value) >= TimeSpan.FromMinutes(5))
        {
            _lastDiagnosticLogAt = now;
            SettingsManager.WriteLog(
                $"監視診断 watched={settings.WatchedProcesses.Count}" +
                $" running={runningProcs.Count} blocked={_isSleepPrevented}" +
                $" runningList=[{string.Join(", ", runningProcs)}]");
        }

        if (procsChanged)
            PublishStatus(settings, runningProcs, now);
    }

    // ── プロセス検索 ─────────────────────────────────────────────
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
                if (TryMatchWatchedProcess(wp, procs))
                    running.Add(wp.DisplayName);
            }
            catch { }
            finally
            {
                if (procs != null)
                    foreach (var p in procs) p.Dispose();
            }
        }
        return running;
    }

    private static bool TryMatchWatchedProcess(WatchedProcess wp, Process[] processes)
    {
        if (processes.Length == 0) return false;

        var expectedPath = NormalizePath(wp.FilePath);
        if (string.IsNullOrWhiteSpace(expectedPath))
            return true;  // パス未設定はプロセス名のみで一致

        foreach (var process in processes)
        {
            try
            {
                var actualPath = NormalizePath(process.MainModule?.FileName);
                if (string.Equals(actualPath, expectedPath, StringComparison.OrdinalIgnoreCase))
                    return true;
            }
            catch { }
        }
        return false;
    }

    private static string NormalizePath(string? path)
    {
        if (string.IsNullOrWhiteSpace(path)) return string.Empty;
        try { return Path.GetFullPath(path.Trim().Trim('"')); }
        catch { return path.Trim().Trim('"'); }
    }

    // ── UI通知 ───────────────────────────────────────────────────
    private void PublishStatus(AppSettings settings, List<string> runningProcs, DateTime now)
    {
        _lastRunningProcs = new List<string>(runningProcs);
        StatusUpdated?.Invoke(new MonitorStatus
        {
            IsSleepPrevented = _isSleepPrevented,
            RunningProcesses = new List<string>(runningProcs),
            WatchedCount     = settings.WatchedProcesses.Count,
            CheckedAt        = now
        });
    }

    public MonitorStatus GetCurrentStatus()
    {
        AppSettings settings;
        lock (_settingsLock) { settings = _settings; }
        var runningProcs = GetRunningWatchedProcesses(settings);
        return new MonitorStatus
        {
            IsSleepPrevented = _isSleepPrevented,
            RunningProcesses = runningProcs,
            WatchedCount     = settings.WatchedProcesses.Count,
            CheckedAt        = DateTime.Now
        };
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        Stop();
        GC.SuppressFinalize(this);
    }
}

/// <summary>監視状態のスナップショット</summary>
public record MonitorStatus
{
    public bool         IsSleepPrevented { get; init; }
    public List<string> RunningProcesses { get; init; } = new();
    public int          WatchedCount     { get; init; }
    public DateTime     CheckedAt        { get; init; }
}
