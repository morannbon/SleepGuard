using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Timers;

namespace SleepGuard;

/// <summary>
/// プロセス監視とスリープ防止を担当するサービス
/// 改善点:
///   - タイマーコールバック毎のディスク読み込み(Load)を廃止しメモリキャッシュを使用
///   - Process.GetProcessesByName後のDispose漏れを修正
///   - UIへの通知を状態変化時のみに絞り不要なDispatch/再描画を削減
///   - タイマー重複実行防止のためのロック追加
///   - Process配列の即時Dispose
/// </summary>
public class MonitorService : IDisposable
{
    // ─── Win32 API ───────────────────────────────────────────────
    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern uint SetThreadExecutionState(uint esFlags);

    private const uint ES_CONTINUOUS       = 0x80000000u;
    private const uint ES_SYSTEM_REQUIRED  = 0x00000001u;
    private const uint ES_DISPLAY_REQUIRED = 0x00000002u;

    // ─── State ───────────────────────────────────────────────────
    private AppSettings _settings = SettingsManager.Load();
    private readonly object _settingsLock = new();
    private readonly object _timerLock = new();
    private bool _timerRunning;

    private System.Timers.Timer? _timer;
    private bool _isSleepPrevented;
    private DateTime? _allStoppedAt;
    private List<string> _lastRunningProcs = new();
    private bool _disposed;

    // ─── Events ──────────────────────────────────────────────────
    public event Action<bool>? StatusChanged;
    public event Action<MonitorStatus>? StatusUpdated;

    // ─── Public API ──────────────────────────────────────────────
    public bool IsSleepPrevented => _isSleepPrevented;

    public void Start()
    {
        _timer = new System.Timers.Timer(GetIntervalMs());
        _timer.Elapsed += OnTimerElapsed;
        _timer.AutoReset = true;
        _timer.Start();
        SettingsManager.WriteLog("SleepGuard 監視開始");
    }

    public void Stop()
    {
        _timer?.Stop();
        _timer?.Dispose();
        _timer = null;
        AllowSleep();
        SettingsManager.WriteLog("SleepGuard 監視停止");
    }

    /// <summary>設定変更時に呼ぶ。ディスクから再読み込みしタイマー間隔を更新。</summary>
    public void ReloadSettings()
    {
        var s = SettingsManager.Load();
        lock (_settingsLock) { _settings = s; }
        if (_timer != null)
            _timer.Interval = GetIntervalMs();
    }

    // ─── Timer callback ─────────────────────────────────────────
    private void OnTimerElapsed(object? sender, ElapsedEventArgs e)
    {
        // 前回のコールバックがまだ実行中なら今回はスキップ（重複防止）
        if (!Monitor.TryEnter(_timerLock)) return;
        try
        {
            Tick();
        }
        finally
        {
            Monitor.Exit(_timerLock);
        }
    }

    private void Tick()
    {
        AppSettings settings;
        lock (_settingsLock) { settings = _settings; }

        // 登録プロセスがなければ何もしない
        if (settings.WatchedProcesses.Count == 0)
        {
            // 念のためスリープ防止を解除
            if (_isSleepPrevented) AllowSleep();
            return;
        }

        var runningProcs = GetRunningWatchedProcesses(settings);
        var anyRunning   = runningProcs.Count > 0;

        if (anyRunning)
        {
            _allStoppedAt = null;

            if (!_isSleepPrevented)
            {
                PreventSleep(settings);
                SettingsManager.WriteLog(
                    $"スリープ防止 有効化 — 実行中: {string.Join(", ", runningProcs)}");
            }
        }
        else
        {
            if (_isSleepPrevented)
            {
                _allStoppedAt ??= DateTime.Now;

                var elapsed   = DateTime.Now - _allStoppedAt.Value;
                var remaining = TimeSpan.FromMinutes(settings.SleepDelayMinutes) - elapsed;

                if (remaining <= TimeSpan.Zero)
                {
                    AllowSleep();
                    _allStoppedAt = null;
                    SettingsManager.WriteLog(
                        $"スリープ防止 解除 — 猶予 {settings.SleepDelayMinutes} 分経過");
                }
            }
        }

        // 状態変化があった時だけUIへ通知（不要な再描画を抑制）
        bool procsChanged = !runningProcs.SequenceEqual(_lastRunningProcs);
        if (procsChanged || _allStoppedAt.HasValue)
        {
            _lastRunningProcs = runningProcs;

            var status = new MonitorStatus
            {
                IsSleepPrevented = _isSleepPrevented,
                RunningProcesses = runningProcs,
                CountdownStartAt = _allStoppedAt,
                DelayMinutes     = settings.SleepDelayMinutes,
                WatchedCount     = settings.WatchedProcesses.Count,
                CheckedAt        = DateTime.Now
            };
            StatusUpdated?.Invoke(status);
        }
    }

    // ─── Process detection ───────────────────────────────────────
    /// <summary>
    /// 監視対象プロセスのうち現在実行中のもの一覧を返す。
    /// Process配列は使用後即Disposeしてハンドルを解放する。
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
                if (procs.Length > 0)
                    running.Add(wp.DisplayName);
            }
            catch { /* アクセス拒否・権限不足は無視 */ }
            finally
            {
                // Processオブジェクトはハンドルを保持するので必ずDispose
                if (procs != null)
                    foreach (var p in procs) p.Dispose();
            }
        }
        return running;
    }

    // ─── Sleep control ───────────────────────────────────────────
    private void PreventSleep(AppSettings settings)
    {
        var flags = ES_CONTINUOUS | ES_SYSTEM_REQUIRED;
        if (settings.PreventDisplaySleep)
            flags |= ES_DISPLAY_REQUIRED;

        SetThreadExecutionState(flags);
        _isSleepPrevented = true;
        StatusChanged?.Invoke(true);
    }

    private void AllowSleep()
    {
        SetThreadExecutionState(ES_CONTINUOUS);
        if (_isSleepPrevented)
        {
            _isSleepPrevented = false;
            StatusChanged?.Invoke(false);
        }
    }

    // ─── Helpers ─────────────────────────────────────────────────
    private double GetIntervalMs()
    {
        lock (_settingsLock)
        {
            return Math.Max(5, _settings.CheckIntervalSeconds) * 1000.0;
        }
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
    public bool IsSleepPrevented { get; init; }
    public List<string> RunningProcesses { get; init; } = new();
    public DateTime? CountdownStartAt { get; init; }
    public int DelayMinutes { get; init; }
    public int WatchedCount { get; init; }
    public DateTime CheckedAt { get; init; }

    /// <summary>カウントダウン残り秒 (null = カウントダウン中でない)</summary>
    public double? RemainingSeconds =>
        CountdownStartAt.HasValue
            ? Math.Max(0, (DelayMinutes * 60) - (CheckedAt - CountdownStartAt.Value).TotalSeconds)
            : null;
}
