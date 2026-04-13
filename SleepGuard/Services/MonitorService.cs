using System.Collections.Concurrent;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Timers;

namespace SleepGuard;

/// <summary>
/// プロセス監視とスリープ防止を担当するサービス。
///
/// SetThreadExecutionState はスレッドに紐付く API のため、
/// Prevent/Allow を必ず同一スレッドから呼ぶ専用スレッドを用意している。
/// これにより「別スレッドから Allow しても効かない」問題を根本解決している。
/// </summary>
public class MonitorService : IDisposable
{
    // ── Win32 API ────────────────────────────────────────────────
    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern uint SetThreadExecutionState(uint esFlags);

    private const uint ES_CONTINUOUS       = 0x80000000u;
    private const uint ES_SYSTEM_REQUIRED  = 0x00000001u;
    private const uint ES_DISPLAY_REQUIRED = 0x00000002u;

    // ── SetThreadExecutionState 専用スレッド ─────────────────────
    private readonly Thread _sleepControlThread;
    private readonly BlockingCollection<Action> _sleepControlQueue = new();

    // ── 状態 ─────────────────────────────────────────────────────
    private AppSettings _settings = SettingsManager.Load();
    private readonly object _settingsLock = new();
    private readonly object _timerLock   = new();

    private System.Timers.Timer? _timer;
    private bool      _isSleepPrevented;
    private DateTime? _allStoppedAt;
    private List<string> _lastRunningProcs = new();
    private bool _disposed;

    // ── イベント ─────────────────────────────────────────────────
    public event Action<bool>?         StatusChanged;
    public event Action<MonitorStatus>? StatusUpdated;

    // ── パブリック ────────────────────────────────────────────────
    public bool IsSleepPrevented => _isSleepPrevented;

    public MonitorService()
    {
        _sleepControlThread = new Thread(() =>
        {
            foreach (var action in _sleepControlQueue.GetConsumingEnumerable())
                action();
        })
        {
            IsBackground = true,
            Name = "SleepControlThread"
        };
        _sleepControlThread.Start();
    }

    public void Start()
    {
        _timer?.Stop();
        _timer?.Dispose();

        _timer = new System.Timers.Timer(GetIntervalMs())
        {
            AutoReset = true
        };
        _timer.Elapsed += OnTimerElapsed;
        _timer.Start();
        SettingsManager.WriteLog("SleepGuard 監視開始");
    }

    public void Stop()
    {
        _timer?.Stop();
        _timer?.Dispose();
        _timer = null;
        AllowSleepSync();  // 専用スレッド経由で確実に解除してから戻る
        SettingsManager.WriteLog("SleepGuard 監視停止");
    }

    /// <summary>設定変更時に呼ぶ。タイマー間隔も即時反映する。</summary>
    public void ReloadSettings()
    {
        var s = SettingsManager.Load();
        lock (_settingsLock) { _settings = s; }
        if (_timer != null)
            _timer.Interval = GetIntervalMs();
    }

    // ── タイマー ─────────────────────────────────────────────────
    private void OnTimerElapsed(object? sender, ElapsedEventArgs e)
    {
        if (!Monitor.TryEnter(_timerLock)) return;  // 前回が未完了なら今回はスキップ
        try { Tick(); }
        finally { Monitor.Exit(_timerLock); }
    }

    private void Tick()
    {
        AppSettings settings;
        lock (_settingsLock) { settings = _settings; }

        if (settings.WatchedProcesses.Count == 0)
        {
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
                var remaining = TimeSpan.FromMinutes(settings.SleepDelayMinutes)
                                - (DateTime.Now - _allStoppedAt.Value);
                if (remaining <= TimeSpan.Zero)
                {
                    AllowSleep();
                    _allStoppedAt = null;
                    SettingsManager.WriteLog(
                        $"スリープ防止 解除 — 猶予 {settings.SleepDelayMinutes} 分経過");
                }
            }
        }

        // UI 通知は状態変化時のみ（不要な再描画を抑制）
        bool procsChanged = !runningProcs.SequenceEqual(_lastRunningProcs);
        if (procsChanged || _allStoppedAt.HasValue)
        {
            _lastRunningProcs = runningProcs;
            StatusUpdated?.Invoke(new MonitorStatus
            {
                IsSleepPrevented = _isSleepPrevented,
                RunningProcesses = runningProcs,
                CountdownStartAt = _allStoppedAt,
                DelayMinutes     = settings.SleepDelayMinutes,
                WatchedCount     = settings.WatchedProcesses.Count,
                CheckedAt        = DateTime.Now
            });
        }
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
                if (procs.Length > 0)
                    running.Add(wp.DisplayName);
            }
            catch { /* アクセス拒否・権限不足は無視 */ }
            finally
            {
                if (procs != null)
                    foreach (var p in procs) p.Dispose();
            }
        }
        return running;
    }

    // ── スリープ制御 ─────────────────────────────────────────────
    private void PreventSleep(AppSettings settings)
    {
        var flags = ES_CONTINUOUS | ES_SYSTEM_REQUIRED;
        if (settings.PreventDisplaySleep)
            flags |= ES_DISPLAY_REQUIRED;

        if (_sleepControlQueue.IsAddingCompleted) return;
        _sleepControlQueue.Add(() => SetThreadExecutionState(flags));
        _isSleepPrevented = true;
        StatusChanged?.Invoke(true);
    }

    private void AllowSleep()
    {
        if (_sleepControlQueue.IsAddingCompleted) return;
        _sleepControlQueue.Add(AllowSleepCore);
        if (_isSleepPrevented)
        {
            _isSleepPrevented = false;
            StatusChanged?.Invoke(false);
        }
    }

    /// <summary>Stop/Dispose 時に専用スレッドの完了を待つ同期版。</summary>
    private void AllowSleepSync()
    {
        if (_sleepControlQueue.IsAddingCompleted) return;
        var tcs = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
        _sleepControlQueue.Add(() =>
        {
            AllowSleepCore();
            tcs.SetResult(true);
        });
        tcs.Task.GetAwaiter().GetResult();
    }

    private void AllowSleepCore() => SetThreadExecutionState(ES_CONTINUOUS);

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
            CountdownStartAt = _allStoppedAt,
            DelayMinutes     = settings.SleepDelayMinutes,
            WatchedCount     = settings.WatchedProcesses.Count,
            CheckedAt        = DateTime.Now
        };
    }

    // ── ヘルパー ─────────────────────────────────────────────────
    private double GetIntervalMs()
    {
        lock (_settingsLock)
            return Math.Max(5, _settings.CheckIntervalSeconds) * 1000.0;
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        Stop();
        _sleepControlQueue.CompleteAdding();
        GC.SuppressFinalize(this);
    }
}

/// <summary>監視状態のスナップショット</summary>
public record MonitorStatus
{
    public bool          IsSleepPrevented { get; init; }
    public List<string>  RunningProcesses { get; init; } = new();
    public DateTime?     CountdownStartAt { get; init; }
    public int           DelayMinutes     { get; init; }
    public int           WatchedCount     { get; init; }
    public DateTime      CheckedAt        { get; init; }

    /// <summary>カウントダウン残り秒。カウントダウン中でなければ null。</summary>
    public double? RemainingSeconds =>
        CountdownStartAt.HasValue
            ? Math.Max(0, (DelayMinutes * 60) - (CheckedAt - CountdownStartAt.Value).TotalSeconds)
            : null;
}
