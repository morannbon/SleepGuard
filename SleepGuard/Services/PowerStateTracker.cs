using System.Diagnostics.Eventing.Reader;
using System.IO;
using Microsoft.Win32;

namespace SleepGuard;

/// <summary>
/// Windows のスリープ・復帰・OS起動履歴を診断ログへ記録する。
/// スリープ抑止制御には一切関与しない。
/// </summary>
public sealed class PowerStateTracker : IDisposable
{
    private static readonly TimeSpan HeartbeatInterval = TimeSpan.FromSeconds(30);
    private static readonly TimeSpan ResumeGapThreshold = TimeSpan.FromSeconds(90);
    private static readonly TimeSpan ResumeHistoryLookback = TimeSpan.FromMinutes(30);

    private readonly string _bootStatePath =
        Path.Combine(SettingsManager.ConfigDirectory, "last-os-boot.txt");
    private readonly object _stateLock = new();
    private readonly System.Threading.Timer _heartbeatTimer;

    private DateTimeOffset _lastHeartbeatAt = DateTimeOffset.Now;
    private DateTimeOffset? _lastSuspendAt;
    private DateTimeOffset? _lastRecordedResumeAt;
    private bool _disposed;

    public PowerStateTracker()
    {
        LogStartupState();
        ScanRecentPowerEvents("startup", TimeSpan.FromHours(12));

        SystemEvents.PowerModeChanged += OnPowerModeChanged;
        _heartbeatTimer = new System.Threading.Timer(
            OnHeartbeat,
            null,
            HeartbeatInterval,
            HeartbeatInterval);
    }

    private void LogStartupState()
    {
        try
        {
            var now = DateTimeOffset.Now;
            var uptime = TimeSpan.FromMilliseconds(Environment.TickCount64);
            var currentBoot = now - uptime;

            DateTimeOffset? previousBoot = null;
            if (File.Exists(_bootStatePath))
            {
                var raw = File.ReadAllText(_bootStatePath).Trim();
                if (DateTimeOffset.TryParse(raw, out var parsed))
                    previousBoot = parsed;
            }

            var rebootDetected = previousBoot.HasValue &&
                Math.Abs((currentBoot - previousBoot.Value).TotalMinutes) > 2;

            SettingsManager.WriteLog(
                $"システム状態 action=startup osBoot={currentBoot:yyyy-MM-dd HH:mm:ss zzz} " +
                $"uptimeSec={(long)uptime.TotalSeconds} rebootDetected={rebootDetected} " +
                $"previousBoot={(previousBoot.HasValue ? previousBoot.Value.ToString("yyyy-MM-dd HH:mm:ss zzz") : "none")}");

            Directory.CreateDirectory(SettingsManager.ConfigDirectory);
            File.WriteAllText(_bootStatePath, currentBoot.ToString("O"));
        }
        catch (Exception ex)
        {
            SettingsManager.WriteLog(
                $"システム状態 action=startup-detect success=False error={ex.GetType().Name}:{ex.Message}");
        }
    }

    private void OnPowerModeChanged(object? sender, PowerModeChangedEventArgs e)
    {
        try
        {
            if (e.Mode == PowerModes.Suspend)
            {
                var now = DateTimeOffset.Now;
                lock (_stateLock)
                {
                    if (_disposed)
                        return;

                    _lastSuspendAt = now;
                }

                SettingsManager.WriteLog(
                    "システム状態 action=suspend source=SystemEvents.PowerModeChanged");
            }
            else if (e.Mode == PowerModes.Resume)
            {
                // PowerModeChanged の Resume は環境によって実際の復帰より遅れて届くことがある。
                // 通知時刻をそのまま復帰時刻にせず、Windows の復帰履歴を正本として確定する。
                if (!TryRecordResumeFromHistory("SystemEvents.PowerModeChanged"))
                    TryRecordResumeFallback(DateTimeOffset.Now, "SystemEvents.PowerModeChanged");
            }
        }
        catch (Exception ex)
        {
            SettingsManager.WriteLog(
                $"システム状態 action=power-event success=False error={ex.GetType().Name}:{ex.Message}");
        }
    }

    private void OnHeartbeat(object? state)
    {
        try
        {
            DateTimeOffset now;
            DateTimeOffset previousHeartbeat;
            bool resumePending;

            lock (_stateLock)
            {
                if (_disposed)
                    return;

                now = DateTimeOffset.Now;
                previousHeartbeat = _lastHeartbeatAt;
                _lastHeartbeatAt = now;
                resumePending = IsResumePendingLocked();
            }

            if (!resumePending)
                return;

            // スリープ中は Timer が進まないため、復帰後の最初の Tick では大きな時間差が出る。
            // Resume 通知を待たずに Windows 電源履歴から実際の復帰時刻を確定する。
            if (TryRecordResumeFromHistory("heartbeat"))
                return;

            if (now - previousHeartbeat >= ResumeGapThreshold)
                TryRecordResumeFallback(now, "heartbeat-gap");
        }
        catch (Exception ex)
        {
            SettingsManager.WriteLog(
                $"システム状態 action=resume-detect success=False error={ex.GetType().Name}:{ex.Message}");
        }
    }

    private bool TryRecordResumeFromHistory(string detectedBy)
    {
        DateTimeOffset? suspendAt;
        DateTimeOffset? lastResumeAt;

        lock (_stateLock)
        {
            if (_disposed || !IsResumePendingLocked())
                return false;

            suspendAt = _lastSuspendAt;
            lastResumeAt = _lastRecordedResumeAt;
        }

        try
        {
            var sinceUtc = DateTime.UtcNow - ResumeHistoryLookback;
            if (suspendAt.HasValue)
            {
                var suspendSinceUtc = suspendAt.Value.UtcDateTime.AddMinutes(-1);
                if (suspendSinceUtc > sinceUtc)
                    sinceUtc = suspendSinceUtc;
            }

            var events = new List<PowerEventInfo>();
            ReadEvents(
                "*[System[Provider[@Name='Microsoft-Windows-Power-Troubleshooter'] and EventID=1]]",
                sinceUtc,
                events);

            var actualResume = events
                .Where(x => !suspendAt.HasValue || x.TimeCreated >= suspendAt.Value)
                .Where(x => !lastResumeAt.HasValue || x.TimeCreated > lastResumeAt.Value)
                .OrderBy(x => x.TimeCreated)
                .LastOrDefault();

            if (actualResume == null)
                return false;

            return TryCommitResume(
                actualResume.TimeCreated,
                "WindowsEventLog.Power-Troubleshooter",
                detectedBy);
        }
        catch (Exception ex)
        {
            SettingsManager.WriteLog(
                $"システム状態 action=resume-history success=False error={ex.GetType().Name}:{ex.Message}");
            return false;
        }
    }

    private void TryRecordResumeFallback(DateTimeOffset detectedAt, string source)
    {
        lock (_stateLock)
        {
            if (_disposed || !IsResumePendingLocked())
                return;
        }

        TryCommitResume(detectedAt, source, source);
    }

    private bool TryCommitResume(DateTimeOffset actualResumeAt, string source, string detectedBy)
    {
        lock (_stateLock)
        {
            if (_disposed || !IsResumePendingLocked())
                return false;

            if (_lastSuspendAt.HasValue && actualResumeAt < _lastSuspendAt.Value)
                return false;

            if (_lastRecordedResumeAt.HasValue && actualResumeAt <= _lastRecordedResumeAt.Value)
                return false;

            _lastRecordedResumeAt = actualResumeAt;
        }

        SettingsManager.WriteLog(
            $"システム状態 action=resume time={actualResumeAt:yyyy-MM-dd HH:mm:ss zzz}");
        ScanRecentPowerEvents("resume", ResumeHistoryLookback);
        return true;
    }

    private bool IsResumePendingLocked()
    {
        return _lastSuspendAt.HasValue &&
            (!_lastRecordedResumeAt.HasValue || _lastRecordedResumeAt.Value < _lastSuspendAt.Value);
    }

    private static void ScanRecentPowerEvents(string reason, TimeSpan lookback)
    {
        try
        {
            var sinceUtc = DateTime.UtcNow - lookback;
            var found = new List<PowerEventInfo>();

            ReadEvents(
                "*[System[Provider[@Name='Microsoft-Windows-Kernel-Power'] and (EventID=41 or EventID=42 or EventID=107)]]",
                sinceUtc, found);
            ReadEvents(
                "*[System[Provider[@Name='Microsoft-Windows-Power-Troubleshooter'] and EventID=1]]",
                sinceUtc, found);
            ReadEvents(
                "*[System[Provider[@Name='EventLog'] and (EventID=6005 or EventID=6006 or EventID=6008)]]",
                sinceUtc, found);

            foreach (var ev in found.OrderBy(x => x.TimeCreated).TakeLast(12))
            {
                SettingsManager.WriteLog(
                    $"Windows電源履歴 reason={reason} time={ev.TimeCreated:yyyy-MM-dd HH:mm:ss} " +
                    $"provider={ev.Provider} eventId={ev.EventId}");
            }

            SettingsManager.WriteLog(
                $"Windows電源履歴 reason={reason} success=True events={found.Count} lookbackMin={(int)lookback.TotalMinutes}");
        }
        catch (Exception ex)
        {
            SettingsManager.WriteLog(
                $"Windows電源履歴 reason={reason} success=False error={ex.GetType().Name}:{ex.Message}");
        }
    }

    private static void ReadEvents(string queryText, DateTime sinceUtc, List<PowerEventInfo> output)
    {
        var query = new EventLogQuery("System", PathType.LogName, queryText)
        {
            ReverseDirection = true,
            TolerateQueryErrors = true
        };

        using var reader = new EventLogReader(query);
        for (var i = 0; i < 64; i++)
        {
            using var record = reader.ReadEvent();
            if (record == null)
                break;

            if (!record.TimeCreated.HasValue)
                continue;

            var created = record.TimeCreated.Value;
            if (created.ToUniversalTime() < sinceUtc)
                break;

            output.Add(new PowerEventInfo(
                new DateTimeOffset(created),
                record.ProviderName ?? "unknown",
                record.Id));
        }
    }

    public void Dispose()
    {
        lock (_stateLock)
        {
            if (_disposed)
                return;

            _disposed = true;
        }

        _heartbeatTimer.Dispose();
        SystemEvents.PowerModeChanged -= OnPowerModeChanged;
    }

    private sealed record PowerEventInfo(DateTimeOffset TimeCreated, string Provider, int EventId);
}
