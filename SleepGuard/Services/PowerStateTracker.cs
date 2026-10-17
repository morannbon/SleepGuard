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
    private readonly string _bootStatePath =
        Path.Combine(SettingsManager.ConfigDirectory, "last-os-boot.txt");

    private bool _disposed;

    public PowerStateTracker()
    {
        LogStartupState();
        ScanRecentPowerEvents("startup", TimeSpan.FromHours(12));
        SystemEvents.PowerModeChanged += OnPowerModeChanged;
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
                SettingsManager.WriteLog(
                    "システム状態 action=suspend source=SystemEvents.PowerModeChanged");
            }
            else if (e.Mode == PowerModes.Resume)
            {
                SettingsManager.WriteLog(
                    "システム状態 action=resume source=SystemEvents.PowerModeChanged");
                ScanRecentPowerEvents("resume", TimeSpan.FromMinutes(30));
            }
        }
        catch (Exception ex)
        {
            SettingsManager.WriteLog(
                $"システム状態 action=power-event success=False error={ex.GetType().Name}:{ex.Message}");
        }
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
                created,
                record.ProviderName ?? "unknown",
                record.Id));
        }
    }

    public void Dispose()
    {
        if (_disposed)
            return;

        SystemEvents.PowerModeChanged -= OnPowerModeChanged;
        _disposed = true;
    }

    private sealed record PowerEventInfo(DateTime TimeCreated, string Provider, int EventId);
}
