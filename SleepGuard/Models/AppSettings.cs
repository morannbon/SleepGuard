using System.IO;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace SleepGuard;

/// <summary>監視対象プロセスの情報</summary>
public class WatchedProcess
{
    [JsonPropertyName("id")]
    public string Id { get; set; } = Guid.NewGuid().ToString();

    [JsonPropertyName("displayName")]
    public string DisplayName { get; set; } = string.Empty;

    /// <summary>プロセス名（拡張子なし。例: notepad）</summary>
    [JsonPropertyName("processName")]
    public string ProcessName { get; set; } = string.Empty;

    /// <summary>元のファイルパス（参照用）</summary>
    [JsonPropertyName("filePath")]
    public string FilePath { get; set; } = string.Empty;
}

/// <summary>アプリ設定</summary>
public class AppSettings
{
    [JsonPropertyName("watchedProcesses")]
    public List<WatchedProcess> WatchedProcesses { get; set; } = new();

    /// <summary>スリープ解除までの猶予時間（分）</summary>
    [JsonPropertyName("sleepDelayMinutes")]
    public int SleepDelayMinutes { get; set; } = 5;

    /// <summary>プロセス確認間隔（秒）</summary>
    [JsonPropertyName("checkIntervalSeconds")]
    public int CheckIntervalSeconds { get; set; } = 10;

    /// <summary>ディスプレイのスリープも防止するか</summary>
    [JsonPropertyName("preventDisplaySleep")]
    public bool PreventDisplaySleep { get; set; } = false;

    /// <summary>起動時にタスクトレイに格納するか</summary>
    [JsonPropertyName("startMinimized")]
    public bool StartMinimized { get; set; } = false;

    /// <summary>動作ログを記録するか</summary>
    [JsonPropertyName("logEnabled")]
    public bool LogEnabled { get; set; } = true;

    /// <summary>タスクトレイに常駐するか（false = × ボタンで完全終了）</summary>
    [JsonPropertyName("residentMode")]
    public bool ResidentMode { get; set; } = true;
}

/// <summary>
/// 設定の読み書きとログ管理を担当。
/// WriteLog はディスク I/O を最小化するため LogEnabled フラグをキャッシュする。
/// </summary>
public static class SettingsManager
{
    private static readonly string ConfigDir =
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "SleepGuard");

    private static readonly string ConfigPath = Path.Combine(ConfigDir, "settings.json");
    private static readonly string LogPath    = Path.Combine(ConfigDir, "sleepguard.log");

    private static readonly JsonSerializerOptions JsonOpts = new()
    {
        WriteIndented = true,
        Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
    };

    // WriteLog のたびにディスクを読まないようフラグをキャッシュ
    private static bool _logEnabled    = true;
    private static int  _logWriteCount = 0;

    public static AppSettings Load()
    {
        try
        {
            if (File.Exists(ConfigPath))
            {
                var json = File.ReadAllText(ConfigPath);
                var s    = JsonSerializer.Deserialize<AppSettings>(json, JsonOpts) ?? new AppSettings();
                _logEnabled = s.LogEnabled;
                return s;
            }
        }
        catch { }
        return new AppSettings();
    }

    public static void Save(AppSettings settings)
    {
        _logEnabled = settings.LogEnabled;
        Directory.CreateDirectory(ConfigDir);
        var json = JsonSerializer.Serialize(settings, JsonOpts);
        File.WriteAllText(ConfigPath, json, System.Text.Encoding.UTF8);
    }

    /// <summary>
    /// ログを 1 行追記する。
    /// 100 回に 1 回だけファイルサイズをチェックし、1 MB 超で末尾 500 行に切り詰める。
    /// </summary>
    public static void WriteLog(string message)
    {
        if (!_logEnabled) return;
        try
        {
            Directory.CreateDirectory(ConfigDir);
            if (++_logWriteCount >= 100)
            {
                _logWriteCount = 0;
                RotateLogIfNeeded();
            }
            var line = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss}  {message}{Environment.NewLine}";
            File.AppendAllText(LogPath, line, System.Text.Encoding.UTF8);
        }
        catch { }
    }

    private static void RotateLogIfNeeded()
    {
        try
        {
            if (!File.Exists(LogPath)) return;
            if (new FileInfo(LogPath).Length < 1024 * 1024) return;

            var lines = File.ReadAllLines(LogPath, System.Text.Encoding.UTF8);
            if (lines.Length <= 500) return;

            File.WriteAllLines(LogPath,
                lines.Skip(lines.Length - 500).ToArray(),
                System.Text.Encoding.UTF8);
        }
        catch { }
    }

    public static string LogFilePath     => LogPath;
    public static string ConfigDirectory => ConfigDir;
}

/// <summary>Windows スタートアップ登録（HKCU\...\Run）を管理する。</summary>
public static class StartupManager
{
    private const string RegKey  = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Run";
    private const string AppName = "SleepGuard";

    public static bool IsRegistered()
    {
        try
        {
            using var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(RegKey, false);
            return key?.GetValue(AppName) is string;
        }
        catch { return false; }
    }

    public static void Register()
    {
        try
        {
            var exe = Environment.ProcessPath
                      ?? System.Diagnostics.Process.GetCurrentProcess().MainModule?.FileName
                      ?? string.Empty;
            if (string.IsNullOrEmpty(exe)) return;
            using var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(RegKey, true);
            key?.SetValue(AppName, $"\"{exe}\"");
        }
        catch { }
    }

    public static void Unregister()
    {
        try
        {
            using var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(RegKey, true);
            key?.DeleteValue(AppName, throwOnMissingValue: false);
        }
        catch { }
    }
}
