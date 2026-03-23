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

    /// <summary>プロセス名 (拡張子なし, 例: notepad)</summary>
    [JsonPropertyName("processName")]
    public string ProcessName { get; set; } = string.Empty;

    /// <summary>元のファイルパス (参照用)</summary>
    [JsonPropertyName("filePath")]
    public string FilePath { get; set; } = string.Empty;
}

/// <summary>アプリ設定</summary>
public class AppSettings
{
    [JsonPropertyName("watchedProcesses")]
    public List<WatchedProcess> WatchedProcesses { get; set; } = new();

    /// <summary>スリープ解除までの猶予時間 (分)</summary>
    [JsonPropertyName("sleepDelayMinutes")]
    public int SleepDelayMinutes { get; set; } = 5;

    /// <summary>プロセス確認間隔 (秒)</summary>
    [JsonPropertyName("checkIntervalSeconds")]
    public int CheckIntervalSeconds { get; set; } = 10;

    /// <summary>ディスプレイのスリープも防止するか</summary>
    [JsonPropertyName("preventDisplaySleep")]
    public bool PreventDisplaySleep { get; set; } = false;

    /// <summary>起動時にタスクトレイに格納するか</summary>
    [JsonPropertyName("startMinimized")]
    public bool StartMinimized { get; set; } = false;

    /// <summary>ログを有効にするか</summary>
    [JsonPropertyName("logEnabled")]
    public bool LogEnabled { get; set; } = true;

    /// <summary>タスクトレイに常駐するか (false = ×ボタンで完全終了)</summary>
    [JsonPropertyName("residentMode")]
    public bool ResidentMode { get; set; } = true;
}

/// <summary>
/// 設定の読み書きを担当。
/// 改善点: WriteLogでLoad()を毎回呼ばずフラグをキャッシュ。
/// </summary>
public class SettingsManager
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

    // ログ有効フラグのキャッシュ（WriteLogのたびにディスクを読まないため）
    private static bool _logEnabled = true;

    public static AppSettings Load()
    {
        try
        {
            if (File.Exists(ConfigPath))
            {
                var json = File.ReadAllText(ConfigPath);
                var s = JsonSerializer.Deserialize<AppSettings>(json, JsonOpts) ?? new AppSettings();
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
    /// ログ書き込み。毎回Load()しないのでディスクI/Oを削減。
    /// ログファイルが1MBを超えたら古い行を削除してローテーション。
    /// </summary>
    public static void WriteLog(string message)
    {
        if (!_logEnabled) return;
        try
        {
            Directory.CreateDirectory(ConfigDir);
            RotateLogIfNeeded();
            var line = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss}  {message}{Environment.NewLine}";
            File.AppendAllText(LogPath, line, System.Text.Encoding.UTF8);
        }
        catch { }
    }

    /// <summary>
    /// ログが1MBを超えたら末尾500行だけ残してローテーション。
    /// </summary>
    private static void RotateLogIfNeeded()
    {
        try
        {
            if (!File.Exists(LogPath)) return;
            var info = new FileInfo(LogPath);
            if (info.Length < 1024 * 1024) return; // 1MB未満はスキップ

            var lines = File.ReadAllLines(LogPath, System.Text.Encoding.UTF8);
            if (lines.Length <= 500) return;

            // 末尾500行を残す
            var kept = lines.Skip(lines.Length - 500).ToArray();
            File.WriteAllLines(LogPath, kept, System.Text.Encoding.UTF8);
        }
        catch { }
    }

    public static string LogFilePath    => LogPath;
    public static string ConfigDirectory => ConfigDir;
}

/// <summary>Windowsスタートアップ登録を管理 (HKCU\...\Run)</summary>
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
