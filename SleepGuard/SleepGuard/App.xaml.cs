using System.Runtime.InteropServices;
using System.Windows;
using Hardcodet.Wpf.TaskbarNotification;

namespace SleepGuard;

public partial class App : Application
{
    private TaskbarIcon?     _trayIcon;
    private MonitorService?  _monitorService;
    private MainWindow?      _mainWindow;

    // トレイアイコンをキャッシュ（状態ごとに一度だけ生成）
    private System.Drawing.Icon? _iconActive;
    private System.Drawing.Icon? _iconInactive;

    // 一時停止フラグ
    private bool _isPaused;

    internal void RequestExit()
    {
        _monitorService?.Stop();
        _trayIcon?.Dispose();
        _iconActive?.Dispose();
        _iconInactive?.Dispose();
        Shutdown();
    }

    private System.Windows.Controls.ContextMenu BuildContextMenu()
    {
        var menu = new System.Windows.Controls.ContextMenu();
        PopulateContextMenu(menu);
        return menu;
    }

    private void RefreshContextMenu()
    {
        if (_trayIcon?.ContextMenu == null) return;
        _trayIcon.ContextMenu.Items.Clear();
        PopulateContextMenu(_trayIcon.ContextMenu);
    }

    private void PopulateContextMenu(System.Windows.Controls.ContextMenu menu)
    {
        // ── 監視プロセス ──────────────────────────────────────────
        menu.Items.Add(new System.Windows.Controls.MenuItem
        {
            Header    = "監視プロセス",
            IsEnabled = false,
            FontWeight = System.Windows.FontWeights.Bold
        });

        var settings     = SettingsManager.Load();
        var runningNames = _monitorService?.GetRunningWatchedProcesses() ?? new List<string>();

        if (settings.WatchedProcesses.Count == 0)
        {
            menu.Items.Add(new System.Windows.Controls.MenuItem
            {
                Header = "  (登録なし)",
                IsEnabled = false
            });
        }
        else
        {
            foreach (var wp in settings.WatchedProcesses)
            {
                var isRunning = runningNames.Contains(wp.DisplayName);
                menu.Items.Add(new System.Windows.Controls.MenuItem
                {
                    Header     = $"  {(isRunning ? "● " : "○ ")}{wp.DisplayName}",
                    IsEnabled  = false,
                    Foreground = isRunning
                        ? System.Windows.Media.Brushes.LightGreen
                        : System.Windows.Media.Brushes.Gray
                });
            }
        }

        menu.Items.Add(new System.Windows.Controls.Separator());

        // ── 抑止 停止 / 再開 ─────────────────────────────────────
        var pauseItem = new System.Windows.Controls.MenuItem
        {
            Header = _isPaused ? "▶  抑止を再開" : "⏸  抑止を停止"
        };
        pauseItem.Click += (_, _) => TogglePause();
        menu.Items.Add(pauseItem);

        menu.Items.Add(new System.Windows.Controls.Separator());

        // ── 設定を開く ───────────────────────────────────────────
        var openItem = new System.Windows.Controls.MenuItem { Header = "設定を開く" };
        openItem.Click += (_, _) => ShowMainWindow();
        menu.Items.Add(openItem);

        menu.Items.Add(new System.Windows.Controls.Separator());

        // ── 終了 ─────────────────────────────────────────────────
        var exitItem = new System.Windows.Controls.MenuItem { Header = "終了" };
        exitItem.Click += (_, _) => RequestExit();
        menu.Items.Add(exitItem);
    }

    private void TogglePause()
    {
        _isPaused = !_isPaused;
        if (_isPaused)
        {
            _monitorService?.Stop();
            if (_trayIcon != null)
            {
                _trayIcon.Icon        = _iconInactive;
                _trayIcon.ToolTipText = "SleepGuard - 停止中";
            }
            SettingsManager.WriteLog("監視を一時停止しました");
        }
        else
        {
            _monitorService?.Start();
            SettingsManager.WriteLog("監視を再開しました");
        }
    }

    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);

        AppDomain.CurrentDomain.UnhandledException += (_, args) =>
        {
            var msg = args.ExceptionObject is Exception ex ? ex.ToString() : args.ExceptionObject?.ToString();
            MessageBox.Show($"起動エラー:\n\n{msg}", "SleepGuard", MessageBoxButton.OK, MessageBoxImage.Error);
        };
        DispatcherUnhandledException += (_, args) =>
        {
            MessageBox.Show($"エラー:\n\n{args.Exception}", "SleepGuard", MessageBoxButton.OK, MessageBoxImage.Error);
            args.Handled = true;
        };

        try
        {
            // アイコンを起動時に一度だけ生成してキャッシュ
            _iconActive   = CreateIcon(active: true);
            _iconInactive = CreateIcon(active: false);

            _monitorService = new MonitorService();

            _trayIcon = new TaskbarIcon
            {
                ToolTipText = "SleepGuard - 待機中",
                Icon        = _iconInactive,
                Visibility  = Visibility.Visible
            };

            // コンテキストメニューを構築
            _trayIcon.ContextMenu = BuildContextMenu();

            _trayIcon.TrayMouseDoubleClick  += (_, _) => ShowMainWindow();
            _trayIcon.TrayRightMouseDown    += (_, _) => RefreshContextMenu();

            _monitorService.StatusChanged  += OnMonitorStatusChanged;
            _monitorService.Start();

            var settings = SettingsManager.Load();
            if (settings.StartMinimized)
            {
                // 起動時トレイ格納: ウィンドウを表示しない
                SettingsManager.WriteLog("SleepGuard 起動 (トレイに格納)");
            }
            else
            {
                ShowMainWindow();
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"起動に失敗しました:\n\n{ex}", "SleepGuard",
                MessageBoxButton.OK, MessageBoxImage.Error);
            Shutdown(1);
        }
    }

    private void OnMonitorStatusChanged(bool isSleepPrevented)
    {
        Dispatcher.Invoke(() =>
        {
            if (_trayIcon == null) return;
            // キャッシュ済みアイコンを切り替えるだけ（GDI描画なし）
            _trayIcon.Icon       = isSleepPrevented ? _iconActive : _iconInactive;
            _trayIcon.ToolTipText = isSleepPrevented
                ? "SleepGuard - スリープ防止中"
                : "SleepGuard - 待機中";
        });
    }

    internal void ShowMainWindow()
    {
        if (_mainWindow == null || !_mainWindow.IsLoaded)
            _mainWindow = new MainWindow(_monitorService!);
        _mainWindow.Show();
        _mainWindow.WindowState = WindowState.Normal;
        _mainWindow.Activate();
    }

    private static System.Drawing.Icon CreateIcon(bool active)
    {
        const int size = 16;
        using var bmp = new System.Drawing.Bitmap(
            size, size, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
        using (var g = System.Drawing.Graphics.FromImage(bmp))
        {
            g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
            g.Clear(System.Drawing.Color.Transparent);
            var outer = active
                ? System.Drawing.Color.FromArgb(180, 22, 211, 165)
                : System.Drawing.Color.FromArgb(120, 100, 100, 120);
            using var b1 = new System.Drawing.SolidBrush(outer);
            g.FillEllipse(b1, 0, 0, size - 1, size - 1);
            var inner = active
                ? System.Drawing.Color.FromArgb(255, 34, 211, 165)
                : System.Drawing.Color.FromArgb(200, 90, 88, 110);
            using var b2 = new System.Drawing.SolidBrush(inner);
            g.FillEllipse(b2, 2, 2, size - 5, size - 5);
        }
        // GetHicon→Clone→DestroyIcon の正しいパターン
        var hIcon = bmp.GetHicon();
        try
        {
            using var tmp = System.Drawing.Icon.FromHandle(hIcon);
            return (System.Drawing.Icon)tmp.Clone();
        }
        finally
        {
            DestroyIcon(hIcon);
        }
    }

    [DllImport("user32.dll")]
    private static extern bool DestroyIcon(IntPtr hIcon);

    protected override void OnExit(ExitEventArgs e)
    {
        _monitorService?.Stop();
        _trayIcon?.Dispose();
        _iconActive?.Dispose();
        _iconInactive?.Dispose();
        base.OnExit(e);
    }
}
