using System.Runtime.InteropServices;
using System.Windows;
using Hardcodet.Wpf.TaskbarNotification;

namespace SleepGuard;

public partial class App : Application
{
    private TaskbarIcon?     _trayIcon;
    private MonitorService?  _monitorService;
    private MainWindow?      _mainWindow;
    private System.Threading.Mutex? _mutex;

    private System.Drawing.Icon? _iconActive;
    private System.Drawing.Icon? _iconInactive;
    private bool _isPaused;

    internal void RequestExit()
    {
        // イベント購読を先に解除してDispose후の呼び出しを防ぐ
        if (_monitorService != null)
            _monitorService.StatusChanged -= OnMonitorStatusChanged;
        _monitorService?.Stop();
        _trayIcon?.Dispose();
        _iconActive?.Dispose();
        _iconInactive?.Dispose();
        _iconActive   = null;
        _iconInactive = null;
        _mutex?.ReleaseMutex();
        _mutex?.Dispose();
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

        // 多重起動防止
        _mutex = new System.Threading.Mutex(true, "SleepGuard_SingleInstance", out bool createdNew);
        if (!createdNew)
        {
            // 既に起動中 → そちらのウィンドウを前面に出して終了
            MessageBox.Show("SleepGuard は既に起動しています。\nタスクトレイを確認してください。",
                "SleepGuard", MessageBoxButton.OK, MessageBoxImage.Information);
            _mutex.Dispose();
            Shutdown();
            return;
        }

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
        Dispatcher.BeginInvoke(() =>
        {
            // Dispose済みの場合はスキップ
            if (_trayIcon == null || _iconActive == null || _iconInactive == null) return;
            try
            {
                _trayIcon.Icon        = isSleepPrevented ? _iconActive : _iconInactive;
                _trayIcon.ToolTipText = isSleepPrevented
                    ? "SleepGuard - スリープ防止中"
                    : "SleepGuard - 待機中";
            }
            catch (ObjectDisposedException) { /* 終了処理中は無視 */ }
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
        try
        {
            var uri    = new Uri("pack://application:,,,/SleepGuard;component/Resources/icon.ico");
            var stream = System.Windows.Application.GetResourceStream(uri)?.Stream;
            if (stream != null)
            {
                using (stream)
                using (var icon = new System.Drawing.Icon(stream, 16, 16))
                {
                    if (active)
                        return (System.Drawing.Icon)icon.Clone();
                    return ToGrayscaleIcon(icon);
                }
            }
        }
        catch { }
        return CreateFallbackIcon(active);
    }

    private static System.Drawing.Icon ToGrayscaleIcon(System.Drawing.Icon icon)
    {
        using var src = icon.ToBitmap();
        using var dst = new System.Drawing.Bitmap(src.Width, src.Height,
            System.Drawing.Imaging.PixelFormat.Format32bppArgb);

        // LockBitsで高速ピクセル処理
        var rect = new System.Drawing.Rectangle(0, 0, src.Width, src.Height);
        var fmt  = System.Drawing.Imaging.PixelFormat.Format32bppArgb;
        var srcData = src.LockBits(rect, System.Drawing.Imaging.ImageLockMode.ReadOnly,  fmt);
        var dstData = dst.LockBits(rect, System.Drawing.Imaging.ImageLockMode.WriteOnly, fmt);
        try
        {
            int bytes = Math.Abs(srcData.Stride) * src.Height;
            var buf   = new byte[bytes];
            System.Runtime.InteropServices.Marshal.Copy(srcData.Scan0, buf, 0, bytes);
            for (int i = 0; i < bytes; i += 4)
            {
                byte b = buf[i], g = buf[i + 1], r = buf[i + 2], a = buf[i + 3];
                byte gs = (byte)(r * 0.299 + g * 0.587 + b * 0.114);
                buf[i] = buf[i + 1] = buf[i + 2] = gs;
                buf[i + 3] = (byte)(a / 2); // 半透明で非アクティブ感を表現
            }
            System.Runtime.InteropServices.Marshal.Copy(buf, 0, dstData.Scan0, bytes);
        }
        finally
        {
            src.UnlockBits(srcData);
            dst.UnlockBits(dstData);
        }

        var hIcon = dst.GetHicon();
        try
        {
            using var tmp = System.Drawing.Icon.FromHandle(hIcon);
            return (System.Drawing.Icon)tmp.Clone();
        }
        finally { DestroyIcon(hIcon); }
    }

    private static System.Drawing.Icon CreateFallbackIcon(bool active)
    {
        const int size = 16;
        using var bmp = new System.Drawing.Bitmap(
            size, size, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
        using (var g = System.Drawing.Graphics.FromImage(bmp))
        {
            g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
            g.Clear(System.Drawing.Color.Transparent);
            var color = active
                ? System.Drawing.Color.FromArgb(255, 34, 211, 165)
                : System.Drawing.Color.FromArgb(180, 100, 100, 120);
            using var b = new System.Drawing.SolidBrush(color);
            g.FillEllipse(b, 1, 1, size - 2, size - 2);
        }
        var hIcon = bmp.GetHicon();
        try
        {
            using var tmp = System.Drawing.Icon.FromHandle(hIcon);
            return (System.Drawing.Icon)tmp.Clone();
        }
        finally { DestroyIcon(hIcon); }
    }

    [DllImport("user32.dll")]
    private static extern bool DestroyIcon(IntPtr hIcon);

    protected override void OnExit(ExitEventArgs e)
    {
        if (_monitorService != null)
            _monitorService.StatusChanged -= OnMonitorStatusChanged;
        _monitorService?.Stop();
        _trayIcon?.Dispose();
        _iconActive?.Dispose();
        _iconInactive?.Dispose();
        _iconActive   = null;
        _iconInactive = null;
        try { _mutex?.ReleaseMutex(); } catch { }
        _mutex?.Dispose();
        base.OnExit(e);
    }
}
