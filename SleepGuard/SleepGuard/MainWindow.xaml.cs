using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Shapes;
using Microsoft.Win32;

namespace SleepGuard;

public partial class MainWindow : Window
{
    private readonly MonitorService _monitor;
    private bool _suppressSettingsChange;
    private List<ProcessInfo> _allRunningProcesses = new();

    // ブラシキャッシュ（Freeze済みでGC負荷軽減・スレッドセーフ）
    private static readonly Dictionary<string, SolidColorBrush> _brushCache = new();
    private static SolidColorBrush Brush(string hex)
    {
        if (!_brushCache.TryGetValue(hex, out var b))
        {
            b = new SolidColorBrush((Color)ColorConverter.ConvertFromString(hex));
            b.Freeze();
            _brushCache[hex] = b;
        }
        return b;
    }

    // FontFamilyキャッシュ
    private static readonly FontFamily _consolasFont = new("Consolas");

    public MainWindow(MonitorService monitor)
    {
        InitializeComponent();
        _monitor = monitor;

        // イベントはInitializeComponent後にコードで登録（XAMLパース中の発火を防ぐ）
        SleepDelaySlider.ValueChanged    += SleepDelaySlider_ValueChanged;
        CheckIntervalSlider.ValueChanged += CheckIntervalSlider_ValueChanged;
        PreventDisplayChk.Checked        += Settings_Changed;
        PreventDisplayChk.Unchecked      += Settings_Changed;
        StartMinimizedChk.Checked        += Settings_Changed;
        StartMinimizedChk.Unchecked      += Settings_Changed;
        ResidentModeChk.Checked          += Settings_Changed;
        ResidentModeChk.Unchecked        += Settings_Changed;
        LogEnabledChk.Checked            += Settings_Changed;
        LogEnabledChk.Unchecked          += Settings_Changed;
        StartupChk.Checked               += Startup_Changed;
        StartupChk.Unchecked             += Startup_Changed;
        ProcessSearchBox.TextChanged     += ProcessSearchBox_TextChanged;

        _monitor.StatusUpdated += OnStatusUpdated;

        // 3行スクロール設定 (1行約42px × 3 = 126px)
        const double threeLineScroll = 126.0;
        ProcessListScrollViewer.ScrollChanged += (_, _) => { };
        ProcessListScrollViewer.PreviewMouseWheel += (s, e) =>
        {
            var sv = (ScrollViewer)s;
            sv.ScrollToVerticalOffset(sv.VerticalOffset - Math.Sign(e.Delta) * threeLineScroll);
            e.Handled = true;
        };
        RunningListScrollViewer.PreviewMouseWheel += (s, e) =>
        {
            var sv = (ScrollViewer)s;
            sv.ScrollToVerticalOffset(sv.VerticalOffset - Math.Sign(e.Delta) * threeLineScroll);
            e.Handled = true;
        };
        LogScrollViewer.PreviewMouseWheel += (s, e) =>
        {
            var sv = (ScrollViewer)s;
            sv.ScrollToVerticalOffset(sv.VerticalOffset - Math.Sign(e.Delta) * threeLineScroll);
            e.Handled = true;
        };

        LoadSettings();
        RefreshProcessList();
        UpdateStatusUI(new MonitorStatus());
        LoadRunningProcesses();
    }

    // ─── Settings ────────────────────────────────────────────────
    private void LoadSettings()
    {
        _suppressSettingsChange = true;
        var s = SettingsManager.Load();
        SleepDelaySlider.Value      = s.SleepDelayMinutes;
        CheckIntervalSlider.Value   = s.CheckIntervalSeconds;
        PreventDisplayChk.IsChecked = s.PreventDisplaySleep;
        StartMinimizedChk.IsChecked = s.StartMinimized;
        ResidentModeChk.IsChecked   = s.ResidentMode;
        LogEnabledChk.IsChecked     = s.LogEnabled;
        StartupChk.IsChecked        = StartupManager.IsRegistered();
        SleepDelayLabel.Text        = $"{s.SleepDelayMinutes} 分";
        CheckIntervalLabel.Text     = $"{s.CheckIntervalSeconds} 秒";
        StatDelay.Text              = $"{s.SleepDelayMinutes} 分";
        _suppressSettingsChange = false;
    }

    private bool IsUiReady =>
        SleepDelaySlider  != null && CheckIntervalSlider != null &&
        PreventDisplayChk != null && StartMinimizedChk   != null &&
        ResidentModeChk   != null && LogEnabledChk        != null;

    private void SaveCurrentSettings()
    {
        if (!IsUiReady) return;
        // WatchedProcessesは変更しないのでマージのためLoadは1回のみ
        var s = SettingsManager.Load();
        s.SleepDelayMinutes    = (int)SleepDelaySlider.Value;
        s.CheckIntervalSeconds = (int)CheckIntervalSlider.Value;
        s.PreventDisplaySleep  = PreventDisplayChk.IsChecked == true;
        s.StartMinimized       = StartMinimizedChk.IsChecked  == true;
        s.ResidentMode         = ResidentModeChk.IsChecked    == true;
        s.LogEnabled           = LogEnabledChk.IsChecked      == true;
        SettingsManager.Save(s);
        _monitor.ReloadSettings();
    }

    // ─── Status update ───────────────────────────────────────────
    private void OnStatusUpdated(MonitorStatus status)
    {
        // BeginInvokeで非同期ディスパッチ（タイマースレッドをブロックしない）
        Dispatcher.BeginInvoke(() => UpdateStatusUI(status));
    }

    private void UpdateStatusUI(MonitorStatus status)
    {
        // MonitorStatusの値を直接使用（ディスクI/Oなし）
        StatRunningCount.Text = status.RunningProcesses.Count.ToString();
        StatWatchedCount.Text = $"/ {status.WatchedCount} 登録";
        StatDelay.Text        = $"{status.DelayMinutes} 分";

        if (status.IsSleepPrevented)
        {
            SetStatusBadge("スリープ防止中", "#22D3A5", "#1522D3A5", "#5022D3A5");
            StatStatusText.Text       = "防止中";
            StatStatusText.Foreground = Brush("#22D3A5");
            StatStatusSub.Text        = "スリープをブロック中";
            StatStatusSub.Foreground  = Brush("#1A9E7A");
        }
        else if (status.CountdownStartAt.HasValue)
        {
            SetStatusBadge("カウントダウン中", "#F59E0B", "#15F59E0B", "#50F59E0B");
            StatStatusText.Text       = "猶予中";
            StatStatusText.Foreground = Brush("#F59E0B");
            StatStatusSub.Text        = "解除まで待機中";
            StatStatusSub.Foreground  = Brush("#A87010");
        }
        else
        {
            SetStatusBadge("待機中", "#B0AABF", "#20FFFFFF", "#33FFFFFF");
            StatStatusText.Text       = "待機中";
            StatStatusText.Foreground = Brush("#706C90");
            StatStatusSub.Text        = "監視中";
            StatStatusSub.Foreground  = Brush("#706C90");
        }

        if (status.CountdownStartAt.HasValue && status.RemainingSeconds.HasValue)
        {
            CountdownBanner.Visibility = Visibility.Visible;
            var rem = TimeSpan.FromSeconds(status.RemainingSeconds.Value);
            CountdownText.Text =
                $"スリープ解除まで {rem.Minutes:D2}:{rem.Seconds:D2} — 全監視プロセスが停止しています";
        }
        else
        {
            CountdownBanner.Visibility = Visibility.Collapsed;
        }

        RunningChipsPanel.Children.Clear();
        foreach (var name in status.RunningProcesses)
            RunningChipsPanel.Children.Add(CreateChip(name));

        // 監視リストの実行状態をStatusから更新（追加のプロセス列挙なし）
        RefreshProcessListWithRunning(status.RunningProcesses);
    }

    private void SetStatusBadge(string text, string textHex, string bgHex, string borderHex)
    {
        StatusText.Text         = text;
        StatusText.Foreground   = Brush(textHex);
        StatusDot.Fill          = Brush(textHex);
        StatusBadge.Background  = Brush(bgHex);
        StatusBadge.BorderBrush = Brush(borderHex);
    }

    // ─── Watched process list ────────────────────────────────────
    /// <summary>プロセス追加・削除時など、設定から再構築する場合に使用</summary>
    private void RefreshProcessList()
    {
        var runningNames = _monitor.GetRunningWatchedProcesses();
        var settings     = SettingsManager.Load();
        RenderProcessList(settings, runningNames);
    }

    /// <summary>StatusUpdateから呼ぶ場合。Loadは1回で済む。</summary>
    private void RefreshProcessListWithRunning(List<string> runningNames)
    {
        var settings = SettingsManager.Load();
        RenderProcessList(settings, runningNames);
    }

    private void RenderProcessList(AppSettings settings, List<string> runningNames)
    {
        ProcessListBox.Items.Clear();
        EmptyState.Visibility  = settings.WatchedProcesses.Count == 0
            ? Visibility.Visible : Visibility.Collapsed;
        ProcessCountLabel.Text = $"{settings.WatchedProcesses.Count} 件";
        foreach (var wp in settings.WatchedProcesses)
            ProcessListBox.Items.Add(
                CreateProcessItem(wp, runningNames.Contains(wp.DisplayName)));
    }

    private UIElement CreateProcessItem(WatchedProcess wp, bool isRunning)
    {
        var border = new Border
        {
            CornerRadius    = new CornerRadius(8),
            BorderThickness = new Thickness(1),
            Padding         = new Thickness(14, 10, 14, 10),
            BorderBrush     = isRunning ? Brush("#4022D3A5") : Brush("#28FFFFFF"),
            Background      = isRunning ? Brush("#0E22D3A5") : Brush("#12101E"),
        };
        var grid = new Grid();
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(40) });
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

        var iconBd = new Border
        {
            Width = 32, Height = 32, CornerRadius = new CornerRadius(7),
            Background = Brush("#252235"), Margin = new Thickness(0, 0, 10, 0)
        };
        iconBd.Child = new TextBlock
        {
            Text = wp.DisplayName.Length > 0
                ? wp.DisplayName[0].ToString().ToUpper() : "?",
            Foreground          = Brush("#9A96B8"),
            FontSize            = 14,
            FontWeight          = FontWeights.Bold,
            HorizontalAlignment = HorizontalAlignment.Center,
            VerticalAlignment   = VerticalAlignment.Center
        };
        Grid.SetColumn(iconBd, 0);

        var info = new StackPanel { VerticalAlignment = VerticalAlignment.Center };
        info.Children.Add(new TextBlock
        {
            Text       = wp.DisplayName,
            Foreground = Brush("#EAE6F5"),
            FontSize   = 13,
            FontWeight = FontWeights.SemiBold
        });
        info.Children.Add(new TextBlock
        {
            Text       = wp.ProcessName + ".exe",
            Foreground = Brush("#706C90"),
            FontFamily = _consolasFont,
            FontSize   = 10,
            Margin     = new Thickness(0, 2, 0, 0)
        });
        Grid.SetColumn(info, 1);

        var chip = new Border
        {
            CornerRadius      = new CornerRadius(10),
            BorderThickness   = new Thickness(1),
            Padding           = new Thickness(8, 3, 8, 3),
            Margin            = new Thickness(0, 0, 8, 0),
            VerticalAlignment = VerticalAlignment.Center,
            Background        = isRunning ? Brush("#1522D3A5") : Brush("#0AFFFFFF"),
            BorderBrush       = isRunning ? Brush("#4022D3A5") : Brush("#25FFFFFF"),
        };
        chip.Child = new TextBlock
        {
            Text       = isRunning ? "実行中" : "停止中",
            FontFamily = _consolasFont,
            FontSize   = 10,
            Foreground = isRunning ? Brush("#22D3A5") : Brush("#706C90")
        };
        Grid.SetColumn(chip, 2);

        var removeBtn = new Button
        {
            Content         = "×",
            Width           = 28,
            Height          = 28,
            Background      = Brushes.Transparent,
            BorderThickness = new Thickness(0),
            Foreground      = Brush("#706C90"),
            FontSize        = 15,
            Cursor          = Cursors.Hand,
            VerticalAlignment = VerticalAlignment.Center,
            Tag             = wp.Id
        };
        removeBtn.Click      += RemoveProcess_Click;
        removeBtn.MouseEnter += (s, _) => ((Button)s).Foreground = Brush("#F87171");
        removeBtn.MouseLeave += (s, _) => ((Button)s).Foreground = Brush("#706C90");
        Grid.SetColumn(removeBtn, 3);

        grid.Children.Add(iconBd);
        grid.Children.Add(info);
        grid.Children.Add(chip);
        grid.Children.Add(removeBtn);
        border.Child = grid;
        return border;
    }

    private UIElement CreateChip(string name)
    {
        var bd = new Border
        {
            CornerRadius    = new CornerRadius(12),
            BorderThickness = new Thickness(1),
            BorderBrush     = Brush("#4022D3A5"),
            Background      = Brush("#1522D3A5"),
            Padding         = new Thickness(10, 4, 10, 4),
            Margin          = new Thickness(0, 0, 6, 6)
        };
        var sp = new StackPanel { Orientation = Orientation.Horizontal };
        sp.Children.Add(new Ellipse
        {
            Width             = 5,
            Height            = 5,
            Margin            = new Thickness(0, 0, 6, 0),
            Fill              = Brush("#22D3A5"),
            VerticalAlignment = VerticalAlignment.Center
        });
        sp.Children.Add(new TextBlock
        {
            Text              = name,
            Foreground        = Brush("#22D3A5"),
            FontFamily        = _consolasFont,
            FontSize          = 11,
            VerticalAlignment = VerticalAlignment.Center
        });
        bd.Child = sp;
        return bd;
    }

    // ─── Running process picker ───────────────────────────────────
    private record ProcessInfo(string Name, int Pid);

    private void LoadRunningProcesses()
    {
        _allRunningProcesses = Process.GetProcesses()
            .Where(p => { try { return p.SessionId > 0; } catch { return false; } })
            .Select(p => { var n = p.ProcessName; p.Dispose(); return new ProcessInfo(n, 0); })
            .DistinctBy(p => p.Name.ToLowerInvariant())
            .Where(p => !string.IsNullOrWhiteSpace(p.Name))
            .OrderBy(p => p.Name)
            .ToList();
        ApplyProcessFilter(ProcessSearchBox?.Text ?? "");
    }

    private void ApplyProcessFilter(string keyword)
    {
        RunningProcessListBox.Items.Clear();
        var settings   = SettingsManager.Load();
        var registered = settings.WatchedProcesses
            .Select(p => p.ProcessName.ToLowerInvariant())
            .ToHashSet();

        var filtered = string.IsNullOrWhiteSpace(keyword)
            ? _allRunningProcesses
            : _allRunningProcesses.Where(p =>
                p.Name.Contains(keyword, StringComparison.OrdinalIgnoreCase));

        foreach (var pi in filtered)
            RunningProcessListBox.Items.Add(
                CreateRunningProcessRow(pi, registered.Contains(pi.Name.ToLowerInvariant())));

        if (SearchPlaceholder != null)
            SearchPlaceholder.Visibility = string.IsNullOrEmpty(ProcessSearchBox?.Text)
                ? Visibility.Visible : Visibility.Collapsed;
    }

    private UIElement CreateRunningProcessRow(ProcessInfo pi, bool alreadyRegistered)
    {
        var border = new Border
        {
            CornerRadius    = new CornerRadius(7),
            BorderThickness = new Thickness(1),
            BorderBrush     = Brush("#25FFFFFF"),
            Background      = Brush("#12101E"),
            Padding         = new Thickness(12, 8, 12, 8)
        };
        var grid = new Grid();
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

        var nameBlock = new TextBlock
        {
            Text              = pi.Name + ".exe",
            Foreground        = alreadyRegistered ? Brush("#706C90") : Brush("#D8D4EC"),
            FontFamily        = _consolasFont,
            FontSize          = 14,
            VerticalAlignment = VerticalAlignment.Center
        };
        Grid.SetColumn(nameBlock, 0);

        UIElement right;
        if (alreadyRegistered)
        {
            right = new TextBlock
            {
                Text              = "登録済",
                Foreground        = Brush("#706C90"),
                FontFamily        = _consolasFont,
                FontSize          = 10,
                VerticalAlignment = VerticalAlignment.Center
            };
        }
        else
        {
            var btn = new Button { Content = "+ 監視に追加", Tag = pi };
            btn.Style  = (Style)FindResource("AddFromListBtn");
            btn.Click += AddFromRunningProcess_Click;
            right = btn;
        }
        Grid.SetColumn(right, 1);

        grid.Children.Add(nameBlock);
        grid.Children.Add(right);
        border.Child = grid;
        return border;
    }

    private void AddFromRunningProcess_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not Button btn || btn.Tag is not ProcessInfo pi) return;

        var settings = SettingsManager.Load();
        if (settings.WatchedProcesses.Any(p =>
            p.ProcessName.Equals(pi.Name, StringComparison.OrdinalIgnoreCase)))
        {
            AddLogLine($"すでに登録済み: {pi.Name}", "warn");
            return;
        }

        settings.WatchedProcesses.Add(new WatchedProcess
        {
            DisplayName = pi.Name,
            ProcessName = pi.Name,
            FilePath    = string.Empty
        });
        SettingsManager.Save(settings);
        _monitor.ReloadSettings();
        RefreshProcessList();
        ApplyProcessFilter(ProcessSearchBox?.Text ?? "");
        AddLogLine($"追加: {pi.Name} (実行中リストより)");
    }

    private void RefreshRunningProcesses_Click(object sender, RoutedEventArgs e)
    {
        LoadRunningProcesses();
        AddLogLine("実行中プロセス一覧を更新しました");
    }

    private void ProcessSearchBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        ApplyProcessFilter(ProcessSearchBox.Text);
        if (SearchPlaceholder != null)
            SearchPlaceholder.Visibility = string.IsNullOrEmpty(ProcessSearchBox.Text)
                ? Visibility.Visible : Visibility.Collapsed;
    }

    // ─── Log ─────────────────────────────────────────────────────
    private void AddLogLine(string message, string type = "normal")
    {
        var sp = new StackPanel { Orientation = Orientation.Horizontal };
        sp.Children.Add(new TextBlock
        {
            Text       = DateTime.Now.ToString("HH:mm:ss"),
            Foreground = Brush("#5C5880"),
            FontFamily = _consolasFont,
            FontSize   = 11,
            Margin     = new Thickness(0, 0, 10, 0)
        });
        var msgColor = type switch
        {
            "active"   => "#22D3A5",
            "released" => "#F59E0B",
            "error"    => "#F87171",
            "warn"     => "#F59E0B",
            _          => "#C8C4DC"
        };
        sp.Children.Add(new TextBlock
        {
            Text       = message,
            Foreground = Brush(msgColor),
            FontFamily = _consolasFont,
            FontSize   = 11
        });
        LogPanel.Children.Add(sp);
        if (LogPanel.Children.Count > 60)
            LogPanel.Children.RemoveAt(0);
        LogScrollViewer.ScrollToEnd();
    }

    // ─── Event handlers ──────────────────────────────────────────
    private void RootScroll_PreviewMouseWheel(object sender, MouseWheelEventArgs e)
    {
        RootScroll.ScrollToVerticalOffset(RootScroll.VerticalOffset - e.Delta * 0.5);
        e.Handled = true;
    }

    private void BrowseButton_Click(object sender, RoutedEventArgs e)
    {
        var dlg = new OpenFileDialog
        {
            Filter = "実行ファイル (*.exe)|*.exe|すべてのファイル (*.*)|*.*",
            Title  = "監視するプロセスのEXEファイルを選択"
        };
        if (dlg.ShowDialog() == true)
        {
            ExePathBox.Text = dlg.FileName;
            if (string.IsNullOrWhiteSpace(DisplayNameBox.Text))
                DisplayNameBox.Text =
                    System.IO.Path.GetFileNameWithoutExtension(dlg.FileName);
        }
    }

    private void AddProcessButton_Click(object sender, RoutedEventArgs e)
    {
        AddErrorText.Visibility = Visibility.Collapsed;
        var path        = ExePathBox.Text.Trim();
        var displayName = DisplayNameBox.Text.Trim();

        if (string.IsNullOrEmpty(path))
        {
            ShowAddError("EXEファイルのパスを入力してください"); return;
        }
        if (!System.IO.Path.GetFileName(path)
                .EndsWith(".exe", StringComparison.OrdinalIgnoreCase))
        {
            ShowAddError(".exe ファイルを指定してください"); return;
        }

        var processName = System.IO.Path.GetFileNameWithoutExtension(path);
        if (string.IsNullOrEmpty(displayName)) displayName = processName;

        var settings = SettingsManager.Load();
        if (settings.WatchedProcesses.Any(p =>
            p.ProcessName.Equals(processName, StringComparison.OrdinalIgnoreCase)))
        {
            ShowAddError($"'{displayName}' はすでに登録されています"); return;
        }

        settings.WatchedProcesses.Add(new WatchedProcess
        {
            DisplayName = displayName,
            ProcessName = processName,
            FilePath    = path
        });
        SettingsManager.Save(settings);
        _monitor.ReloadSettings();

        ExePathBox.Text = DisplayNameBox.Text = string.Empty;
        RefreshProcessList();
        ApplyProcessFilter(ProcessSearchBox?.Text ?? "");
        AddLogLine($"追加: {displayName} ({processName}.exe)");
    }

    private void RemoveProcess_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not Button btn) return;
        var settings = SettingsManager.Load();
        var target   = settings.WatchedProcesses
            .FirstOrDefault(p => p.Id == btn.Tag as string);
        if (target == null) return;

        settings.WatchedProcesses.Remove(target);
        SettingsManager.Save(settings);
        _monitor.ReloadSettings();
        RefreshProcessList();
        ApplyProcessFilter(ProcessSearchBox?.Text ?? "");
        AddLogLine($"削除: {target.DisplayName}");
    }

    private void SleepDelaySlider_ValueChanged(
        object sender, RoutedPropertyChangedEventArgs<double> e)
    {
        if (SleepDelayLabel == null) return;
        var val = (int)e.NewValue;
        SleepDelayLabel.Text = $"{val} 分";
        if (StatDelay != null) StatDelay.Text = $"{val} 分";
        if (!_suppressSettingsChange && IsUiReady) SaveCurrentSettings();
    }

    private void CheckIntervalSlider_ValueChanged(
        object sender, RoutedPropertyChangedEventArgs<double> e)
    {
        if (CheckIntervalLabel == null) return;
        CheckIntervalLabel.Text = $"{(int)e.NewValue} 秒";
        if (!_suppressSettingsChange && IsUiReady) SaveCurrentSettings();
    }

    private void Settings_Changed(object sender, RoutedEventArgs e)
    {
        if (!_suppressSettingsChange && IsUiReady) SaveCurrentSettings();
    }

    private void Startup_Changed(object sender, RoutedEventArgs e)
    {
        if (_suppressSettingsChange) return;
        if (StartupChk.IsChecked == true)
        {
            StartupManager.Register();
            AddLogLine("スタートアップに登録しました");
        }
        else
        {
            StartupManager.Unregister();
            AddLogLine("スタートアップから削除しました");
        }
    }

    private void OpenLogButton_Click(object sender, RoutedEventArgs e)
    {
        var logPath = SettingsManager.LogFilePath;
        if (File.Exists(logPath))
            Process.Start(new ProcessStartInfo(logPath) { UseShellExecute = true });
        else
            MessageBox.Show("ログファイルがまだ作成されていません。", "SleepGuard",
                MessageBoxButton.OK, MessageBoxImage.Information);
    }

    private void ClearLogButton_Click(object sender, RoutedEventArgs e) =>
        LogPanel.Children.Clear();

    private void ShowAddError(string message)
    {
        AddErrorText.Text       = message;
        AddErrorText.Visibility = Visibility.Visible;
    }

    private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
    {
        // UIのチェックボックス値を参照（ディスクI/Oなし）
        if (ResidentModeChk?.IsChecked == true)
        {
            e.Cancel = true;
            Hide();
        }
        else
        {
            e.Cancel = false;
            ((App)Application.Current).RequestExit();
        }
    }
}
