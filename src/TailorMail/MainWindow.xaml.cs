using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Threading;
using CommunityToolkit.Mvvm.ComponentModel;
using TailorMail.Views;

namespace TailorMail;

public partial class MainWindow
{
    private int _currentStep;
    private int _previousStep = -1;

    private enum NavDirection { Forward, Back }
    private NavDirection _navDirection = NavDirection.Forward;

    private ObservableCollection<StepItem> _steps = null!;

    private RecipientsPage? _step1;
    private VariablesPage? _step2;
    private MailComposePage? _step3;
    private AttachmentPage? _step4;
    private PreviewPage? _step5;
    private SendPage? _step6;

    private static readonly string[] StepNames = ["收件选择", "变量配置", "模板撰写", "附件匹配", "效果预览", "批量发送"];

    private static readonly string[] StepDescs = [
        "选择收件对象",
        "定义模板变量，发送时自动替换",
        "编辑邮件主题和正文",
        "配置公共附件和专有附件",
        "查看每个收件人的邮件效果",
        "确认后开始批量发送"
    ];

    private static readonly string[] StepShortcuts = ["Ctrl+F 搜索", "", "Ctrl+B 粗体 · Ctrl+I 斜体", "", "Ctrl+F 搜索 · ↑ ↓ 切换", ""];

    private static string ShortcutHintForStep(int step)
    {
        var hint = StepShortcuts[step];
        return string.IsNullOrEmpty(hint) ? "" : hint;
    }

    private readonly HashSet<int> _visitedSteps = [];

    private StepCompletionCache? _stepCompletionCache;
    private int _cacheCurrentStep = -1;

    private class StepCompletionCache
    {
        public int SelectedCount { get; set; }
        public int SelectedWithVariablesCount { get; set; }
        public bool HasSubject { get; set; }
        public bool HasBody { get; set; }
        public int CommonAttachments { get; set; }
        public int SpecialAttachments { get; set; }
        public DateTime CachedAt { get; set; }
    }

    private void InvalidateStepCompletionCache()
    {
        _stepCompletionCache = null;
    }

    private void EnsureStepCompletionCache()
    {
        if (_stepCompletionCache != null && _cacheCurrentStep == _currentStep)
            return;

        var settings = App.DataService.LoadSettings();
        var groups = App.DataService.LoadRecipientGroups();
        var allRecipients = groups.SelectMany(g => g.Recipients).ToList();
        var selectedRecipients = allRecipients.Where(r => r.IsSelected).ToList();
        var attachConfig = App.DataService.LoadAttachmentConfig();

        _stepCompletionCache = new StepCompletionCache
        {
            SelectedCount = selectedRecipients.Count,
            SelectedWithVariablesCount = selectedRecipients.Count(r => r.Variables.Count > 0),
            HasSubject = !string.IsNullOrEmpty(settings.LastSubject),
            HasBody = !string.IsNullOrEmpty(settings.LastBody),
            CommonAttachments = attachConfig.CommonAttachments.Count,
            SpecialAttachments = attachConfig.RecipientAttachments.Count(ua => ua.Files.Count > 0),
            CachedAt = DateTime.Now
        };
        _cacheCurrentStep = _currentStep;
    }

    public MainWindow()
    {
        InitializeComponent();
        Loaded += OnLoaded;
        Closing += OnMainWindowClosing;
        KeyDown += OnWindowKeyDown;
    }

    private void OnWindowKeyDown(object sender, KeyEventArgs e)
    {
        if (Keyboard.Modifiers == ModifierKeys.Control)
        {
            switch (e.Key)
            {
                case Key.S:
                    e.Handled = true;
                    SaveCurrentStep();
                    break;
                case Key.N:
                    e.Handled = true;
                    BtnNewMail_Click(sender, e);
                    break;
                case Key.F:
                    e.Handled = true;
                    FocusCurrentSearch();
                    break;
            }
        }
        else if (Keyboard.Modifiers == ModifierKeys.Alt && e.Key == Key.Left)
        {
            e.Handled = true;
            SaveCurrentStep();
            _navDirection = NavDirection.Back;
            if (_currentStep > 0) NavigateToStep(_currentStep - 1);
        }
        else if (Keyboard.Modifiers == ModifierKeys.Alt && e.Key == Key.Right)
        {
            e.Handled = true;
            SaveCurrentStep();
            _navDirection = NavDirection.Forward;
            if (_currentStep < 5) NavigateToStep(_currentStep + 1);
        }
        else if (e.Key == Key.F1)
        {
            e.Handled = true;
            BtnAbout_Click(sender, e);
        }
        else if (e.Key == Key.Escape)
        {
            var owned = OwnedWindows.Cast<Window>().FirstOrDefault(w => w.IsVisible);
            if (owned != null)
            {
                e.Handled = true;
                owned.Close();
            }
        }
    }

    private void FocusCurrentSearch()
    {
        switch (_currentStep)
        {
            case 0: _step1?.FocusSearch(); break;
            case 4: _step5?.FocusSearch(); break;
        }
    }

    private void OnLoaded(object sender, RoutedEventArgs e)
    {
        _steps = new ObservableCollection<StepItem>(
            StepNames.Select((name, i) => new StepItem(i, name)));
        StepItems.ItemsSource = _steps;

        RestoreWindowPosition();

        var settings = App.DataService.LoadSettings();
        if (!settings.HasSeenWelcome)
        {
            settings.HasSeenWelcome = true;
            App.DataService.SaveSettings(settings);
            ShowWelcome();
        }
        else
        {
            NavigateToStep(0);
        }
    }

    private void ShowSkeletonForStep(int step)
    {
        SkeletonStep0.Visibility = Visibility.Collapsed;
        SkeletonStep1.Visibility = Visibility.Collapsed;
        SkeletonStep3.Visibility = Visibility.Collapsed;
        SkeletonStep4.Visibility = Visibility.Collapsed;
        SkeletonGeneric.Visibility = Visibility.Collapsed;

        switch (step)
        {
            case 0: SkeletonStep0.Visibility = Visibility.Visible; break;
            case 1: SkeletonStep1.Visibility = Visibility.Visible; break;
            case 3: SkeletonStep3.Visibility = Visibility.Visible; break;
            case 4: SkeletonStep4.Visibility = Visibility.Visible; break;
            default: SkeletonGeneric.Visibility = Visibility.Visible; break;
        }
    }

    private void ShowWelcome()
    {
        var welcomePage = new WelcomePage();
        MainContent.Content = welcomePage;
        SkeletonHost.Visibility = Visibility.Collapsed;
        MainContent.Visibility = Visibility.Visible;
        MainContent.Opacity = 1;
        MainContentTransform.Y = 0;
        TxtStepHint.Text = "欢迎";
        TxtStepDesc.Text = "了解 TailorMail 功能流程";
        TxtShortcutHint.Text = "";
        BtnPrev.Visibility = Visibility.Collapsed;
        BtnSend.Visibility = Visibility.Collapsed;
        BtnStop.Visibility = Visibility.Collapsed;
        BtnNext.Visibility = Visibility.Collapsed;

        welcomePage.StartClicked += () =>
        {
            _currentStep = 0;
            _visitedSteps.Add(0);
            NavigateToStep(0);
        };
    }

    private void OnStepClick(object sender, RoutedEventArgs e)
    {
        if (sender is Button btn && btn.Tag is int step)
        {
            SaveCurrentStep();
            _navDirection = step >= _currentStep ? NavDirection.Forward : NavDirection.Back;
            NavigateToStep(step);
        }
    }

    private void NavigateToStep(int step)
    {
        if (_previousStep >= 0 && MainContent.Visibility == Visibility.Visible && MainContent.Opacity > 0)
        {
            var settings = App.DataService.LoadSettings();
            if (!settings.ReducedMotion)
            {
                var fadeOut = new DoubleAnimation(1, 0, TimeSpan.FromMilliseconds(120))
                {
                    EasingFunction = new CubicEase { EasingMode = EasingMode.EaseIn }
                };
                fadeOut.Completed += (_, _) => LoadStepContent(step);
                MainContent.BeginAnimation(UIElement.OpacityProperty, fadeOut);
                _previousStep = _currentStep;
                _currentStep = step;
                _visitedSteps.Add(step);
                UpdateStepIndicator();
                UpdateButtons();
                TxtStepHint.Text = $"步骤 {step + 1}/6 · {StepNames[step]}";
                TxtShortcutHint.Text = ShortcutHintForStep(step);
                UnsubscribeDynamicDesc();
                return;
            }
        }

        _previousStep = _currentStep;
        _currentStep = step;
        _visitedSteps.Add(step);
        UpdateStepIndicator();
        UpdateButtons();
        TxtStepHint.Text = $"步骤 {step + 1}/6 · {StepNames[step]}";
        TxtShortcutHint.Text = ShortcutHintForStep(step);
        UnsubscribeDynamicDesc();
        LoadStepContent(step);
    }

    private void LoadStepContent(int step)
    {
        SkeletonHost.Visibility = Visibility.Visible;
        MainContent.Visibility = Visibility.Collapsed;
        ShowSkeletonForStep(step);

        Dispatcher.BeginInvoke(() =>
        {
            switch (step)
            {
                case 0: _step1 ??= new RecipientsPage(); _step1.RefreshData(); MainContent.Content = _step1; break;
                case 1: _step2 ??= new VariablesPage(); _step2.RefreshData(); MainContent.Content = _step2; break;
                case 2: _step3 ??= new MailComposePage(); _step3.RefreshData(); MainContent.Content = _step3; break;
                case 3: _step4 ??= new AttachmentPage(); _step4.RefreshData(); MainContent.Content = _step4; break;
                case 4: _step5 ??= new PreviewPage(); _step5.RefreshData(); MainContent.Content = _step5; break;
                case 5: _step6 ??= new SendPage(); _step6.RefreshData(); MainContent.Content = _step6; break;
            }

            SubscribeDynamicDesc();
            UpdateStepDesc();

            SkeletonHost.Visibility = Visibility.Collapsed;
            MainContent.Visibility = Visibility.Visible;
            AnimateContentIn();
        }, System.Windows.Threading.DispatcherPriority.Loaded);
    }

    private void UnsubscribeDynamicDesc()
    {
        if (_step1 is IDynamicStepDesc d1) d1.StepDescriptionChanged -= OnDynamicDescChanged;
        if (_step6 is IDynamicStepDesc d6) d6.StepDescriptionChanged -= OnDynamicDescChanged;
        if (_step6 != null) _step6.SendStateChanged -= OnSendStateChanged;
    }

    private void SubscribeDynamicDesc()
    {
        if (_currentStep == 0 && _step1 is IDynamicStepDesc d1) d1.StepDescriptionChanged += OnDynamicDescChanged;
        if (_currentStep == 5 && _step6 is IDynamicStepDesc d6) d6.StepDescriptionChanged += OnDynamicDescChanged;
        if (_step6 != null) _step6.SendStateChanged += OnSendStateChanged;
    }

    private void OnDynamicDescChanged()
    {
        Dispatcher.BeginInvoke(() => UpdateStepDesc());
    }

    private void UpdateStepDesc()
    {
        if (_currentStep == 5 && _step6 is IDynamicStepDesc d5)
        {
            var desc = d5.GetStepDescription();
            TxtStepDesc.Inlines.Clear();
            ParseAndRenderLinks(desc);
        }
        else if (_currentStep == 0 && _step1 is IDynamicStepDesc d0)
        {
            TxtStepDesc.Text = d0.GetStepDescription();
        }
        else
        {
            TxtStepDesc.Text = StepDescs[_currentStep];
        }
    }

    private void ParseAndRenderLinks(string desc)
    {
        int pos = 0;
        while (pos < desc.Length)
        {
            var linkStart = desc.IndexOf("[LINK:", pos, StringComparison.Ordinal);
            if (linkStart < 0)
            {
                TxtStepDesc.Inlines.Add(new Run(desc[pos..]));
                return;
            }

            if (linkStart > pos)
                TxtStepDesc.Inlines.Add(new Run(desc[pos..linkStart]));

            var linkEnd = desc.IndexOf(']', linkStart + 6);
            if (linkEnd < 0)
            {
                TxtStepDesc.Inlines.Add(new Run(desc[linkStart..]));
                return;
            }

            var linkText = desc[(linkStart + 6)..linkEnd];
            var hyperlink = new Hyperlink(new Run(linkText))
            {
                NavigateUri = new Uri("tailormail://settings"),
                Foreground = (Brush)FindResource("AccentBrush"),
                TextDecorations = TextDecorations.Underline
            };
            hyperlink.RequestNavigate += OnSettingsLinkClick;
            TxtStepDesc.Inlines.Add(hyperlink);

            pos = linkEnd + 1;
        }
    }

    private void OnSettingsLinkClick(object sender, System.Windows.Navigation.RequestNavigateEventArgs e)
    {
        if (sender is Hyperlink hl)
        {
            hl.Foreground = (Brush)FindResource("AccentHoverBrush");
            var timer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(300) };
            timer.Tick += (_, _) =>
            {
                timer.Stop();
                hl.Foreground = (Brush)FindResource("AccentBrush");
            };
            timer.Start();
        }
        BtnSettings_Click(sender, new RoutedEventArgs());
    }

    private void UpdateButtons()
    {
        BtnPrev.Visibility = _currentStep > 0 ? Visibility.Visible : Visibility.Collapsed;
        BtnSend.Visibility = _currentStep == 5 ? Visibility.Visible : Visibility.Collapsed;
        BtnStop.Visibility = Visibility.Collapsed;
        BtnNext.Visibility = _currentStep == 5 ? Visibility.Collapsed : Visibility.Visible;

        if (_currentStep == 5)
        {
            UpdateSendButton();
        }
        else
        {
            BtnNext.Content = "下一步";
            BtnNext.Appearance = Wpf.Ui.Controls.ControlAppearance.Primary;
            BtnNext.IsEnabled = true;
        }
    }

    private void UpdateStepIndicator()
    {
        for (int i = 0; i < _steps.Count; i++)
        {
            if (i == _currentStep)
                _steps[i].State = StepState.Current;
            else if (IsStepCompleted(i))
                _steps[i].State = StepState.Completed;
            else if (_visitedSteps.Contains(i))
                _steps[i].State = StepState.Visited;
            else
                _steps[i].State = StepState.Upcoming;
        }

        Dispatcher.BeginInvoke(DispatcherPriority.Loaded, UpdateAccentBar);
    }

    private bool IsStepCompleted(int step)
    {
        EnsureStepCompletionCache();
        var cache = _stepCompletionCache!;

        return step switch
        {
            0 => cache.SelectedCount > 0,
            1 => cache.SelectedCount > 0 && cache.SelectedWithVariablesCount > 0,
            2 => cache.HasSubject || cache.HasBody,
            3 => cache.CommonAttachments > 0 || cache.SpecialAttachments > 0,
            4 => cache.SelectedCount > 0,
            5 => false,
            _ => false
        };
    }

    private double _currentAccentCenter = 0.5;
    private double _targetAccentCenter = 0.5;
    private System.Windows.Threading.DispatcherTimer? _accentBarTimer;

    private void UpdateAccentBar()
    {
        if (AccentBar == null || StepItems.ItemContainerGenerator.Status != GeneratorStatus.ContainersGenerated)
            return;

        var container = StepItems.ItemContainerGenerator.ContainerFromIndex(_currentStep);
        if (container is not ContentPresenter contentContainer)
        {
            return;
        }
        var button = FindVisualChild<Button>(contentContainer);
        if (button == null)
            return;

        try
        {
            var pillBounds = button.TransformToAncestor(this).TransformBounds(new Rect(0, 0, button.ActualWidth, button.ActualHeight));
            var barBounds = AccentBar.TransformToAncestor(this).TransformBounds(new Rect(0, 0, AccentBar.ActualWidth, AccentBar.ActualHeight));

            var pillCenterScreenX = pillBounds.Left + pillBounds.Width / 2;
            var barLeftScreenX = barBounds.Left;
            var barWidth = barBounds.Width;

            if (barWidth <= 0) return;
            var centerX = (pillCenterScreenX - barLeftScreenX) / barWidth;
            centerX = Math.Max(0.05, Math.Min(0.95, centerX));

            _targetAccentCenter = centerX;
            StartAccentBarAnimation();
        }
        catch
        {
        }
    }

    private void StartAccentBarAnimation()
    {
        var settings = App.DataService.LoadSettings();
        if (settings.ReducedMotion)
        {
            _currentAccentCenter = _targetAccentCenter;
            ApplyAccentBarGradient(_currentAccentCenter);
            return;
        }

        _accentBarTimer?.Stop();
        var startCenter = _currentAccentCenter;
        var targetCenter = _targetAccentCenter;
        var startTime = DateTime.UtcNow;
        var durationMs = 300.0;

        _accentBarTimer = new System.Windows.Threading.DispatcherTimer
        {
            Interval = TimeSpan.FromMilliseconds(16)
        };
        _accentBarTimer.Tick += (_, _) =>
        {
            var elapsed = (DateTime.UtcNow - startTime).TotalMilliseconds;
            var progress = Math.Min(1.0, elapsed / durationMs);
            var eased = 1 - Math.Pow(1 - progress, 3);
            _currentAccentCenter = startCenter + (targetCenter - startCenter) * eased;
            ApplyAccentBarGradient(_currentAccentCenter);
            if (progress >= 1.0)
            {
                _accentBarTimer.Stop();
                _currentAccentCenter = targetCenter;
            }
        };
        _accentBarTimer.Start();
    }

    private void ApplyAccentBarGradient(double centerX)
    {
        var halfSpread = 0.44;
        var midLeft = Math.Max(0, centerX - halfSpread * 0.5);
        var midRight = Math.Min(1, centerX + halfSpread * 0.5);
        AccentBar.Background = new LinearGradientBrush
        {
            StartPoint = new Point(0, 0.5),
            EndPoint = new Point(1, 0.5),
            GradientStops = new GradientStopCollection
            {
                new GradientStop(Colors.Transparent, Math.Max(0, centerX - halfSpread)),
                new GradientStop((Color)FindResource("AccentLightColor"), midLeft),
                new GradientStop((Color)FindResource("AccentColor"), centerX),
                new GradientStop((Color)FindResource("AccentLightColor"), midRight),
                new GradientStop(Colors.Transparent, Math.Min(1, centerX + halfSpread))
            }
        };
    }

    private static T? FindVisualChild<T>(DependencyObject parent) where T : DependencyObject
    {
        for (int i = 0; i < VisualTreeHelper.GetChildrenCount(parent); i++)
        {
            var child = VisualTreeHelper.GetChild(parent, i);
            if (child is T result)
                return result;
            var descendant = FindVisualChild<T>(child);
            if (descendant != null)
                return descendant;
        }
        return null;
    }

    private void BtnPrev_Click(object sender, RoutedEventArgs e)
    {
        SaveCurrentStep();
        _navDirection = NavDirection.Back;
        if (_currentStep > 0) NavigateToStep(_currentStep - 1);
    }

    private void BtnNext_Click(object sender, RoutedEventArgs e)
    {
        SaveCurrentStep();
        _navDirection = NavDirection.Forward;
        if (_currentStep < 5) NavigateToStep(_currentStep + 1);
    }

    private void OnSendStateChanged()
    {
        Dispatcher.BeginInvoke(() => UpdateSendButton());
    }

    private void UpdateSendButton()
    {
        if (_currentStep != 5 || _step6 == null) return;

        if (_step6.IsSending)
        {
            BtnSend.Content = "发送中...";
            BtnSend.Appearance = Wpf.Ui.Controls.ControlAppearance.Secondary;
            BtnSend.IsEnabled = false;
            BtnStop.Visibility = Visibility.Visible;
            BtnStop.IsEnabled = true;
        }
        else
        {
            BtnSend.IsEnabled = true;
            BtnStop.Visibility = Visibility.Collapsed;
            if (_step6.HasPendingRecipients)
            {
                BtnSend.Content = _step6.HasAnyResults ? "继续发送" : "开始发送";
                BtnSend.Appearance = Wpf.Ui.Controls.ControlAppearance.Primary;
            }
            else
            {
                BtnSend.Content = "已完成";
                BtnSend.Appearance = Wpf.Ui.Controls.ControlAppearance.Secondary;
                BtnSend.IsEnabled = false;
            }
        }
    }

    private void BtnSend_Click(object sender, RoutedEventArgs e)
    {
        if (_currentStep == 5 && _step6 != null)
        {
            _step6.StartSend();
        }
    }

    private void BtnStop_Click(object sender, RoutedEventArgs e)
    {
        if (_currentStep == 5 && _step6 != null)
        {
            _step6.StopSend();
        }
    }

    private bool HasEditedContent()
    {
        var settings = App.DataService.LoadSettings();
        return !string.IsNullOrEmpty(settings.LastSubject) ||
               !string.IsNullOrEmpty(settings.LastBody) ||
               !string.IsNullOrEmpty(settings.LastBodyXaml);
    }

    private void BtnNewMail_Click(object sender, RoutedEventArgs e)
    {
        if (HasEditedContent())
        {
            var dlg = new Views.ConfirmDialog { Title = "确认", Message = "新建邮件将清除已撰写内容，是否继续？" };
            if (dlg.ShowDialog() != true) return;
        }

        SaveCurrentStep();
        var settings = App.DataService.LoadSettings();
        settings.LastSubject = string.Empty;
        settings.LastBody = string.Empty;
        settings.LastBodyXaml = string.Empty;
        App.DataService.SaveSettings(settings);

        App.DataService.SaveAttachmentConfig(new Models.AttachmentConfig());

        _step3 = null;
        _step4 = null;
        _step5 = null;
        _step6 = null;

        NavigateToStep(1);
    }

    private void BtnSettings_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            var win = new SettingsWindow { Owner = this };
            win.ShowDialog();
        }
        catch (Exception ex)
        {
            Services.AppLogger.Error("打开设置窗口失败", ex);
            var detail = ex.ToString();
            if (ex.InnerException != null)
                detail += "\n\n--- InnerException ---\n" + ex.InnerException.ToString();
            App.ShowError("打开设置窗口失败，请查看日志");
        }
    }

    private void BtnAbout_Click(object sender, RoutedEventArgs e)
    {
        ShowWelcome();
    }

    private void SaveCurrentStep()
    {
        switch (_currentStep)
        {
            case 0: _step1?.SaveAll(); break;
            case 1: _step2?.SaveAll(); break;
            case 2: _step3?.SaveCurrent(); break;
        }
    }

    private void AnimateContentIn()
    {
        var settings = App.DataService.LoadSettings();
        if (settings.ReducedMotion)
        {
            MainContent.Opacity = 1;
            MainContentTransform.Y = 0;
            MainContentTransform.X = 0;
            return;
        }

        double offsetX = _navDirection == NavDirection.Forward ? 40 : -40;
        MainContent.Opacity = 0;
        MainContentTransform.X = offsetX;
        MainContentTransform.Y = 0;

        var duration = TimeSpan.FromMilliseconds(DesignTokens.DurationNormal);
        var opacityAnimation = new DoubleAnimation(0, 1, duration)
        {
            EasingFunction = new CubicEase { EasingMode = EasingMode.EaseOut }
        };
        var slideAnimation = new DoubleAnimation(offsetX, 0, duration)
        {
            EasingFunction = new CubicEase { EasingMode = EasingMode.EaseOut }
        };

        MainContent.BeginAnimation(UIElement.OpacityProperty, opacityAnimation);
        MainContentTransform.BeginAnimation(TranslateTransform.XProperty, slideAnimation);
    }

    private void OnMainWindowClosing(object? sender, CancelEventArgs e)
    {
        SaveCurrentStep();

        if (_step6 is { IsSending: true })
        {
            var dlg = new ConfirmDialog { Title = "确认", Message = "正在发送邮件，确定要关闭窗口吗？" };
            if (dlg.ShowDialog() != true)
            {
                e.Cancel = true;
                return;
            }
        }

        SaveWindowPosition();
    }

    private void RestoreWindowPosition()
    {
        var settings = App.DataService.LoadSettings();
        if (settings.WindowLeft.HasValue && settings.WindowTop.HasValue)
        {
            Left = settings.WindowLeft.Value;
            Top = settings.WindowTop.Value;
        }
        if (settings.WindowWidth.HasValue && settings.WindowWidth.Value > 0)
            Width = settings.WindowWidth.Value;
        if (settings.WindowHeight.HasValue && settings.WindowHeight.Value > 0)
            Height = settings.WindowHeight.Value;
        if (settings.WindowState == "Maximized")
            WindowState = System.Windows.WindowState.Maximized;
    }

    private void SaveWindowPosition()
    {
        var settings = App.DataService.LoadSettings();
        if (WindowState == System.Windows.WindowState.Maximized)
        {
            settings.WindowState = "Maximized";
            var bounds = RestoreBounds;
            settings.WindowLeft = bounds.Left;
            settings.WindowTop = bounds.Top;
            settings.WindowWidth = bounds.Width;
            settings.WindowHeight = bounds.Height;
        }
        else
        {
            settings.WindowState = "Normal";
            settings.WindowLeft = Left;
            settings.WindowTop = Top;
            settings.WindowWidth = Width;
            settings.WindowHeight = Height;
        }
        App.DataService.SaveSettings(settings);
    }
}

public enum StepState { Upcoming, Current, Completed, Visited }

public class StepItem : ObservableObject
{
    public int Index { get; }
    public string Label { get; }

    private StepState _state;
    public StepState State
    {
        get => _state;
        set => SetProperty(ref _state, value);
    }

    public string DisplayIndex => State == StepState.Completed ? "✓" : (Index + 1).ToString();

    public StepItem(int index, string label)
    {
        Index = index;
        Label = label;
    }

    protected override void OnPropertyChanged(PropertyChangedEventArgs e)
    {
        base.OnPropertyChanged(e);
        if (e.PropertyName == nameof(State))
            OnPropertyChanged(nameof(DisplayIndex));
    }
}

public interface IRefreshable
{
    void RefreshData();
}

public interface IDynamicStepDesc
{
    string GetStepDescription();
    event Action? StepDescriptionChanged;
}
