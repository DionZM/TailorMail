using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Documents;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Threading;
using CommunityToolkit.Mvvm.ComponentModel;
using TailorMail.Views;

namespace TailorMail;

public partial class MainWindow
{
    private int _currentStep;

    private ObservableCollection<StepItem> _steps = null!;

    private RecipientsPage? _step1;
    private VariablesPage? _step2;
    private MailComposePage? _step3;
    private AttachmentPage? _step4;
    private PreviewPage? _step5;
    private SendPage? _step6;

    private static readonly string[] StepNames = ["收件选择", "变量配置", "模板撰写", "附件匹配", "效果预览", "批量发送"];

    private static readonly string[] StepDescs = ["选择收件对象", "定义模板变量，发送时自动替换", "编辑邮件主题和正文", "配置公共附件和专有附件", "查看每个收件人的邮件效果", "确认后开始批量发送"];

    public MainWindow()
    {
        InitializeComponent();
        Loaded += OnLoaded;
        Closed += OnMainWindowClosed;
    }

    private void OnLoaded(object sender, RoutedEventArgs e)
    {
        _steps = new ObservableCollection<StepItem>(
            StepNames.Select((name, i) => new StepItem(i, name)));
        StepItems.ItemsSource = _steps;

        NavigateToStep(0);
    }

    private void OnStepClick(object sender, RoutedEventArgs e)
    {
        if (sender is Button btn && btn.Tag is int step)
        {
            SaveCurrentStep();
            NavigateToStep(step);
        }
    }

    private void NavigateToStep(int step)
    {
        _currentStep = step;
        UpdateStepIndicator();
        UpdateButtons();
        TxtStepHint.Text = $"步骤 {step + 1}/6 · {StepNames[step]}";

        UnsubscribeDynamicDesc();

        SkeletonPanel.Visibility = Visibility.Visible;
        MainContent.Visibility = Visibility.Collapsed;

        Dispatcher.BeginInvoke(() =>
        {
            switch (step)
            {
                case 0: _step1 ??= new RecipientsPage(); MainContent.Content = _step1; break;
                case 1: _step2 ??= new VariablesPage(); _step2.RefreshData(); MainContent.Content = _step2; break;
                case 2: _step3 ??= new MailComposePage(); _step3.RefreshData(); MainContent.Content = _step3; break;
                case 3: _step4 ??= new AttachmentPage(); _step4.RefreshData(); MainContent.Content = _step4; break;
                case 4: _step5 ??= new PreviewPage(); _step5.RefreshData(); MainContent.Content = _step5; break;
                case 5: _step6 ??= new SendPage(); _step6.RefreshData(); MainContent.Content = _step6; break;
            }

            SubscribeDynamicDesc();
            UpdateStepDesc();

            SkeletonPanel.Visibility = Visibility.Collapsed;
            MainContent.Visibility = Visibility.Visible;
            AnimateContentIn();
        }, System.Windows.Threading.DispatcherPriority.Loaded);
    }

    private void UnsubscribeDynamicDesc()
    {
        if (_step1 is IDynamicStepDesc d1) d1.StepDescriptionChanged -= OnDynamicDescChanged;
        if (_step6 is IDynamicStepDesc d6) d6.StepDescriptionChanged -= OnDynamicDescChanged;
    }

    private void SubscribeDynamicDesc()
    {
        if (_currentStep == 0 && _step1 is IDynamicStepDesc d1) d1.StepDescriptionChanged += OnDynamicDescChanged;
        if (_currentStep == 5 && _step6 is IDynamicStepDesc d6) d6.StepDescriptionChanged += OnDynamicDescChanged;
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
            var linkIndex = desc.IndexOf("设置");
            if (linkIndex >= 0)
            {
                TxtStepDesc.Inlines.Clear();
                TxtStepDesc.Inlines.Add(new Run(desc[..linkIndex]));
                var hyperlink = new Hyperlink(new Run("设置"))
                {
                    NavigateUri = new Uri("tailormail://settings"),
                    Foreground = (Brush)FindResource("AccentBrush"),
                    TextDecorations = null
                };
                hyperlink.RequestNavigate += OnSettingsLinkClick;
                TxtStepDesc.Inlines.Add(hyperlink);
                TxtStepDesc.Inlines.Add(new Run(desc[(linkIndex + 2)..]));
            }
            else
            {
                TxtStepDesc.Text = desc;
            }
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
        BtnNewMail.Visibility = _currentStep == 0 ? Visibility.Visible : Visibility.Collapsed;

        if (_currentStep == 5)
        {
            BtnNext.Content = "开始发送";
            BtnNext.Appearance = Wpf.Ui.Controls.ControlAppearance.Primary;
        }
        else
        {
            BtnNext.Content = "下一步";
            BtnNext.Appearance = _currentStep == 0
                ? Wpf.Ui.Controls.ControlAppearance.Secondary
                : Wpf.Ui.Controls.ControlAppearance.Primary;
        }
    }

    private void UpdateStepIndicator()
    {
        for (int i = 0; i < _steps.Count; i++)
        {
            _steps[i].State = i < _currentStep ? StepState.Completed
                : i == _currentStep ? StepState.Current
                : StepState.Upcoming;
        }

        Dispatcher.BeginInvoke(DispatcherPriority.Loaded, UpdateAccentBar);
    }

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

            var halfSpread = 0.44;
            AccentBar.Background = new LinearGradientBrush
            {
                StartPoint = new Point(0, 0.5),
                EndPoint = new Point(1, 0.5),
                GradientStops = new GradientStopCollection
                {
                    new GradientStop(Colors.Transparent, Math.Max(0, centerX - halfSpread)),
                    new GradientStop((Color)FindResource("AccentColor"), centerX),
                    new GradientStop(Colors.Transparent, Math.Min(1, centerX + halfSpread))
                }
            };
        }
        catch
        {
        }
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
        if (_currentStep > 0) NavigateToStep(_currentStep - 1);
    }

    private void BtnNext_Click(object sender, RoutedEventArgs e)
    {
        SaveCurrentStep();
        if (_currentStep < 5) NavigateToStep(_currentStep + 1);
        else
        {
            if (_step6 != null) _step6.StartSend();
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
            if (MessageBox.Show("新建邮件将清除已撰写内容，是否继续？", "确认",
                MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes) return;
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
            MessageBox.Show($"打开设置窗口失败：\n\n{detail}", "错误",
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void BtnAbout_Click(object sender, RoutedEventArgs e)
    {
        var win = new AboutWindow { Owner = this };
        win.ShowDialog();
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
        MainContent.Opacity = 0;
        MainContentTransform.Y = 12;

        var opacityAnimation = new DoubleAnimation(0, 1, TimeSpan.FromMilliseconds(250))
        {
            EasingFunction = new CubicEase { EasingMode = EasingMode.EaseOut }
        };
        var slideAnimation = new DoubleAnimation(12, 0, TimeSpan.FromMilliseconds(250))
        {
            EasingFunction = new CubicEase { EasingMode = EasingMode.EaseOut }
        };

        MainContent.BeginAnimation(UIElement.OpacityProperty, opacityAnimation);
        MainContentTransform.BeginAnimation(TranslateTransform.YProperty, slideAnimation);
    }

    private void OnMainWindowClosed(object? sender, EventArgs e)
    {
        Application.Current.Shutdown();
    }
}

public enum StepState { Upcoming, Current, Completed }

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
