﻿﻿using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Shapes;
using TailorMail.Views;

namespace TailorMail;

/// <summary>
/// 应用程序主窗口，采用 6 步骤向导式导航设计：
/// 步骤1 收件选择 → 步骤2 变量配置 → 步骤3 模板撰写 → 步骤4 附件匹配 → 步骤5 效果预览 → 步骤6 批量发送。
/// 各步骤页面采用延迟初始化策略，仅在首次访问时创建实例。
/// </summary>
public partial class MainWindow
{
    /// <summary>当前步骤索引（0-5）。</summary>
    private int _currentStep;

    /// <summary>步骤导航按钮数组。</summary>
    private readonly Button[] _stepBtns = new Button[6];

    /// <summary>步骤指示器圆形图标数组。</summary>
    private readonly Border[] _stepCircles = new Border[6];

    /// <summary>步骤标签文本数组。</summary>
    private readonly TextBlock[] _stepLabels = new TextBlock[6];

    /// <summary>步骤之间的连接线数组。</summary>
    private readonly Rectangle[] _stepLines = new Rectangle[5];

    /// <summary>步骤1：收件选择页面（延迟初始化）。</summary>
    private RecipientsPage? _step1;

    /// <summary>步骤2：变量管理页面（延迟初始化）。</summary>
    private VariablesPage? _step2;

    /// <summary>步骤3：模板撰写页面（延迟初始化）。</summary>
    private MailComposePage? _step3;

    /// <summary>步骤4：附件管理页面（延迟初始化）。</summary>
    private AttachmentPage? _step4;

    /// <summary>步骤5：邮件预览页面（延迟初始化）。</summary>
    private PreviewPage? _step5;

    /// <summary>步骤6：邮件发送页面（延迟初始化）。</summary>
    private SendPage? _step6;

    /// <summary>6 个步骤的显示名称。</summary>
    private static readonly string[] StepNames = ["收件选择", "变量配置", "模板撰写", "附件匹配", "效果预览", "批量发送"];

    public MainWindow()
    {
        InitializeComponent();
        Loaded += OnLoaded;
    }

    /// <summary>
    /// 窗口加载时初始化步骤导航控件映射并导航到第一步。
    /// </summary>
    private void OnLoaded(object sender, RoutedEventArgs e)
    {
        _stepBtns[0] = StepBtn0; _stepBtns[1] = StepBtn1; _stepBtns[2] = StepBtn2;
        _stepBtns[3] = StepBtn3; _stepBtns[4] = StepBtn4; _stepBtns[5] = StepBtn5;

        _stepCircles[0] = StepCircle0; _stepCircles[1] = StepCircle1; _stepCircles[2] = StepCircle2;
        _stepCircles[3] = StepCircle3; _stepCircles[4] = StepCircle4; _stepCircles[5] = StepCircle5;

        _stepLabels[0] = StepLabel0; _stepLabels[1] = StepLabel1; _stepLabels[2] = StepLabel2;
        _stepLabels[3] = StepLabel3; _stepLabels[4] = StepLabel4; _stepLabels[5] = StepLabel5;

        _stepLines[0] = StepLine0; _stepLines[1] = StepLine1; _stepLines[2] = StepLine2;
        _stepLines[3] = StepLine3; _stepLines[4] = StepLine4;

        NavigateToStep(0);
    }

    /// <summary>
    /// 步骤按钮点击处理：保存当前步骤数据并导航到目标步骤。
    /// </summary>
    private void OnStepClick(object sender, RoutedEventArgs e)
    {
        if (sender is Button btn && btn.Tag is string tag && int.TryParse(tag, out var step))
        {
            SaveCurrentStep();
            NavigateToStep(step);
        }
    }

    /// <summary>
    /// 导航到指定步骤。采用延迟初始化策略，页面仅在首次访问时创建。
    /// 导航时会刷新页面数据并更新步骤指示器。
    /// </summary>
    /// <param name="step">目标步骤索引（0-5）。</param>
    private void NavigateToStep(int step)
    {
        _currentStep = step;
        UpdateStepIndicator();
        UpdateButtons();
        TxtStepHint.Text = $"步骤 {step + 1}/6 · {StepNames[step]}";

        switch (step)
        {
            case 0: _step1 ??= new RecipientsPage(); MainContent.Content = _step1; break;
            case 1: _step2 ??= new VariablesPage(); _step2.RefreshData(); MainContent.Content = _step2; break;
            case 2: _step3 ??= new MailComposePage(); _step3.RefreshData(); MainContent.Content = _step3; break;
            case 3: _step4 ??= new AttachmentPage(); _step4.RefreshData(); MainContent.Content = _step4; break;
            case 4: _step5 ??= new PreviewPage(); _step5.RefreshData(); MainContent.Content = _step5; break;
            case 5: _step6 ??= new SendPage(); _step6.RefreshData(); MainContent.Content = _step6; break;
        }
    }

    /// <summary>
    /// 根据当前步骤更新导航按钮的可见性和文字。
    /// </summary>
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

    /// <summary>
    /// 更新步骤指示器的视觉状态。三种状态：
    /// - 已完成（i &lt; currentStep）：浅色背景 + 强调色边框 + ✓ 标记
    /// - 当前步骤（i == currentStep）：强调色背景 + 白色文字
    /// - 未到达（i &gt; currentStep）：灰色背景 + 灰色文字
    /// </summary>
    private void UpdateStepIndicator()
    {
        var accentBrush = (Brush)FindResource("AccentBrush");
        var accentLightBrush = (Brush)FindResource("AccentLightBrush");
        var surfaceBrush = (Brush)FindResource("SurfaceBrush");
        var borderSubtleBrush = (Brush)FindResource("BorderSubtleBrush");
        var primaryBrush = (Brush)FindResource("TextPrimaryBrush");
        var tertiaryBrush = (Brush)FindResource("TextTertiaryBrush");

        for (int i = 0; i < 6; i++)
        {
            var circle = _stepCircles[i];
            var label = _stepLabels[i];
            var circleText = circle.Child as TextBlock;

            if (i < _currentStep)
            {
                // 已完成步骤
                circle.Background = accentLightBrush;
                circle.BorderBrush = accentBrush;
                circle.BorderThickness = new Thickness(1.5);
                if (circleText != null) { circleText.Foreground = accentBrush; circleText.Text = "✓"; }
                if (label != null) { label.Foreground = accentBrush; label.FontWeight = FontWeights.Medium; }
            }
            else if (i == _currentStep)
            {
                // 当前步骤
                circle.Background = accentBrush;
                circle.BorderBrush = null;
                circle.BorderThickness = new Thickness(0);
                if (circleText != null) { circleText.Foreground = Brushes.White; }
                if (label != null) { label.Foreground = primaryBrush; label.FontWeight = FontWeights.SemiBold; }
            }
            else
            {
                // 未到达步骤
                circle.Background = surfaceBrush;
                circle.BorderBrush = borderSubtleBrush;
                circle.BorderThickness = new Thickness(1.5);
                if (circleText != null) { circleText.Foreground = tertiaryBrush; }
                if (label != null) { label.Foreground = tertiaryBrush; label.FontWeight = FontWeights.Medium; }
            }
        }

        // 更新连接线颜色
        for (int i = 0; i < 5; i++)
        {
            _stepLines[i].Fill = i < _currentStep ? accentBrush : borderSubtleBrush;
        }
    }

    /// <summary>
    /// "上一步"按钮点击处理。
    /// </summary>
    private void BtnPrev_Click(object sender, RoutedEventArgs e)
    {
        SaveCurrentStep();
        if (_currentStep > 0) NavigateToStep(_currentStep - 1);
    }

    /// <summary>
    /// "下一步"按钮点击处理。若当前在最后一步则触发发送。
    /// </summary>
    private void BtnNext_Click(object sender, RoutedEventArgs e)
    {
        SaveCurrentStep();
        if (_currentStep < 5) NavigateToStep(_currentStep + 1);
        else
        {
            if (_step6 != null) _step6.StartSend();
        }
    }

    /// <summary>
    /// 检查是否存在已编辑的邮件内容。
    /// </summary>
    /// <returns>若存在已编辑的主题或正文则返回 true。</returns>
    private bool HasEditedContent()
    {
        var settings = App.DataService.LoadSettings();
        return !string.IsNullOrEmpty(settings.LastSubject) ||
               !string.IsNullOrEmpty(settings.LastBody) ||
               !string.IsNullOrEmpty(settings.LastBodyXaml);
    }

    /// <summary>
    /// "新建邮件"按钮点击处理。清除已撰写内容和附件配置，重置后续步骤页面实例。
    /// </summary>
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

        // 重置后续步骤页面实例，强制重新初始化
        _step3 = null;
        _step4 = null;
        _step5 = null;
        _step6 = null;

        NavigateToStep(1);
    }

    /// <summary>
    /// "设置"按钮点击处理，打开设置窗口。
    /// </summary>
    private void BtnSettings_Click(object sender, RoutedEventArgs e)
    {
        var win = new SettingsWindow { Owner = this };
        win.ShowDialog();
    }

    /// <summary>
    /// "关于"按钮点击处理，打开关于窗口。
    /// </summary>
    private void BtnAbout_Click(object sender, RoutedEventArgs e)
    {
        var win = new AboutWindow { Owner = this };
        win.ShowDialog();
    }

    /// <summary>
    /// 保存当前步骤页面的数据。仅对需要保存的步骤（0-2）执行保存操作。
    /// </summary>
    private void SaveCurrentStep()
    {
        switch (_currentStep)
        {
            case 0: _step1?.SaveAll(); break;
            case 1: _step2?.SaveAll(); break;
            case 2: _step3?.SaveCurrent(); break;
        }
    }
}

/// <summary>
/// 可刷新数据接口，由各步骤页面实现，用于在导航到该步骤时刷新数据。
/// </summary>
public interface IRefreshable
{
    /// <summary>
    /// 刷新页面数据。
    /// </summary>
    void RefreshData();
}
