using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using TailorMail.Models;
using TailorMail.ViewModels;

namespace TailorMail.Views;

public partial class SendPage : UserControl, IRefreshable, IDynamicStepDesc
{
    private readonly SendViewModel _viewModel;
    private List<Recipient> _allRecipients = [];
    private string? _smtpPassword;
    private DateTime? _sendStartTime;
    private int _doneAtResume;
    private bool _hasInitialized;
    private System.ComponentModel.PropertyChangedEventHandler? _propertyChangedHandler;

    public event Action? StepDescriptionChanged;
    public event Action? SendStateChanged;

    public bool IsSending => _viewModel.IsSending;

    public bool HasPendingRecipients =>
        _viewModel.SendResults.Any(r => r.Status == SendStatus.Pending);

    public bool HasAnyResults =>
        _viewModel.SendResults.Any(r => r.Status == SendStatus.Success || r.Status == SendStatus.Failed);

    public SendPage()
    {
        InitializeComponent();
        _viewModel = new SendViewModel(App.DataService);
        DataContext = _viewModel;
    }

    public void RefreshData()
    {
        _viewModel.ReloadSettings();
        _allRecipients = _viewModel.GetSelectedRecipients();

        // Only clear results on first load, preserve on subsequent visits
        if (!_hasInitialized)
        {
            _viewModel.SendResults.Clear();
            foreach (var r in _allRecipients)
                _viewModel.SendResults.Add(new SendResult { RecipientId = r.Id, RecipientName = r.Name });
            _viewModel.TotalCount = _allRecipients.Count;
            _viewModel.SuccessCount = 0;
            _viewModel.FailedCount = 0;
            _viewModel.ProgressValue = 0;
            _hasInitialized = true;
        }

        _viewModel.StatusText = _allRecipients.Count > 0
            ? $"共 {_allRecipients.Count} 封待发送"
            : "未选择发送对象";
        DataGridResults.ItemsSource = _viewModel.SendResults;
        UpdateProgressDisplay();
        StepDescriptionChanged?.Invoke();
        SendStateChanged?.Invoke();
    }

    public string GetStepDescription()
    {
        var channelName = _viewModel.SendMethod == SendMethod.Smtp ? "SMTP" : "Outlook";
        var total = _allRecipients.Count;

        if (_viewModel.SendResults.Count > 0)
        {
            var pending = _viewModel.SendResults.Count(r => r.Status == SendStatus.Pending);
            var failed = _viewModel.SendResults.Count(r => r.Status == SendStatus.Failed);
            if (_viewModel.IsSending)
                return $"发送中: 剩余 {pending} 封, 通道 [LINK:{channelName}]";
            if (pending > 0 && HasAnyResults)
                return $"已暂停: {_viewModel.SuccessCount} 成 / {failed} 败, 待发 {pending}, 通道 [LINK:{channelName}]";
            if (failed > 0)
                return $"已完成: {_viewModel.SuccessCount} 成 / {failed} 败, 通道 [LINK:{channelName}]";
            return $"已完成: {_viewModel.SuccessCount} 封, 通道 [LINK:{channelName}]";
        }

        return total > 0
            ? $"待发送 {total} 封，通道 [LINK:{channelName}]"
            : $"当前发送通道为 [LINK:{channelName}]，可在[LINK:设置]中修改";
    }

    public void StartSend()
    {
        if (_allRecipients.Count == 0)
        {
            _allRecipients = _viewModel.GetSelectedRecipients();
        }

        if (_allRecipients.Count == 0)
        {
            App.ShowWarning("请先在「对象选择」步骤中选择要发送的对象");
            return;
        }

        if (HasAnyResults && !HasPendingRecipients)
        {
            App.ShowNotification("所有邮件已发送完毕");
            return;
        }

        // Confirmation dialog before sending
        var pendingCount = HasAnyResults
            ? _viewModel.SendResults.Count(r => r.Status == SendStatus.Pending)
            : _allRecipients.Count;

        if (MessageBox.Show($"确认发送 {pendingCount} 封邮件？", "确认发送",
            MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes) return;

        OnStartOrContinue();
    }

    private async void OnStartOrContinue()
    {
        if (_viewModel.SendMethod == SendMethod.Smtp && _smtpPassword == null)
        {
            var dialog = new SmtpPasswordDialog { Owner = Window.GetWindow(this) };
            if (dialog.ShowDialog() != true) return;
            _smtpPassword = dialog.Password;
        }

        List<Recipient> toSend;
        if (HasAnyResults)
        {
            var pendingIds = _viewModel.SendResults
                .Where(r => r.Status == SendStatus.Pending)
                .Select(r => r.RecipientId)
                .ToHashSet();
            toSend = _allRecipients.Where(r => pendingIds.Contains(r.Id)).ToList();
        }
        else
        {
            toSend = _allRecipients;
        }

        if (toSend.Count == 0) return;

        DataGridResults.ItemsSource = _viewModel.SendResults;
        _sendStartTime = DateTime.Now;
        _doneAtResume = _viewModel.SuccessCount + _viewModel.FailedCount;

        try
        {
            _propertyChangedHandler = (_, _) => Dispatcher.BeginInvoke(UpdateProgressDisplay);
            _viewModel.PropertyChanged += _propertyChangedHandler;
            var sendTask = _viewModel.ExecuteSend(toSend, _smtpPassword, append: HasAnyResults);
            NotifyStateChanged();
            await sendTask;
        }
        finally
        {
            if (_propertyChangedHandler != null)
            {
                _viewModel.PropertyChanged -= _propertyChangedHandler;
                _propertyChangedHandler = null;
            }
            UpdateProgressDisplay();
            NotifyStateChanged();
            _sendStartTime = null;
        }
    }

    public void StopSend()
    {
        _viewModel.CancelSend();
    }

    private async void OnRetryOne(object sender, RoutedEventArgs e)
    {
        if (sender is FrameworkElement fe && fe.Tag is string id)
        {
            try
            {
                await _viewModel.RetryOne(id);
            }
            finally
            {
                UpdateProgressDisplay();
                NotifyStateChanged();
            }
        }
    }

    private async void OnRetryAll(object sender, RoutedEventArgs e)
    {
        try
        {
            await _viewModel.RetryAllFailed();
        }
        finally
        {
            UpdateProgressDisplay();
            NotifyStateChanged();
        }
    }

    private void OnExportResults(object sender, RoutedEventArgs e)
    {
        if (_viewModel.SendResults.Count == 0) return;

        var dialog = new Microsoft.Win32.SaveFileDialog
        {
            Filter = "Excel文件|*.xlsx",
            FileName = "发送结果.xlsx"
        };
        if (dialog.ShowDialog() != true) return;

        try
        {
            OfficeOpenXml.ExcelPackage.License.SetNonCommercialPersonal("TailorMail");
            using var package = new OfficeOpenXml.ExcelPackage();
            var ws = package.Workbook.Worksheets.Add("发送结果");
            ws.Cells[1, 1].Value = "收件人";
            ws.Cells[1, 2].Value = "状态";
            ws.Cells[1, 3].Value = "发送时间";
            ws.Cells[1, 4].Value = "错误信息";

            using (var range = ws.Cells[1, 1, 1, 4])
            {
                range.Style.Font.Bold = true;
                range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
            }

            for (int i = 0; i < _viewModel.SendResults.Count; i++)
            {
                var r = _viewModel.SendResults[i];
                ws.Cells[i + 2, 1].Value = r.RecipientName;
                ws.Cells[i + 2, 2].Value = r.StatusText;
                ws.Cells[i + 2, 3].Value = r.SendTime?.ToString("yyyy-MM-dd HH:mm:ss") ?? "";
                ws.Cells[i + 2, 4].Value = r.ErrorMessage ?? "";
            }
            ws.Cells[ws.Dimension.Address].AutoFitColumns();
            package.SaveAs(new System.IO.FileInfo(dialog.FileName));
            App.ShowSuccess("导出完成");
        }
        catch (Exception ex)
        {
            App.ShowError($"导出失败: {ex.Message}");
        }
    }

    private void UpdateProgressDisplay()
    {
        // Show current recipient during sending
        if (_viewModel.IsSending && !string.IsNullOrEmpty(_viewModel.StatusText))
        {
            TxtCurrentRecipient.Text = _viewModel.StatusText;
            TxtCurrentRecipient.Visibility = Visibility.Visible;
        }
        else
        {
            TxtCurrentRecipient.Visibility = Visibility.Collapsed;
        }

        var done = _viewModel.SuccessCount + _viewModel.FailedCount;
        var total = _viewModel.TotalCount;
        TxtProgress.Text = $"{done}/{total}";
        ProgressBar.Value = _viewModel.ProgressValue;
        TxtPercent.Text = total > 0 ? $"{(int)_viewModel.ProgressValue}%" : "0%";

        var hasFailed = _viewModel.SendResults.Any(r => r.Status == SendStatus.Failed);
        BtnRetryAll.Visibility = hasFailed ? Visibility.Visible : Visibility.Collapsed;
        BtnExport.Visibility = HasAnyResults ? Visibility.Visible : Visibility.Collapsed;

        // ETA calculation
        if (_viewModel.IsSending && _sendStartTime.HasValue && done > 0)
        {
            var doneSinceResume = done - _doneAtResume;
            if (doneSinceResume > 0)
            {
                var elapsed = DateTime.Now - _sendStartTime.Value;
                var avgTime = elapsed.TotalSeconds / doneSinceResume;
                var remaining = total - done;
                var eta = TimeSpan.FromSeconds(avgTime * remaining);
                TxtEta.Text = eta.TotalMinutes >= 1
                    ? $"预计剩余 {Math.Ceiling(eta.TotalMinutes)} 分钟"
                    : $"预计剩余 {Math.Ceiling(eta.TotalSeconds)} 秒";
            }
            else
            {
                TxtEta.Text = "正在计算...";
            }
        }
        else
        {
            TxtEta.Text = "";
        }
    }

    private void NotifyStateChanged()
    {
        StepDescriptionChanged?.Invoke();
        SendStateChanged?.Invoke();
    }

    private void OnFilterChanged(object sender, RoutedEventArgs e)
    {
        var view = CollectionViewSource.GetDefaultView(DataGridResults.ItemsSource);
        if (view == null) return;

        if (FilterSuccess.IsChecked == true)
        {
            view.Filter = item => item is SendResult s && s.Status == SendStatus.Success;
        }
        else if (FilterFailed.IsChecked == true)
        {
            view.Filter = item => item is SendResult s && s.Status == SendStatus.Failed;
        }
        else if (FilterPending.IsChecked == true)
        {
            view.Filter = item => item is SendResult s && s.Status == SendStatus.Pending;
        }
        else
        {
            view.Filter = null;
        }
    }
}
