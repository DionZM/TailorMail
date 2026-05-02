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
        _viewModel.SendResults.Clear();
        foreach (var r in _allRecipients)
            _viewModel.SendResults.Add(new SendResult { RecipientId = r.Id, RecipientName = r.Name });
        _viewModel.TotalCount = _allRecipients.Count;
        _viewModel.SuccessCount = 0;
        _viewModel.FailedCount = 0;
        _viewModel.ProgressValue = 0;
        _viewModel.StatusText = _allRecipients.Count > 0
            ? $"共 {_allRecipients.Count} 封待发送"
            : "未选择发送对象";
        DataGridResults.ItemsSource = _viewModel.SendResults;
        TxtProgress.Text = $"0/{_allRecipients.Count}";
        BtnRetryAll.Visibility = Visibility.Collapsed;
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
            MessageBox.Show("请先在「对象选择」步骤中选择要发送的对象", "提示",
                MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        if (HasAnyResults && !HasPendingRecipients)
        {
            MessageBox.Show("所有邮件已发送完毕", "提示",
                MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

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

        try
        {
            var sendTask = _viewModel.ExecuteSend(toSend, _smtpPassword, append: HasAnyResults);
            NotifyStateChanged();
            await sendTask;
        }
        finally
        {
            UpdateProgressDisplay();
            NotifyStateChanged();
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

    private void UpdateProgressDisplay()
    {
        var done = _viewModel.SuccessCount + _viewModel.FailedCount;
        TxtProgress.Text = $"{done}/{_viewModel.TotalCount}";
        ProgressBar.Value = _viewModel.ProgressValue;

        var hasFailed = _viewModel.SendResults.Any(r => r.Status == SendStatus.Failed);
        BtnRetryAll.Visibility = hasFailed ? Visibility.Visible : Visibility.Collapsed;
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
