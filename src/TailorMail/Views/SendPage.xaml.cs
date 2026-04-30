using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using TailorMail.Models;
using TailorMail.ViewModels;

namespace TailorMail.Views;

public partial class SendPage : UserControl, IRefreshable, IDynamicStepDesc
{
    private readonly SendViewModel _viewModel;

    public event Action? StepDescriptionChanged;

    public SendPage()
    {
        InitializeComponent();
        _viewModel = new SendViewModel(App.DataService);
        DataContext = _viewModel;
    }

    public void RefreshData()
    {
        _viewModel.ReloadSettings();
        LoadResults();
        StepDescriptionChanged?.Invoke();
    }

    public string GetStepDescription()
    {
        var channelName = _viewModel.SendMethod == SendMethod.Smtp ? "SMTP" : "Outlook";
        return $"当前发送通道为 {channelName}，可在设置中修改";
    }

    private void LoadResults()
    {
        var recipients = _viewModel.GetSelectedRecipients();
        var results = recipients.Select(r => new
        {
            r.Id,
            RecipientName = r.Name,
            StatusText = "等待",
            Status = SendStatus.Pending,
            SendTime = (DateTime?)null,
            ErrorMessage = ""
        }).ToList();
        DataGridResults.ItemsSource = results;
        TxtProgress.Text = $"0/{results.Count}";
    }

    public void StartSend()
    {
        OnStartSend(this, new RoutedEventArgs());
    }

    private async void OnStartSend(object sender, RoutedEventArgs e)
    {
        var selectedRecipients = _viewModel.GetSelectedRecipients();
        if (selectedRecipients.Count == 0)
        {
            MessageBox.Show("请先在「对象选择」步骤中选择要发送的对象", "提示",
                MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        string? smtpPassword = null;
        if (_viewModel.SendMethod == SendMethod.Smtp)
        {
            var dialog = new SmtpPasswordDialog { Owner = Window.GetWindow(this) };
            if (dialog.ShowDialog() != true) return;
            smtpPassword = dialog.Password;
        }

        LoadingOverlay.Visibility = Visibility.Visible;
        LoadingText.Text = "正在发送邮件...";

        try
        {
            await _viewModel.ExecuteSend(selectedRecipients, smtpPassword);

            DataGridResults.ItemsSource = _viewModel.SendResults;
            var done = _viewModel.SuccessCount + _viewModel.FailedCount;
            ProgressBar.Value = _viewModel.ProgressValue;
            TxtProgress.Text = $"{done}/{_viewModel.TotalCount}";
        }
        finally
        {
            LoadingOverlay.Visibility = Visibility.Collapsed;
        }
    }

    private async void OnRetry(object sender, RoutedEventArgs e)
    {
        if (sender is FrameworkElement fe && fe.Tag is string id)
        {
            LoadingOverlay.Visibility = Visibility.Visible;
            LoadingText.Text = "正在重试...";

            try
            {
                await _viewModel.RetryFailed();
                DataGridResults.ItemsSource = _viewModel.SendResults;
            }
            finally
            {
                LoadingOverlay.Visibility = Visibility.Collapsed;
            }
        }
    }

    private void OnFilterChanged(object sender, RoutedEventArgs e)
    {
        var view = CollectionViewSource.GetDefaultView(DataGridResults.ItemsSource);
        if (view == null) return;

        if (FilterSuccess.IsChecked == true)
        {
            view.Filter = item =>
            {
                var status = item?.GetType().GetProperty("Status")?.GetValue(item);
                return status is SendStatus s && s == SendStatus.Success;
            };
        }
        else if (FilterFailed.IsChecked == true)
        {
            view.Filter = item =>
            {
                var status = item?.GetType().GetProperty("Status")?.GetValue(item);
                return status is SendStatus s && s == SendStatus.Failed;
            };
        }
        else if (FilterPending.IsChecked == true)
        {
            view.Filter = item =>
            {
                var status = item?.GetType().GetProperty("Status")?.GetValue(item);
                return status is SendStatus s && s == SendStatus.Pending;
            };
        }
        else
        {
            view.Filter = null;
        }
    }
}
