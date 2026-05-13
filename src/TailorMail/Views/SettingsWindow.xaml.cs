using System.Threading;
using System.Windows;
using System.Windows.Controls;
using TailorMail.Models;
using TailorMail.ViewModels;
using TailorMail.Views;

namespace TailorMail.Views;

public partial class SettingsWindow
{
    private readonly SettingsViewModel _vm;
    private bool _isLoading;
    private CancellationTokenSource? _portHintCts;

    public SettingsWindow()
    {
        try
        {
            InitializeComponent();
        }
        catch (Exception ex)
        {
            Services.AppLogger.Error("SettingsWindow.InitializeComponent 失败", ex);
            throw;
        }

        try
        {
            _vm = new SettingsViewModel(App.DataService);
        }
        catch (Exception ex)
        {
            Services.AppLogger.Error("SettingsViewModel 创建失败", ex);
            throw;
        }

        try
        {
            LoadSettings();
            UpdateOutlookStatus();
        }
        catch (Exception ex)
        {
            Services.AppLogger.Error("SettingsWindow.LoadSettings 失败", ex);
            throw;
        }
    }

    private void LoadSettings()
    {
        _isLoading = true;
        try
        {
            CmbChannel.SelectedIndex = _vm.SendMethod == SendMethod.Smtp ? 1 : 0;
            UpdateSmtpExpander();
            TxtServer.Text = _vm.SmtpHost;
            TxtPort.Value = _vm.SmtpPort;
            TxtUsername.Text = _vm.SmtpUserName;
            ChkSsl.IsChecked = _vm.SmtpUseSsl;
            TxtSenderName.Text = _vm.SmtpDisplayName;
            TxtSenderEmail.Text = !string.IsNullOrEmpty(_vm.SmtpSenderEmail)
                ? _vm.SmtpSenderEmail
                : _vm.SmtpUserName;
            TxtPassword.Password = _vm.StoredPassword;
            TxtSendInterval.Value = _vm.SendIntervalMs / 1000.0;
            TxtSmtpConcurrency.Value = _vm.SmtpConcurrency;
            TxtSignature.Text = _vm.Signature;
        }
        finally
        {
            _isLoading = false;
        }
    }

    private void UpdateOutlookStatus()
    {
        if (_vm == null) return;
        var isAvailable = _vm.IsOutlookAvailable;
        OutlookAvailableBar.IsOpen = isAvailable && CmbChannel.SelectedIndex == 0;
        OutlookUnavailableBar.IsOpen = !isAvailable && CmbChannel.SelectedIndex == 0;
    }

    private void OnChannelChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_vm == null) return; // fires during InitializeComponent before _vm is set
        UpdateOutlookStatus();
        UpdateSmtpExpander();
    }

    private void UpdateSmtpExpander()
    {
        var isSmtp = CmbChannel.SelectedIndex == 1;
        if (isSmtp)
        {
            SmtpExpander.IsEnabled = true;
            SmtpExpander.Header = "SMTP 服务器";
            SmtpExpander.IsExpanded = true;
        }
        else
        {
            SmtpExpander.IsEnabled = false;
            SmtpExpander.Header = "SMTP 服务器（切换到 SMTP 模式后可用）";
            SmtpExpander.IsExpanded = false;
        }
        SmtpConcurrencySection.Visibility = isSmtp ? Visibility.Visible : Visibility.Collapsed;
        SmtpConcurrencyWarning.Visibility = isSmtp ? Visibility.Visible : Visibility.Collapsed;
    }

    private void OnPresetChanged(object sender, SelectionChangedEventArgs e)
    {
        if (CmbPreset?.SelectedIndex <= 0) return;

        var (server, port, ssl) = CmbPreset.SelectedIndex switch
        {
            1 => ("smtp.qq.com", 465, true),
            2 => ("smtp.163.com", 465, true),
            3 => ("smtp.gmail.com", 587, false),
            4 => ("smtp-mail.outlook.com", 587, false),
            _ => (null as string, 0, false)
        };

        if (server != null)
        {
            TxtServer.Text = server;
            TxtPort.Value = port;
            ChkSsl.IsChecked = ssl;
        }
    }

    private void OnSslToggleChanged(object sender, RoutedEventArgs e)
    {
        if (_isLoading) return;
        var oldPort = (int)(TxtPort.Value ?? 587);
        if (ChkSsl.IsChecked == true)
        {
            if (oldPort == 25 || oldPort == 587)
            {
                TxtPort.Value = 465;
                ShowPortHint("端口已切换为 465 (SSL)");
            }
        }
        else
        {
            if (oldPort == 465)
            {
                TxtPort.Value = 587;
                ShowPortHint("端口已切换为 587 (TLS)");
            }
        }
    }

    private async void ShowPortHint(string message)
    {
        _portHintCts?.Cancel();
        var cts = new CancellationTokenSource();
        _portHintCts = cts;
        PortHintText.Text = message;
        try
        {
            await Task.Delay(5000, cts.Token);
        }
        catch (TaskCanceledException)
        {
            return;
        }
        if (!cts.IsCancellationRequested)
        {
            PortHintText.Text = "";
        }
    }

    private async void OnTestConnection(object sender, RoutedEventArgs e)
    {
        BtnTestConnection.IsEnabled = false;
        TestProgressRing.Visibility = Visibility.Visible;
        TestResultBar.IsOpen = false;

        try
        {
            if (CmbChannel.SelectedIndex == 0)
            {
                var isAvailable = _vm.IsOutlookAvailable;
                TestResultBar.Title = isAvailable ? "成功" : "错误";
                TestResultBar.Message = isAvailable ? "Outlook 连接正常" : "未检测到 Outlook";
                TestResultBar.Severity = isAvailable
                    ? Wpf.Ui.Controls.InfoBarSeverity.Success
                    : Wpf.Ui.Controls.InfoBarSeverity.Error;
                TestResultBar.IsOpen = true;
                return;
            }

            var host = TxtServer.Text?.Trim() ?? "";
            var port = (int)(TxtPort.Value ?? 587);
            var useSsl = ChkSsl.IsChecked ?? true;
            var username = TxtUsername.Text?.Trim() ?? "";

            if (string.IsNullOrEmpty(host))
            {
                TestResultBar.Title = "错误";
                TestResultBar.Message = "请填写 SMTP 服务器地址";
                TestResultBar.Severity = Wpf.Ui.Controls.InfoBarSeverity.Error;
                TestResultBar.IsOpen = true;
                return;
            }

            if (string.IsNullOrEmpty(username))
            {
                TestResultBar.Title = "错误";
                TestResultBar.Message = "请填写 SMTP 用户名";
                TestResultBar.Severity = Wpf.Ui.Controls.InfoBarSeverity.Error;
                TestResultBar.IsOpen = true;
                return;
            }

            TestResultBar.Title = "测试中";
            TestResultBar.Message = "正在测试连接...";
            TestResultBar.Severity = Wpf.Ui.Controls.InfoBarSeverity.Informational;
            TestResultBar.IsOpen = true;

            try
            {
                using var client = new MailKit.Net.Smtp.SmtpClient();
                await client.ConnectAsync(host, port, useSsl
                    ? MailKit.Security.SecureSocketOptions.SslOnConnect
                    : MailKit.Security.SecureSocketOptions.StartTlsWhenAvailable);

                if (!string.IsNullOrEmpty(username))
                {
                    var storedPassword = TxtPassword.Password?.Trim();
                    var password = !string.IsNullOrEmpty(storedPassword) ? storedPassword : PromptForPassword();
                    if (password != null)
                    {
                        try
                        {
                            await client.AuthenticateAsync(username, password);
                            await client.DisconnectAsync(true);
                            TestResultBar.Title = "成功";
                            TestResultBar.Message = "连接成功，认证通过";
                            TestResultBar.Severity = Wpf.Ui.Controls.InfoBarSeverity.Success;
                            TestResultBar.IsOpen = true;
                            return;
                        }
                        catch (Exception authEx)
                        {
                            await client.DisconnectAsync(true);
                            TestResultBar.Title = "警告";
                            TestResultBar.Message = $"连接成功但认证失败：{authEx.Message}";
                            TestResultBar.Severity = Wpf.Ui.Controls.InfoBarSeverity.Warning;
                            TestResultBar.IsOpen = true;
                            return;
                        }
                    }
                }

                await client.DisconnectAsync(true);
                TestResultBar.Title = "成功";
                TestResultBar.Message = "连接成功（未测试认证）";
                TestResultBar.Severity = Wpf.Ui.Controls.InfoBarSeverity.Success;
                TestResultBar.IsOpen = true;
            }
            catch (Exception ex)
            {
                TestResultBar.Title = "错误";
                TestResultBar.Message = $"连接失败：{ex.Message}";
                TestResultBar.Severity = Wpf.Ui.Controls.InfoBarSeverity.Error;
                TestResultBar.IsOpen = true;
            }
        }
        finally
        {
            BtnTestConnection.IsEnabled = true;
            TestProgressRing.Visibility = Visibility.Collapsed;
        }
    }

    private string? PromptForPassword()
    {
        var dialog = new SmtpPasswordDialog { Owner = this };
        return dialog.ShowDialog() == true ? dialog.Password : null;
    }

    private void BtnSave_Click(object sender, RoutedEventArgs e)
    {
        if (CmbChannel.SelectedIndex == 1)
        {
            if (string.IsNullOrWhiteSpace(TxtServer.Text))
            {
                App.ShowWarning("请填写 SMTP 服务器地址");
                return;
            }
            if (string.IsNullOrWhiteSpace(TxtUsername.Text))
            {
                App.ShowWarning("请填写 SMTP 用户名");
                return;
            }
        }

        _vm.SendMethod = CmbChannel.SelectedIndex == 1 ? SendMethod.Smtp : SendMethod.Outlook;
        _vm.SmtpHost = TxtServer.Text;
        _vm.SmtpPort = (int)(TxtPort.Value ?? 587);
        _vm.SmtpUserName = TxtUsername.Text;
        _vm.SmtpUseSsl = ChkSsl.IsChecked ?? true;
        _vm.SmtpDisplayName = TxtSenderName.Text;
        _vm.SmtpSenderEmail = TxtSenderEmail.Text;
        _vm.SmtpPassword = TxtPassword.Password?.Trim() ?? "";
        _vm.SendIntervalMs = (int)((TxtSendInterval.Value ?? 1) * 1000);
        _vm.SmtpConcurrency = (int)(TxtSmtpConcurrency.Value ?? 1);
        _vm.Signature = TxtSignature.Text;
        _vm.SaveCommand.Execute(null);
        Close();
    }

    private void BtnCancel_Click(object sender, RoutedEventArgs e)
    {
        Close();
    }
}