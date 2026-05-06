using System.Windows;
using System.Windows.Controls;
using TailorMail.Models;
using TailorMail.ViewModels;
using TailorMail.Views;

namespace TailorMail.Views;

public partial class SettingsWindow
{
    private readonly SettingsViewModel _vm;

    public SettingsWindow()
    {
        InitializeComponent();
        _vm = new SettingsViewModel(App.DataService);
        LoadSettings();
        UpdateOutlookStatus();
    }

    private void LoadSettings()
    {
        CmbChannel.SelectedIndex = _vm.SendMethod == SendMethod.Smtp ? 1 : 0;
        TxtServer.Text = _vm.SmtpHost;
        TxtPort.Value = _vm.SmtpPort;
        TxtUsername.Text = _vm.SmtpUserName;
        ChkSsl.IsChecked = _vm.SmtpUseSsl;
        TxtSenderName.Text = _vm.SmtpDisplayName;
        TxtSenderEmail.Text = !string.IsNullOrEmpty(_vm.SmtpSenderEmail)
            ? _vm.SmtpSenderEmail
            : _vm.SmtpUserName;
        TxtSignature.Text = _vm.Signature;
    }

    private void UpdateOutlookStatus()
    {
        if (_vm == null) return;
        var isAvailable = _vm.IsOutlookAvailable;
        OutlookAvailableBar.IsOpen = isAvailable;
        OutlookUnavailableBar.IsOpen = !isAvailable;
    }

    private void OnChannelChanged(object sender, SelectionChangedEventArgs e)
    {
        UpdateOutlookStatus();
    }

    private void OnSslToggleChanged(object sender, RoutedEventArgs e)
    {
        if (ChkSsl.IsChecked == true && TxtPort.Value == 25)
            TxtPort.Value = 465;
        else if (ChkSsl.IsChecked == false && TxtPort.Value == 465)
            TxtPort.Value = 587;
    }

    private async void OnTestConnection(object sender, RoutedEventArgs e)
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
                var passwordDlg = new SmtpPasswordDialog { Owner = this };
                if (passwordDlg.ShowDialog() == true)
                {
                    try
                    {
                        await client.AuthenticateAsync(username, passwordDlg.Password);
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
            TestResultBar.Message = $"连接失败: {ex.Message}";
            TestResultBar.Severity = Wpf.Ui.Controls.InfoBarSeverity.Error;
            TestResultBar.IsOpen = true;
        }
    }

    private void BtnSave_Click(object sender, RoutedEventArgs e)
    {
        // Validate SMTP fields when SMTP is selected
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
        _vm.Signature = TxtSignature.Text;
        _vm.SaveCommand.Execute(null);
        Close();
    }

    private void BtnCancel_Click(object sender, RoutedEventArgs e)
    {
        Close();
    }
}
