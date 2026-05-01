using System.Windows;
using System.Windows.Controls;
using TailorMail.Models;
using TailorMail.ViewModels;

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

    private void BtnSave_Click(object sender, RoutedEventArgs e)
    {
        _vm.SendMethod = CmbChannel.SelectedIndex == 1 ? SendMethod.Smtp : SendMethod.Outlook;
        _vm.SmtpHost = TxtServer.Text;
        _vm.SmtpPort = (int)(TxtPort.Value ?? 587);
        _vm.SmtpUserName = TxtUsername.Text;
        _vm.SmtpUseSsl = ChkSsl.IsChecked ?? true;
        _vm.SmtpDisplayName = TxtSenderName.Text;
        _vm.SmtpSenderEmail = TxtSenderEmail.Text;
        _vm.SaveCommand.Execute(null);
        Close();
    }

    private void BtnCancel_Click(object sender, RoutedEventArgs e)
    {
        Close();
    }
}
