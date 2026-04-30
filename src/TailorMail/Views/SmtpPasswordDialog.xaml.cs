using System.Windows;
using Wpf.Ui.Controls;

namespace TailorMail.Views;

/// <summary>
/// SMTP 密码输入对话框，用于在 SMTP 方式发送邮件时临时获取密码。
/// </summary>
public partial class SmtpPasswordDialog : FluentWindow
{
    public string Password => PasswordBox.Password;

    public SmtpPasswordDialog()
    {
        InitializeComponent();
    }

    private void OnConfirm(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrWhiteSpace(PasswordBox.Password))
        {
            System.Windows.MessageBox.Show("密码不能为空", "提示",
                System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
            return;
        }
        DialogResult = true;
        Close();
    }

    private void OnCancel(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
        Close();
    }
}
