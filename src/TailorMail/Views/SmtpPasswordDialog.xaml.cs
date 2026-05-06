using System.Windows;
using System.Windows.Controls;
using TailorMail.ViewModels;
using Wpf.Ui.Controls;

namespace TailorMail.Views;

public partial class SmtpPasswordDialog : FluentWindow
{
    public string Password => PasswordBox.Visibility == Visibility.Visible
        ? PasswordBox.Password
        : PasswordTextBox.Text;

    public SmtpPasswordDialog()
    {
        InitializeComponent();
    }

    private void OnConfirm(object sender, RoutedEventArgs e)
    {
        var pw = Password;
        if (string.IsNullOrWhiteSpace(pw))
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

    private void OnToggleShowPassword(object sender, RoutedEventArgs e)
    {
        if (ChkShowPassword.IsChecked == true)
        {
            PasswordTextBox.Text = PasswordBox.Password;
            PasswordBox.Visibility = Visibility.Collapsed;
            PasswordTextBox.Visibility = Visibility.Visible;
            PasswordTextBox.Focus();
        }
        else
        {
            PasswordBox.Password = PasswordTextBox.Text;
            PasswordTextBox.Visibility = Visibility.Collapsed;
            PasswordBox.Visibility = Visibility.Visible;
            PasswordBox.Focus();
        }
    }

    private void OnRevealedTextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
    {
        PasswordBox.Password = PasswordTextBox.Text;
    }
}
