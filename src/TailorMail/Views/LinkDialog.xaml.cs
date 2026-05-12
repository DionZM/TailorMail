using System.Windows;

namespace TailorMail.Views;

public partial class LinkDialog
{
    public string LinkText => TxtLinkText.Text?.Trim() ?? "";
    public string LinkUrl => TxtLinkUrl.Text?.Trim() ?? "";

    public LinkDialog(string? defaultText = null)
    {
        InitializeComponent();
        Owner = Application.Current.MainWindow;
        if (!string.IsNullOrEmpty(defaultText))
            TxtLinkText.Text = defaultText;

        Loaded += (_, _) =>
        {
            if (!string.IsNullOrEmpty(defaultText))
                TxtLinkUrl.Focus();
            else
                TxtLinkText.Focus();
        };
    }

    private void BtnOk_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrWhiteSpace(LinkText))
        {
            System.Windows.MessageBox.Show("请输入链接文字", "提示",
                System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
            return;
        }
        if (string.IsNullOrWhiteSpace(LinkUrl))
        {
            System.Windows.MessageBox.Show("请输入链接地址", "提示",
                System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
            return;
        }
        DialogResult = true;
        Close();
    }

    private void BtnCancel_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
        Close();
    }
}