using System.Reflection;
using System.Windows;
using System.Windows.Navigation;

namespace TailorMail.Views;

/// <summary>
/// 关于窗口，展示软件名称、版本、用途、项目地址和授权信息。
/// </summary>
public partial class AboutWindow
{
    public AboutWindow()
    {
        InitializeComponent();
        var version = Assembly.GetExecutingAssembly().GetName().Version;
        TxtVersion.Text = $"v{version?.Major ?? 0}.{version?.Minor ?? 9}.{version?.Build ?? 0}";
    }

    private void BtnClose_Click(object sender, RoutedEventArgs e)
    {
        Close();
    }

    private void OnLinkClick(object sender, RequestNavigateEventArgs e)
    {
        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(e.Uri.AbsoluteUri)
        {
            UseShellExecute = true
        });
        e.Handled = true;
    }
}
