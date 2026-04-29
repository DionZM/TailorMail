using System.Windows.Controls;
using TailorMail.ViewModels;

namespace TailorMail.Views;

/// <summary>
/// 设置页面，提供发送方式选择和 SMTP 服务器配置。
/// </summary>
public partial class SettingsPage : Page, IRefreshable
{
    private readonly SettingsViewModel _viewModel;

    public SettingsPage()
    {
        InitializeComponent();
        _viewModel = new SettingsViewModel(App.DataService);
        DataContext = _viewModel;
    }

    public void RefreshData()
    {
        _viewModel.ReloadSettings();
    }
}
