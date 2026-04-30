using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using TailorMail.Models;
using TailorMail.Services;

namespace TailorMail.ViewModels;

/// <summary>
/// 设置视图模型，负责邮件发送方式、SMTP 服务器配置的管理和持久化，
/// 以及 Outlook 可用性检测。
/// </summary>
public partial class SettingsViewModel : ObservableObject
{
    private readonly IDataService _dataService;

    /// <summary>
    /// 获取或设置邮件发送方式（Outlook 或 SMTP）。
    /// </summary>
    [ObservableProperty]
    private SendMethod _sendMethod;

    /// <summary>
    /// 获取或设置 SMTP 服务器主机地址。
    /// </summary>
    [ObservableProperty]
    private string _smtpHost = string.Empty;

    /// <summary>
    /// 获取或设置 SMTP 服务器端口号。默认 587。
    /// </summary>
    [ObservableProperty]
    private int _smtpPort = 587;

    /// <summary>
    /// 获取或设置是否使用 SSL/TLS 加密连接。
    /// </summary>
    [ObservableProperty]
    private bool _smtpUseSsl = true;

    /// <summary>
    /// 获取或设置 SMTP 认证用户名。
    /// </summary>
    [ObservableProperty]
    private string _smtpUserName = string.Empty;

    /// <summary>
    /// 获取或设置发件人显示名称。
    /// </summary>
    [ObservableProperty]
    private string _smtpDisplayName = string.Empty;

    /// <summary>
    /// 获取或设置发件人邮箱地址。
    /// </summary>
    [ObservableProperty]
    private string _smtpSenderEmail = string.Empty;

    /// <summary>
    /// 获取或设置本地是否检测到 Outlook 安装。
    /// </summary>
    [ObservableProperty]
    private bool _isOutlookAvailable;

    public SettingsViewModel(IDataService dataService)
    {
        _dataService = dataService;
        LoadSettings();
        CheckOutlookAvailability();
    }

    /// <summary>
    /// 从数据服务加载设置到视图属性。
    /// </summary>
    private void LoadSettings()
    {
        var settings = _dataService.LoadSettings();
        SendMethod = settings.SendMethod;
        SmtpHost = settings.Smtp.Host;
        SmtpPort = settings.Smtp.Port;
        SmtpUseSsl = settings.Smtp.UseSsl;
        SmtpUserName = settings.Smtp.UserName;
        SmtpDisplayName = settings.Smtp.DisplayName;
        SmtpSenderEmail = settings.Smtp.SenderEmail;
    }

    /// <summary>
    /// 重新加载设置（用于外部修改后刷新）。
    /// </summary>
    public void ReloadSettings()
    {
        LoadSettings();
    }

    /// <summary>
    /// 检测本地是否安装了 Microsoft Outlook。
    /// 通过 COM ProgID "Outlook.Application" 是否可获取类型来判断。
    /// </summary>
    private void CheckOutlookAvailability()
    {
        try
        {
            var outlookType = Type.GetTypeFromProgID("Outlook.Application");
            IsOutlookAvailable = outlookType != null;
        }
        catch { IsOutlookAvailable = false; }
    }

    /// <summary>
    /// 保存当前设置到数据服务。
    /// </summary>
    [RelayCommand]
    private void Save()
    {
        var settings = _dataService.LoadSettings();
        settings.SendMethod = SendMethod;
        settings.Smtp.Host = SmtpHost;
        settings.Smtp.Port = SmtpPort;
        settings.Smtp.UseSsl = SmtpUseSsl;
        settings.Smtp.UserName = SmtpUserName;
        settings.Smtp.DisplayName = SmtpDisplayName;
        settings.Smtp.SenderEmail = SmtpSenderEmail;
        _dataService.SaveSettings(settings);
    }
}
