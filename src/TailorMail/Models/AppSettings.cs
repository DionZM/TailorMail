namespace TailorMail.Models;

/// <summary>
/// 应用程序全局配置模型，包含邮件发送方式、SMTP 服务器设置
/// 以及上次使用的邮件主题和正文等持久化信息。
/// 该配置会在应用启动时加载、关闭时保存，用于恢复用户上次的编辑状态。
/// </summary>
public class AppSettings
{
    /// <summary>
    /// 获取或设置邮件发送方式。默认值为 <see cref="SendMethod.Outlook"/>，即通过 Outlook 客户端发送。
    /// </summary>
    public SendMethod SendMethod { get; set; } = SendMethod.Outlook;

    /// <summary>
    /// 获取或设置 SMTP 服务器配置信息。
    /// 当 <see cref="SendMethod"/> 为 <see cref="SendMethod.Smtp"/> 时使用此配置。
    /// </summary>
    public SmtpSettings Smtp { get; set; } = new();

    /// <summary>
    /// 获取或设置上次使用的邮件主题。用于在应用重启后恢复用户上次的编辑内容。
    /// </summary>
    public string LastSubject { get; set; } = string.Empty;

    /// <summary>
    /// 获取或设置上次使用的邮件正文（纯文本或 HTML 格式）。
    /// 用于在应用重启后恢复用户上次的编辑内容。
    /// </summary>
    public string LastBody { get; set; } = string.Empty;

    /// <summary>
    /// 获取或设置上次使用的邮件正文的 XAML 序列化内容。
    /// 用于在 WPF 富文本编辑器中恢复格式化的邮件正文。
    /// </summary>
    public string LastBodyXaml { get; set; } = string.Empty;
}

/// <summary>
/// SMTP 服务器配置模型，包含连接 SMTP 服务器所需的全部参数。
/// 当选择 SMTP 方式发送邮件时，将使用此配置建立与邮件服务器的连接。
/// </summary>
public class SmtpSettings
{
    /// <summary>
    /// 获取或设置 SMTP 服务器主机地址。例如："smtp.qq.com"、"smtp.gmail.com" 等。
    /// </summary>
    public string Host { get; set; } = string.Empty;

    /// <summary>
    /// 获取或设置 SMTP 服务器端口号。默认值为 587（TLS 加密端口）。
    /// 常见端口：25（非加密）、465（SSL）、587（TLS）。
    /// </summary>
    public int Port { get; set; } = 587;

    /// <summary>
    /// 获取或设置是否使用 SSL/TLS 加密连接。默认值为 <c>true</c>，建议在生产环境中保持启用以确保通信安全。
    /// </summary>
    public bool UseSsl { get; set; } = true;

    /// <summary>
    /// 获取或设置 SMTP 服务器认证的用户名。通常为邮箱地址，例如 "user@example.com"。
    /// </summary>
    public string UserName { get; set; } = string.Empty;

    /// <summary>
    /// 获取或设置发件人显示名称。收件人看到的发件人名称，例如"公司名称"或"张三"。
    /// </summary>
    public string DisplayName { get; set; } = string.Empty;

    /// <summary>
    /// 获取或设置发件人邮箱地址。即实际发送邮件的邮箱地址，需与 <see cref="UserName"/> 对应的邮箱一致。
    /// </summary>
    public string SenderEmail { get; set; } = string.Empty;
}

/// <summary>
/// 邮件发送方式枚举，定义应用程序支持的邮件发送途径。
/// </summary>
public enum SendMethod
{
    /// <summary>
    /// 通过本地 Outlook 客户端发送邮件。需要用户已安装并配置 Outlook 应用程序。
    /// </summary>
    Outlook,

    /// <summary>
    /// 通过 SMTP 协议直接发送邮件。需要配置 <see cref="SmtpSettings"/> 中的服务器连接信息。
    /// </summary>
    Smtp
}
