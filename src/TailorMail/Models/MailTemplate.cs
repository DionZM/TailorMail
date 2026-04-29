﻿﻿﻿namespace TailorMail.Models;

/// <summary>
/// 邮件模板模型，用于定义可复用的邮件内容模板。
/// 包含模板的标识、名称、邮件主题、邮件正文及创建时间。
/// 模板中可使用占位符（如 {变量名}），发送时结合收件人的自定义变量进行替换，
/// 实现邮件内容的个性化定制。
/// </summary>
public class MailTemplate
{
    /// <summary>
    /// 获取或设置模板的唯一标识符。默认值为新生成的 GUID 字符串。
    /// </summary>
    public string Id { get; set; } = Guid.NewGuid().ToString();

    /// <summary>
    /// 获取或设置模板名称。用于在界面上标识和选择不同的邮件模板。
    /// </summary>
    public string Name { get; set; } = string.Empty;

    /// <summary>
    /// 获取或设置邮件主题模板。
    /// 可包含占位符变量，发送时将根据收件人的变量字典进行替换。
    /// </summary>
    public string Subject { get; set; } = string.Empty;

    /// <summary>
    /// 获取或设置邮件正文模板（XAML 格式，用于富文本编辑器状态保存）。
    /// 可包含占位符变量，发送时将根据收件人的变量字典进行替换。
    /// </summary>
    public string Body { get; set; } = string.Empty;

    /// <summary>
    /// 获取或设置模板的创建时间。默认值为当前系统时间。
    /// </summary>
    public DateTime CreatedAt { get; set; } = DateTime.Now;
}
