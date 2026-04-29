﻿﻿using CommunityToolkit.Mvvm.ComponentModel;

namespace TailorMail.Models;

/// <summary>
/// 收件人分组模型，用于将收件人按组进行分类管理。
/// 每个分组包含唯一的标识、分组名称以及该分组下的收件人列表。
/// </summary>
public class RecipientGroup
{
    /// <summary>
    /// 获取或设置分组的唯一标识符。默认值为新生成的 GUID 字符串。
    /// </summary>
    public string Id { get; set; } = Guid.NewGuid().ToString();

    /// <summary>
    /// 获取或设置分组名称。默认值为"默认分组"。
    /// </summary>
    public string Name { get; set; } = "默认分组";

    /// <summary>
    /// 获取或设置该分组下的收件人列表。默认为空列表。
    /// </summary>
    public List<Recipient> Recipients { get; set; } = [];
}

/// <summary>
/// 收件人模型，表示一个邮件收件人。
/// 继承自 <see cref="ObservableObject"/>，支持 MVVM 属性变更通知。
/// 包含收件人的基本信息（名称、简称）、多种邮件地址（收件人、抄送、密送）、
/// 备注信息、选中状态以及用于邮件模板替换的自定义变量字典。
/// </summary>
public partial class Recipient : ObservableObject
{
    /// <summary>
    /// 获取或设置收件人的唯一标识符。默认值为新生成的 GUID 字符串。
    /// </summary>
    public string Id { get; set; } = Guid.NewGuid().ToString();

    /// <summary>
    /// 获取或设置收件人姓名。用于在界面上显示及邮件模板变量替换。
    /// </summary>
    [ObservableProperty]
    private string _name = string.Empty;

    /// <summary>
    /// 获取或设置收件人简称。可用于附件自动匹配等场景，作为匹配关键字使用。
    /// </summary>
    [ObservableProperty]
    private string _shortName = string.Empty;

    /// <summary>
    /// 获取或设置收件人邮箱地址（主收件人）。
    /// 多个邮箱地址之间可使用分号、逗号或换行符分隔。
    /// </summary>
    [ObservableProperty]
    private string _toEmails = string.Empty;

    /// <summary>
    /// 获取或设置抄送邮箱地址。多个邮箱地址之间可使用分号、逗号或换行符分隔。
    /// </summary>
    [ObservableProperty]
    private string _ccEmails = string.Empty;

    /// <summary>
    /// 获取或设置密送邮箱地址。多个邮箱地址之间可使用分号、逗号或换行符分隔。
    /// </summary>
    [ObservableProperty]
    private string _bccEmails = string.Empty;

    /// <summary>
    /// 获取或设置收件人备注信息。可用于记录与该收件人相关的补充说明。
    /// </summary>
    [ObservableProperty]
    private string _remark = string.Empty;

    /// <summary>
    /// 获取或设置该收件人是否被选中。
    /// 用于在发送邮件时标记哪些收件人需要接收邮件。
    /// </summary>
    [ObservableProperty]
    private bool _isSelected;

    /// <summary>
    /// 获取或设置自定义变量字典。
    /// 键为变量名，值为变量值，用于邮件模板中的占位符替换，实现每封邮件内容的个性化定制。
    /// </summary>
    public Dictionary<string, string> Variables { get; set; } = [];

    /// <summary>
    /// 将主收件人邮箱字符串拆分为邮箱地址列表。
    /// </summary>
    /// <returns>拆分后的邮箱地址列表；若 <see cref="ToEmails"/> 为空则返回空列表。</returns>
    public List<string> GetToList() => SplitEmails(ToEmails);

    /// <summary>
    /// 将抄送邮箱字符串拆分为邮箱地址列表。
    /// </summary>
    /// <returns>拆分后的邮箱地址列表；若 <see cref="CcEmails"/> 为空则返回空列表。</returns>
    public List<string> GetCcList() => SplitEmails(CcEmails);

    /// <summary>
    /// 将密送邮箱字符串拆分为邮箱地址列表。
    /// </summary>
    /// <returns>拆分后的邮箱地址列表；若 <see cref="BccEmails"/> 为空则返回空列表。</returns>
    public List<string> GetBccList() => SplitEmails(BccEmails);

    /// <summary>
    /// 将包含多个邮箱地址的字符串按分隔符拆分为独立的邮箱地址列表。
    /// 支持的分隔符包括：英文分号(;)、英文逗号(,)、中文分号(；)、中文逗号(，)、换行符(\n)和回车符(\r)。
    /// 拆分后会自动去除每个地址的首尾空白字符，并过滤掉空字符串。
    /// </summary>
    /// <param name="emails">包含一个或多个邮箱地址的原始字符串。</param>
    /// <returns>拆分并清理后的邮箱地址列表；若输入为空或空白则返回空列表。</returns>
    private static List<string> SplitEmails(string emails)
    {
        if (string.IsNullOrWhiteSpace(emails)) return [];
        return emails.Split(';', ',', '；', '，', '\n', '\r')
            .Select(e => e.Trim())
            .Where(e => !string.IsNullOrEmpty(e))
            .ToList();
    }
}
