﻿﻿using System.Collections.Generic;
using TailorMail.Models;

namespace TailorMail.Services;

/// <summary>
/// 邮件发送器接口，定义邮件发送的统一契约。
/// 不同的实现类支持通过不同的邮件发送通道（如 Outlook COM、SMTP 协议）发送邮件。
/// </summary>
public interface IEmailSender
{
    /// <summary>
    /// 获取发送方式的显示名称（如 "Outlook"、"SMTP"）。
    /// </summary>
    string Name { get; }

    /// <summary>
    /// 异步发送邮件给指定收件人。
    /// </summary>
    /// <param name="subject">邮件主题（已进行变量替换后的最终内容）。</param>
    /// <param name="body">邮件正文 HTML 内容（已进行变量替换后的最终内容）。</param>
    /// <param name="recipient">收件人对象，包含 TO/CC/BCC 邮箱地址。</param>
    /// <param name="attachments">附件文件路径列表（公共附件与收件人专属附件的合集）。</param>
    /// <param name="smtpPassword">
    /// SMTP 密码，仅在 SMTP 发送方式下使用。
    /// 密码不持久化存储，每次发送时由用户临时输入。
    /// </param>
    /// <returns>包含发送状态和结果信息的 <see cref="SendResult"/> 对象。</returns>
    Task<SendResult> SendAsync(
        string subject,
        string body,
        Recipient recipient,
        List<string> attachments,
        string? smtpPassword = null);
}
