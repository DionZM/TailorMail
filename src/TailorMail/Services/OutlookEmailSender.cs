﻿using System.Collections.Generic;
using System.Runtime.InteropServices;
using TailorMail.Models;

namespace TailorMail.Services;

/// <summary>
/// 基于 Outlook COM 自动化的邮件发送器，实现 <see cref="IEmailSender"/> 接口。
/// 通过 COM Interop 调用本地安装的 Microsoft Outlook 应用程序发送邮件。
/// 需要用户已安装并配置好 Outlook 桌面客户端。
/// </summary>
public class OutlookEmailSender : IEmailSender
{
    /// <inheritdoc/>
    public string Name => "Outlook";

    /// <inheritdoc/>
    /// <remarks>
    /// 实现流程：
    /// <list type="number">
    ///   <item>通过 ProgID "Outlook.Application" 获取 Outlook COM 类型</item>
    ///   <item>创建 Outlook Application 实例和 MailItem 对象</item>
    ///   <item>设置邮件主题、HTML 正文、收件人/抄送/密送地址</item>
    ///   <item>添加附件文件（仅添加存在的文件）</item>
    ///   <item>调用 MailItem.Send() 发送邮件</item>
    ///   <item>在 finally 块中释放所有 COM 对象，防止内存泄漏</item>
    /// </list>
    /// 整个发送过程在 <see cref="Task.Run"/> 中执行，避免阻塞 UI 线程。
    /// </remarks>
    public async Task<SendResult> SendAsync(
        string subject,
        string body,
        Recipient recipient,
        List<string> attachments,
        string? smtpPassword = null)
    {
        var result = new SendResult
        {
            RecipientId = recipient.Id,
            RecipientName = recipient.Name,
            Status = SendStatus.Sending
        };

        try
        {
            await Task.Run(() =>
            {
                // 通过 ProgID 获取 Outlook COM 类型，若未安装则抛出异常
                var outlookType = Type.GetTypeFromProgID("Outlook.Application")
                    ?? throw new InvalidOperationException("未检测到 Outlook，请确认已安装 Microsoft Outlook。");

                dynamic? outlookApp = null;
                dynamic? mailItem = null;
                try
                {
                    // 创建 Outlook 应用实例和邮件项
                    outlookApp = Activator.CreateInstance(outlookType)!;
                    mailItem = outlookApp.CreateItem(0); // 0 = olMailItem

                    // 设置邮件基本信息
                    mailItem.Subject = subject;
                    mailItem.HTMLBody = body;

                    // 设置收件人地址（支持 TO/CC/BCC）
                    var toList = recipient.GetToList();
                    var ccList = recipient.GetCcList();
                    var bccList = recipient.GetBccList();

                    if (toList.Count > 0) mailItem.To = string.Join(";", toList);
                    if (ccList.Count > 0) mailItem.CC = string.Join(";", ccList);
                    if (bccList.Count > 0) mailItem.BCC = string.Join(";", bccList);

                    // 添加附件（1 = olByValue，以值类型嵌入附件）
                    foreach (var filePath in attachments)
                    {
                        if (System.IO.File.Exists(filePath))
                            mailItem.Attachments.Add(filePath, 1);
                    }

                    // 发送邮件
                    mailItem.Send();
                }
                finally
                {
                    // 释放 COM 对象，防止 Outlook 进程残留
                    if (mailItem != null) Marshal.ReleaseComObject(mailItem);
                    if (outlookApp != null) Marshal.ReleaseComObject(outlookApp);
                }
            });

            result.Status = SendStatus.Success;
            result.SendTime = DateTime.Now;
        }
        catch (Exception ex)
        {
            AppLogger.Error($"Outlook发送失败 - {recipient.Name}", ex);
            result.Status = SendStatus.Failed;
            result.ErrorMessage = ex.Message;
        }

        return result;
    }
}
