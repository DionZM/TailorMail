using System.Collections.Generic;
using TailorMail.Models;
using MailKit.Net.Smtp;
using MimeKit;

namespace TailorMail.Services;

/// <summary>
/// 基于 SMTP 协议的邮件发送器，实现 <see cref="IEmailSender"/> 接口。
/// 使用 MailKit 库通过 SMTP 协议直接与邮件服务器通信发送邮件。
/// 需要用户在设置中配置 SMTP 服务器地址、端口、加密方式、用户名等信息，
/// 密码在每次发送时由用户临时输入，不持久化存储。
/// </summary>
public class SmtpEmailSender : IEmailSender
{
    /// <inheritdoc/>
    public string Name => "SMTP";

    /// <inheritdoc/>
    /// <remarks>
    /// 实现流程：
    /// <list type="number">
    ///   <item>验证 SMTP 密码不为空</item>
    ///   <item>从数据服务加载 SMTP 服务器配置</item>
    ///   <item>构建 MimeMessage 邮件对象，设置发件人、收件人、主题</item>
    ///   <item>构建 HTML 正文和附件</item>
    ///   <item>连接 SMTP 服务器、认证、发送邮件、断开连接</item>
    /// </list>
    /// 所有 SMTP 操作均使用异步方法，避免阻塞 UI 线程。
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
            // SMTP 密码由用户在发送时临时输入，不持久化
            if (string.IsNullOrEmpty(smtpPassword))
                throw new InvalidOperationException("SMTP 密码不能为空");

            // 从数据服务获取 SMTP 配置
            var dataService = Helpers.ServiceHelper.GetRequiredService<IDataService>();
            var settings = dataService.LoadSettings().Smtp;

            // 构建邮件消息
            var message = new MimeMessage();

            // 设置发件人：优先使用 SenderEmail，若为空则使用 UserName
            var senderEmail = !string.IsNullOrEmpty(settings.SenderEmail) ? settings.SenderEmail : settings.UserName;
            message.From.Add(new MailboxAddress(settings.DisplayName, senderEmail));

            // 设置收件人/抄送/密送
            foreach (var to in recipient.GetToList())
                message.To.Add(MailboxAddress.Parse(to));
            foreach (var cc in recipient.GetCcList())
                message.Cc.Add(MailboxAddress.Parse(cc));
            foreach (var bcc in recipient.GetBccList())
                message.Bcc.Add(MailboxAddress.Parse(bcc));

            message.Subject = subject;

            // 构建邮件正文和附件
            var bodyBuilder = new BodyBuilder { HtmlBody = body };
            foreach (var filePath in attachments)
            {
                if (System.IO.File.Exists(filePath))
                    await bodyBuilder.Attachments.AddAsync(filePath);
            }

            message.Body = bodyBuilder.ToMessageBody();

            // 连接 SMTP 服务器并发送邮件
            using var client = new SmtpClient();
            await client.ConnectAsync(settings.Host, settings.Port, settings.UseSsl);
            await client.AuthenticateAsync(settings.UserName, smtpPassword);
            await client.SendAsync(message);
            await client.DisconnectAsync(true);

            result.Status = SendStatus.Success;
            result.SendTime = DateTime.Now;
        }
        catch (Exception ex)
        {
            result.Status = SendStatus.Failed;
            result.ErrorMessage = ex.Message;
        }

        return result;
    }
}
