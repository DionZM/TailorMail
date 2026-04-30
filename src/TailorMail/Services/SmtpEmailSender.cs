using System.Collections.Generic;
using TailorMail.Models;
using MailKit.Net.Smtp;
using MimeKit;

namespace TailorMail.Services;

public class SmtpEmailSender : IEmailSender, IDisposable
{
    private SmtpClient? _client;
    private string? _currentHost;
    private int _currentPort;
    private bool _disposed;

    public string Name => "SMTP";

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
            if (string.IsNullOrEmpty(smtpPassword))
                throw new InvalidOperationException("SMTP 密码不能为空");

            var dataService = Helpers.ServiceHelper.GetRequiredService<IDataService>();
            var settings = dataService.LoadSettings().Smtp;

            await EnsureConnectedAsync(settings, smtpPassword);

            var message = new MimeMessage();

            var senderEmail = !string.IsNullOrEmpty(settings.SenderEmail) ? settings.SenderEmail : settings.UserName;
            message.From.Add(new MailboxAddress(settings.DisplayName, senderEmail));

            foreach (var to in recipient.GetToList())
                message.To.Add(MailboxAddress.Parse(to));
            foreach (var cc in recipient.GetCcList())
                message.Cc.Add(MailboxAddress.Parse(cc));
            foreach (var bcc in recipient.GetBccList())
                message.Bcc.Add(MailboxAddress.Parse(bcc));

            message.Subject = subject;

            var bodyBuilder = new BodyBuilder { HtmlBody = body };
            foreach (var filePath in attachments)
            {
                if (System.IO.File.Exists(filePath))
                    await bodyBuilder.Attachments.AddAsync(filePath);
            }

            message.Body = bodyBuilder.ToMessageBody();

            await _client!.SendAsync(message);

            result.Status = SendStatus.Success;
            result.SendTime = DateTime.Now;
        }
        catch (Exception ex)
        {
            result.Status = SendStatus.Failed;
            result.ErrorMessage = ex.Message;

            try { await DisconnectAsync(); }
            catch { }
        }

        return result;
    }

    private async Task EnsureConnectedAsync(Models.SmtpSettings settings, string password)
    {
        if (_client != null && _client.IsConnected && _currentHost == settings.Host && _currentPort == settings.Port)
            return;

        await DisconnectAsync();

        _client = new SmtpClient();
        await _client.ConnectAsync(settings.Host, settings.Port, settings.UseSsl);
        await _client.AuthenticateAsync(settings.UserName, password);
        _currentHost = settings.Host;
        _currentPort = settings.Port;
    }

    private async Task DisconnectAsync()
    {
        if (_client != null)
        {
            try
            {
                if (_client.IsConnected)
                    await _client.DisconnectAsync(true);
            }
            catch { }
            _client.Dispose();
            _client = null;
            _currentHost = null;
            _currentPort = 0;
        }
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;

        if (_client != null)
        {
            try
            {
                if (_client.IsConnected)
                    _client.Disconnect(true);
            }
            catch { }
            _client.Dispose();
            _client = null;
        }
    }
}
