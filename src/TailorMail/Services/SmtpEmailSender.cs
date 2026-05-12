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

    // Bulk send cache
    private Models.SmtpSettings? _cachedSettings;
    private string? _cachedPassword;
    private MimePart[]? _cachedCommonAttachments;

    public string Name => "SMTP";

    /// <summary>
    /// Pre-establish connection and pre-read common attachments for bulk sending.
    /// Call once before the send loop to avoid per-recipient overhead.
    /// </summary>
    public async Task PrepareForBulkSend(Models.SmtpSettings settings, string password, List<string> commonAttachmentPaths)
    {
        _cachedSettings = settings;
        _cachedPassword = password;
        await EnsureConnectedAsync(settings, password);

        // Pre-read common attachments into MimePart objects
        var parts = new List<MimePart>();
        foreach (var filePath in commonAttachmentPaths)
        {
            if (System.IO.File.Exists(filePath))
            {
                var part = await CreateMimePartAsync(filePath);
                if (part != null) parts.Add(part);
            }
        }
        _cachedCommonAttachments = parts.ToArray();
    }

    /// <summary>
    /// Release cached bulk-send resources.
    /// </summary>
    public void CleanupBulkSend()
    {
        _cachedCommonAttachments = null;
        _cachedSettings = null;
        _cachedPassword = null;
    }

    public async Task<SendResult> SendAsync(
        string subject,
        string body,
        Recipient recipient,
        List<string> attachments,
        string? smtpPassword = null,
        Models.SmtpSettings? smtpSettings = null)
    {
        var result = new SendResult
        {
            RecipientId = recipient.Id,
            RecipientName = recipient.Name,
            Status = SendStatus.Sending
        };

        try
        {
            var settings = smtpSettings ?? _cachedSettings;
            var password = smtpPassword ?? _cachedPassword;
            if (settings == null)
                throw new InvalidOperationException("SMTP 设置未提供");
            if (string.IsNullOrEmpty(password))
                throw new InvalidOperationException("SMTP 密码不能为空");

            await EnsureConnectedAsync(settings, password);

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

            // Add pre-cached common attachments
            if (_cachedCommonAttachments != null)
            {
                foreach (var part in _cachedCommonAttachments)
                    bodyBuilder.Attachments.Add(part);
            }

            // Add per-recipient attachments (not cached, varies per recipient)
            foreach (var filePath in attachments)
            {
                if (System.IO.File.Exists(filePath))
                {
                    var part = await CreateMimePartAsync(filePath);
                    if (part != null) bodyBuilder.Attachments.Add(part);
                }
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

    private static async Task<MimePart?> CreateMimePartAsync(string filePath)
    {
        try
        {
            var fileName = System.IO.Path.GetFileName(filePath);
            var contentType = MimeTypes.GetMimeType(filePath);
            var mediaType = contentType.Split('/');
            var bytes = await System.IO.File.ReadAllBytesAsync(filePath);
            return new MimePart(mediaType[0], mediaType[1])
            {
                Content = new MimeContent(new System.IO.MemoryStream(bytes)),
                ContentDisposition = new ContentDisposition(ContentDisposition.Attachment),
                ContentTransferEncoding = ContentEncoding.Base64,
                FileName = fileName
            };
        }
        catch
        {
            return null;
        }
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
        CleanupBulkSend();

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

    public bool SendTest(string fromAddress, string displayName, string password, string toAddress,
        string subject, string bodyHtml, string[] attachments, out string error)
    {
        try
        {
            var dataService = Helpers.ServiceHelper.GetRequiredService<IDataService>();
            var settings = dataService.LoadSettings().Smtp;

            using var client = new SmtpClient();
            client.Connect(settings.Host, settings.Port, settings.UseSsl);
            client.Authenticate(settings.UserName, password);

            var message = new MimeMessage();
            message.From.Add(new MailboxAddress(displayName, fromAddress));
            message.To.Add(MailboxAddress.Parse(toAddress));
            message.Subject = subject;

            var bodyBuilder = new BodyBuilder { HtmlBody = bodyHtml };
            foreach (var filePath in attachments)
            {
                if (!string.IsNullOrEmpty(filePath) && System.IO.File.Exists(filePath))
                    bodyBuilder.Attachments.Add(filePath);
            }

            message.Body = bodyBuilder.ToMessageBody();

            client.Send(message);
            client.Disconnect(true);

            error = string.Empty;
            return true;
        }
        catch (Exception ex)
        {
            error = ex.Message;
            return false;
        }
    }
}
