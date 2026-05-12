using System.Collections.Generic;
using System.Threading;
using TailorMail.Models;
using MailKit.Net.Smtp;
using MimeKit;

namespace TailorMail.Services;

public class SmtpEmailSender : IEmailSender, IDisposable
{
    private SmtpClient? _client;
    private string? _currentHost;
    private int _currentPort;
    private volatile int _disposed; // T-04: Interlocked for thread-safe disposal

    // Bulk send cache
    private Models.SmtpSettings? _cachedSettings;
    private string? _cachedPassword;
    private MimePart[]? _cachedCommonAttachments;

    // S-01: Configurable timeout (ms)
    private const int ConnectTimeoutMs = 30_000;
    private const int SendTimeoutMs = 120_000;

    // S-02: Retry configuration
    private const int MaxRetries = 3;
    private static readonly int[] RetryDelaysMs = [1000, 2000, 4000];

    public string Name => "SMTP";

    /// <summary>
    /// Pre-establish connection and pre-read common attachments for bulk sending.
    /// </summary>
    public async Task PrepareForBulkSend(Models.SmtpSettings settings, string password, List<string> commonAttachmentPaths)
    {
        _cachedSettings = settings;
        _cachedPassword = password;
        await EnsureConnectedAsync(settings, password);

        // P-04: Read attachments into byte[] to avoid FileStream leak
        var parts = new List<MimePart>();
        foreach (var filePath in commonAttachmentPaths)
        {
            if (System.IO.File.Exists(filePath))
            {
                var part = await CreateMimePartFromFileAsync(filePath);
                if (part != null) parts.Add(part);
            }
        }
        _cachedCommonAttachments = parts.ToArray();
    }

    /// <summary>
    /// Set pre-read common attachments for bulk sending.
    /// </summary>
    public void SetCachedCommonAttachments(MimePart[] attachments)
    {
        _cachedCommonAttachments = attachments;
    }

    public MimePart[]? GetCachedCommonAttachments()
    {
        return _cachedCommonAttachments;
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

        // S-02: Retry with exponential backoff for transient errors
        for (int attempt = 0; attempt <= MaxRetries; attempt++)
        {
            try
            {
                var settings = smtpSettings ?? _cachedSettings;
                var password = smtpPassword ?? _cachedPassword;
                if (settings == null)
                    throw new InvalidOperationException("SMTP 设置未提供");
                if (string.IsNullOrEmpty(password))
                    throw new InvalidOperationException("SMTP 密码不能为空");

                await EnsureConnectedAsync(settings, password);

                // M-01: Set send timeout before sending
                if (_client != null) _client.Timeout = SendTimeoutMs;

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

                // H-08: Clone cached MimeParts for each message to avoid stream reuse issues
                if (_cachedCommonAttachments != null)
                {
                    foreach (var part in _cachedCommonAttachments)
                    {
                        var cloned = await CloneMimePartAsync(part);
                        if (cloned != null) bodyBuilder.Attachments.Add(cloned);
                    }
                }

                // P-04: Add per-recipient attachments with byte[] to avoid FileStream leak
                foreach (var filePath in attachments)
                {
                    if (System.IO.File.Exists(filePath))
                    {
                        var part = await CreateMimePartFromFileAsync(filePath);
                        if (part != null) bodyBuilder.Attachments.Add(part);
                    }
                }

                message.Body = bodyBuilder.ToMessageBody();

                await _client!.SendAsync(message);

                result.Status = SendStatus.Success;
                result.SendTime = DateTime.Now;
                return result; // Success, no retry needed
            }
            catch (Exception ex) when (IsTransientError(ex) && attempt < MaxRetries)
            {
                // S-02: Retry on transient errors (network, timeout)
                AppLogger.Warning($"SMTP发送瞬态错误 (尝试 {attempt + 1}/{MaxRetries + 1}): {recipient.Name} - {ex.Message}");
                try { await DisconnectAsync(); } catch { }
                await Task.Delay(RetryDelaysMs[attempt]);
            }
            catch (Exception ex)
            {
                result.Status = SendStatus.Failed;
                result.ErrorMessage = ex.Message;

                try { await DisconnectAsync(); }
                catch (Exception dcEx) { AppLogger.Warning($"发送失败后断开连接出错: {dcEx.Message}"); }
                return result;
            }
        }

        // C-02: Ensure status is Failed after all retries exhausted
        result.Status = SendStatus.Failed;
        result.ErrorMessage ??= "发送失败：所有重试已耗尽";
        return result;
    }

    /// <summary>
    /// H-08: Clone a MimePart by reading its content into a new MemoryStream.
    /// </summary>
    private static async Task<MimePart?> CloneMimePartAsync(MimePart source)
    {
        try
        {
            var ms = new System.IO.MemoryStream();
            await source.Content.DecodeToAsync(ms);
            ms.Position = 0;
            return new MimePart(source.ContentType)
            {
                Content = new MimeContent(ms),
                ContentDisposition = source.ContentDisposition,
                ContentTransferEncoding = source.ContentTransferEncoding,
                FileName = source.FileName
            };
        }
        catch { return null; }
    }

    /// <summary>
    /// S-02: Determine if an exception is transient.
    /// </summary>
    private static bool IsTransientError(Exception ex)
    {
        return ex is System.IO.IOException
            or System.Net.Sockets.SocketException
            or MailKit.ServiceNotConnectedException
            or MailKit.ProtocolException
            or TimeoutException
            || (ex is MailKit.Net.Smtp.SmtpCommandException cmdEx
                && cmdEx.StatusCode != SmtpStatusCode.MailboxUnavailable
                && cmdEx.StatusCode != SmtpStatusCode.MailboxNameNotAllowed);
    }

    /// <summary>
    /// P-04: Read file into byte[] then wrap in MemoryStream to avoid FileStream leak.
    /// </summary>
    private static async Task<MimePart?> CreateMimePartFromFileAsync(string filePath)
    {
        try
        {
            var fileName = System.IO.Path.GetFileName(filePath);
            var fileBytes = await System.IO.File.ReadAllBytesAsync(filePath);
            var stream = new System.IO.MemoryStream(fileBytes);
            var contentType = MimeTypes.GetMimeType(fileName);
            var contentTypeParsed = ContentType.Parse(contentType);
            return await Task.FromResult<MimePart?>(new MimePart(contentTypeParsed)
            {
                Content = new MimeContent(stream),
                ContentDisposition = new ContentDisposition(ContentDisposition.Attachment),
                ContentTransferEncoding = ContentEncoding.Base64,
                FileName = fileName
            });
        }
        catch (Exception ex)
        {
            AppLogger.Warning($"创建附件MimePart失败: {filePath} - {ex.Message}");
            return await Task.FromResult<MimePart?>(null);
        }
    }

    private async Task EnsureConnectedAsync(Models.SmtpSettings settings, string password)
    {
        // S-03: Check IsConnected AND IsAuthenticated
        if (_client != null && _client.IsConnected && _client.IsAuthenticated
            && _currentHost == settings.Host && _currentPort == settings.Port)
            return;

        await DisconnectAsync();

        _client = new SmtpClient { Timeout = ConnectTimeoutMs }; // S-01: Set timeout
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
            catch (Exception ex)
            {
                AppLogger.Warning($"SMTP断开连接时出错: {ex.Message}");
            }
            _client.Dispose();
            _client = null;
            _currentHost = null;
            _currentPort = 0;
        }
    }

    public void Dispose()
    {
        // T-04: Thread-safe disposal using Interlocked
        if (Interlocked.Exchange(ref _disposed, 1) != 0)
            return;

        CleanupBulkSend();

        if (_client != null)
        {
            try
            {
                if (_client.IsConnected)
                    _client.Disconnect(true);
            }
            catch (Exception ex)
            {
                AppLogger.Warning($"SMTP Dispose断开连接时出错: {ex.Message}");
            }
            _client.Dispose();
            _client = null;
        }
    }

    /// <summary>
    /// S-04: Async test send method. M-01: Accept settings directly instead of ServiceLocator.
    /// </summary>
    public async Task<(bool success, string error)> SendTestAsync(
        string fromAddress, string displayName, string password,
        string toAddress, string subject, string bodyHtml,
        string[] attachments, Models.SmtpSettings settings)
    {
        try
        {
            using var client = new SmtpClient { Timeout = ConnectTimeoutMs }; // S-01
            await client.ConnectAsync(settings.Host, settings.Port, settings.UseSsl);
            await client.AuthenticateAsync(settings.UserName, password);

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

            await client.SendAsync(message);
            await client.DisconnectAsync(true);

            return (true, string.Empty);
        }
        catch (Exception ex)
        {
            return (false, ex.Message);
        }
    }

    /// <summary>
    /// Legacy synchronous test send (kept for backward compatibility).
    /// </summary>
    public bool SendTest(string fromAddress, string displayName, string password, string toAddress,
        string subject, string bodyHtml, string[] attachments, out string error)
    {
        try
        {
            var dataService = Helpers.ServiceHelper.GetRequiredService<IDataService>();
            var settings = dataService.LoadSettings().Smtp;

            using var client = new SmtpClient { Timeout = ConnectTimeoutMs }; // S-01
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
