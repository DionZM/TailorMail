using System.Collections.ObjectModel;
using System.Threading;
using CommunityToolkit.Mvvm.ComponentModel;
using TailorMail.Models;
using TailorMail.Services;
using Microsoft.Extensions.DependencyInjection;

namespace TailorMail.ViewModels;

public partial class SendViewModel : ObservableObject
{
    private readonly IDataService _dataService;
    private CancellationTokenSource? _cts;
    // T-05: volatile for thread-safe visibility
    private volatile List<Recipient>? _cachedSelectedRecipients;

    [ObservableProperty]
    private ObservableCollection<SendResult> _sendResults = [];

    [ObservableProperty]
    private bool _isSending;

    [ObservableProperty]
    private int _totalCount;

    [ObservableProperty]
    private int _successCount;

    [ObservableProperty]
    private int _failedCount;

    [ObservableProperty]
    private double _progressValue;

    [ObservableProperty]
    private string _statusText = "就绪";

    [ObservableProperty]
    private SendMethod _sendMethod;

    private readonly System.Windows.Threading.Dispatcher _dispatcher;

    public SendViewModel(IDataService dataService)
    {
        _dataService = dataService;
        _dispatcher = System.Windows.Application.Current.Dispatcher;
        var settings = _dataService.LoadSettings();
        SendMethod = settings.SendMethod;
    }

    public void ReloadSettings()
    {
        var settings = _dataService.LoadSettings();
        SendMethod = settings.SendMethod;
    }

    public List<Recipient> GetSelectedRecipients()
    {
        if (_cachedSelectedRecipients != null) return _cachedSelectedRecipients;
        var groups = _dataService.LoadRecipientGroups();
        _cachedSelectedRecipients = groups.SelectMany(g => g.Recipients).Where(r => r.IsSelected).ToList();
        return _cachedSelectedRecipients;
    }

    public Recipient? FindRecipient(string id)
    {
        if (_cachedSelectedRecipients != null)
            return _cachedSelectedRecipients.FirstOrDefault(r => r.Id == id);
        var groups = _dataService.LoadRecipientGroups();
        return groups.SelectMany(g => g.Recipients).FirstOrDefault(r => r.Id == id);
    }

    public async Task ExecuteSend(List<Recipient> selectedRecipients, string? smtpPassword = null, bool append = false)
    {
        IsSending = true;

        try
        {
            await ExecuteSendCore(selectedRecipients, smtpPassword, append);
        }
        catch (Exception ex)
        {
            AppLogger.Error("ExecuteSend 未处理异常", ex);
            StatusText = $"发送出错：{ex.Message}";
        }
        finally
        {
            IsSending = false;
            _cts?.Dispose();
            _cts = null;
            _cachedSelectedRecipients = null;
        }
    }

    private async Task ExecuteSendCore(List<Recipient> selectedRecipients, string? smtpPassword, bool append)
    {
        AppLogger.Info($"ExecuteSendCore 开始: {selectedRecipients.Count} 个收件人, 方式={SendMethod}, append={append}");

        // --- Build result map for O(1) lookup ---
        var resultMap = new Dictionary<string, SendResult>(selectedRecipients.Count);

        if (!append)
        {
            TotalCount = selectedRecipients.Count;
            SuccessCount = 0;
            FailedCount = 0;
            ProgressValue = 0;
            SendResults.Clear();
        }
        else
        {
            foreach (var sr in SendResults)
                resultMap[sr.RecipientId] = sr;
        }

        // Populate SendResults + resultMap
        foreach (var r in selectedRecipients)
        {
            if (append && resultMap.ContainsKey(r.Id))
                continue;

            var sr = new SendResult { RecipientId = r.Id, RecipientName = r.Name };
            SendResults.Add(sr);
            resultMap[r.Id] = sr;
        }

        // H-07: Snapshot variables for thread-safe access during send
        foreach (var r in selectedRecipients)
        {
            if (r.Variables.Count > 0)
                r.Variables = new Dictionary<string, string>(r.Variables);
        }

        _cts = new CancellationTokenSource();

        // --- Load settings and dependencies once ---
        var settings = _dataService.LoadSettings();
        var attachConfig = _dataService.LoadAttachmentConfig();
        var recipientAttachMap = attachConfig.RecipientAttachments
            .ToDictionary(ua => ua.RecipientId);

        // --- Resolve sender and prepare for bulk send ---
        IEmailSender sender;
        SmtpEmailSender? smtpForCleanup = null;
        var useConcurrentSmtp = SendMethod == SendMethod.Smtp && selectedRecipients.Count > 1 && settings.SmtpConcurrency > 1;

        if (SendMethod == SendMethod.Outlook)
        {
            AppLogger.Info("发送方式: Outlook, 准备批量发送");
            var outlook = App.Services.GetRequiredService<OutlookEmailSender>();
            outlook.PrepareForBulkSend();
            sender = outlook;
        }
        else if (useConcurrentSmtp)
        {
            AppLogger.Info($"发送方式: SMTP 并发, {selectedRecipients.Count} 个收件人");
            // C-03: Concurrent path creates its own senders; no need to prepare here
            sender = null!; // not used in concurrent path
        }
        else
        {
            AppLogger.Info("发送方式: SMTP 单线程");
            var smtp = App.Services.GetRequiredService<SmtpEmailSender>();
            smtpForCleanup = smtp;
            await smtp.PrepareForBulkSend(
                settings.Smtp,
                smtpPassword ?? "",
                attachConfig.CommonAttachments);
            sender = smtp;
        }

        // U-01: Build HTML body on current thread (UI thread at entry point)
        string htmlBody;
        try
        {
            if (!string.IsNullOrEmpty(settings.LastBodyXaml))
            {
                try
                {
                    var doc = new System.Windows.Documents.FlowDocument();
                    Helpers.FlowDocumentHelper.LoadFromXaml(doc, settings.LastBodyXaml);
                    var bodyContent = Helpers.FlowDocumentHelper.ToHtml(doc);
                    htmlBody = Helpers.FlowDocumentHelper.WrapAsEmailDocument(bodyContent);
                }
                catch
                {
                    var bodyContent = Helpers.FlowDocumentHelper.PlainTextToHtml(settings.LastBody);
                    htmlBody = Helpers.FlowDocumentHelper.WrapAsEmailDocument(bodyContent);
                }
            }
            else
            {
                var bodyContent = Helpers.FlowDocumentHelper.PlainTextToHtml(settings.LastBody);
                htmlBody = Helpers.FlowDocumentHelper.WrapAsEmailDocument(bodyContent);
            }
        }
        catch (Exception ex)
        {
            AppLogger.Error("构建邮件HTML失败", ex);
            htmlBody = $"<html><body><p>{System.Net.WebUtility.HtmlEncode(settings.LastBody)}</p></body></html>";
        }

        // Append signature
        var signature = settings.Signature?.Trim();
        if (!string.IsNullOrEmpty(signature))
        {
            var sigHtml = System.Net.WebUtility.HtmlEncode(signature).Replace("\n", "<br/>");
            htmlBody = htmlBody.Replace("</body>",
                $"<br/><br/><div style='border-top:1px solid #ccc;padding-top:8px;margin-top:8px;color:#666;font-size:13px;'>{sigHtml}</div></body>");
        }

        // --- Progress throttling ---
        var lastProgressUpdate = DateTime.UtcNow;
        var progressUpdateInterval = TimeSpan.FromMilliseconds(200);

        // --- Send loop ---
        if (useConcurrentSmtp)
        {
            // Concurrent SMTP sending
            await SendConcurrentSmtp(selectedRecipients, resultMap, recipientAttachMap,
                settings, htmlBody, smtpPassword, lastProgressUpdate, progressUpdateInterval,
                attachConfig.CommonAttachments);
        }
        else
        {
            // Sequential sending (Outlook or single recipient)
            await SendSequential(selectedRecipients, resultMap, recipientAttachMap,
                sender, settings, htmlBody, smtpPassword, lastProgressUpdate, progressUpdateInterval);
        }

        AppLogger.Info($"发送循环结束: 成功={SuccessCount}, 失败={FailedCount}");

        StatusText = _cts.IsCancellationRequested
            ? $"已取消: 成功 {SuccessCount} 封, 失败 {FailedCount} 封"
            : $"发送完成: 成功 {SuccessCount} 封, 失败 {FailedCount} 封";

        // R-03: Properly dispose sender
        if (smtpForCleanup != null)
            smtpForCleanup.Dispose();
        else if (sender is IDisposable disposable)
            disposable.Dispose();
    }

    private async Task SendSequential(
        List<Recipient> recipients,
        Dictionary<string, SendResult> resultMap,
        Dictionary<string, RecipientAttachment> recipientAttachMap,
        IEmailSender sender,
        AppSettings settings,
        string htmlBody,
        string? smtpPassword,
        DateTime lastProgressUpdate,
        TimeSpan progressUpdateInterval)
    {
        for (int i = 0; i < recipients.Count; i++)
        {
            if (_cts!.IsCancellationRequested) break;

            var recipient = recipients[i];
            StatusText = $"正在发送: {recipient.Name} ({i + 1}/{recipients.Count})...";

            var perRecipientAttachments = GetPerRecipientAttachments(recipient.Id, recipientAttachMap);
            var subject = VariablesViewModel.ProcessBodyFast(settings.LastSubject, recipient);
            var body = VariablesViewModel.ProcessBodyFast(htmlBody, recipient);

            var existing = resultMap.GetValueOrDefault(recipient.Id);
            if (existing != null) existing.Status = SendStatus.Sending;

            AppLogger.Info($"正在发送第 {i + 1}/{recipients.Count} 封: {recipient.Name}");
            var result = await sender.SendAsync(subject, body, recipient, perRecipientAttachments, smtpPassword, settings.Smtp);
            AppLogger.Info($"发送结果: {recipient.Name} => {result.Status}");

            // T-01/T-03: Update UI-bound properties on UI thread
            if (existing != null)
            {
                _dispatcher.Invoke(() =>
                {
                    existing.Status = result.Status;
                    existing.ErrorMessage = result.ErrorMessage;
                    existing.SendTime = result.SendTime;
                });
            }

            if (result.Status == SendStatus.Success) SuccessCount++;
            else FailedCount++;

            ProgressValue = (i + 1) * 100.0 / recipients.Count;

            var interval = settings.SendIntervalMs;
            if (interval > 0 && i < recipients.Count - 1)
                await Task.Delay(interval, _cts.Token);
        }
    }

    private async Task SendConcurrentSmtp(
        List<Recipient> recipients,
        Dictionary<string, SendResult> resultMap,
        Dictionary<string, RecipientAttachment> recipientAttachMap,
        AppSettings settings,
        string htmlBody,
        string? smtpPassword,
        DateTime lastProgressUpdate,
        TimeSpan progressUpdateInterval,
        List<string> commonAttachments)
    {
        var concurrency = Math.Max(1, settings.SmtpConcurrency);
        var semaphore = new SemaphoreSlim(concurrency, concurrency);
        var completedCount = 0;
        var lockObj = new object();

        AppLogger.Info($"SMTP 并发发送开始: {recipients.Count} 个收件人, 并发数={concurrency}");

        // P-05: Each sender gets its own connection (no shared MimePart with FileStream)
        var senders = new SmtpEmailSender[concurrency];
        for (int s = 0; s < concurrency; s++)
        {
            senders[s] = new SmtpEmailSender();
            // Only first sender reads common attachments, rest get empty list
            await senders[s].PrepareForBulkSend(
                settings.Smtp, smtpPassword ?? "",
                s == 0 ? commonAttachments : new List<string>());
        }

        // Share the byte[]-backed MimeParts (P-04/P-05: MemoryStream is safe to share)
        var sharedAttachments = senders[0].GetCachedCommonAttachments();
        for (int s = 1; s < concurrency; s++)
        {
            if (sharedAttachments != null)
                senders[s].SetCachedCommonAttachments(sharedAttachments);
        }

        // T-02: Use atomic counter for sender assignment
        var senderCounter = -1;

        try
        {
            var tasks = recipients.Select(async (recipient, _) =>
            {
                if (_cts!.IsCancellationRequested) return;

                await semaphore.WaitAsync(_cts.Token);
                // T-02: Each task gets its own sender via atomic increment
                var taskSenderIndex = Interlocked.Increment(ref senderCounter) % concurrency;
                var sender = senders[taskSenderIndex];
                try
                {
                    var perRecipientAttachments = GetPerRecipientAttachments(recipient.Id, recipientAttachMap);
                    var subject = VariablesViewModel.ProcessBodyFast(settings.LastSubject, recipient);
                    var body = VariablesViewModel.ProcessBodyFast(htmlBody, recipient);

                    var existing = resultMap.GetValueOrDefault(recipient.Id);
                    if (existing != null)
                    {
                        // H-06: Use BeginInvoke to avoid blocking concurrent threads
                        _dispatcher.BeginInvoke(() => existing.Status = SendStatus.Sending);
                    }

                    var result = await sender.SendAsync(subject, body, recipient, perRecipientAttachments, smtpPassword, settings.Smtp);

                    if (existing != null)
                    {
                        // H-06: Use BeginInvoke to avoid blocking concurrent threads
                        _dispatcher.BeginInvoke(() =>
                        {
                            existing.Status = result.Status;
                            existing.ErrorMessage = result.ErrorMessage;
                            existing.SendTime = result.SendTime;
                        });
                    }

                    lock (lockObj)
                    {
                        if (result.Status == SendStatus.Success) SuccessCount++;
                        else FailedCount++;
                        completedCount++;
                        ProgressValue = completedCount * 100.0 / recipients.Count;
                    }
                }
                finally
                {
                    semaphore.Release();
                }
            });

            await Task.WhenAll(tasks);
        }
        finally
        {
            foreach (var s in senders)
                s.Dispose();
        }
    }

    private static List<string> GetPerRecipientAttachments(
        string recipientId,
        Dictionary<string, RecipientAttachment> recipientAttachMap)
    {
        if (recipientAttachMap.TryGetValue(recipientId, out var ua))
            return new List<string>(ua.Files);
        return [];
    }

    public async Task RetryAllFailed(string? smtpPassword = null)
    {
        var failedIds = SendResults.Where(r => r.Status == SendStatus.Failed)
            .Select(r => r.RecipientId).ToHashSet();
        if (failedIds.Count == 0) return;
        var groups = _dataService.LoadRecipientGroups();
        var failedRecipients = groups.SelectMany(g => g.Recipients).Where(r => failedIds.Contains(r.Id)).ToList();
        await ExecuteSend(failedRecipients, smtpPassword, append: true);
    }

    public async Task RetryOne(string recipientId, string? smtpPassword = null)
    {
        var recipient = FindRecipient(recipientId);
        if (recipient == null) return;
        await ExecuteSend(new List<Recipient> { recipient }, smtpPassword, append: true);
    }

    public void CancelSend()
    {
        _cts?.Cancel();
    }
}
