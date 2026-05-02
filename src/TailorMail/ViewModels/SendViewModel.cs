using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using TailorMail.Models;
using TailorMail.Services;
using Microsoft.Extensions.DependencyInjection;

namespace TailorMail.ViewModels;

public partial class SendViewModel : ObservableObject
{
    private readonly IDataService _dataService;
    private CancellationTokenSource? _cts;

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

    public SendViewModel(IDataService dataService)
    {
        _dataService = dataService;
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
        var groups = _dataService.LoadRecipientGroups();
        return groups.SelectMany(g => g.Recipients).Where(r => r.IsSelected).ToList();
    }

    public Recipient? FindRecipient(string id)
    {
        var groups = _dataService.LoadRecipientGroups();
        return groups.SelectMany(g => g.Recipients).FirstOrDefault(r => r.Id == id);
    }

    public async Task ExecuteSend(List<Recipient> selectedRecipients, string? smtpPassword = null, bool append = false)
    {
        IsSending = true;

        if (!append)
        {
            TotalCount = selectedRecipients.Count;
            SuccessCount = 0;
            FailedCount = 0;
            ProgressValue = 0;
            SendResults.Clear();
            foreach (var r in selectedRecipients)
                SendResults.Add(new SendResult { RecipientId = r.Id, RecipientName = r.Name });
        }
        else
        {
            foreach (var r in selectedRecipients)
            {
                if (!SendResults.Any(sr => sr.RecipientId == r.Id))
                    SendResults.Add(new SendResult { RecipientId = r.Id, RecipientName = r.Name });
            }
        }

        _cts = new CancellationTokenSource();

        var settings = _dataService.LoadSettings();
        var attachConfig = _dataService.LoadAttachmentConfig();
        var varVm = new VariablesViewModel(_dataService);

        IEmailSender sender = SendMethod == SendMethod.Outlook
            ? App.Services.GetRequiredService<OutlookEmailSender>()
            : App.Services.GetRequiredService<SmtpEmailSender>();

        string htmlBody;
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

        for (int i = 0; i < selectedRecipients.Count; i++)
        {
            if (_cts.IsCancellationRequested) break;

            var recipient = selectedRecipients[i];
            StatusText = $"正在发送: {recipient.Name} ({i + 1}/{selectedRecipients.Count})...";

            var allAttachments = new List<string>(attachConfig.CommonAttachments);
            var unitAtt = attachConfig.RecipientAttachments.FirstOrDefault(ua => ua.RecipientId == recipient.Id);
            if (unitAtt != null) allAttachments.AddRange(unitAtt.Files);

            var subject = varVm.ProcessBody(settings.LastSubject, recipient);
            var body = varVm.ProcessBody(htmlBody, recipient);

            var existing = SendResults.FirstOrDefault(r => r.RecipientId == recipient.Id);
            if (existing != null)
            {
                existing.Status = SendStatus.Sending;
            }

            var result = await sender.SendAsync(subject, body, recipient, allAttachments, smtpPassword);

            if (existing != null)
            {
                existing.Status = result.Status;
                existing.ErrorMessage = result.ErrorMessage;
                existing.SendTime = result.SendTime;
            }

            if (result.Status == SendStatus.Success) SuccessCount++;
            else FailedCount++;

            var done = SuccessCount + FailedCount;
            if (append)
                ProgressValue = done * 100.0 / TotalCount;
            else
                ProgressValue = (i + 1) * 100.0 / selectedRecipients.Count;
        }

        StatusText = _cts.IsCancellationRequested
            ? $"已取消: 成功 {SuccessCount} 封, 失败 {FailedCount} 封"
            : $"发送完成: 成功 {SuccessCount} 封, 失败 {FailedCount} 封";
        IsSending = false;
        _cts = null;

        if (sender is IDisposable disposable)
            disposable.Dispose();
    }

    public async Task RetryAllFailed()
    {
        var failedIds = SendResults.Where(r => r.Status == SendStatus.Failed).Select(r => r.RecipientId).ToList();
        if (failedIds.Count == 0) return;
        var groups = _dataService.LoadRecipientGroups();
        var failedRecipients = groups.SelectMany(g => g.Recipients).Where(r => failedIds.Contains(r.Id)).ToList();
        await ExecuteSend(failedRecipients, append: true);
    }

    public async Task RetryOne(string recipientId)
    {
        var recipient = FindRecipient(recipientId);
        if (recipient == null) return;
        await ExecuteSend(new List<Recipient> { recipient }, append: true);
    }

    public void CancelSend()
    {
        _cts?.Cancel();
    }
}
