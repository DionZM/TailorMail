using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using TailorMail.Models;
using TailorMail.Services;
using Microsoft.Extensions.DependencyInjection;

namespace TailorMail.ViewModels;

/// <summary>
/// 邮件发送视图模型，负责批量发送邮件、跟踪发送进度、支持失败重试和取消操作。
/// 根据 <see cref="SendMethod"/> 设置选择 Outlook 或 SMTP 发送通道。
/// 发送过程中逐个处理收件人，每封邮件独立进行变量替换和附件合并。
/// </summary>
public partial class SendViewModel : ObservableObject
{
    private readonly IDataService _dataService;

    /// <summary>
    /// 取消令牌源，用于支持发送过程中取消操作。
    /// </summary>
    private CancellationTokenSource? _cts;

    /// <summary>
    /// 获取或设置发送结果列表，记录每个收件人的发送状态。
    /// </summary>
    [ObservableProperty]
    private ObservableCollection<SendResult> _sendResults = [];

    /// <summary>
    /// 获取或设置是否正在发送中。
    /// </summary>
    [ObservableProperty]
    private bool _isSending;

    /// <summary>
    /// 获取或设置待发送邮件总数。
    /// </summary>
    [ObservableProperty]
    private int _totalCount;

    /// <summary>
    /// 获取或设置已成功发送的邮件数。
    /// </summary>
    [ObservableProperty]
    private int _successCount;

    /// <summary>
    /// 获取或设置发送失败的邮件数。
    /// </summary>
    [ObservableProperty]
    private int _failedCount;

    /// <summary>
    /// 获取或设置发送进度百分比（0-100）。
    /// </summary>
    [ObservableProperty]
    private double _progressValue;

    /// <summary>
    /// 获取或设置当前发送状态文本。
    /// </summary>
    [ObservableProperty]
    private string _statusText = "就绪";

    /// <summary>
    /// 获取或设置当前邮件发送方式。
    /// </summary>
    [ObservableProperty]
    private SendMethod _sendMethod;

    public SendViewModel(IDataService dataService)
    {
        _dataService = dataService;
        var settings = _dataService.LoadSettings();
        SendMethod = settings.SendMethod;
    }

    /// <summary>
    /// 重新从数据服务加载发送方式设置。
    /// </summary>
    public void ReloadSettings()
    {
        var settings = _dataService.LoadSettings();
        SendMethod = settings.SendMethod;
    }

    /// <summary>
    /// 获取所有选中的收件人列表。
    /// </summary>
    /// <returns>选中收件人列表。</returns>
    public List<Recipient> GetSelectedRecipients()
    {
        var groups = _dataService.LoadRecipientGroups();
        return groups.SelectMany(g => g.Recipients).Where(r => r.IsSelected).ToList();
    }

    /// <summary>
    /// 执行批量邮件发送。
    /// 逐个处理收件人：合并附件、替换变量、调用邮件发送器发送。
    /// 支持通过 <see cref="CancellationTokenSource"/> 取消发送。
    /// </summary>
    /// <param name="selectedRecipients">待发送的收件人列表。</param>
    /// <param name="smtpPassword">SMTP 密码（仅 SMTP 方式使用），由用户临时输入。</param>
    public async Task ExecuteSend(List<Recipient> selectedRecipients, string? smtpPassword = null)
    {
        IsSending = true;
        TotalCount = selectedRecipients.Count;
        SuccessCount = 0;
        FailedCount = 0;
        ProgressValue = 0;
        _cts = new CancellationTokenSource();

        // 初始化发送结果列表
        SendResults = new ObservableCollection<SendResult>(
            selectedRecipients.Select(r => new SendResult { RecipientId = r.Id, RecipientName = r.Name }));

        var settings = _dataService.LoadSettings();
        var attachConfig = _dataService.LoadAttachmentConfig();
        var varVm = new VariablesViewModel(_dataService);

        // 根据发送方式选择对应的发送器
        IEmailSender sender = SendMethod == SendMethod.Outlook
            ? App.Services.GetRequiredService<OutlookEmailSender>()
            : App.Services.GetRequiredService<SmtpEmailSender>();

        // 将邮件正文转换为 HTML 格式
        string htmlBody;
        if (!string.IsNullOrEmpty(settings.LastBodyXaml))
        {
            try
            {
                var doc = new System.Windows.Documents.FlowDocument();
                Helpers.FlowDocumentHelper.LoadFromXaml(doc, settings.LastBodyXaml);
                htmlBody = Helpers.FlowDocumentHelper.ToHtml(doc);
            }
            catch
            {
                // XAML 转换失败时回退为纯文本
                htmlBody = settings.LastBody.Replace("\n", "<br/>");
            }
        }
        else
        {
            htmlBody = settings.LastBody.Replace("\n", "<br/>");
        }

        // 逐个发送邮件
        for (int i = 0; i < selectedRecipients.Count; i++)
        {
            // 检查是否已取消
            if (_cts.IsCancellationRequested) break;

            var recipient = selectedRecipients[i];
            StatusText = $"正在发送: {recipient.Name} ({i + 1}/{selectedRecipients.Count})...";

            // 合并公共附件和单位专属附件
            var allAttachments = new List<string>(attachConfig.CommonAttachments);
            var unitAtt = attachConfig.UnitAttachments.FirstOrDefault(ua => ua.RecipientId == recipient.Id);
            if (unitAtt != null) allAttachments.AddRange(unitAtt.Files);

            // 对每个收件人独立进行变量替换
            var subject = varVm.ProcessBody(settings.LastSubject, recipient);
            var body = varVm.ProcessBody(htmlBody, recipient);

            // 发送邮件
            var result = await sender.SendAsync(subject, body, recipient, allAttachments, smtpPassword);

            // 更新发送结果
            var existing = SendResults.FirstOrDefault(r => r.RecipientId == recipient.Id);
            if (existing != null)
            {
                existing.Status = result.Status;
                existing.ErrorMessage = result.ErrorMessage;
                existing.SendTime = result.SendTime;
            }

            if (result.Status == SendStatus.Success) SuccessCount++;
            else FailedCount++;

            ProgressValue = (i + 1) * 100.0 / selectedRecipients.Count;
        }

        StatusText = _cts.IsCancellationRequested
            ? $"已取消: 成功 {SuccessCount} 封, 失败 {FailedCount} 封"
            : $"发送完成: 成功 {SuccessCount} 封, 失败 {FailedCount} 封";
        IsSending = false;
        _cts = null;
    }

    /// <summary>
    /// 重试所有发送失败的收件人。
    /// </summary>
    public async Task RetryFailed()
    {
        var failedIds = SendResults.Where(r => r.Status == SendStatus.Failed).Select(r => r.RecipientId).ToList();
        if (failedIds.Count == 0) return;
        var groups = _dataService.LoadRecipientGroups();
        var failedRecipients = groups.SelectMany(g => g.Recipients).Where(r => failedIds.Contains(r.Id)).ToList();
        await ExecuteSend(failedRecipients);
    }

    /// <summary>
    /// 取消正在进行的发送操作。
    /// </summary>
    public void CancelSend()
    {
        _cts?.Cancel();
    }
}
