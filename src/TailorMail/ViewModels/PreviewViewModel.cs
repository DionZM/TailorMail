using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using TailorMail.Helpers;
using TailorMail.Models;
using TailorMail.Services;

namespace TailorMail.ViewModels;

/// <summary>
/// 邮件预览视图模型，提供发送前的邮件效果预览功能。
/// 展示变量替换后的邮件主题、收件人邮箱、附件列表和 HTML 正文预览。
/// 正文预览优先使用 XAML 转 HTML，失败时回退为纯文本。
/// </summary>
public partial class PreviewViewModel : ObservableObject
{
    private readonly IDataService _dataService;

    private AppSettings? _cachedSettings;
    private VariablesViewModel? _cachedVarVm;
    private AttachmentConfig? _cachedAttachConfig;

    [ObservableProperty]
    private ObservableCollection<Recipient> _selectedRecipients = [];

    [ObservableProperty]
    private Recipient? _selectedRecipient;

    [ObservableProperty]
    private string _previewSubject = string.Empty;

    [ObservableProperty]
    private string _previewTo = string.Empty;

    [ObservableProperty]
    private string _previewCc = string.Empty;

    [ObservableProperty]
    private string _previewBcc = string.Empty;

    [ObservableProperty]
    private ObservableCollection<string> _previewAttachments = [];

    public event Action? SelectedRecipientChanged;

    public PreviewViewModel(IDataService dataService)
    {
        _dataService = dataService;
    }

    private AppSettings CachedSettings => _cachedSettings ??= _dataService.LoadSettings();
    private VariablesViewModel CachedVarVm => _cachedVarVm ??= new VariablesViewModel(_dataService);
    private AttachmentConfig CachedAttachConfig => _cachedAttachConfig ??= _dataService.LoadAttachmentConfig();

    public void InvalidateCache()
    {
        _cachedSettings = null;
        _cachedVarVm = null;
        _cachedAttachConfig = null;
    }

    public AttachmentConfig GetCachedAttachmentConfig() => CachedAttachConfig;

    public void LoadData()
    {
        InvalidateCache();
        var groups = _dataService.LoadRecipientGroups();
        SelectedRecipients = new ObservableCollection<Recipient>(
            groups.SelectMany(g => g.Recipients).Where(r => r.IsSelected));
        if (SelectedRecipients.Count > 0 && SelectedRecipient == null)
            SelectedRecipient = SelectedRecipients[0];
        UpdatePreview();
    }

    /// <summary>
    /// 当预览收件人变化时更新预览内容并触发事件。
    /// </summary>
    partial void OnSelectedRecipientChanged(Recipient? value)
    {
        UpdatePreview();
        SelectedRecipientChanged?.Invoke();
    }

    /// <summary>
    /// 更新当前收件人的预览信息，包括主题、邮箱和附件。
    /// </summary>
    private void UpdatePreview()
    {
        if (SelectedRecipient == null)
        {
            PreviewTo = PreviewCc = PreviewBcc = string.Empty;
            PreviewAttachments = [];
            return;
        }

        var settings = CachedSettings;
        var varVm = CachedVarVm;
        PreviewSubject = VariablesViewModel.ProcessBodyFast(settings.LastSubject, SelectedRecipient);
        PreviewTo = SelectedRecipient.ToEmails;
        PreviewCc = SelectedRecipient.CcEmails;
        PreviewBcc = SelectedRecipient.BccEmails;

        var attachments = new List<string>();
        var config = CachedAttachConfig;
        attachments.AddRange(config.CommonAttachments);
        var unitAtt = config.RecipientAttachments.FirstOrDefault(ua => ua.RecipientId == SelectedRecipient.Id);
        if (unitAtt != null) attachments.AddRange(unitAtt.Files);
        PreviewAttachments = new ObservableCollection<string>(attachments);
    }

    /// <summary>
    /// 选择指定收件人进行预览。
    /// </summary>
    /// <param name="r">要预览的收件人。</param>
    public void SelectRecipient(Recipient r)
    {
        SelectedRecipient = r;
    }

    /// <summary>
    /// 获取当前收件人的预览 FlowDocument。
    /// 优先使用保存的 XAML 直接加载（保留格式），失败时回退为纯文本。
    /// 变量替换在文本元素上逐段执行。
    /// </summary>
    public System.Windows.Documents.FlowDocument? GetPreviewDocument()
    {
        if (SelectedRecipient == null) return null;

        var settings = CachedSettings;
        var varVm = CachedVarVm;

        System.Windows.Documents.FlowDocument doc;

        if (!string.IsNullOrEmpty(settings.LastBodyXaml))
        {
            try
            {
                doc = new System.Windows.Documents.FlowDocument();
                Helpers.FlowDocumentHelper.LoadFromXaml(doc, settings.LastBodyXaml);
            }
            catch (Exception ex)
            {
                AppLogger.Error("预览文档加载失败", ex);
                doc = CreatePlainTextDoc(VariablesViewModel.ProcessBodyFast(settings.LastBody, SelectedRecipient));
            }
        }
        else
        {
            doc = CreatePlainTextDoc(VariablesViewModel.ProcessBodyFast(settings.LastBody, SelectedRecipient));
        }

        foreach (var run in GetAllRuns(doc).ToList())
        {
            if (!string.IsNullOrEmpty(run.Text))
                run.Text = VariablesViewModel.ProcessBodyFast(run.Text, SelectedRecipient);
        }

        return doc;
    }

    private static IEnumerable<System.Windows.Documents.Run> GetAllRuns(System.Windows.Documents.FlowDocument doc)
    {
        foreach (var block in doc.Blocks)
        {
            if (block is System.Windows.Documents.Paragraph para)
            {
                foreach (var run in CollectRuns(para.Inlines))
                    yield return run;
            }
            else if (block is System.Windows.Documents.Table table)
            {
                foreach (var rowGroup in table.RowGroups)
                    foreach (var row in rowGroup.Rows)
                        foreach (var cell in row.Cells)
                            foreach (var cellBlock in cell.Blocks)
                                if (cellBlock is System.Windows.Documents.Paragraph cellPara)
                                    foreach (var run in CollectRuns(cellPara.Inlines))
                                        yield return run;
            }
        }
    }

    private static IEnumerable<System.Windows.Documents.Run> CollectRuns(System.Windows.Documents.InlineCollection inlines)
    {
        foreach (var inline in inlines)
        {
            if (inline is System.Windows.Documents.Run run)
                yield return run;
            else if (inline is System.Windows.Documents.Span span)
            {
                foreach (var child in CollectRuns(span.Inlines))
                    yield return child;
            }
        }
    }

    private static System.Windows.Documents.FlowDocument CreatePlainTextDoc(string text)
    {
        var doc = new System.Windows.Documents.FlowDocument();
        doc.FontFamily = new System.Windows.Media.FontFamily("Segoe UI, Microsoft YaHei UI");
        doc.FontSize = 15;
        if (string.IsNullOrEmpty(text)) return doc;
        foreach (var line in text.Split('\n'))
        {
            var para = new System.Windows.Documents.Paragraph(new System.Windows.Documents.Run(line.TrimEnd('\r')))
            {
                Margin = new System.Windows.Thickness(0, 0, 0, 8)
            };
            doc.Blocks.Add(para);
        }
        return doc;
    }

    public async Task<(bool success, string error)> SendTestEmailAsync(AppSettings settings, string senderEmail)
    {
        if (SelectedRecipient == null)
            return (false, "未选择收件人");

        try
        {
            var subject = VariablesViewModel.ProcessBodyFast(settings.LastSubject, SelectedRecipient);

            string bodyHtml;
            if (!string.IsNullOrEmpty(settings.LastBodyXaml))
            {
                try
                {
                    var doc = new System.Windows.Documents.FlowDocument();
                    Helpers.FlowDocumentHelper.LoadFromXaml(doc, settings.LastBodyXaml);
                    foreach (var run in GetAllRuns(doc).ToList())
                    {
                        if (!string.IsNullOrEmpty(run.Text))
                            run.Text = VariablesViewModel.ProcessBodyFast(run.Text, SelectedRecipient);
                    }
                    bodyHtml = Helpers.FlowDocumentHelper.ToHtml(doc);
                    bodyHtml = Helpers.FlowDocumentHelper.WrapAsEmailDocument(bodyHtml);
                }
                catch
                {
                    bodyHtml = $"<pre>{System.Net.WebUtility.HtmlEncode(VariablesViewModel.ProcessBodyFast(settings.LastBody, SelectedRecipient))}</pre>";
                }
            }
            else
            {
                bodyHtml = $"<pre>{System.Net.WebUtility.HtmlEncode(VariablesViewModel.ProcessBodyFast(settings.LastBody, SelectedRecipient))}</pre>";
            }

            if (!string.IsNullOrEmpty(settings.Signature))
                bodyHtml += $"<br/><br/><span style='color:#666;'>{System.Net.WebUtility.HtmlEncode(VariablesViewModel.ProcessBodyFast(settings.Signature, SelectedRecipient))}</span>";

            var attachments = PreviewAttachments?.ToArray() ?? [];

            if (settings.SendMethod == Models.SendMethod.Outlook)
            {
                using var outlookSender = new OutlookEmailSender();
                var result = outlookSender.SendTest(senderEmail, senderEmail, subject, bodyHtml, attachments, out string error);
                return (result, error);
            }
            else
            {
                var smtpSettings = settings.Smtp;
                var password = CredentialHelper.Unprotect(smtpSettings.EncryptedPassword);
                if (string.IsNullOrEmpty(password))
                    return (false, "未配置 SMTP 密码，请在设置中保存密码");

                var displayName = !string.IsNullOrWhiteSpace(smtpSettings.DisplayName) ? smtpSettings.DisplayName : smtpSettings.UserName;
                var from = !string.IsNullOrWhiteSpace(smtpSettings.SenderEmail) ? smtpSettings.SenderEmail : smtpSettings.UserName;

                // H-05: Dispose SmtpEmailSender after use
                using var smtpSender = new SmtpEmailSender();
                return await smtpSender.SendTestAsync(from, displayName, password, senderEmail, subject, bodyHtml, attachments, settings.Smtp);
            }
        }
        catch (Exception ex)
        {
            return (false, ex.Message);
        }
    }
}
