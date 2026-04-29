using System.Collections.ObjectModel;
using System.Text;
using CommunityToolkit.Mvvm.ComponentModel;
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

    /// <summary>
    /// 获取或设置选中的收件人列表（仅包含 IsSelected 为 true 的收件人）。
    /// </summary>
    [ObservableProperty]
    private ObservableCollection<Recipient> _selectedRecipients = [];

    /// <summary>
    /// 获取或设置当前预览的收件人。
    /// </summary>
    [ObservableProperty]
    private Recipient? _selectedRecipient;

    /// <summary>
    /// 获取或设置变量替换后的邮件主题预览文本。
    /// </summary>
    [ObservableProperty]
    private string _previewSubject = string.Empty;

    /// <summary>
    /// 获取或设置收件人邮箱地址预览文本。
    /// </summary>
    [ObservableProperty]
    private string _previewTo = string.Empty;

    /// <summary>
    /// 获取或设置抄送邮箱地址预览文本。
    /// </summary>
    [ObservableProperty]
    private string _previewCc = string.Empty;

    /// <summary>
    /// 获取或设置密送邮箱地址预览文本。
    /// </summary>
    [ObservableProperty]
    private string _previewBcc = string.Empty;

    /// <summary>
    /// 获取或设置附件文件路径预览列表（公共附件 + 单位专属附件）。
    /// </summary>
    [ObservableProperty]
    private ObservableCollection<string> _previewAttachments = [];

    /// <summary>
    /// 当预览的收件人发生变化时触发，用于通知界面刷新 HTML 预览。
    /// </summary>
    public event Action? SelectedRecipientChanged;

    public PreviewViewModel(IDataService dataService)
    {
        _dataService = dataService;
    }

    /// <summary>
    /// 加载选中的收件人列表，默认选中第一个收件人并更新预览。
    /// </summary>
    public void LoadData()
    {
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

        var settings = _dataService.LoadSettings();
        var varVm = new VariablesViewModel(_dataService);
        PreviewSubject = varVm.ProcessBody(settings.LastSubject, SelectedRecipient);
        PreviewTo = SelectedRecipient.ToEmails;
        PreviewCc = SelectedRecipient.CcEmails;
        PreviewBcc = SelectedRecipient.BccEmails;

        // 合并公共附件和单位专属附件
        var attachments = new List<string>();
        var config = _dataService.LoadAttachmentConfig();
        attachments.AddRange(config.CommonAttachments);
        var unitAtt = config.UnitAttachments.FirstOrDefault(ua => ua.RecipientId == SelectedRecipient.Id);
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
    /// 获取当前收件人的完整 HTML 预览内容。
    /// 优先使用 XAML 转 HTML 方式生成格式化正文，失败时回退为纯文本转 HTML。
    /// </summary>
    /// <returns>完整的 HTML 文档字符串；若未选中收件人则返回空字符串。</returns>
    public string GetPreviewHtml()
    {
        if (SelectedRecipient == null) return string.Empty;

        var settings = _dataService.LoadSettings();
        var varVm = new VariablesViewModel(_dataService);
        var processedBody = varVm.ProcessBody(settings.LastBody, SelectedRecipient);

        var sb = new StringBuilder();
        sb.Append("<!DOCTYPE html><html><head>");
        sb.Append("<meta charset='utf-8'>");
        sb.Append("<style>body{font-family:'Microsoft YaHei UI',sans-serif;font-size:16px;line-height:1.8;padding:12px;margin:0;}</style>");
        sb.Append("</head><body>");

        // 优先使用 XAML 转 HTML 生成格式化正文
        if (!string.IsNullOrEmpty(settings.LastBodyXaml))
        {
            try
            {
                var doc = new System.Windows.Documents.FlowDocument();
                Helpers.FlowDocumentHelper.LoadFromXaml(doc, settings.LastBodyXaml);
                var html = Helpers.FlowDocumentHelper.ToHtml(doc);
                var processedHtml = varVm.ProcessBody(html, SelectedRecipient);
                sb.Append(processedHtml);
            }
            catch (Exception ex)
            {
                // XAML 转换失败时回退为纯文本
                AppLogger.Error("预览HTML生成失败", ex);
                sb.Append(processedBody.Replace("\n", "<br/>"));
            }
        }
        else
        {
            sb.Append(processedBody.Replace("\n", "<br/>"));
        }

        sb.Append("</body></html>");
        return sb.ToString();
    }
}
