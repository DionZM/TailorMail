using System.Collections.ObjectModel;
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
    /// 获取或设置附件文件路径预览列表（公共附件 + 收件人专属附件）。
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

        // 合并公共附件和收件人专属附件
        var attachments = new List<string>();
        var config = _dataService.LoadAttachmentConfig();
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

        var settings = _dataService.LoadSettings();
        var varVm = new VariablesViewModel(_dataService);

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
                doc = CreatePlainTextDoc(varVm.ProcessBody(settings.LastBody, SelectedRecipient));
            }
        }
        else
        {
            doc = CreatePlainTextDoc(varVm.ProcessBody(settings.LastBody, SelectedRecipient));
        }

        // 替换正文中的变量
        foreach (var run in GetAllRuns(doc).ToList())
        {
            if (!string.IsNullOrEmpty(run.Text))
                run.Text = varVm.ProcessBody(run.Text, SelectedRecipient);
        }

        return doc;
    }

    private static IEnumerable<System.Windows.Documents.Run> GetAllRuns(System.Windows.Documents.FlowDocument doc)
    {
        foreach (var block in doc.Blocks)
        {
            if (block is System.Windows.Documents.Paragraph para)
            {
                foreach (var inline in para.Inlines)
                {
                    if (inline is System.Windows.Documents.Run run)
                        yield return run;
                    else if (inline is System.Windows.Documents.Span span)
                    {
                        foreach (var child in span.Inlines)
                        {
                            if (child is System.Windows.Documents.Run childRun)
                                yield return childRun;
                        }
                    }
                }
            }
            else if (block is System.Windows.Documents.Table table)
            {
                foreach (var rowGroup in table.RowGroups)
                    foreach (var row in rowGroup.Rows)
                        foreach (var cell in row.Cells)
                            foreach (var cellBlock in cell.Blocks)
                                if (cellBlock is System.Windows.Documents.Paragraph cellPara)
                                    foreach (var inline in cellPara.Inlines)
                                        if (inline is System.Windows.Documents.Run run)
                                            yield return run;
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
}
