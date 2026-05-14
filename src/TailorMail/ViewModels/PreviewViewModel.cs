using System.Collections.ObjectModel;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Documents;
using System.Windows.Media;
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
        var varNames = varVm.VariableNames;
        PreviewSubject = VariablesViewModel.ProcessBodyFast(settings.LastSubject, SelectedRecipient, varNames);
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
        var varNames = varVm.VariableNames;

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
                doc = CreatePlainTextDoc(VariablesViewModel.ProcessBodyFast(settings.LastBody, SelectedRecipient, varNames));
            }
        }
        else
        {
            doc = CreatePlainTextDoc(VariablesViewModel.ProcessBodyFast(settings.LastBody, SelectedRecipient, varNames));
        }

        foreach (var run in GetAllRuns(doc).ToList())
        {
            if (!string.IsNullOrEmpty(run.Text))
                HighlightVariablesInRun(run, SelectedRecipient, varNames);
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

    private static void HighlightVariablesInRun(Run run, Recipient recipient, IReadOnlyList<string>? varNames)
    {
        var text = run.Text;
        if (string.IsNullOrEmpty(text) || !text.Contains('{'))
        {
            run.Text = VariablesViewModel.ProcessBodyFast(text, recipient, varNames);
            return;
        }

        var parent = run.Parent as Paragraph;
        if (parent == null)
        {
            run.Text = VariablesViewModel.ProcessBodyFast(text, recipient, varNames);
            return;
        }

        var replacements = BuildReplacementMap(recipient, varNames);
        var pattern = @"\{([^}]+)\}";
        var matches = Regex.Matches(text, pattern);
        if (matches.Count == 0)
        {
            run.Text = VariablesViewModel.ProcessBodyFast(text, recipient, varNames);
            return;
        }

        var nextSibling = run.NextInline;
        int lastPos = 0;

        for (int i = 0; i < matches.Count; i++)
        {
            var match = matches[i];
            if (match.Index > lastPos)
            {
                var plainText = text[lastPos..match.Index];
                var newRun = new Run(VariablesViewModel.ProcessBodyFast(plainText, recipient, varNames))
                {
                    FontFamily = run.FontFamily,
                    FontSize = run.FontSize,
                    FontWeight = run.FontWeight,
                    FontStyle = run.FontStyle,
                    Foreground = run.Foreground
                };
                if (nextSibling != null)
                    parent.Inlines.InsertBefore(nextSibling, newRun);
                else
                    parent.Inlines.Add(newRun);
            }

            var varName = match.Groups[1].Value;
            var value = replacements.TryGetValue(varName, out var v) ? v : match.Value;
            var highlightRun = new Run(value)
            {
                Background = (Brush)Application.Current.FindResource("AccentLightBrush"),
                Foreground = (Brush)Application.Current.FindResource("AccentBrush"),
                FontFamily = run.FontFamily,
                FontSize = run.FontSize,
                FontWeight = run.FontWeight,
                FontStyle = run.FontStyle
            };
            if (nextSibling != null)
                parent.Inlines.InsertBefore(nextSibling, highlightRun);
            else
                parent.Inlines.Add(highlightRun);

            lastPos = match.Index + match.Length;
        }

        if (lastPos < text.Length)
        {
            var trailing = text[lastPos..];
            var trailingRun = new Run(VariablesViewModel.ProcessBodyFast(trailing, recipient, varNames))
            {
                FontFamily = run.FontFamily,
                FontSize = run.FontSize,
                FontWeight = run.FontWeight,
                FontStyle = run.FontStyle,
                Foreground = run.Foreground
            };
            if (nextSibling != null)
                parent.Inlines.InsertBefore(nextSibling, trailingRun);
            else
                parent.Inlines.Add(trailingRun);
        }

        parent.Inlines.Remove(run);
    }

    private static Dictionary<string, string> BuildReplacementMap(Recipient recipient, IReadOnlyList<string>? varNames)
    {
        var replacements = new Dictionary<string, string>(StringComparer.Ordinal)
        {
            ["名称"] = recipient.Name ?? "",
            ["简称"] = recipient.ShortName ?? ""
        };
        foreach (var kvp in recipient.Variables)
            replacements[kvp.Key] = kvp.Value ?? "";

        if (varNames != null)
        {
            for (int i = 0; i < varNames.Count; i++)
            {
                var name = varNames[i];
                if (!replacements.ContainsKey(name))
                    replacements[name] = "";
            }
        }
        return replacements;
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
