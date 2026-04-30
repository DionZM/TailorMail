using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Documents;
using System.Windows.Media;
using System.Text.RegularExpressions;
using TailorMail.Services;

namespace TailorMail.Helpers;

/// <summary>
/// FlowDocument 与 HTML/XAML 之间的转换工具类。
/// 提供 FlowDocument 转 HTML（用于邮件发送）、XAML 序列化/反序列化（用于富文本编辑器状态保存）等功能。
/// HTML 输出采用极简主义设计规范：温暖的纯白色调背景、精致的排版层级、宽裕的留白间距。
/// </summary>
public static class FlowDocumentHelper
{
    /// <summary>
    /// 邮件 HTML 容器的基础内联 CSS，基于极简主义设计规范。
    /// </summary>
    private const string EmailBaseStyle = @"
        background:#FFFFFF;
        font-family:'Helvetica Neue','PingFang SC','Microsoft YaHei UI','SF Pro Display',sans-serif;
        font-size:15px;
        line-height:1.75;
        color:#2F3437;
        max-width:640px;
        margin:0 auto;
        padding:40px 32px;";

    /// <summary>
    /// 段落样式：底部间距 16px。
    /// </summary>
    private const string ParagraphStyle = "margin:0 0 16px 0;";

    /// <summary>
    /// 标题样式常量：用于大号字体段落的编辑风格式排版。
    /// 根据 FlowDocument 中段落的字号大小，映射为 h1 ~ h4 四级标题。
    /// 采用衬线体 + 紧凑字距 + 深色强调，营造编辑风格的层次感。
    /// </summary>
    private static readonly (double MinSize, string Tag, string Style)[] HeadingLevels =
    [
        (28.0, "h1", "font-family:'Helvetica Neue','PingFang SC','Microsoft YaHei UI',sans-serif;font-size:26px;font-weight:700;color:#111111;margin:32px 0 12px 0;line-height:1.25;letter-spacing:-0.02em;"),
        (22.0, "h2", "font-family:'Helvetica Neue','PingFang SC','Microsoft YaHei UI',sans-serif;font-size:20px;font-weight:600;color:#1A1A1A;margin:28px 0 10px 0;line-height:1.3;letter-spacing:-0.01em;"),
        (19.0, "h3", "font-family:'Helvetica Neue','PingFang SC','Microsoft YaHei UI',sans-serif;font-size:17px;font-weight:600;color:#2F3437;margin:24px 0 8px 0;line-height:1.4;"),
        (17.0, "h4", "font-family:'Helvetica Neue','PingFang SC','Microsoft YaHei UI',sans-serif;font-size:15px;font-weight:600;color:#2F3437;margin:20px 0 6px 0;line-height:1.5;"),
    ];

    /// <summary>
    /// 无序列表样式：左侧缩进 20px，列表项之间间距 4px。
    /// </summary>
    private const string UnorderedListStyle = "margin:0 0 16px 0;padding-left:20px;";

    /// <summary>
    /// 有序列表样式：左侧缩进 20px，列表项之间间距 4px。
    /// </summary>
    private const string OrderedListStyle = "margin:0 0 16px 0;padding-left:20px;";

    /// <summary>
    /// 列表项样式。
    /// </summary>
    private const string ListItemStyle = "margin:0 0 4px 0;line-height:1.75;";

    /// <summary>
    /// 超链接样式：柔蓝色 #1F6C9F，无下划线（邮件客户端兼容）。
    /// </summary>
    private const string LinkStyle = "color:#1F6C9F;text-decoration:none;";

    /// <summary>
    /// 表格样式：全宽、边框合并、极浅灰边框。
    /// </summary>
    private const string TableStyle = "width:100%;border-collapse:collapse;margin:0 0 16px 0;";

    /// <summary>
    /// 表格单元格样式：极浅灰边框、内边距 10px 12px、顶部对齐。
    /// </summary>
    private const string TableCellStyle = "border:1px solid #EAEAEA;padding:10px 12px;text-align:left;vertical-align:top;";

    /// <summary>
    /// 表头单元格样式：在 <see cref="TableCellStyle"/> 基础上增加深色背景和加粗文字。
    /// </summary>
    private const string TableHeaderCellStyle = "border:1px solid #EAEAEA;padding:10px 12px;text-align:left;vertical-align:top;background:#F7F6F3;font-weight:600;color:#2F3437;";

    /// <summary>
    /// 引用块样式：左侧 3px 深色竖线、浅灰背景、内边距。
    /// </summary>
    private const string BlockquoteStyle = "margin:0 0 16px 0;padding:12px 16px;border-left:3px solid #D1D1D1;background:#F9F9F8;color:#555555;";

    /// <summary>
    /// 签名分隔线样式：顶部极浅灰边框，上方留出 32px 间距。
    /// </summary>
    private const string SignatureSeparatorStyle = "margin-top:32px;padding-top:16px;border-top:1px solid #EAEAEA;";

    /// <summary>
    /// 默认段落字号，用于判断段落是否为标题级别。
    /// 与 MailComposePage.xaml 中 RichTextBox 的 FontSize="16" 对应。
    /// </summary>
    private const double DefaultFontSize = 16.0;

    /// <summary>
    /// 将 <see cref="FlowDocument"/> 转换为美观的邮件 HTML 字符串。
    /// 采用极简主义排版设计：宽裕的留白、精致的字体层级、温暖的纯白色调。
    /// 支持 Paragraph（含标题检测）、List、Table、Section、BlockUIContainer 等元素。
    /// 支持内联元素：Bold、Italic、Underline、Hyperlink、Span、Run、LineBreak。
    /// </summary>
    public static string ToHtml(FlowDocument doc)
    {
        var sb = new StringBuilder();
        sb.Append($"<div style=\"{EmailBaseStyle}\">");

        WriteBlocks(sb, doc.Blocks, isTopLevel: true);

        sb.Append($"<div style=\"{SignatureSeparatorStyle}\"></div>");
        sb.Append("</div>");
        return sb.ToString();
    }

    /// <summary>
    /// 递归遍历块级元素集合，将 Block 转换为对应的 HTML 标签。
    /// 支持段落、列表、表格、分区等块级元素。
    /// </summary>
    private static void WriteBlocks(StringBuilder sb, BlockCollection blocks, bool isTopLevel = false)
    {
        var isFirst = true;
        foreach (Block block in blocks)
        {
            switch (block)
            {
                case Paragraph para:
                    WriteParagraph(sb, para, isTopLevel && isFirst);
                    break;
                case List list:
                    WriteList(sb, list);
                    break;
                case Table table:
                    WriteTable(sb, table);
                    break;
                case Section section:
                    WriteBlocks(sb, section.Blocks, isTopLevel && isFirst);
                    break;
            }
            isFirst = false;
        }
    }

    /// <summary>
    /// 将段落转换为 HTML。根据字号大小自动检测是否为标题：
    /// 字号 >= 17px 的段落映射为 h1 ~ h4，否则为普通 <p>。
    /// 首段去除顶部间距。
    /// </summary>
    private static void WriteParagraph(StringBuilder sb, Paragraph para, bool isFirst)
    {
        var fontSize = para.FontSize;
        if (double.IsNaN(fontSize)) fontSize = DefaultFontSize;

        if (fontSize >= HeadingLevels[^1].MinSize)
        {
            foreach (var (minSize, tag, style) in HeadingLevels)
            {
                if (fontSize >= minSize)
                {
                    sb.Append($"<{tag} style=\"{style}\">");
                    WriteInlines(sb, para.Inlines);
                    sb.Append($"</{tag}>");
                    return;
                }
            }
        }

        var extraStyle = isFirst ? "margin-top:0;" : "";
        sb.Append($"<p style=\"{ParagraphStyle}{extraStyle}\">");
        WriteInlines(sb, para.Inlines);
        sb.Append("</p>");
    }

    /// <summary>
    /// 将 FlowDocument List 转换为 HTML 的 &lt;ul&gt; 或 &lt;ol&gt; 列表。
    /// 支持 MarkerStyle 的子弹/数字/大小写字母/罗马数字等样式映射。
    /// 嵌套列表通过递归处理 ListItem 的子块实现。
    /// </summary>
    private static void WriteList(StringBuilder sb, List list)
    {
        var tag = list.MarkerStyle switch
        {
            TextMarkerStyle.Decimal => "ol",
            TextMarkerStyle.LowerLatin => "ol",
            TextMarkerStyle.UpperLatin => "ol",
            TextMarkerStyle.LowerRoman => "ol",
            TextMarkerStyle.UpperRoman => "ol",
            _ => "ul"
        };

        var listStyle = tag == "ol" ? OrderedListStyle : UnorderedListStyle;

        var listTypeAttr = list.MarkerStyle switch
        {
            TextMarkerStyle.LowerLatin => " type=\"a\"",
            TextMarkerStyle.UpperLatin => " type=\"A\"",
            TextMarkerStyle.LowerRoman => " type=\"i\"",
            TextMarkerStyle.UpperRoman => " type=\"I\"",
            _ => ""
        };

        sb.Append($"<{tag} style=\"{listStyle}\"{listTypeAttr}>");

        foreach (var listItem in list.ListItems)
        {
            sb.Append($"<li style=\"{ListItemStyle}\">");
            foreach (var childBlock in listItem.Blocks)
            {
                if (childBlock is Paragraph childPara)
                {
                    WriteInlines(sb, childPara.Inlines);
                }
                else if (childBlock is List nestedList)
                {
                    WriteList(sb, nestedList);
                }
            }
            sb.Append("</li>");
        }

        sb.Append($"</{tag}>");
    }

    /// <summary>
    /// 将 FlowDocument Table 转换为 HTML 表格。
    /// 第一行作为表头（&lt;thead&gt;），其余行作为表体（&lt;tbody&gt;）。
    /// </summary>
    private static void WriteTable(StringBuilder sb, Table table)
    {
        sb.Append($"<table style=\"{TableStyle}\">");

        var isFirstRowGroup = true;
        foreach (var rowGroup in table.RowGroups)
        {
            if (isFirstRowGroup && rowGroup.Rows.Count > 0)
            {
                sb.Append("<thead>");
                foreach (var row in rowGroup.Rows)
                {
                    sb.Append("<tr>");
                    foreach (var cell in row.Cells)
                    {
                        sb.Append($"<th style=\"{TableHeaderCellStyle}\">");
                        foreach (var cellBlock in cell.Blocks)
                        {
                            if (cellBlock is Paragraph cellPara)
                                WriteInlines(sb, cellPara.Inlines);
                        }
                        sb.Append("</th>");
                    }
                    sb.Append("</tr>");
                }
                sb.Append("</thead>");
                isFirstRowGroup = false;
            }
            else
            {
                sb.Append("<tbody>");
                foreach (var row in rowGroup.Rows)
                {
                    sb.Append("<tr>");
                    foreach (var cell in row.Cells)
                    {
                        sb.Append($"<td style=\"{TableCellStyle}\">");
                        foreach (var cellBlock in cell.Blocks)
                        {
                            if (cellBlock is Paragraph cellPara)
                                WriteInlines(sb, cellPara.Inlines);
                        }
                        sb.Append("</td>");
                    }
                    sb.Append("</tr>");
                }
                sb.Append("</tbody>");
                isFirstRowGroup = false;
            }
        }

        sb.Append("</table>");
    }

    /// <summary>
    /// 将纯文本转换为美观的邮件 HTML，保留段落结构并智能识别常见的文本模式：
    /// - 空行分隔段落
    /// - # 开头的行转换为标题
    /// - - / * / • 开头的行转换为无序列表
    /// - 数字. 开头的行转换为有序列表
    /// - > 开头的行转换为引用块
    /// </summary>
    public static string PlainTextToHtml(string plainText)
    {
        if (string.IsNullOrEmpty(plainText)) return string.Empty;

        var sb = new StringBuilder();
        sb.Append($"<div style=\"{EmailBaseStyle}\">");

        var lines = plainText.Split('\n');
        var isFirstBlock = true;
        var i = 0;

        while (i < lines.Length)
        {
            var line = lines[i].TrimEnd('\r');
            var trimmed = line.Trim();

            if (string.IsNullOrEmpty(trimmed))
            {
                i++;
                continue;
            }

            if (trimmed.StartsWith("### "))
            {
                EmitBlockOpen(sb, ref isFirstBlock);
                sb.Append($"<h3 style=\"{HeadingLevels[2].Style}\">");
                sb.Append(HtmlEncode(trimmed[4..]));
                sb.Append("</h3>");
                i++;
            }
            else if (trimmed.StartsWith("## "))
            {
                EmitBlockOpen(sb, ref isFirstBlock);
                sb.Append($"<h2 style=\"{HeadingLevels[1].Style}\">");
                sb.Append(HtmlEncode(trimmed[3..]));
                sb.Append("</h2>");
                i++;
            }
            else if (trimmed.StartsWith("# "))
            {
                EmitBlockOpen(sb, ref isFirstBlock);
                sb.Append($"<h1 style=\"{HeadingLevels[0].Style}\">");
                sb.Append(HtmlEncode(trimmed[2..]));
                sb.Append("</h1>");
                i++;
            }
            else if (trimmed.StartsWith("> "))
            {
                EmitBlockOpen(sb, ref isFirstBlock);
                var quoteBuilder = new StringBuilder();
                while (i < lines.Length)
                {
                    var qLine = lines[i].TrimEnd('\r').Trim();
                    if (!qLine.StartsWith("> ")) break;
                    if (quoteBuilder.Length > 0) quoteBuilder.Append("<br/>");
                    quoteBuilder.Append(HtmlEncode(qLine[2..]));
                    i++;
                }
                sb.Append($"<blockquote style=\"{BlockquoteStyle}\">");
                sb.Append($"<p style=\"{ParagraphStyle}\">{quoteBuilder}</p>");
                sb.Append("</blockquote>");
            }
            else if (IsUnorderedListItem(trimmed))
            {
                EmitBlockOpen(sb, ref isFirstBlock);
                sb.Append($"<ul style=\"{UnorderedListStyle}\">");
                while (i < lines.Length)
                {
                    var liLine = lines[i].TrimEnd('\r').Trim();
                    if (!IsUnorderedListItem(liLine)) break;
                    var content = ExtractListItemContent(liLine);
                    sb.Append($"<li style=\"{ListItemStyle}\">{HtmlEncode(content)}</li>");
                    i++;
                }
                sb.Append("</ul>");
            }
            else if (IsOrderedListItem(trimmed))
            {
                EmitBlockOpen(sb, ref isFirstBlock);
                sb.Append($"<ol style=\"{OrderedListStyle}\">");
                while (i < lines.Length)
                {
                    var liLine = lines[i].TrimEnd('\r').Trim();
                    if (!IsOrderedListItem(liLine)) break;
                    var content = ExtractOrderedListItemContent(liLine);
                    sb.Append($"<li style=\"{ListItemStyle}\">{HtmlEncode(content)}</li>");
                    i++;
                }
                sb.Append("</ol>");
            }
            else
            {
                EmitBlockOpen(sb, ref isFirstBlock);
                var extraStyle = isFirstBlock ? "margin-top:0;" : "";
                isFirstBlock = false;
                var paragraphText = new StringBuilder(trimmed);
                i++;
                while (i < lines.Length)
                {
                    var nextLine = lines[i].TrimEnd('\r');
                    var nextTrimmed = nextLine.Trim();
                    if (string.IsNullOrEmpty(nextTrimmed)) break;
                    if (nextTrimmed.StartsWith("# ") || nextTrimmed.StartsWith("## ") || nextTrimmed.StartsWith("### ")) break;
                    if (IsUnorderedListItem(nextTrimmed) || IsOrderedListItem(nextTrimmed)) break;
                    if (nextTrimmed.StartsWith("> ")) break;
                    paragraphText.Append(' ');
                    paragraphText.Append(nextTrimmed);
                    i++;
                }
                var pStyle = isFirstBlock ? $"{ParagraphStyle}{extraStyle}" : ParagraphStyle;
                sb.Append($"<p style=\"{pStyle}\">{HtmlEncode(paragraphText.ToString())}</p>");
            }
        }

        sb.Append($"<div style=\"{SignatureSeparatorStyle}\"></div>");
        sb.Append("</div>");
        return sb.ToString();
    }

    /// <summary>
    /// 输出块级元素的首次标记。首个块级元素前不输出任何分隔，
    /// 后续块级元素之间由各自的 margin 自然分隔。
    /// </summary>
    private static void EmitBlockOpen(StringBuilder sb, ref bool isFirstBlock)
    {
        isFirstBlock = false;
    }

    /// <summary>
    /// 判断文本行是否为无序列表项（以 -、*、• 开头）。
    /// </summary>
    private static bool IsUnorderedListItem(string trimmed)
    {
        if (trimmed.Length < 2) return false;
        var first = trimmed[0];
        return (first == '-' || first == '*' || first == '\u2022') && char.IsWhiteSpace(trimmed[1]);
    }

    /// <summary>
    /// 从无序列表项中提取正文内容（去除前面的标记符号和空格）。
    /// </summary>
    private static string ExtractListItemContent(string trimmed)
    {
        return trimmed[1..].TrimStart();
    }

    /// <summary>
    /// 判断文本行是否为有序列表项（以 "数字." 或 "数字)" 开头）。
    /// </summary>
    private static bool IsOrderedListItem(string trimmed)
    {
        if (trimmed.Length < 2) return false;
        var match = Regex.Match(trimmed, @"^(\d+)[.)]\s");
        return match.Success;
    }

    /// <summary>
    /// 从有序列表项中提取正文内容（去除前面的序号标记和空格）。
    /// </summary>
    private static string ExtractOrderedListItemContent(string trimmed)
    {
        var match = Regex.Match(trimmed, @"^(\d+)[.)]\s*(.*)$");
        return match.Success ? match.Groups[2].Value : trimmed;
    }

    /// <summary>
    /// 递归遍历内联元素集合，将每个内联元素转换为对应的 HTML 标签。
    /// 支持 Run、Bold、Italic、Underline、Hyperlink、LineBreak、Span。
    /// </summary>
    private static void WriteInlines(StringBuilder sb, InlineCollection inlines)
    {
        foreach (var inline in inlines)
        {
            switch (inline)
            {
                case Run run:
                    WriteRun(sb, run);
                    break;
                case Bold bold:
                    sb.Append("<strong style=\"font-weight:600;color:#111111;\">");
                    WriteInlines(sb, bold.Inlines);
                    sb.Append("</strong>");
                    break;
                case Italic italic:
                    sb.Append("<em>");
                    WriteInlines(sb, italic.Inlines);
                    sb.Append("</em>");
                    break;
                case Underline underline:
                    sb.Append("<u>");
                    WriteInlines(sb, underline.Inlines);
                    sb.Append("</u>");
                    break;
                case Hyperlink hyperlink:
                    WriteHyperlink(sb, hyperlink);
                    break;
                case LineBreak:
                    sb.Append("<br/>");
                    break;
                case Span span:
                    WriteSpan(sb, span);
                    break;
            }
        }
    }

    /// <summary>
    /// 将 Hyperlink 元素转换为带样式的 HTML &lt;a&gt; 标签。
    /// 若超链接有 NavigateUri 则作为 href，否则使用空链接。
    /// </summary>
    private static void WriteHyperlink(StringBuilder sb, Hyperlink hyperlink)
    {
        var href = hyperlink.NavigateUri?.ToString() ?? "#";
        sb.Append($"<a href=\"{HtmlEncode(href)}\" style=\"{LinkStyle}\" target=\"_blank\">");
        WriteInlines(sb, hyperlink.Inlines);
        sb.Append("</a>");
    }

    /// <summary>
    /// 将 <see cref="Run"/> 元素转换为 HTML。若存在字号或颜色样式则包裹在 span 标签中。
    /// </summary>
    private static void WriteRun(StringBuilder sb, Run run)
    {
        var styles = GetInlineStyles(run);
        if (styles.Length > 0)
            sb.Append($"<span style=\"{styles}\">{HtmlEncode(run.Text)}</span>");
        else
            sb.Append(HtmlEncode(run.Text));
    }

    /// <summary>
    /// 将 <see cref="Span"/> 元素转换为 HTML。若存在样式则包裹在 span 标签中，否则直接输出子元素。
    /// </summary>
    private static void WriteSpan(StringBuilder sb, Span span)
    {
        var styles = GetInlineStyles(span);
        if (styles.Length > 0)
        {
            sb.Append($"<span style=\"{styles}\">");
            WriteInlines(sb, span.Inlines);
            sb.Append("</span>");
        }
        else
        {
            WriteInlines(sb, span.Inlines);
        }
    }

    /// <summary>
    /// 提取内联元素的字号和颜色样式。
    /// 字号仅在与默认值差异超过 0.5px 时输出；
    /// 颜色仅在该属性被显式设置时输出。
    /// </summary>
    private static string GetInlineStyles(Inline inline)
    {
        var parts = new List<string>();
        if (inline.FontSize != double.NaN && inline.FontSize != 0 && Math.Abs(inline.FontSize - DefaultFontSize) > 0.5)
            parts.Add($"font-size:{(int)inline.FontSize}px");
        if (inline.ReadLocalValue(System.Windows.Documents.TextElement.ForegroundProperty) != DependencyProperty.UnsetValue)
        {
            if (inline.Foreground is SolidColorBrush colorBrush)
                parts.Add($"color:{colorBrush.Color.ToString().Substring(0, 7)}");
        }
        return string.Join(";", parts);
    }

    /// <summary>
    /// 生成完整的邮件 HTML 文档包装。
    /// 将正文内容包裹在标准的 &lt;!DOCTYPE html&gt; 结构中，
    /// 包含 charset 声明、邮件客户端兼容的 body 样式、基础链接颜色。
    /// </summary>
    /// <param name="bodyContent">邮件正文 HTML 片段（通常是 ToHtml 或 PlainTextToHtml 的输出）。</param>
    /// <returns>完整的 HTML 文档字符串，可直接用于邮件发送。</returns>
    public static string WrapAsEmailDocument(string bodyContent)
    {
        var sb = new StringBuilder();
        sb.Append("<!DOCTYPE html>");
        sb.Append("<html lang=\"zh-CN\">");
        sb.Append("<head>");
        sb.Append("<meta charset=\"utf-8\">");
        sb.Append("<meta http-equiv=\"X-UA-Compatible\" content=\"IE=edge\">");
        sb.Append("<meta name=\"viewport\" content=\"width=device-width,initial-scale=1.0\">");
        sb.Append("<style>");
        sb.Append("body{background:#F7F6F3;margin:0;padding:24px;}");
        sb.Append("a{color:#1F6C9F;text-decoration:underline;}");
        sb.Append("img{max-width:100%;height:auto;}");
        sb.Append("</style>");
        sb.Append("</head>");
        sb.Append("<body>");
        sb.Append(bodyContent);
        sb.Append("</body>");
        sb.Append("</html>");
        return sb.ToString();
    }

    /// <summary>
    /// 生成轻量的预览用 HTML 文档包装。
    /// 与 <see cref="WrapAsEmailDocument"/> 不同，此方法不添加背景色、外边距和居中布局，
    /// 而是覆盖内部邮件容器的 max-width 和 padding，让内容自然填满 WebBrowser 宽度。
    /// 仅用于预览页面，不用于邮件发送。
    /// </summary>
    public static string WrapAsPreviewHtml(string bodyContent)
    {
        var sb = new StringBuilder();
        sb.Append("<!DOCTYPE html>");
        sb.Append("<html>");
        sb.Append("<head>");
        sb.Append("<meta http-equiv=\"X-UA-Compatible\" content=\"IE=edge\">");
        sb.Append("<style>");
        sb.Append("html,body{margin:0;padding:0;width:100%;overflow-x:hidden;}");
        sb.Append("body{font-family:'Segoe UI','Microsoft YaHei UI',sans-serif;font-size:15px;line-height:1.7;color:#2F3437;word-wrap:break-word;overflow-wrap:break-word;padding:8px 12px;}");
        sb.Append("body>div{max-width:100%!important;margin:0!important;padding:4px 0!important;background:none!important;}");
        sb.Append("a{color:#1F6C9F;}");
        sb.Append("p,h1,h2,h3,h4,ul,ol,blockquote,table{max-width:100%;overflow-wrap:break-word;word-wrap:break-word;}");
        sb.Append("td{word-wrap:break-word;overflow-wrap:break-word;}");
        sb.Append("</style>");
        sb.Append("</head>");
        sb.Append("<body>");
        sb.Append(bodyContent);
        sb.Append("</body>");
        sb.Append("</html>");
        return sb.ToString();
    }

    /// <summary>
    /// 将 <see cref="FlowDocument"/> 序列化为 XAML 字符串，用于富文本编辑器状态保存。
    /// </summary>
    public static string SaveToXaml(FlowDocument doc)
    {
        var range = new TextRange(doc.ContentStart, doc.ContentEnd);
        using var stream = new MemoryStream();
        range.Save(stream, DataFormats.Xaml);
        return Encoding.UTF8.GetString(stream.ToArray());
    }

    /// <summary>
    /// 从 XAML 字符串反序列化内容到 <see cref="FlowDocument"/>。
    /// 若 XAML 加载失败，则降级为纯文本加载。
    /// </summary>
    public static void LoadFromXaml(FlowDocument doc, string xaml)
    {
        if (string.IsNullOrEmpty(xaml)) return;
        try
        {
            var range = new TextRange(doc.ContentStart, doc.ContentEnd);
            using var stream = new MemoryStream(Encoding.UTF8.GetBytes(xaml));
            range.Load(stream, DataFormats.Xaml);
        }
        catch (Exception ex)
        {
            AppLogger.Error("加载XAML到FlowDocument失败", ex);
            var range = new TextRange(doc.ContentStart, doc.ContentEnd);
            range.Text = xaml;
        }
    }

    /// <summary>
    /// 对文本进行 HTML 特殊字符编码，防止 XSS 和格式错误。
    /// </summary>
    private static string HtmlEncode(string text)
    {
        return text.Replace("&", "&amp;")
                   .Replace("<", "&lt;")
                   .Replace(">", "&gt;")
                   .Replace("\"", "&quot;")
                   .Replace("\n", "<br/>");
    }
}
