using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Documents;
using System.Windows.Media;
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
    /// 邮件 HTML 的默认内联 CSS 样式，基于极简主义设计规范：
    /// - 背景: 纯白 #FFFFFF
    /// - 正文: 炭灰色 #2F3437, 字号 15px, 行高 1.75
    /// - 次要文字: 柔灰 #787774
    /// - 分隔线: 极浅灰 #EAEAEA
    /// - 链接: 柔蓝 #1F6C9F
    /// - 最大宽度: 640px 居中，模拟 A4 信纸阅读体验
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
    /// 段落样式：底部间距 16px，首段无顶部间距。
    /// </summary>
    private const string ParagraphStyle = "margin:0 0 16px 0;";

    /// <summary>
    /// 将 <see cref="FlowDocument"/> 转换为美观的邮件 HTML 字符串。
    /// 采用极简主义排版设计：宽裕的留白、精致的字体层级、温暖的纯白色调。
    /// 支持 Paragraph、Bold、Italic、Underline、Span、Run、LineBreak 等元素。
    /// </summary>
    /// <param name="doc">要转换的 FlowDocument 对象。</param>
    /// <returns>带有精致排版的 HTML 格式字符串。</returns>
    public static string ToHtml(FlowDocument doc)
    {
        var sb = new StringBuilder();
        sb.Append($"<div style=\"{EmailBaseStyle}\">");

        var isFirstParagraph = true;
        foreach (Block block in doc.Blocks)
        {
            if (block is Paragraph para)
            {
                var extraStyle = isFirstParagraph ? "margin-top:0;" : "";
                isFirstParagraph = false;

                sb.Append($"<p style=\"{ParagraphStyle}{extraStyle}\">");
                WriteInlines(sb, para.Inlines);
                sb.Append("</p>");
            }
            else if (block is Section section)
            {
                foreach (var child in section.Blocks)
                {
                    if (child is Paragraph childPara)
                    {
                        sb.Append($"<p style=\"{ParagraphStyle}\">");
                        WriteInlines(sb, childPara.Inlines);
                        sb.Append("</p>");
                    }
                }
            }
        }

        // 底部签名分隔线
        sb.Append("<div style=\"margin-top:32px;padding-top:16px;border-top:1px solid #EAEAEA;\"></div>");
        sb.Append("</div>");
        return sb.ToString();
    }

    /// <summary>
    /// 将纯文本转换为美观的邮件 HTML，保留换行并应用精致的默认样式。
    /// 用于没有 XAML 格式信息时的回退场景。
    /// </summary>
    /// <param name="plainText">纯文本内容。</param>
    /// <returns>带有精致排版的 HTML 格式字符串。</returns>
    public static string PlainTextToHtml(string plainText)
    {
        if (string.IsNullOrEmpty(plainText)) return string.Empty;

        var sb = new StringBuilder();
        sb.Append($"<div style=\"{EmailBaseStyle}\">");

        var lines = plainText.Split('\n');
        var isFirstLine = true;
        var currentParagraph = new StringBuilder();

        foreach (var line in lines)
        {
            var trimmed = line.TrimEnd('\r');

            if (string.IsNullOrEmpty(trimmed))
            {
                // 空行触发段落分隔
                if (currentParagraph.Length > 0)
                {
                    var extraStyle = isFirstLine ? "margin-top:0;" : "";
                    isFirstLine = false;
                    sb.Append($"<p style=\"{ParagraphStyle}{extraStyle}\">");
                    sb.Append(HtmlEncode(currentParagraph.ToString()));
                    sb.Append("</p>");
                    currentParagraph.Clear();
                }
            }
            else
            {
                if (currentParagraph.Length > 0) currentParagraph.Append(' ');
                currentParagraph.Append(trimmed);
            }
        }

        // 输出最后一段
        if (currentParagraph.Length > 0)
        {
            var extraStyle = isFirstLine ? "margin-top:0;" : "";
            sb.Append($"<p style=\"{ParagraphStyle}{extraStyle}\">");
            sb.Append(HtmlEncode(currentParagraph.ToString()));
            sb.Append("</p>");
        }

        // 底部签名分隔线
        sb.Append("<div style=\"margin-top:32px;padding-top:16px;border-top:1px solid #EAEAEA;\"></div>");
        sb.Append("</div>");
        return sb.ToString();
    }

    /// <summary>
    /// 递归遍历内联元素集合，将每个内联元素转换为对应的 HTML 标签。
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
        if (inline.FontSize != double.NaN && inline.FontSize != 0 && Math.Abs(inline.FontSize - 16.0) > 0.5)
            parts.Add($"font-size:{(int)inline.FontSize}px");
        if (inline.ReadLocalValue(System.Windows.Documents.TextElement.ForegroundProperty) != DependencyProperty.UnsetValue)
        {
            if (inline.Foreground is SolidColorBrush colorBrush)
                parts.Add($"color:{colorBrush.Color.ToString().Substring(0, 7)}");
        }
        return string.Join(";", parts);
    }

    /// <summary>
    /// 将 <see cref="FlowDocument"/> 序列化为 XAML 字符串，用于富文本编辑器状态保存。
    /// </summary>
    /// <param name="doc">要序列化的 FlowDocument 对象。</param>
    /// <returns>XAML 格式的字符串。</returns>
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
    /// <param name="doc">目标 FlowDocument 对象。</param>
    /// <param name="xaml">XAML 格式的字符串。</param>
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
