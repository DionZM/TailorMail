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
/// </summary>
public static class FlowDocumentHelper
{
    /// <summary>
    /// 将 <see cref="FlowDocument"/> 转换为 HTML 字符串。
    /// 支持 Paragraph、Bold、Italic、Underline、Span、Run、LineBreak 等元素，
    /// 以及字号和颜色的内联样式输出。
    /// </summary>
    /// <param name="doc">要转换的 FlowDocument 对象。</param>
    /// <returns>HTML 格式的字符串。</returns>
    public static string ToHtml(FlowDocument doc)
    {
        var sb = new StringBuilder();
        sb.Append("<div style=\"font-family:'Microsoft YaHei UI',sans-serif;font-size:16px;line-height:1.8;\">");
        foreach (Block block in doc.Blocks)
        {
            if (block is Paragraph para)
            {
                sb.Append("<p style=\"margin:0 0 8px 0;\">");
                WriteInlines(sb, para.Inlines);
                sb.Append("</p>");
            }
        }
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
                    sb.Append("<b>");
                    WriteInlines(sb, bold.Inlines);
                    sb.Append("</b>");
                    break;
                case Italic italic:
                    sb.Append("<i>");
                    WriteInlines(sb, italic.Inlines);
                    sb.Append("</i>");
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
    /// 字号仅在与默认值（16px）差异超过 0.5px 时输出；
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
                parts.Add($"color:{colorBrush.Color.ToString().Substring(0,7)}");
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
            // XAML 加载失败时降级为纯文本
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
