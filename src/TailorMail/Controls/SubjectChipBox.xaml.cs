using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;

namespace TailorMail.Controls;

public partial class SubjectChipBox : UserControl
{
    private bool _isUpdating;
    private HashSet<string> _variableNames = [];
    private static readonly Regex VariablePattern = new(@"\{([^}]+)\}", RegexOptions.Compiled);

    public static readonly DependencyProperty SubjectProperty =
        DependencyProperty.Register(nameof(Subject), typeof(string), typeof(SubjectChipBox),
            new FrameworkPropertyMetadata(string.Empty, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, OnSubjectChanged));

    public static readonly DependencyProperty PlaceholderTextProperty =
        DependencyProperty.Register(nameof(PlaceholderText), typeof(string), typeof(SubjectChipBox),
            new PropertyMetadata(""));

    public static readonly DependencyProperty VariablesProperty =
        DependencyProperty.Register(nameof(Variables), typeof(IEnumerable<string>), typeof(SubjectChipBox),
            new PropertyMetadata(null, OnVariablesChanged));

    public string Subject
    {
        get => (string)GetValue(SubjectProperty);
        set => SetValue(SubjectProperty, value);
    }

    public string PlaceholderText
    {
        get => (string)GetValue(PlaceholderTextProperty);
        set => SetValue(PlaceholderTextProperty, value);
    }

    public IEnumerable<string>? Variables
    {
        get => (IEnumerable<string>?)GetValue(VariablesProperty);
        set => SetValue(VariablesProperty, value);
    }

    public RichTextBox GetInternalEditor() => InternalEditor;

    public SubjectChipBox()
    {
        InitializeComponent();
    }

    private static void OnSubjectChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
    {
        if (d is SubjectChipBox box && !box._isUpdating)
            box.RenderFromPlainText(e.NewValue as string ?? "");
    }

    private static void OnVariablesChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
    {
        if (d is SubjectChipBox box && e.NewValue is IEnumerable<string> vars)
            box._variableNames = new HashSet<string>(vars);
    }

    private void OnEditorTextChanged(object sender, TextChangedEventArgs e)
    {
        if (_isUpdating) return;
        _isUpdating = true;

        var plainText = GetPlainText();
        if (plainText != Subject)
        {
            Subject = plainText;
        }

        _isUpdating = false;
    }

    private void OnEditorSelectionChanged(object sender, RoutedEventArgs e) { }

    private void OnEditorPreviewKeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Enter)
        {
            e.Handled = true;
        }
    }

    private string GetPlainText()
    {
        var doc = InternalEditor.Document;
        var sb = new System.Text.StringBuilder();
        foreach (var block in doc.Blocks)
        {
            if (block is Paragraph para)
                AppendInlineText(para.Inlines, sb);
        }
        return sb.ToString();
    }

    private void AppendInlineText(InlineCollection inlines, System.Text.StringBuilder sb)
    {
        foreach (var inline in inlines)
        {
            if (inline is Run run)
                sb.Append(run.Text);
            else if (inline is InlineUIContainer ui && ui.Child is Border border)
            {
                foreach (var child in LogicalTreeHelper.GetChildren(border))
                {
                    if (child is TextBlock tb)
                        sb.Append(tb.Text);
                }
            }
            else if (inline is Span span)
                AppendInlineText(span.Inlines, sb);
        }
    }

    public void RenderFromPlainText(string text)
    {
        _isUpdating = true;
        var doc = InternalEditor.Document;
        doc.Blocks.Clear();

        var para = new Paragraph { Margin = new Thickness(0) };

        if (string.IsNullOrEmpty(text))
        {
            doc.Blocks.Add(para);
            _isUpdating = false;
            return;
        }

        var matches = VariablePattern.Matches(text);
        int lastPos = 0;

        for (int i = 0; i < matches.Count; i++)
        {
            var match = matches[i];
            if (match.Index > lastPos)
                para.Inlines.Add(new Run(text[lastPos..match.Index]));

            var varName = match.Groups[1].Value;
            var chip = CreateChip(varName);
            para.Inlines.Add(new InlineUIContainer(chip) { BaselineAlignment = BaselineAlignment.Center });

            lastPos = match.Index + match.Length;
        }

        if (lastPos < text.Length)
            para.Inlines.Add(new Run(text[lastPos..]));

        doc.Blocks.Add(para);
        _isUpdating = false;
    }

    private Border CreateChip(string varName)
    {
        var textBlock = new TextBlock
        {
            Text = varName,
            FontSize = 12,
            FontWeight = FontWeights.Medium,
            Foreground = (Brush)FindResource("AccentBrush"),
            VerticalAlignment = VerticalAlignment.Center
        };

        return new Border
        {
            Background = (Brush)FindResource("AccentLightBrush"),
            CornerRadius = new CornerRadius(3),
            Padding = new Thickness(4, 1, 4, 1),
            Margin = new Thickness(1, 0, 1, 0),
            Child = textBlock
        };
    }

    public void InsertVariableAtSelection(string variableName)
    {
        var text = $"{{{variableName}}}";
        var caret = InternalEditor.CaretPosition;
        var chip = CreateChip(variableName);
        var container = new InlineUIContainer(chip, caret) { BaselineAlignment = BaselineAlignment.Center };

        _isUpdating = true;
        Subject = GetPlainText();
        _isUpdating = false;

        InternalEditor.CaretPosition = container.ElementEnd;
        InternalEditor.Focus();
    }
}
