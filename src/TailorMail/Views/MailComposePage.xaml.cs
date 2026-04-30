using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using TailorMail.Helpers;
using TailorMail.ViewModels;

namespace TailorMail.Views;

/// <summary>
/// 邮件撰写页面，提供富文本编辑器、格式工具栏、变量插入、模板保存/加载等功能。
/// </summary>
public partial class MailComposePage : UserControl, IRefreshable
{
    private readonly MailComposeViewModel _vm;
    private bool _isUpdating;

    private static readonly double[] FontSizeSteps = [8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 28, 32, 36, 48, 72];
    private const double DefaultFontSize = 16;
    private const double MinFontSize = 8;
    private const double MaxFontSize = 72;

    public MailComposePage()
    {
        InitializeComponent();
        _vm = new MailComposeViewModel(App.DataService);
        DataContext = _vm;
        LoadContentToEditor();
        LoadVariables();
    }

    public void RefreshData() => LoadVariables();

    private void LoadContentToEditor()
    {
        _isUpdating = true;
        if (!string.IsNullOrEmpty(_vm.BodyXaml))
            FlowDocumentHelper.LoadFromXaml(Editor.Document, _vm.BodyXaml);
        _isUpdating = false;
    }

    private void LoadVariables()
    {
        var varVm = new VariablesViewModel(App.DataService);
        var placeholders = varVm.GetAllVariablePlaceholders();
        VariableCombo.ItemsSource = placeholders;
        if (placeholders.Count > 0) VariableCombo.SelectedIndex = 0;
    }

    private void OnEditorTextChanged(object sender, TextChangedEventArgs e)
    {
        if (_isUpdating) return;
        _vm.BodyXaml = FlowDocumentHelper.SaveToXaml(Editor.Document);
        _vm.Body = new TextRange(Editor.Document.ContentStart, Editor.Document.ContentEnd).Text.TrimEnd('\r', '\n');
    }

    private void OnEditorSelectionChanged(object sender, RoutedEventArgs e)
    {
        UpdateToolbarState();
    }

    private void UpdateToolbarState()
    {
        if (Editor == null || Editor.Selection == null) return;

        var sel = Editor.Selection;
        if (sel.IsEmpty)
        {
            var parent = Editor.CaretPosition.Parent as FrameworkContentElement;
            BoldBtn.IsChecked = ReadProperty<bool>(parent, TextElement.FontWeightProperty, v => v is FontWeight fw && fw == FontWeights.Bold);
            ItalicBtn.IsChecked = ReadProperty<bool>(parent, TextElement.FontStyleProperty, v => v is FontStyle fs && fs == FontStyles.Italic);
            UnderlineBtn.IsChecked = ReadProperty<bool>(parent, Inline.TextDecorationsProperty, v => v is TextDecorationCollection tdc && tdc == TextDecorations.Underline);
            var fontSize = ReadProperty<double>(parent, TextElement.FontSizeProperty, v => v is double d ? d : DefaultFontSize);
            FontSizeLabel.Text = ((int)fontSize).ToString();
        }
        else
        {
            BoldBtn.IsChecked = sel.GetPropertyValue(TextElement.FontWeightProperty) as FontWeight? == FontWeights.Bold;
            ItalicBtn.IsChecked = sel.GetPropertyValue(TextElement.FontStyleProperty) as FontStyle? == FontStyles.Italic;
            UnderlineBtn.IsChecked = sel.GetPropertyValue(Inline.TextDecorationsProperty) == TextDecorations.Underline;
            var fs = sel.GetPropertyValue(TextElement.FontSizeProperty) as double? ?? DefaultFontSize;
            FontSizeLabel.Text = ((int)fs).ToString();
        }
    }

    private T ReadProperty<T>(FrameworkContentElement? element, DependencyProperty prop, Func<object?, T> extract)
    {
        while (element != null)
        {
            var val = element.GetValue(prop);
            if (val != null && !Equals(val, DependencyProperty.UnsetValue))
                return extract(val);
            element = element.Parent as FrameworkContentElement;
        }
        return extract(null);
    }

    private double GetNextFontSize(double current, bool increase)
    {
        if (increase)
        {
            foreach (var s in FontSizeSteps)
            {
                if (s > current + 0.5) return s;
            }
            return MaxFontSize;
        }
        else
        {
            for (int i = FontSizeSteps.Length - 1; i >= 0; i--)
            {
                if (FontSizeSteps[i] < current - 0.5) return FontSizeSteps[i];
            }
            return MinFontSize;
        }
    }

    private void OnIncreaseFontSize(object sender, RoutedEventArgs e)
    {
        var currentSize = GetCurrentFontSize();
        var newSize = GetNextFontSize(currentSize, true);
        ApplyFontSize(newSize);
    }

    private void OnDecreaseFontSize(object sender, RoutedEventArgs e)
    {
        var currentSize = GetCurrentFontSize();
        var newSize = GetNextFontSize(currentSize, false);
        ApplyFontSize(newSize);
    }

    private double GetCurrentFontSize()
    {
        if (Editor.Selection != null && !Editor.Selection.IsEmpty)
        {
            return Editor.Selection.GetPropertyValue(TextElement.FontSizeProperty) as double? ?? DefaultFontSize;
        }
        var parent = Editor.CaretPosition.Parent as FrameworkContentElement;
        return ReadProperty<double>(parent, TextElement.FontSizeProperty, v => v is double d ? d : DefaultFontSize);
    }

    private void ApplyFontSize(double size)
    {
        if (Editor.Selection != null && !Editor.Selection.IsEmpty)
        {
            Editor.Selection.ApplyPropertyValue(TextElement.FontSizeProperty, size);
        }
        else
        {
            var tp = Editor.CaretPosition;
            if (tp.Parent is Inline inline)
            {
                inline.FontSize = size;
            }
            else if (tp.Paragraph != null)
            {
                tp.Paragraph.FontSize = size;
            }
        }
        FontSizeLabel.Text = ((int)size).ToString();
        Editor.Focus();
    }

    private void OnToggleBold(object sender, RoutedEventArgs e)
    {
        Editor.Selection.ApplyPropertyValue(TextElement.FontWeightProperty,
            BoldBtn.IsChecked == true ? FontWeights.Bold : FontWeights.Normal);
        Editor.Focus();
    }

    private void OnToggleItalic(object sender, RoutedEventArgs e)
    {
        Editor.Selection.ApplyPropertyValue(TextElement.FontStyleProperty,
            ItalicBtn.IsChecked == true ? FontStyles.Italic : FontStyles.Normal);
        Editor.Focus();
    }

    private void OnToggleUnderline(object sender, RoutedEventArgs e)
    {
        Editor.Selection.ApplyPropertyValue(Inline.TextDecorationsProperty,
            UnderlineBtn.IsChecked == true ? TextDecorations.Underline : null);
        Editor.Focus();
    }

    private void OnColorPickerClick(object sender, RoutedEventArgs e)
    {
        ColorPopup.IsOpen = !ColorPopup.IsOpen;
    }

    private void OnColorSelected(object sender, MouseButtonEventArgs e)
    {
        if (sender is FrameworkElement fe && fe.Tag is string colorStr)
        {
            ApplyColor(colorStr);
        }
    }

    private void OnApplyCustomColor(object sender, RoutedEventArgs e)
    {
        var input = CustomColorInput?.Text?.Trim();
        if (string.IsNullOrEmpty(input)) return;
        if (!input.StartsWith("#")) input = "#" + input;
        try
        {
            var color = (Color)ColorConverter.ConvertFromString(input);
            ApplyColor(input);
            if (CustomColorInput != null) CustomColorInput.Text = "";
        }
        catch
        {
            if (CustomColorInput != null) CustomColorInput.Text = "";
        }
    }

    private void ApplyColor(string colorStr)
    {
        var color = (Color)ColorConverter.ConvertFromString(colorStr);
        var brush = new SolidColorBrush(color);
        Editor.Selection.ApplyPropertyValue(TextElement.ForegroundProperty, brush);
        ColorPreview.Background = brush;
        ColorPopup.IsOpen = false;
        Editor.Focus();
    }

    private void OnInsertVariable(object sender, RoutedEventArgs e)
    {
        if (VariableCombo.SelectedItem is string varName)
        {
            Editor.CaretPosition.InsertTextInRun(varName);
            Editor.Focus();
        }
    }

    public void SaveCurrent() => _vm.SaveCurrent();
}
