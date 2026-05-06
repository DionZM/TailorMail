using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using TailorMail.Helpers;
using TailorMail.ViewModels;
using UiTextBox = Wpf.Ui.Controls.TextBox;

namespace TailorMail.Views;

public partial class MailComposePage : UserControl, IRefreshable
{
    private readonly MailComposeViewModel _vm;
    private bool _isUpdating;
    private List<string> _allVariablePlaceholders = [];
    private static readonly List<string> _recentColors = [];
    private const int MaxRecentColors = 7;

    private static readonly double[] FontSizeSteps = [8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 28, 32, 36, 48, 72];
    private const double DefaultFontSize = 16;
    private const double MinFontSize = 8;
    private const double MaxFontSize = 72;

    private bool _acActive;
    private int _acStartIndex;
    private TextPointer? _acStartPointer;
    private Control? _acSource;

    public MailComposePage()
    {
        InitializeComponent();
        _vm = new MailComposeViewModel(App.DataService);
        DataContext = _vm;
        LoadContentToEditor();
        LoadVariables();
        InitFontSizeCombo();
        AttachSubjectAutoComplete();
        AttachEditorAutoComplete();
        AutoCompleteList.PreviewKeyDown += OnAutoCompleteListKeyDown;
    }

    private void InitFontSizeCombo()
    {
        foreach (var size in FontSizeSteps)
            FontSizeCombo.Items.Add((int)size);
        FontSizeCombo.SelectedItem = (int)DefaultFontSize;
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
        _allVariablePlaceholders = placeholders;
        VariableCombo.ItemsSource = placeholders;
        if (placeholders.Count > 0) VariableCombo.SelectedIndex = 0;
        UpdateVariableHint();
    }

    private void UpdateVariableHint()
    {
        var varVm = new VariablesViewModel(App.DataService);
        var vars = varVm.GetAllVariablePlaceholders();
        if (vars.Count == 0)
        {
            VarHintText.Text = "提示：请先在「变量配置」步骤中创建变量";
        }
        else
        {
            VarHintText.Text = $"提示：使用 {{变量名}} 格式插入变量，如 {string.Join("、", vars.Take(3))}，系统将自动替换为对应值";
        }
    }

    #region AutoComplete

    private void AttachSubjectAutoComplete()
    {
        SubjectBox.PreviewTextInput += OnSubjectPreviewTextInput;
        SubjectBox.PreviewKeyDown += OnSubjectPreviewKeyDown;
    }

    private void AttachEditorAutoComplete()
    {
        Editor.PreviewTextInput += OnEditorPreviewTextInput;
        Editor.PreviewKeyDown += OnEditorPreviewKeyDown;
    }

    private void OnSubjectPreviewTextInput(object sender, TextCompositionEventArgs e)
    {
        HandleAutoCompleteTextInput(e.Text, SubjectBox);
    }

    private void OnSubjectPreviewKeyDown(object sender, KeyEventArgs e)
    {
        if (_acActive && _acSource == SubjectBox)
            HandleAutoCompleteKey(e, SubjectBox);
    }

    private void OnEditorPreviewTextInput(object sender, TextCompositionEventArgs e)
    {
        HandleAutoCompleteTextInput(e.Text, Editor);
    }

    private void OnEditorPreviewKeyDown(object sender, KeyEventArgs e)
    {
        if (_acActive && _acSource == Editor)
            HandleAutoCompleteKey(e, Editor);
    }

    private void HandleAutoCompleteTextInput(string text, Control source)
    {
        if (text == "{")
        {
            _acSource = source;
            if (source is UiTextBox tb)
            {
                _acStartIndex = tb.SelectionStart + 1;
                _acStartPointer = null;
            }
            else if (source is RichTextBox rtb)
            {
                _acStartPointer = rtb.CaretPosition.GetPositionAtOffset(0, LogicalDirection.Forward);
                _acStartIndex = -1;
            }
            ShowAutoComplete(source);
        }
        else if (_acActive && _acSource == source)
        {
            Dispatcher.BeginInvoke(() => UpdateAutoCompleteFilter(source), System.Windows.Threading.DispatcherPriority.Background);
        }
    }

    private void HandleAutoCompleteKey(KeyEventArgs e, Control source)
    {
        switch (e.Key)
        {
            case Key.Down:
                if (AutoCompleteList.Items.Count > 0)
                    AutoCompleteList.SelectedIndex = Math.Min(AutoCompleteList.SelectedIndex + 1, AutoCompleteList.Items.Count - 1);
                e.Handled = true;
                break;
            case Key.Up:
                if (AutoCompleteList.Items.Count > 0)
                    AutoCompleteList.SelectedIndex = Math.Max(AutoCompleteList.SelectedIndex - 1, 0);
                e.Handled = true;
                break;
            case Key.Enter:
                if (AutoCompleteList.SelectedItem is string selected)
                    InsertAutoCompleteValue(selected, source);
                e.Handled = true;
                break;
            case Key.Tab:
                if (AutoCompleteList.SelectedItem is string tabSel)
                {
                    InsertAutoCompleteValue(tabSel, source);
                    e.Handled = true;
                }
                break;
            case Key.Escape:
                CloseAutoComplete();
                e.Handled = true;
                break;
        }
    }

    private void ShowAutoComplete(Control source)
    {
        if (_allVariablePlaceholders.Count == 0) return;
        AutoCompleteList.ItemsSource = _allVariablePlaceholders;
        AutoCompleteList.SelectedIndex = 0;

        AutoCompletePopup.PlacementTarget = source;
        AutoCompletePopup.Placement = PlacementMode.Relative;
        AutoCompletePopup.HorizontalOffset = 0;
        AutoCompletePopup.VerticalOffset = 0;

        if (source is RichTextBox rtb)
        {
            try
            {
                var caretRect = rtb.CaretPosition.GetCharacterRect(LogicalDirection.Forward);
                AutoCompletePopup.HorizontalOffset = caretRect.Left;
                AutoCompletePopup.VerticalOffset = caretRect.Bottom + 2;
            }
            catch { }
        }
        else if (source is UiTextBox tb)
        {
            AutoCompletePopup.VerticalOffset = tb.ActualHeight + 2;
        }

        _acActive = true;
        AutoCompletePopup.IsOpen = true;
    }

    private void UpdateAutoCompleteFilter(Control source)
    {
        if (!_acActive) return;

        string typed;
        if (source is UiTextBox tb)
        {
            typed = tb.Text?.Substring(_acStartIndex) ?? "";
        }
        else if (source is RichTextBox rtb && _acStartPointer != null)
        {
            typed = new TextRange(_acStartPointer, rtb.CaretPosition).Text;
        }
        else
            return;

        var filtered = _allVariablePlaceholders
            .Where(v => v.Contains(typed, StringComparison.OrdinalIgnoreCase))
            .ToList();

        if (filtered.Count == 0 || typed.Contains('}'))
        {
            CloseAutoComplete();
            return;
        }

        AutoCompleteList.ItemsSource = filtered;
        AutoCompleteList.SelectedIndex = 0;
    }

    private void InsertAutoCompleteValue(string variable, Control source)
    {
        var varName = variable.Trim('{', '}');
        var insertText = varName + "}";
        if (source is UiTextBox tb)
        {
            var before = tb.Text?.Substring(0, _acStartIndex) ?? "";
            var after = tb.Text?.Substring(tb.SelectionStart) ?? "";
            tb.Text = before + insertText + after;
            tb.SelectionStart = _acStartIndex + insertText.Length;
        }
        else if (source is RichTextBox rtb && _acStartPointer != null)
        {
            var replaceRange = new TextRange(_acStartPointer, rtb.CaretPosition);
            replaceRange.Text = insertText;
            rtb.CaretPosition = replaceRange.End;
        }
        CloseAutoComplete();
        Editor?.Focus();
    }

    private void CloseAutoComplete()
    {
        _acActive = false;
        _acSource = null;
        AutoCompletePopup.IsOpen = false;
    }

    private void OnAutoCompleteSelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        // Selection highlight only; insertion is handled by Enter/Tab key
    }

    private void OnAutoCompleteListKeyDown(object sender, KeyEventArgs e)
    {
        if (!_acActive || _acSource == null) return;
        if (e.Key == Key.Enter || e.Key == Key.Tab)
        {
            if (AutoCompleteList.SelectedItem is string selected)
                InsertAutoCompleteValue(selected, _acSource);
            e.Handled = true;
        }
        else if (e.Key == Key.Escape)
        {
            CloseAutoComplete();
            e.Handled = true;
        }
    }

    #endregion

    private void OnEditorTextChanged(object sender, TextChangedEventArgs e)
    {
        if (_isUpdating) return;
        _vm.BodyXaml = FlowDocumentHelper.SaveToXaml(Editor.Document);
        _vm.Body = new TextRange(Editor.Document.ContentStart, Editor.Document.ContentEnd).Text.TrimEnd('\r', '\n');
        ValidateVariablesInEditor();
    }

    private void ValidateVariablesInEditor()
    {
        var text = new TextRange(Editor.Document.ContentStart, Editor.Document.ContentEnd).Text;
        var varVm = new VariablesViewModel(App.DataService);
        var validNames = new HashSet<string>(varVm.GetAllVariablePlaceholders());

        // Check for {xxx} patterns that aren't valid variables
        var openIdx = 0;
        var unknownVars = new List<string>();
        while ((openIdx = text.IndexOf('{', openIdx)) >= 0)
        {
            var closeIdx = text.IndexOf('}', openIdx);
            if (closeIdx < 0) break;
            var candidate = text.Substring(openIdx, closeIdx - openIdx + 1);
            if (!string.IsNullOrWhiteSpace(candidate.Trim('{', '}')) && !validNames.Contains(candidate))
                unknownVars.Add(candidate);
            openIdx = closeIdx + 1;
        }

        if (unknownVars.Count > 0)
        {
            VarWarningText.Text = $"未识别的变量：{string.Join("、", unknownVars.Distinct())}";
            VarWarningText.Visibility = Visibility.Visible;
            VarHintText.Visibility = Visibility.Collapsed;
        }
        else
        {
            VarWarningText.Visibility = Visibility.Collapsed;
            VarHintText.Visibility = Visibility.Visible;
        }
    }

    private void OnEditorSelectionChanged(object sender, RoutedEventArgs e)
    {
        UpdateToolbarState();
    }

    private void UpdateToolbarState()
    {
        if (Editor == null || Editor.Selection == null) return;

        var sel = Editor.Selection;
        double fontSize;
        if (sel.IsEmpty)
        {
            var parent = Editor.CaretPosition.Parent as FrameworkContentElement;
            BoldBtn.IsChecked = ReadProperty<bool>(parent, TextElement.FontWeightProperty, v => v is FontWeight fw && fw == FontWeights.Bold);
            ItalicBtn.IsChecked = ReadProperty<bool>(parent, TextElement.FontStyleProperty, v => v is FontStyle fs && fs == FontStyles.Italic);
            UnderlineBtn.IsChecked = ReadProperty<bool>(parent, Inline.TextDecorationsProperty, v => v is TextDecorationCollection tdc && tdc == TextDecorations.Underline);
            fontSize = ReadProperty<double>(parent, TextElement.FontSizeProperty, v => v is double d ? d : DefaultFontSize);
        }
        else
        {
            BoldBtn.IsChecked = sel.GetPropertyValue(TextElement.FontWeightProperty) as FontWeight? == FontWeights.Bold;
            ItalicBtn.IsChecked = sel.GetPropertyValue(TextElement.FontStyleProperty) as FontStyle? == FontStyles.Italic;
            UnderlineBtn.IsChecked = sel.GetPropertyValue(Inline.TextDecorationsProperty) == TextDecorations.Underline;
            fontSize = sel.GetPropertyValue(TextElement.FontSizeProperty) as double? ?? DefaultFontSize;
        }

        var intSize = (int)fontSize;
        if (FontSizeCombo.Items.Contains(intSize))
            FontSizeCombo.SelectedItem = intSize;
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

    private double GetCurrentFontSize()
    {
        if (Editor.Selection != null && !Editor.Selection.IsEmpty)
            return Editor.Selection.GetPropertyValue(TextElement.FontSizeProperty) as double? ?? DefaultFontSize;
        var parent = Editor.CaretPosition.Parent as FrameworkContentElement;
        return ReadProperty<double>(parent, TextElement.FontSizeProperty, v => v is double d ? d : DefaultFontSize);
    }

    private void ApplyFontSize(double size)
    {
        if (Editor.Selection != null && !Editor.Selection.IsEmpty)
            Editor.Selection.ApplyPropertyValue(TextElement.FontSizeProperty, size);
        else
        {
            var tp = Editor.CaretPosition;
            if (tp.Parent is Inline inline)
                inline.FontSize = size;
            else if (tp.Paragraph != null)
                tp.Paragraph.FontSize = size;
        }
        Editor.Focus();
    }

    private void OnFontSizeComboChanged(object sender, SelectionChangedEventArgs e)
    {
        if (FontSizeCombo.SelectedItem is int size && !_isUpdating)
        {
            ApplyFontSize(size);
        }
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
            ApplyColor(colorStr);
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

        _recentColors.Remove(colorStr);
        _recentColors.Insert(0, colorStr);
        if (_recentColors.Count > MaxRecentColors)
            _recentColors.RemoveAt(_recentColors.Count - 1);
        UpdateRecentColors();
    }

    private void UpdateRecentColors()
    {
        if (RecentColorsPanel == null) return;
        RecentColorsPanel.Children.Clear();
        if (_recentColors.Count == 0)
        {
            RecentColorsPanel.Visibility = Visibility.Collapsed;
            return;
        }
        RecentColorsPanel.Visibility = Visibility.Visible;
        RecentColorsPanel.Children.Add(new TextBlock
        {
            Text = "最近使用",
            FontSize = 11,
            Foreground = (Brush)FindResource("TextTertiaryBrush"),
            VerticalAlignment = VerticalAlignment.Center,
            Margin = new Thickness(0, 0, 8, 4)
        });
        foreach (var colorStr in _recentColors)
        {
            var rect = new System.Windows.Shapes.Rectangle
            {
                Fill = new SolidColorBrush((Color)ColorConverter.ConvertFromString(colorStr)),
                Tag = colorStr,
                Style = (Style)FindResource("ColorSwatchStyle"),
                ToolTip = colorStr
            };
            rect.MouseLeftButtonUp += OnColorSelected;
            RecentColorsPanel.Children.Add(rect);
        }
    }

    private void OnInsertVariable(object sender, RoutedEventArgs e)
    {
        string? varName = VariableCombo.SelectedItem as string
            ?? VariableCombo.Text?.Trim();
        if (!string.IsNullOrEmpty(varName))
        {
            if (_allVariablePlaceholders.Contains(varName))
                Editor.CaretPosition.InsertTextInRun(varName);
            else if (!varName.StartsWith('{'))
                Editor.CaretPosition.InsertTextInRun($"{{{varName}}}");
            else
                Editor.CaretPosition.InsertTextInRun(varName);
            VariableCombo.ItemsSource = _allVariablePlaceholders;
            VariableCombo.Text = "";
            Editor.Focus();
        }
    }

    private void OnInsertLink(object sender, RoutedEventArgs e)
    {
        var dialog = new InputDialog
        {
            Title = "插入链接",
            Prompt = "请输入链接地址："
        };
        if (dialog.ShowDialog() == true && !string.IsNullOrWhiteSpace(dialog.InputText))
        {
            var url = dialog.InputText.Trim();
            try
            {
                var hyperlink = new Hyperlink(Editor.Selection.Start, Editor.Selection.End)
                {
                    NavigateUri = new Uri(url)
                };
            }
            catch (UriFormatException)
            {
                App.ShowWarning("请输入有效的链接地址");
            }
            Editor.Focus();
        }
    }

    public void SaveCurrent() => _vm.SaveCurrent();
}
