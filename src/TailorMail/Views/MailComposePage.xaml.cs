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
    private HashSet<string> _validVariableNames = [];
    private static readonly List<string> _recentColors = [];
    private const int MaxRecentColors = 7;

    private static readonly double[] FontSizeSteps = [8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 28, 32, 36, 48, 72];
    private const double DefaultFontSize = 16;
    private const double MinFontSize = 8;
    private const double MaxFontSize = 72;

    private readonly System.Windows.Threading.DispatcherTimer _xamlSaveTimer;
    private bool _hasPendingXamlSave;

    private bool _acActive;
    private int _acStartIndex;
    private TextPointer? _acStartPointer;
    private Control? _acSource;

    public MailComposePage()
    {
        InitializeComponent();
        _vm = new MailComposeViewModel(App.DataService);
        DataContext = _vm;
        Editor.UndoLimit = 100;
        _xamlSaveTimer = new System.Windows.Threading.DispatcherTimer { Interval = TimeSpan.FromMilliseconds(800) };
        _xamlSaveTimer.Tick += (_, _) =>
        {
            _xamlSaveTimer.Stop();
            FlushXamlSave();
        };
        LoadContentToEditor();
        LoadVariables();
        InitFontSizeCombo();
        AttachSubjectAutoComplete();
        AttachEditorAutoComplete();
        AutoCompleteList.PreviewKeyDown += OnAutoCompleteListKeyDown;
        DataObject.AddPastingHandler(Editor, OnEditorPasting);
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
        Editor.UndoLimit = 0;
        Editor.UndoLimit = 100;
        _isUpdating = false;
    }

    private void LoadVariables()
    {
        var varVm = new VariablesViewModel(App.DataService);
        var placeholders = varVm.GetAllVariablePlaceholders();
        _allVariablePlaceholders = placeholders;
        _validVariableNames = new HashSet<string>(placeholders);
        VariableCombo.ItemsSource = placeholders;
        if (placeholders.Count > 0) VariableCombo.SelectedIndex = 0;
        UpdateVariableHint();
    }

    private void UpdateVariableHint()
    {
        if (_allVariablePlaceholders.Count == 0)
        {
            VarHintText.Text = "提示：请先在「变量配置」步骤中创建变量";
        }
        else
        {
            VarHintText.Text = $"提示：使用 {{变量名}} 格式插入变量，如 {string.Join("、", _allVariablePlaceholders.Take(3))}，系统将自动替换为对应值";
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
                if (!caretRect.IsEmpty)
                {
                    AutoCompletePopup.HorizontalOffset = caretRect.Left;
                    AutoCompletePopup.VerticalOffset = caretRect.Bottom + 2;
                }
                else
                {
                    AutoCompletePopup.HorizontalOffset = 0;
                    AutoCompletePopup.VerticalOffset = 30;
                }
            }
            catch
            {
                AutoCompletePopup.HorizontalOffset = 0;
                AutoCompletePopup.VerticalOffset = 30;
            }
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
            var text = tb.Text ?? "";
            if (_acStartIndex >= text.Length) { CloseAutoComplete(); return; }
            typed = text.Substring(_acStartIndex);
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
            var text = tb.Text ?? "";
            var safeStart = Math.Min(_acStartIndex, text.Length);
            var before = text.Substring(0, safeStart);
            var after = tb.SelectionStart < text.Length ? text.Substring(tb.SelectionStart) : "";
            var newText = before + insertText + after;
            var newCursorPos = safeStart + insertText.Length;

            if (source.DataContext is MailComposeViewModel vm)
                vm.Subject = newText;

            Dispatcher.BeginInvoke(() => tb.SelectionStart = newCursorPos,
                System.Windows.Threading.DispatcherPriority.Loaded);
        }
        else if (source is RichTextBox rtb && _acStartPointer != null)
        {
            var replaceRange = new TextRange(_acStartPointer, rtb.CaretPosition);
            replaceRange.Text = insertText;
            rtb.CaretPosition = replaceRange.End;
        }
        CloseAutoComplete();
        _acSource?.Focus();
    }

    private void CloseAutoComplete()
    {
        _acActive = false;
        _acSource = null;
        AutoCompletePopup.IsOpen = false;
    }

    private void OnAutoCompleteSelectionChanged(object sender, SelectionChangedEventArgs e)
    {
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

    private void OnEditorPasting(object sender, DataObjectPastingEventArgs e)
    {
        var hadRichText = e.DataObject.GetDataPresent(DataFormats.Rtf);
        if (e.DataObject.GetDataPresent(DataFormats.Text))
        {
            var text = e.DataObject.GetData(DataFormats.Text) as string;
            if (!string.IsNullOrEmpty(text))
            {
                var cleanData = new DataObject(DataFormats.Text, text);
                e.DataObject = cleanData;
                if (hadRichText)
                    App.ShowNotification("已粘贴为纯文本");
            }
        }
    }

    private void OnEditorTextChanged(object sender, TextChangedEventArgs e)
    {
        if (_isUpdating) return;
        _vm.Body = new TextRange(Editor.Document.ContentStart, Editor.Document.ContentEnd).Text.TrimEnd('\r', '\n');
        _hasPendingXamlSave = true;
        _xamlSaveTimer.Stop();
        _xamlSaveTimer.Start();
        ValidateVariablesInEditor();
        UpdateCharCount();
    }

    private void UpdateCharCount()
    {
        var text = new TextRange(Editor.Document.ContentStart, Editor.Document.ContentEnd).Text;
        var charCount = text.Length;
        var lineCount = Editor.Document.Blocks.Count;
        TxtCharCount.Text = $"{charCount} 字符 · {lineCount} 行";
    }

    private void FlushXamlSave()
    {
        _xamlSaveTimer.Stop();
        if (!_hasPendingXamlSave) return;
        _hasPendingXamlSave = false;
        _vm.BodyXaml = FlowDocumentHelper.SaveToXaml(Editor.Document);
    }

    private void ValidateVariablesInEditor()
    {
        var text = new TextRange(Editor.Document.ContentStart, Editor.Document.ContentEnd).Text;

        var openIdx = 0;
        var unknownVars = new List<string>();
        while ((openIdx = text.IndexOf('{', openIdx)) >= 0)
        {
            var closeIdx = text.IndexOf('}', openIdx);
            if (closeIdx < 0) break;
            var candidate = text.Substring(openIdx, closeIdx - openIdx + 1);
            if (!string.IsNullOrWhiteSpace(candidate.Trim('{', '}')) && !_validVariableNames.Contains(candidate))
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
        UpdateUndoRedoState();
    }

    private void UpdateUndoRedoState()
    {
        if (Editor == null) return;
        UndoBtn.IsEnabled = Editor.CanUndo;
        UndoBtn.Opacity = Editor.CanUndo ? 1.0 : 0.35;
        RedoBtn.IsEnabled = Editor.CanRedo;
        RedoBtn.Opacity = Editor.CanRedo ? 1.0 : 0.35;
    }

    private void UpdateToolbarState()
    {
        if (Editor == null || Editor.Selection == null) return;

        var sel = Editor.Selection;
        double fontSize;
        Brush? foreground;
        if (sel.IsEmpty)
        {
            var parent = Editor.CaretPosition.Parent as FrameworkContentElement;
            BoldBtn.IsChecked = ReadProperty<bool>(parent, TextElement.FontWeightProperty, v => v is FontWeight fw && fw == FontWeights.Bold);
            ItalicBtn.IsChecked = ReadProperty<bool>(parent, TextElement.FontStyleProperty, v => v is FontStyle fs && fs == FontStyles.Italic);
            UnderlineBtn.IsChecked = ReadProperty<bool>(parent, Inline.TextDecorationsProperty, v => v is TextDecorationCollection tdc && tdc == TextDecorations.Underline);
            fontSize = ReadProperty<double>(parent, TextElement.FontSizeProperty, v => v is double d ? d : DefaultFontSize);
            foreground = ReadProperty<Brush?>(parent, TextElement.ForegroundProperty, v => v as Brush);
        }
        else
        {
            BoldBtn.IsChecked = sel.GetPropertyValue(TextElement.FontWeightProperty) as FontWeight? == FontWeights.Bold;
            ItalicBtn.IsChecked = sel.GetPropertyValue(TextElement.FontStyleProperty) as FontStyle? == FontStyles.Italic;
            UnderlineBtn.IsChecked = sel.GetPropertyValue(Inline.TextDecorationsProperty) == TextDecorations.Underline;
            fontSize = sel.GetPropertyValue(TextElement.FontSizeProperty) as double? ?? DefaultFontSize;
            foreground = sel.GetPropertyValue(TextElement.ForegroundProperty) as Brush;
        }

        if (foreground is SolidColorBrush scb)
            ColorPreview.Background = scb;

        TextAlignment alignment;
        if (sel.IsEmpty)
        {
            var parent = Editor.CaretPosition.Parent as FrameworkContentElement;
            alignment = ReadProperty<TextAlignment>(parent, Block.TextAlignmentProperty, v => v is TextAlignment ta ? ta : TextAlignment.Left);
        }
        else
        {
            alignment = sel.GetPropertyValue(Block.TextAlignmentProperty) as TextAlignment? ?? TextAlignment.Left;
        }
        AlignLeftBtn.IsChecked = alignment == TextAlignment.Left;
        AlignCenterBtn.IsChecked = alignment == TextAlignment.Center;
        AlignRightBtn.IsChecked = alignment == TextAlignment.Right;

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

    private void OnAlignLeft(object sender, RoutedEventArgs e)
    {
        ApplyAlignment(TextAlignment.Left);
        AlignLeftBtn.IsChecked = true;
        AlignCenterBtn.IsChecked = false;
        AlignRightBtn.IsChecked = false;
        Editor.Focus();
    }

    private void OnAlignCenter(object sender, RoutedEventArgs e)
    {
        ApplyAlignment(TextAlignment.Center);
        AlignLeftBtn.IsChecked = false;
        AlignCenterBtn.IsChecked = true;
        AlignRightBtn.IsChecked = false;
        Editor.Focus();
    }

    private void OnAlignRight(object sender, RoutedEventArgs e)
    {
        ApplyAlignment(TextAlignment.Right);
        AlignLeftBtn.IsChecked = false;
        AlignCenterBtn.IsChecked = false;
        AlignRightBtn.IsChecked = true;
        Editor.Focus();
    }

    private void ApplyAlignment(TextAlignment alignment)
    {
        if (Editor.Selection != null && !Editor.Selection.IsEmpty)
        {
            Editor.Selection.ApplyPropertyValue(Block.TextAlignmentProperty, alignment);
        }
        else
        {
            var paragraph = Editor.CaretPosition.Paragraph;
            if (paragraph != null)
                paragraph.TextAlignment = alignment;
        }
    }

    private void OnUndo(object sender, RoutedEventArgs e)
    {
        if (Editor.CanUndo) Editor.Undo();
        Editor.Focus();
    }

    private void OnRedo(object sender, RoutedEventArgs e)
    {
        if (Editor.CanRedo) Editor.Redo();
        Editor.Focus();
    }

    private void OnFontSizeKeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Enter)
        {
            if (FontSizeCombo.SelectedItem is int size)
            {
                ApplyFontSize(size);
                FontSizeError.Visibility = Visibility.Collapsed;
            }
            else if (int.TryParse(FontSizeCombo.Text, out var customSize) && customSize >= MinFontSize && customSize <= MaxFontSize)
            {
                ApplyFontSize(customSize);
                FontSizeError.Visibility = Visibility.Collapsed;
            }
            else
            {
                FontSizeError.Text = "无效字号";
                FontSizeError.Visibility = Visibility.Visible;
            }
            e.Handled = true;
            Editor.Focus();
        }
    }

    private void OnFontSizePreviewTextInput(object sender, TextCompositionEventArgs e)
    {
        Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.Loaded, () =>
        {
            if (FontSizeCombo == null) return;
            var text = FontSizeCombo.Text?.Trim();
            if (string.IsNullOrEmpty(text))
            {
                FontSizeError.Visibility = Visibility.Collapsed;
                return;
            }
            if (!int.TryParse(text, out var size) || size < MinFontSize || size > MaxFontSize)
            {
                FontSizeError.Text = $"字号需{MinFontSize}-{MaxFontSize}";
                FontSizeError.Visibility = Visibility.Visible;
            }
            else
            {
                FontSizeError.Visibility = Visibility.Collapsed;
            }
        });
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
        if (string.IsNullOrEmpty(input))
        {
            ShowColorInputError("请输入颜色值");
            return;
        }
        if (!input.StartsWith("#")) input = "#" + input;
        if (input.Length != 7)
        {
            ShowColorInputError("格式应为 #RRGGBB");
            return;
        }
        try
        {
            var color = (Color)ColorConverter.ConvertFromString(input);
            ApplyColor(input);
            if (CustomColorInput != null)
            {
                CustomColorInput.Text = "";
                CustomColorInput.BorderBrush = null;
            }
            HideColorInputError();
        }
        catch
        {
            ShowColorInputError("无效的颜色值");
        }
    }

    private void ShowColorInputError(string message)
    {
        if (CustomColorInput != null)
        {
            CustomColorInput.BorderBrush = (System.Windows.Media.Brush)FindResource("DangerBrush");
            CustomColorInput.BorderThickness = new Thickness(2);
        }
        var errorText = FindName("CustomColorError") as TextBlock;
        if (errorText != null)
        {
            errorText.Text = message;
            errorText.Visibility = Visibility.Visible;
        }
    }

    private void HideColorInputError()
    {
        if (CustomColorInput != null)
        {
            CustomColorInput.BorderBrush = null;
            CustomColorInput.BorderThickness = new Thickness(1);
        }
        var errorText = FindName("CustomColorError") as TextBlock;
        if (errorText != null)
            errorText.Visibility = Visibility.Collapsed;
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

    private void OnLineHeightChanged(object sender, SelectionChangedEventArgs e)
    {
        if (Editor == null) return;
        if (LineHeightCombo.SelectedItem is ComboBoxItem item && item.Tag is string tagStr && double.TryParse(tagStr, out var lineHeight))
        {
            if (Editor.Selection != null && !Editor.Selection.IsEmpty)
            {
                Editor.Selection.ApplyPropertyValue(Block.LineHeightProperty, lineHeight * Editor.FontSize);
            }
            else
            {
                var paragraph = Editor.CaretPosition.Paragraph;
                if (paragraph != null)
                    paragraph.LineHeight = lineHeight * Editor.FontSize;
            }
        }
    }

    private void OnInsertLink(object sender, RoutedEventArgs e)
    {
        var sel = Editor.Selection;
        var hasSelection = sel != null && !sel.IsEmpty;
        var selectedText = hasSelection
            ? new TextRange(sel!.Start, sel.End).Text
            : "";

        var dlg = new LinkDialog(!string.IsNullOrEmpty(selectedText) ? selectedText : null);
        if (dlg.ShowDialog() != true) return;
        var linkText = dlg.LinkText;
        var url = dlg.LinkUrl;

        try
        {
            var uri = new Uri(url);
            if (hasSelection && sel != null)
            {
                var hyperlink = new Hyperlink(sel.Start, sel.End)
                {
                    NavigateUri = uri
                };
            }
            else
            {
                var paragraph = Editor.CaretPosition.Paragraph;
                if (paragraph != null)
                {
                    var hyperlink = new Hyperlink();
                    hyperlink.NavigateUri = uri;
                    hyperlink.Inlines.Add(linkText);
                    paragraph.Inlines.Add(hyperlink);
                }
            }
        }
        catch (UriFormatException)
        {
            App.ShowWarning("请输入有效的链接地址");
        }
        Editor.Focus();
    }

    public void SaveCurrent()
    {
        FlushXamlSave();
        _vm.SaveCurrent();
    }
}
