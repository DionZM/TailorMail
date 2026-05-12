using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using TailorMail.Models;
using TailorMail.ViewModels;

namespace TailorMail.Views;

public partial class VariablesPage : UserControl, IRefreshable
{
    private readonly VariablesViewModel _vm;
    private readonly System.Windows.Threading.DispatcherTimer _saveTimer;
    private bool _hasPendingSave;
    private int? _restoredScrollIndex;
    private int? _editingRowIndex;
    private int? _editingColIndex;

    public VariablesPage()
    {
        InitializeComponent();
        _vm = new VariablesViewModel(App.DataService);
        DataContext = _vm;
        _saveTimer = new System.Windows.Threading.DispatcherTimer { Interval = TimeSpan.FromMilliseconds(700) };
        _saveTimer.Tick += (_, _) =>
        {
            _saveTimer.Stop();
            FlushSave();
        };
        Loaded += (_, _) => BuildGrid();
    }

    public void RefreshData()
    {
        _vm.LoadData();
        BuildGrid();
    }

    private void BuildGrid()
    {
        var scrollViewer = FindVisualChild<System.Windows.Controls.ScrollViewer>(VariablesGrid);
        _restoredScrollIndex = scrollViewer?.VerticalOffset > 0
            ? VariablesGrid.Items.IndexOf(VariablesGrid.CurrentItem)
            : null;

        if (VariablesGrid.CurrentItem != null)
        {
            _editingRowIndex = VariablesGrid.Items.IndexOf(VariablesGrid.CurrentItem);
            _editingColIndex = VariablesGrid.CurrentColumn?.DisplayIndex;
        }

        VariablesGrid.Columns.Clear();

        VarEmptyState.Visibility = _vm.VariableNames.Count == 0 ? Visibility.Visible : Visibility.Collapsed;

        var nameCol = new DataGridTextColumn
        {
            Header = "名称",
            Binding = new Binding("Name") { Mode = BindingMode.OneWay },
            Width = new DataGridLength(120),
            IsReadOnly = true
        };
        VariablesGrid.Columns.Add(nameCol);

        foreach (var varName in _vm.VariableNames)
        {
            var col = new DataGridTextColumn
            {
                Header = varName,
                Width = new DataGridLength(1, DataGridLengthUnitType.Star),
                MinWidth = 100,
            };
            col.Binding = new Binding($"Variables[{varName}]") { UpdateSourceTrigger = UpdateSourceTrigger.LostFocus };
            VariablesGrid.Columns.Add(col);
        }

        VariablesGrid.ItemsSource = _vm.SelectedRecipients;
        VariablesGrid.FrozenColumnCount = 1;

        if (_editingRowIndex.HasValue && _editingColIndex.HasValue
            && _editingRowIndex.Value < VariablesGrid.Items.Count
            && _editingColIndex.Value < VariablesGrid.Columns.Count)
        {
            var rowIdx = _editingRowIndex.Value;
            var colIdx = _editingColIndex.Value;
            Dispatcher.BeginInvoke(() =>
            {
                try
                {
                    VariablesGrid.UpdateLayout();
                    var col = VariablesGrid.Columns[colIdx];
                    VariablesGrid.CurrentCell = new DataGridCellInfo(VariablesGrid.Items[rowIdx], col);
                    VariablesGrid.ScrollIntoView(VariablesGrid.Items[rowIdx]);
                    VariablesGrid.Focus();
                    VariablesGrid.BeginEdit();
                }
                catch { }
            });
        }
        else if (_restoredScrollIndex.HasValue && _restoredScrollIndex.Value >= 0)
        {
            Dispatcher.BeginInvoke(() =>
            {
                if (_restoredScrollIndex.Value < VariablesGrid.Items.Count)
                    VariablesGrid.ScrollIntoView(VariablesGrid.Items[_restoredScrollIndex.Value]);
            });
        }

        _editingRowIndex = null;
        _editingColIndex = null;
        UpdateUsageHint();
    }

    private void UpdateUsageHint()
    {
        var settings = App.DataService.LoadSettings();
        var subject = settings.LastSubject ?? "";
        var body = settings.LastBody ?? "";
        var bodyXaml = settings.LastBodyXaml ?? "";

        var unusedVars = new List<string>();
        foreach (var name in _vm.VariableNames)
        {
            var placeholder = $"{{{name}}}";
            if (!subject.Contains(placeholder) && !body.Contains(placeholder) && !bodyXaml.Contains(placeholder))
                unusedVars.Add(name);
        }

        if (unusedVars.Count > 0 && _vm.VariableNames.Count > 0)
        {
            VarUsageHint.Message = $"以下变量尚未在邮件模板中使用：{string.Join("、", unusedVars)}";
            VarUsageHint.IsOpen = true;
        }
        else
        {
            VarUsageHint.IsOpen = false;
        }
    }

    private static T? FindVisualChild<T>(DependencyObject parent) where T : DependencyObject
    {
        for (int i = 0; i < System.Windows.Media.VisualTreeHelper.GetChildrenCount(parent); i++)
        {
            var child = System.Windows.Media.VisualTreeHelper.GetChild(parent, i);
            if (child is T result) return result;
            var descendant = FindVisualChild<T>(child);
            if (descendant != null) return descendant;
        }
        return null;
    }

    private void OnAddVariableClick(object sender, RoutedEventArgs e)
    {
        var inputDlg = new InputDialog
        {
            Title = "添加变量",
            Prompt = "请输入变量名称："
        };
        
        if (inputDlg.ShowDialog() == true && !string.IsNullOrWhiteSpace(inputDlg.InputText))
        {
            var newName = inputDlg.InputText.Trim();
            if (_vm.VariableNames.Contains(newName))
            {
                App.ShowWarning("变量名已存在");
                return;
            }
            _vm.NewVariableName = newName;
            _vm.AddVariableAndSave();
            BuildGrid();
        }
    }

    private void OnDeleteVariable(object sender, RoutedEventArgs e)
    {
        if (_vm.VariableNames.Count == 0)
        {
            App.ShowNotification("当前没有可删除的变量");
            return;
        }

        if (_vm.VariableNames.Count == 1)
        {
            var name = _vm.VariableNames[0];
            var msg = GetDeleteConfirmMessage(name);
            var dlg = new ConfirmDialog { Title = "确认删除", Message = msg };
            if (dlg.ShowDialog() != true) return;
            _vm.DeleteVariableAndSave(name);
            BuildGrid();
            return;
        }

        var selectDlg = new VariableSelectDialog(_vm.VariableNames.ToList())
        {
            Owner = Window.GetWindow(this)
        };
        if (selectDlg.ShowDialog() == true && selectDlg.SelectedVariables.Count > 0)
        {
            var namesUsed = selectDlg.SelectedVariables.Where(IsVariableUsedInTemplate).ToList();
            var msg = namesUsed.Count > 0
                ? $"以下变量正在邮件模板中使用：{string.Join("、", namesUsed)}\n\n确定删除选中的 {selectDlg.SelectedVariables.Count} 个变量？"
                : $"确定删除选中的 {selectDlg.SelectedVariables.Count} 个变量？";
            var confirmDlg = new ConfirmDialog { Title = "确认删除", Message = msg };
            if (confirmDlg.ShowDialog() != true) return;
            foreach (var name in selectDlg.SelectedVariables)
                _vm.DeleteVariableAndSave(name);
            BuildGrid();
        }
    }

    private static bool IsVariableUsedInTemplate(string variableName)
    {
        var settings = App.DataService.LoadSettings();
        var placeholder = $"{{{variableName}}}";
        return settings.LastSubject?.Contains(placeholder) == true
            || settings.LastBody?.Contains(placeholder) == true
            || settings.LastBodyXaml?.Contains(placeholder) == true;
    }

    private static string GetDeleteConfirmMessage(string name)
    {
        if (IsVariableUsedInTemplate(name))
            return $"变量「{name}」正在邮件模板中使用，删除后模板中的占位符将失效。\n\n确定删除该变量及其所有数据？";
        return $"确定删除变量「{name}」及其所有数据？";
    }

    private void OnRenameVariable(object sender, RoutedEventArgs e)
    {
        if (_vm.VariableNames.Count == 0)
        {
            App.ShowNotification("当前没有可重命名的变量");
            return;
        }

        var selectDlg = new VariableSelectDialog(_vm.VariableNames.ToList()) { AllowSingle = true, ConfirmText = "选择", HeaderText = "请选择要重命名的变量：" };
        selectDlg.Title = "选择要重命名的变量";
        selectDlg.Owner = Window.GetWindow(this);
        if (selectDlg.ShowDialog() != true || selectDlg.SelectedVariables.Count == 0) return;

        var oldName = selectDlg.SelectedVariables[0];
        var inputDlg = new InputDialog { Title = "重命名变量", Prompt = $"将「{oldName}」重命名为：" };
        if (inputDlg.ShowDialog() != true || string.IsNullOrWhiteSpace(inputDlg.InputText)) return;

        var newName = inputDlg.InputText.Trim();
        if (newName == oldName) return;
        if (_vm.VariableNames.Contains(newName))
        {
            App.ShowWarning($"变量名「{newName}」已存在");
            return;
        }

        _vm.RenameVariableAndSave(oldName, newName);
        BuildGrid();
    }

    private void OnCellEditEnding(object sender, DataGridCellEditEndingEventArgs e) => ScheduleSave();

    private void ScheduleSave()
    {
        _hasPendingSave = true;
        _saveTimer.Stop();
        _saveTimer.Start();
    }

    private void FlushSave()
    {
        _saveTimer.Stop();
        if (!_hasPendingSave) return;
        _hasPendingSave = false;
        _vm.SaveAll();
    }

    private void OnGridPreviewKeyDown(object sender, KeyEventArgs e)
    {
        if (sender is not DataGrid grid) return;

        if (e.Key == Key.Enter)
        {
            e.Handled = true;
            var colIndex = grid.CurrentColumn?.DisplayIndex ?? 0;
            var rowIndex = grid.Items.IndexOf(grid.CurrentItem);

            if (colIndex == grid.Columns.Count - 1)
            {
                grid.CommitEdit(DataGridEditingUnit.Row, true);
                Dispatcher.BeginInvoke(() =>
                {
                    grid.CommitEdit();
                    if (rowIndex + 1 < grid.Items.Count)
                    {
                        grid.CurrentCell = new DataGridCellInfo(grid.Items[rowIndex + 1], grid.Columns[1]);
                        grid.BeginEdit();
                    }
                });
            }
            else
            {
                grid.CommitEdit(DataGridEditingUnit.Cell, true);
                Dispatcher.BeginInvoke(() =>
                {
                    var nextCol = grid.Columns.FirstOrDefault(c => c.DisplayIndex == colIndex + 1);
                    if (nextCol != null)
                    {
                        grid.CurrentCell = new DataGridCellInfo(grid.Items[rowIndex], nextCol);
                        grid.BeginEdit();
                    }
                });
            }
        }
        else if (e.Key == Key.Tab && Keyboard.Modifiers == ModifierKeys.None)
        {
            var colIndex = grid.CurrentColumn?.DisplayIndex ?? 0;
            var rowIndex = grid.Items.IndexOf(grid.CurrentItem);

            if (colIndex == grid.Columns.Count - 1)
            {
                e.Handled = true;
                grid.CommitEdit(DataGridEditingUnit.Row, true);
                Dispatcher.BeginInvoke(() =>
                {
                    grid.CommitEdit();
                    if (rowIndex + 1 < grid.Items.Count)
                    {
                        grid.CurrentCell = new DataGridCellInfo(grid.Items[rowIndex + 1], grid.Columns[1]);
                        grid.BeginEdit();
                    }
                });
            }
        }
        else if (e.Key == Key.V && Keyboard.Modifiers == ModifierKeys.Control)
        {
            PasteFromClipboard(grid);
            e.Handled = true;
        }
    }

    private void PasteFromClipboard(DataGrid grid)
    {
        try
        {
            var text = Clipboard.GetText();
            if (string.IsNullOrEmpty(text)) return;
            var lines = text.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);
            if (lines.Length == 0) return;

            var startRowIndex = grid.Items.IndexOf(grid.CurrentItem);
            if (startRowIndex < 0) startRowIndex = 0;

            for (int i = 0; i < lines.Length && startRowIndex + i < grid.Items.Count; i++)
            {
                var columns = lines[i].Split('\t');
                if (columns.Length < 2) continue;

                if (grid.Items.Count > 1 && columns.Length - 1 == _vm.VariableNames.Count)
                {
                    if (grid.Items[startRowIndex + i] is Recipient record)
                    for (int j = 0; j < _vm.VariableNames.Count && j + 1 < columns.Length; j++)
                    {
                        record.Variables[_vm.VariableNames[j]] = columns[j + 1].Trim();
                    }
                }
            }

            grid.Items.Refresh();
            ScheduleSave();
        }
        catch { }
    }

    public void SaveAll() => FlushSave();
}
