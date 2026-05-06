using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using TailorMail.ViewModels;

namespace TailorMail.Views;

/// <summary>
/// 变量管理页面，提供自定义变量的增删、Excel 导入导出、变量值编辑等功能。
/// </summary>
public partial class VariablesPage : UserControl, IRefreshable
{
    private readonly VariablesViewModel _vm;
    private int? _restoredScrollIndex;

    public VariablesPage()
    {
        InitializeComponent();
        _vm = new VariablesViewModel(App.DataService);
        DataContext = _vm;
        Loaded += (_, _) => BuildGrid();
    }

    public void RefreshData()
    {
        _vm.LoadData();
        BuildGrid();
    }

    private void BuildGrid()
    {
        // Preserve scroll position
        var scrollViewer = FindVisualChild<System.Windows.Controls.ScrollViewer>(VariablesGrid);
        _restoredScrollIndex = scrollViewer?.VerticalOffset > 0
            ? VariablesGrid.Items.IndexOf(VariablesGrid.SelectedItem)
            : null;

        VariablesGrid.Columns.Clear();
        VariablesGrid.ItemsSource = null;

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

        // Restore scroll position
        if (_restoredScrollIndex.HasValue && _restoredScrollIndex.Value >= 0)
        {
            Dispatcher.BeginInvoke(() =>
            {
                if (_restoredScrollIndex.Value < VariablesGrid.Items.Count)
                {
                    VariablesGrid.ScrollIntoView(VariablesGrid.Items[_restoredScrollIndex.Value]);
                }
            });
        }

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
            VarUsageHint.Text = $"提示：以下变量尚未在邮件模板中使用：{string.Join("、", unusedVars)}";
            VarUsageHint.Visibility = Visibility.Visible;
        }
        else
        {
            VarUsageHint.Visibility = Visibility.Collapsed;
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
            if (MessageBox.Show($"确定删除变量「{name}」及其所有数据？", "确认删除",
                MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes) return;
            _vm.DeleteVariableAndSave(name);
            BuildGrid();
            return;
        }

        var dlg = new VariableSelectDialog(_vm.VariableNames.ToList())
        {
            Owner = Window.GetWindow(this)
        };
        if (dlg.ShowDialog() == true && dlg.SelectedVariables.Count > 0)
        {
            foreach (var name in dlg.SelectedVariables)
                _vm.DeleteVariableAndSave(name);
            BuildGrid();
        }
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

    private void OnCellEditEnding(object sender, DataGridCellEditEndingEventArgs e) => _vm.SaveAll();

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
    }

    public void SaveAll() => _vm.SaveAll();
}
