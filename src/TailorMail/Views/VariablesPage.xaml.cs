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
        VariablesGrid.Columns.Clear();
        VariablesGrid.ItemsSource = null;

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
            };
            col.Binding = new Binding($"Variables[{varName}]") { UpdateSourceTrigger = UpdateSourceTrigger.LostFocus };
            VariablesGrid.Columns.Add(col);
        }

        VariablesGrid.ItemsSource = _vm.SelectedRecipients;
    }

    private void OnNewVarKeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Enter)
        {
            e.Handled = true;
            DoAddVariable();
        }
    }

    private void OnAddVariableClick(object sender, RoutedEventArgs e)
    {
        DoAddVariable();
    }

    private void DoAddVariable()
    {
        if (string.IsNullOrWhiteSpace(_vm.NewVariableName)) return;
        if (_vm.VariableNames.Contains(_vm.NewVariableName))
        {
            MessageBox.Show("变量名已存在", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }
        _vm.AddVariableAndSave();
        BuildGrid();
        NewVarTextBox.Focus();
    }

    private void OnDeleteVariable(object sender, RoutedEventArgs e)
    {
        if (_vm.VariableNames.Count == 0)
        {
            MessageBox.Show("当前没有可删除的变量", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
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
