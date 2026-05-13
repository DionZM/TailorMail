using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using TailorMail.Models;
using TailorMail.ViewModels;
using System.Windows.Threading;

namespace TailorMail.Views;

/// <summary>
/// 收件人管理页面，提供收件人分组的切换、收件人的增删改查、Excel 导入导出、全选/取消全选等功能。
/// </summary>
public partial class RecipientsPage : UserControl, IRefreshable, IDynamicStepDesc
{
    private readonly RecipientsViewModel _vm;
    private ListSortDirection _sortDirection;
    private bool _skipNextSort;
    private string? _lastSortProperty;
    private ListSortDirection _lastSortDirection;
    private readonly DispatcherTimer _searchTimer;
    private string _pendingSearchText = "";
    private Point _dragStartPoint;
    private bool _isDragging;

    public event Action? StepDescriptionChanged;

    private const int FirstEditableColIndex = 1;
    private int LastEditableColIndex => RecipientsGrid.Columns.Count - 2;

    public RecipientsPage()
    {
        InitializeComponent();
        _vm = new RecipientsViewModel(App.DataService);
        DataContext = _vm;
        _vm.SelectedGroupChanged += OnSelectedGroupChanged;
        _vm.PropertyChanged += (s, e) =>
        {
            if (e.PropertyName is nameof(_vm.SelectedCount) or nameof(_vm.TotalCount))
                StepDescriptionChanged?.Invoke();
        };
        _searchTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(300) };
        _searchTimer.Tick += OnSearchTimerTick;
        RestorePanelWidth();
        UpdateEmptyState();
        System.ComponentModel.DependencyPropertyDescriptor.FromProperty(ColumnDefinition.WidthProperty, typeof(ColumnDefinition))
            ?.AddValueChanged(LeftPanelColumn, (_, _) => SavePanelWidth());
    }

    public string GetStepDescription() => $"已选 {_vm.SelectedCount} / 共 {_vm.TotalCount}";

    public void RefreshData()
    {
        _vm.LoadGroups();
        UpdateEmptyState();
        RestoreSortState();
    }

    private void OnSelectedGroupChanged()
    {
        UpdateHeaderCheckBox();
        UpdateEmptyState();
        RestoreSortState();
    }

    private void RestoreSortState()
    {
        if (string.IsNullOrEmpty(_lastSortProperty)) return;
        var view = CollectionViewSource.GetDefaultView(_vm.CurrentRecipients);
        view.SortDescriptions.Clear();
        view.SortDescriptions.Add(new SortDescription(_lastSortProperty, _lastSortDirection));

        var col = RecipientsGrid.Columns.FirstOrDefault(c =>
        {
            var header = c.Header as string ?? "";
            return header switch
            {
                "名称" => _lastSortProperty == "Name",
                "简称" => _lastSortProperty == "ShortName",
                "收件人(To)" => _lastSortProperty == "ToEmails",
                "抄送(Cc)" => _lastSortProperty == "CcEmails",
                "密送(Bcc)" => _lastSortProperty == "BccEmails",
                "备注" => _lastSortProperty == "Remark",
                _ => false
            };
        });
        if (col != null) col.SortDirection = _lastSortDirection;
    }

    private void OnSearchTextChanged(object sender, TextChangedEventArgs e)
    {
        _pendingSearchText = ((System.Windows.Controls.TextBox)sender).Text.Trim().ToLower();
        _searchTimer.Stop();
        _searchTimer.Start();
    }

    private void OnSearchTimerTick(object? sender, EventArgs e)
    {
        _searchTimer.Stop();
        var searchText = _pendingSearchText;
        var view = CollectionViewSource.GetDefaultView(_vm.CurrentRecipients);
        view.Filter = item =>
        {
            if (string.IsNullOrEmpty(searchText)) return true;
            if (item is Models.Recipient r)
            {
                return (r.Name?.ToLower().Contains(searchText) == true) ||
                       (r.ToEmails?.ToLower().Contains(searchText) == true) ||
                       (r.ShortName?.ToLower().Contains(searchText) == true) ||
                       (r.CcEmails?.ToLower().Contains(searchText) == true) ||
                       (r.BccEmails?.ToLower().Contains(searchText) == true) ||
                       (r.Remark?.ToLower().Contains(searchText) == true);
            }
            return false;
        };
    }

    private void OnAddRecipientClick(object sender, RoutedEventArgs e)
    {
        AddNewRowAndFocus();
    }

    private void OnDeleteGroup(object sender, RoutedEventArgs e)
    {
        if (_vm.SelectedGroup == null) return;
        var dlg = new ConfirmDialog { Title = "确认删除", Message = $"确定删除分组「{_vm.SelectedGroup.Name}」及其所有收件人？" };
        if (dlg.ShowDialog() != true) return;
        _vm.DeleteGroupCommand.Execute(null);
        UpdateEmptyState();
    }

    private void UpdateEmptyState()
    {
        var isEmpty = _vm.CurrentRecipients.Count == 0;
        EmptyState.Visibility = isEmpty ? Visibility.Visible : Visibility.Collapsed;
    }

    private void OnHeaderCheckBoxClick(object sender, RoutedEventArgs e)
    {
        if (sender is CheckBox cb)
        {
            if (cb.IsChecked == true) _vm.SelectAll();
            else _vm.DeselectAll();
            UpdateHeaderCheckBox();
            UpdateDeleteButton();
        }
    }

    private void OnCheckClick(object sender, RoutedEventArgs e)
    {
        _vm.UpdateCounts();
        UpdateHeaderCheckBox();
        UpdateDeleteButton();
    }

    private void UpdateDeleteButton()
    {
        var count = _vm.CurrentRecipients.Count(r => r.IsSelected);
        DeleteSelectedText.Text = count > 0 ? $"删除已选 ({count})" : "删除已选";
    }

    private void BtnDeleteSelected_Click(object sender, RoutedEventArgs e)
    {
        var count = _vm.CurrentRecipients.Count(r => r.IsSelected);
        if (count == 0) return;

        var dlg = new ConfirmDialog
        {
            Title = "确认删除",
            Message = $"确定删除选中的 {count} 个收件人？此操作不可撤销。"
        };
        if (dlg.ShowDialog() != true) return;

        _vm.DeleteSelectedRecipientsCommand.Execute(null);
        UpdateEmptyState();
    }

    private void UpdateHeaderCheckBox()
    {
        var all = _vm.CurrentRecipients;
        if (all.Count == 0)
        {
            HeaderCheckBox.IsChecked = false;
            HeaderCheckBox.IsThreeState = false;
            return;
        }

        var selectedCount = all.Count(r => r.IsSelected);
        if (selectedCount == 0)
        {
            HeaderCheckBox.IsChecked = false;
            HeaderCheckBox.IsThreeState = false;
        }
        else if (selectedCount == all.Count)
        {
            HeaderCheckBox.IsChecked = true;
            HeaderCheckBox.IsThreeState = false;
        }
        else
        {
            HeaderCheckBox.IsThreeState = true;
            HeaderCheckBox.IsChecked = null;
        }
    }

    private void OnCellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
    {
        if (e.Column.Header as string == "名称" && e.EditAction == DataGridEditAction.Commit)
        {
            var textBox = e.EditingElement as System.Windows.Controls.TextBox;
            var newName = textBox?.Text?.Trim() ?? "";
            var current = e.Row.Item as Recipient;
            if (current != null && !string.IsNullOrEmpty(newName))
            {
                var duplicate = _vm.CurrentRecipients.FirstOrDefault(r =>
                    r.Id != current.Id && r.Name == newName);
                if (duplicate != null)
                {
                    App.ShowWarning($"名称「{newName}」已存在，请使用不同的名称");
                    e.Cancel = true;
                    Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.Loaded, () =>
                    {
                        RecipientsGrid.CurrentCell = new DataGridCellInfo(e.Row.Item, e.Column);
                        RecipientsGrid.BeginEdit();
                    });
                    return;
                }
            }
        }
        _vm.ScheduleSave();
    }

    private void OnRowEditEnding(object sender, DataGridRowEditEndingEventArgs e)
    {
        if (e.EditAction == DataGridEditAction.Commit)
        {
            Dispatcher.BeginInvoke(() => _vm.ScheduleSave());
        }
    }

    private bool IsRowEmpty(Recipient r)
    {
        return string.IsNullOrEmpty(r.Name) &&
               string.IsNullOrEmpty(r.ShortName) &&
               string.IsNullOrEmpty(r.ToEmails) &&
               string.IsNullOrEmpty(r.CcEmails) &&
               string.IsNullOrEmpty(r.BccEmails) &&
               string.IsNullOrEmpty(r.Remark);
    }

    private bool HasEmptyNameRow()
    {
        return _vm.CurrentRecipients.Any(r => string.IsNullOrEmpty(r.Name));
    }

    private void AddNewRowAndFocus()
    {
        if (HasEmptyNameRow()) return;
        var newRecipient = new Recipient { Name = "", IsSelected = false };
        _vm.CurrentRecipients.Add(newRecipient);
        Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.Loaded, () =>
        {
            if (RecipientsGrid.Items.Count > 0)
            {
                var lastIdx = RecipientsGrid.Items.Count - 1;
                RecipientsGrid.UpdateLayout();
                RecipientsGrid.ScrollIntoView(RecipientsGrid.Items[lastIdx]);
                RecipientsGrid.SelectedItem = RecipientsGrid.Items[lastIdx];
                var firstCol = RecipientsGrid.Columns[FirstEditableColIndex];
                RecipientsGrid.CurrentCell = new DataGridCellInfo(RecipientsGrid.Items[lastIdx], firstCol);
                RecipientsGrid.Focus();
                UpdateEmptyState();
                RecipientsGrid.BeginEdit();
            }
        });
    }

    private void OnGridLostFocus(object sender, RoutedEventArgs e)
    {
        Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.Background, () =>
        {
            if (!RecipientsGrid.IsKeyboardFocusWithin)
                CleanupEmptyRows();
        });
    }

    private void OnGridMouseClick(object sender, MouseButtonEventArgs e)
    {
        var hit = VisualTreeHelper.HitTest(RecipientsGrid, e.GetPosition(RecipientsGrid));
        if (hit == null) return;

        var row = FindVisualParent<DataGridRow>(hit.VisualHit);
        if (row == null && _vm.CurrentRecipients.Count == 0)
        {
            AddNewRowAndFocus();
        }
    }

    private void OnGridDoubleClick(object sender, MouseButtonEventArgs e)
    {
        var hit = VisualTreeHelper.HitTest(RecipientsGrid, e.GetPosition(RecipientsGrid));
        if (hit == null) return;

        var row = FindVisualParent<DataGridRow>(hit.VisualHit);
        if (row == null && !HasEmptyNameRow())
        {
            AddNewRowAndFocus();
        }
    }

    private static T? FindVisualParent<T>(DependencyObject? child) where T : DependencyObject
    {
        while (child != null)
        {
            if (child is T result) return result;
            child = System.Windows.Media.VisualTreeHelper.GetParent(child);
        }
        return null;
    }

    private void CleanupEmptyRows()
    {
        var emptyRows = _vm.CurrentRecipients.Where(r => IsRowEmpty(r)).ToList();
        if (emptyRows.Count > 0)
        {
            foreach (var r in emptyRows)
            {
                _vm.CurrentRecipients.Remove(r);
                if (_vm.SelectedGroup != null)
                    _vm.SelectedGroup.Recipients.Remove(r);
            }
            _vm.SaveAll();
            UpdateHeaderCheckBox();
        }
    }

    private DataGridColumn? GetNextEditableColumn(int currentDisplayIndex)
    {
        for (int i = currentDisplayIndex + 1; i <= LastEditableColIndex; i++)
        {
            var col = RecipientsGrid.Columns.FirstOrDefault(c => c.DisplayIndex == i);
            if (col != null && !col.IsReadOnly) return col;
        }
        return null;
    }

    private void MoveToCell(DataGrid grid, int rowIndex, DataGridColumn? column, bool beginEdit)
    {
        if (rowIndex < 0 || rowIndex >= grid.Items.Count || column == null) return;
        grid.SelectedItem = grid.Items[rowIndex];
        grid.CurrentCell = new DataGridCellInfo(grid.Items[rowIndex], column);
        grid.ScrollIntoView(grid.Items[rowIndex]);
        if (beginEdit)
        {
            Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.Loaded, () =>
            {
                grid.BeginEdit();
            });
        }
    }

    private void OnGridPreviewKeyDown(object sender, KeyEventArgs e)
    {
        if (sender is not DataGrid grid) return;

        var colIndex = grid.CurrentColumn?.DisplayIndex ?? 0;
        var rowIndex = grid.Items.IndexOf(grid.CurrentItem);

        if (e.Key == Key.Tab && Keyboard.Modifiers == ModifierKeys.None)
        {
            var nextCol = GetNextEditableColumn(colIndex);
            if (nextCol != null)
            {
                e.Handled = true;
                MoveToCell(grid, rowIndex, nextCol, true);
            }
            else
            {
                e.Handled = true;
                if (rowIndex + 1 < grid.Items.Count)
                {
                    var firstEditCol = RecipientsGrid.Columns[FirstEditableColIndex];
                    MoveToCell(grid, rowIndex + 1, firstEditCol, true);
                }
                else
                {
                    var current = grid.Items[rowIndex] as Recipient;
                    if (current != null && !string.IsNullOrEmpty(current.Name))
                    {
                        AddNewRowAndFocus();
                    }
                }
            }
        }
        else if (e.Key == Key.Tab && Keyboard.Modifiers == ModifierKeys.Shift)
        {
            e.Handled = true;
            int prevColIndex = colIndex - 1;
            while (prevColIndex >= FirstEditableColIndex)
            {
                var prevCol = RecipientsGrid.Columns.FirstOrDefault(c => c.DisplayIndex == prevColIndex);
                if (prevCol != null && !prevCol.IsReadOnly)
                {
                    MoveToCell(grid, rowIndex, prevCol, true);
                    return;
                }
                prevColIndex--;
            }
            if (rowIndex > 0)
            {
                var lastEditCol = RecipientsGrid.Columns.FirstOrDefault(c => c.DisplayIndex == LastEditableColIndex);
                MoveToCell(grid, rowIndex - 1, lastEditCol, true);
            }
        }
        else if (e.Key == Key.Enter)
        {
            e.Handled = true;
            grid.CommitEdit(DataGridEditingUnit.Cell, true);
            
            Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.Loaded, () =>
            {
                var current = grid.Items[rowIndex] as Recipient;
                var isNameColumn = colIndex == FirstEditableColIndex;
                
                if (isNameColumn && current != null && !string.IsNullOrEmpty(current.Name))
                {
                    var duplicate = _vm.CurrentRecipients.FirstOrDefault(r =>
                        r.Id != current.Id && r.Name == current.Name);
                    if (duplicate != null)
                    {
                        App.ShowWarning($"名称「{current.Name}」已存在，请使用不同的名称");
                        return;
                    }
                    
                    AddNewRowAndFocus();
                }
                else if (rowIndex + 1 < grid.Items.Count)
                {
                    var sameCol = grid.CurrentColumn ?? RecipientsGrid.Columns[FirstEditableColIndex];
                    MoveToCell(grid, rowIndex + 1, sameCol, true);
                }
            });
        }
        else if (e.Key == Key.Down && Keyboard.Modifiers == ModifierKeys.None)
        {
            e.Handled = true;
            if (rowIndex + 1 < grid.Items.Count)
            {
                var sameCol = grid.CurrentColumn ?? RecipientsGrid.Columns[FirstEditableColIndex];
                MoveToCell(grid, rowIndex + 1, sameCol, true);
            }
            else
            {
                var current = grid.Items[rowIndex] as Recipient;
                if (current != null && !string.IsNullOrEmpty(current.Name))
                {
                    AddNewRowAndFocus();
                }
            }
        }
        else if (e.Key == Key.Up && Keyboard.Modifiers == ModifierKeys.None)
        {
            if (rowIndex > 0)
            {
                e.Handled = true;
                var sameCol = grid.CurrentColumn ?? RecipientsGrid.Columns[FirstEditableColIndex];
                MoveToCell(grid, rowIndex - 1, sameCol, true);
            }
        }
        else if (e.Key == Key.Escape)
        {
            grid.CancelEdit();
        }
        else if (e.Key == Key.V && Keyboard.Modifiers == ModifierKeys.Control)
        {
            PasteFromClipboard(grid, rowIndex);
            e.Handled = true;
        }
    }

    private void OnGridSorting(object sender, DataGridSortingEventArgs e)
    {
        if (_skipNextSort)
        {
            _skipNextSort = false;
            return;
        }

        var column = e.Column;
        var header = column.Header as string ?? "";

        if (column.SortDirection == null)
        {
            column.SortDirection = ListSortDirection.Ascending;
            _sortDirection = ListSortDirection.Ascending;
        }
        else if (column.SortDirection == ListSortDirection.Ascending)
        {
            column.SortDirection = ListSortDirection.Descending;
            _sortDirection = ListSortDirection.Descending;
        }
        else
        {
            column.SortDirection = null;
            _skipNextSort = true;
            var view = CollectionViewSource.GetDefaultView(_vm.CurrentRecipients);
            view.SortDescriptions.Clear();
            foreach (var col in RecipientsGrid.Columns)
                col.SortDirection = null;
            e.Handled = true;
            return;
        }

        foreach (var col in RecipientsGrid.Columns)
        {
            if (col != column) col.SortDirection = null;
        }

        var propName = header switch
        {
            "名称" => "Name",
            "简称" => "ShortName",
            "收件人(To)" => "ToEmails",
            "抄送(Cc)" => "CcEmails",
            "密送(Bcc)" => "BccEmails",
            "备注" => "Remark",
            _ => ""
        };

        if (!string.IsNullOrEmpty(propName))
        {
            _lastSortProperty = propName;
            _lastSortDirection = _sortDirection;
            var view = CollectionViewSource.GetDefaultView(_vm.CurrentRecipients);
            view.SortDescriptions.Clear();
            view.SortDescriptions.Add(new SortDescription(propName, _sortDirection));
        }

        e.Handled = true;
    }

    private void OnMoveUp(object sender, RoutedEventArgs e)
    {
        if (RecipientsGrid.SelectedItem is Recipient r)
            _vm.MoveUp(r);
    }

    private void OnMoveDown(object sender, RoutedEventArgs e)
    {
        if (RecipientsGrid.SelectedItem is Recipient r)
            _vm.MoveDown(r);
    }

    private void OnDeleteRow(object sender, RoutedEventArgs e)
    {
        if (RecipientsGrid.SelectedItem is Recipient r)
        {
            _vm.DeleteRecipientCommand.Execute(r);
            UpdateHeaderCheckBox();
            UpdateEmptyState();
        }
    }

    private void OnDeleteRowClick(object sender, RoutedEventArgs e)
    {
        if (sender is FrameworkElement fe && fe.Tag is Recipient r)
        {
            _vm.DeleteRecipientCommand.Execute(r);
            UpdateHeaderCheckBox();
            UpdateEmptyState();
        }
    }

    private void RestorePanelWidth()
    {
        var settings = App.DataService.LoadSettings();
        if (settings.RecipientsPanelWidth is > 160 and < 400)
            LeftPanelColumn.Width = new GridLength(settings.RecipientsPanelWidth.Value);
    }

    private void SavePanelWidth()
    {
        var settings = App.DataService.LoadSettings();
        settings.RecipientsPanelWidth = LeftPanelColumn.ActualWidth;
        App.DataService.SaveSettings(settings);
    }

    public void SaveAll() => _vm.SaveAll();

    public void FocusSearch()
    {
        SearchBox?.Focus();
    }

    private void PasteFromClipboard(DataGrid grid, int startRowIndex)
    {
        try
        {
            var text = Clipboard.GetText();
            if (string.IsNullOrEmpty(text)) return;

            var lines = text.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);
            var dataLines = lines.Where(l => !string.IsNullOrWhiteSpace(l)).ToList();
            if (dataLines.Count == 0) return;

            for (int i = 0; i < dataLines.Count; i++)
            {
                var columns = dataLines[i].Split('\t');
                var rowIdx = startRowIndex + i;

                if (rowIdx >= grid.Items.Count)
                {
                    if (HasEmptyNameRow()) break;
                    var newRecipient = new Recipient { Name = columns.Length > 0 ? columns[0].Trim() : "", IsSelected = false };
                    _vm.CurrentRecipients.Add(newRecipient);
                }

                if (rowIdx < grid.Items.Count && grid.Items[rowIdx] is Recipient recipient)
                {
                    if (columns.Length > 0) recipient.Name = columns[0].Trim();
                    if (columns.Length > 1) recipient.ShortName = columns[1].Trim();
                    if (columns.Length > 2) recipient.ToEmails = columns[2].Trim();
                    if (columns.Length > 3) recipient.CcEmails = columns[3].Trim();
                    if (columns.Length > 4) recipient.BccEmails = columns[4].Trim();
                    if (columns.Length > 5) recipient.Remark = columns[5].Trim();
                }
            }

            grid.Items.Refresh();
            _vm.SaveAll();
            UpdateEmptyState();
            UpdateHeaderCheckBox();
        }
        catch { }
    }

    // UI-01: Drag-drop row reordering
    private void OnGridPreviewMouseMove(object sender, MouseEventArgs e)
    {
        if (e.LeftButton != MouseButtonState.Pressed || _isDragging) return;
        var pos = e.GetPosition(RecipientsGrid);
        if (Math.Abs(pos.X - _dragStartPoint.X) > SystemParameters.MinimumHorizontalDragDistance ||
            Math.Abs(pos.Y - _dragStartPoint.Y) > SystemParameters.MinimumVerticalDragDistance)
        {
            var row = FindVisualParent<DataGridRow>(e.OriginalSource as DependencyObject);
            if (row == null) return;
            _isDragging = true;
            var data = new DataObject("RecipientRow", row.Item);
            DragDrop.DoDragDrop(RecipientsGrid, data, DragDropEffects.Move);
            _isDragging = false;
        }
    }

    private void OnGridPreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        _dragStartPoint = e.GetPosition(RecipientsGrid);
    }

    private void OnGridDragOver(object sender, DragEventArgs e)
    {
        e.Effects = e.Data.GetDataPresent("RecipientRow") ? DragDropEffects.Move : DragDropEffects.None;
        e.Handled = true;
    }

    private void OnGridDrop(object sender, DragEventArgs e)
    {
        if (!e.Data.GetDataPresent("RecipientRow")) return;
        var draggedItem = e.Data.GetData("RecipientRow") as Recipient;
        if (draggedItem == null) return;

        var target = FindVisualParent<DataGridRow>(e.OriginalSource as DependencyObject);
        Recipient? targetItem = target?.Item as Recipient;
        if (targetItem == null || targetItem == draggedItem) return;

        var list = _vm.CurrentRecipients;
        int oldIndex = list.IndexOf(draggedItem);
        int newIndex = list.IndexOf(targetItem);
        if (oldIndex < 0 || newIndex < 0) return;

        list.Move(oldIndex, newIndex);
        _vm.SyncRecipientsToGroup();
        _vm.ScheduleSave();
    }

    // UI-02: Group rename via context menu
    private void OnRenameGroup(object sender, RoutedEventArgs e)
    {
        if (_vm.SelectedGroup == null) return;
        var inputDlg = new InputDialog
        {
            Title = "重命名分组",
            Prompt = "请输入新的分组名称：",
            InputText = _vm.SelectedGroup.Name
        };
        if (inputDlg.ShowDialog() == true && !string.IsNullOrWhiteSpace(inputDlg.InputText))
        {
            var newName = inputDlg.InputText.Trim();
            if (newName == _vm.SelectedGroup.Name) return;
            if (_vm.Groups.Any(g => g.Name == newName))
            {
                App.ShowWarning("分组名已存在");
                return;
            }
            _vm.SelectedGroup.Name = newName;
            _vm.SaveAll();
        }
    }
}
