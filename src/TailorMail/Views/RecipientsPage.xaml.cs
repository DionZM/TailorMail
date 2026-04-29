﻿using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using TailorMail.Models;
using TailorMail.ViewModels;

namespace TailorMail.Views;

/// <summary>
/// 收件人管理页面，提供收件人分组的切换、收件人的增删改查、Excel 导入导出、全选/取消全选等功能。
/// </summary>
public partial class RecipientsPage : UserControl, IRefreshable
{
    private readonly RecipientsViewModel _vm;
    private ListSortDirection _sortDirection;
    private bool _skipNextSort;

    private const int FirstEditableColIndex = 1;
    private int LastEditableColIndex => RecipientsGrid.Columns.Count - 2;

    public RecipientsPage()
    {
        InitializeComponent();
        _vm = new RecipientsViewModel(App.DataService);
        DataContext = _vm;
        _vm.SelectedGroupChanged += OnSelectedGroupChanged;
    }

    public void RefreshData() => _vm.LoadGroups();

    private void OnSelectedGroupChanged()
    {
        UpdateHeaderCheckBox();
        UpdateEmptyState();
    }

    private void OnSearchTextChanged(object sender, TextChangedEventArgs e)
    {
        var searchText = ((System.Windows.Controls.TextBox)sender).Text.Trim().ToLower();
        var view = CollectionViewSource.GetDefaultView(_vm.CurrentRecipients);
        view.Filter = item =>
        {
            if (string.IsNullOrEmpty(searchText)) return true;
            if (item is Models.Recipient r)
            {
                return (r.Name?.ToLower().Contains(searchText) == true) ||
                       (r.ToEmails?.ToLower().Contains(searchText) == true) ||
                       (r.ShortName?.ToLower().Contains(searchText) == true);
            }
            return false;
        };
    }

    private void OnDeleteGroup(object sender, RoutedEventArgs e)
    {
        if (_vm.SelectedGroup == null) return;
        if (MessageBox.Show($"确定删除分组「{_vm.SelectedGroup.Name}」及其所有收件人？", "确认删除",
            MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes) return;
        _vm.DeleteGroupCommand.Execute(null);
        UpdateEmptyState();
    }

    private void UpdateEmptyState()
    {
        EmptyState.Visibility = _vm.CurrentRecipients.Count == 0 ? Visibility.Visible : Visibility.Collapsed;
    }

    private void OnHeaderCheckBoxClick(object sender, RoutedEventArgs e)
    {
        if (sender is CheckBox cb)
        {
            if (cb.IsChecked == true) _vm.SelectAll();
            else _vm.DeselectAll();
            UpdateHeaderCheckBox();
        }
    }

    private void OnCheckClick(object sender, RoutedEventArgs e)
    {
        _vm.UpdateCounts();
        UpdateHeaderCheckBox();
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
                    System.Windows.MessageBox.Show($"名称「{newName}」已存在，请使用不同的名称", "名称重复",
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                    if (textBox != null) textBox.Text = "";
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
        _vm.SaveAll();
    }

    private void OnRowEditEnding(object sender, DataGridRowEditEndingEventArgs e)
    {
        if (e.EditAction == DataGridEditAction.Commit)
        {
            Dispatcher.BeginInvoke(() => _vm.SaveAll());
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
                var firstCol = RecipientsGrid.Columns[FirstEditableColIndex];
                RecipientsGrid.CurrentCell = new DataGridCellInfo(RecipientsGrid.Items[lastIdx], firstCol);
                RecipientsGrid.Focus();
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

    private static T? FindVisualParent<T>(DependencyObject child) where T : DependencyObject
    {
        var parent = System.Windows.Media.VisualTreeHelper.GetParent(child);
        while (parent != null)
        {
            if (parent is T result) return result;
            parent = System.Windows.Media.VisualTreeHelper.GetParent(parent);
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
            var isNameColumn = colIndex == FirstEditableColIndex;
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
                    if (isNameColumn)
                    {
                        _vm.SaveAll();
                        AddNewRowAndFocus();
                    }
                    else
                    {
                        AddNewRowAndFocus();
                    }
                }
            }
        }
        else if (e.Key == Key.Down && Keyboard.Modifiers == ModifierKeys.None)
        {
            if (rowIndex == grid.Items.Count - 1)
            {
                var current = grid.Items[rowIndex] as Recipient;
                if (current != null && !string.IsNullOrEmpty(current.Name))
                {
                    e.Handled = true;
                    AddNewRowAndFocus();
                }
            }
        }
        else if (e.Key == Key.Escape)
        {
            grid.CancelEdit();
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
        }
    }

    private void OnDeleteRowClick(object sender, RoutedEventArgs e)
    {
        if (sender is FrameworkElement fe && fe.Tag is Recipient r)
        {
            _vm.DeleteRecipientCommand.Execute(r);
            UpdateHeaderCheckBox();
        }
    }

    public void SaveAll() => _vm.SaveAll();
}
