using System.Windows;
using System.Windows.Controls;
using TailorMail.Models;
using TailorMail.ViewModels;

namespace TailorMail.Views;

public partial class AttachmentPage : UserControl, IRefreshable
{
    private readonly AttachmentViewModel _vm;

    public AttachmentPage()
    {
        InitializeComponent();
        _vm = new AttachmentViewModel(App.DataService,
            App.GetRequiredService<Services.AttachmentMatchService>());
        DataContext = _vm;
        LoadLists();
        RestorePanelWidth();
        System.ComponentModel.DependencyPropertyDescriptor.FromProperty(ColumnDefinition.WidthProperty, typeof(ColumnDefinition))
            ?.AddValueChanged(CommonPanelColumn, (_, _) => SavePanelWidth());
    }

    public void RefreshData()
    {
        _vm.LoadData();
        LoadLists();
    }

    private void LoadLists()
    {
        CommonList.ItemsSource = _vm.CommonAttachments;
        RecipientGrid.ItemsSource = _vm.RecipientAttachments;
        TxtFolder.Text = string.IsNullOrEmpty(_vm.MatchDirectory) ? "" : _vm.MatchDirectory;
        TxtFolder.ToolTip = string.IsNullOrEmpty(_vm.MatchDirectory) ? null : _vm.MatchDirectory;
        UpdateEmptyStates();
    }

    private void UpdateEmptyStates()
    {
        var hasCommon = _vm.CommonAttachments.Count > 0;
        CommonEmptyState.Visibility = hasCommon ? Visibility.Collapsed : Visibility.Visible;
        SpecialEmptyState.Visibility = _vm.RecipientAttachments.Count == 0 ? Visibility.Visible : Visibility.Collapsed;
    }

    private void OnAddCommon(object sender, RoutedEventArgs e)
    {
        _vm.AddCommonAttachmentCommand.Execute(null);
        CommonList.ItemsSource = null;
        CommonList.ItemsSource = _vm.CommonAttachments;
        UpdateEmptyStates();

        if (string.IsNullOrEmpty(_vm.MatchDirectory) && _vm.CommonAttachments.Count > 0)
        {
            var dir = System.IO.Path.GetDirectoryName(_vm.CommonAttachments[0]);
            if (dir != null && System.IO.Directory.Exists(dir))
            {
                _vm.MatchDirectory = dir;
                TxtFolder.Text = dir;
                TxtFolder.ToolTip = dir;
            }
        }

        RefreshGrid();
    }

    private void OnClearCommon(object sender, RoutedEventArgs e)
    {
        var dlg = new ConfirmDialog { Title = "确认清空", Message = "确定清空所有公共附件？" };
        if (dlg.ShowDialog() != true) return;
        _vm.CommonAttachments.Clear();
        _vm.SaveConfig();
        CommonList.ItemsSource = null;
        UpdateEmptyStates();
    }

    private void OnRemoveCommon(object sender, RoutedEventArgs e)
    {
        if (sender is FrameworkElement fe && fe.Tag is string file)
        {
            var fileName = System.IO.Path.GetFileName(file);
            var index = _vm.CommonAttachments.IndexOf(file);
            _vm.CommonAttachments.Remove(file);
            _vm.SaveConfig();
            CommonList.ItemsSource = null;
            CommonList.ItemsSource = _vm.CommonAttachments;
            UpdateEmptyStates();

            // UI-37: Show undo notification
            App.ShowNotification($"已删除「{fileName}」，可通过重新添加恢复");
        }
    }

    private void OnSelectDirectory(object sender, RoutedEventArgs e)
    {
        var path = Helpers.FolderPicker.PickFolder("选择专有附件所在目录");
        if (path != null)
        {
            _vm.MatchDirectory = path;
            TxtFolder.Text = path;
            TxtFolder.ToolTip = path;
            _vm.AutoMatchCommand.Execute(null);
            RefreshGrid();
        }
    }

    private void OnAutoMatch(object sender, RoutedEventArgs e)
    {
        BtnAutoMatch.IsEnabled = false;
        var originalContent = BtnAutoMatch.Content;
        BtnAutoMatch.Content = "匹配中...";

        var beforeCount = _vm.RecipientAttachments.Sum(ra => ra.Files.Count);
        _vm.AutoMatchCommand.Execute(null);
        var afterCount = _vm.RecipientAttachments.Sum(ra => ra.Files.Count);
        var matched = afterCount - beforeCount;

        RefreshGrid();
        UpdateEmptyStates();

        if (matched > 0)
        {
            TxtMatchResult.Text = $"自动匹配完成，新增 {matched} 个附件";
            TxtMatchResult.Foreground = (System.Windows.Media.Brush)FindResource("SuccessBrush");
            TxtMatchResult.Visibility = Visibility.Visible;
            var timer = new System.Windows.Threading.DispatcherTimer { Interval = TimeSpan.FromSeconds(3) };
            timer.Tick += (_, _) =>
            {
                timer.Stop();
                TxtMatchResult.Visibility = Visibility.Collapsed;
            };
            timer.Start();
        }
        else if (beforeCount == afterCount && _vm.RecipientAttachments.Count > 0)
        {
            TxtMatchResult.Text = "未匹配到新的附件，请检查文件名是否与收件人名称匹配";
            TxtMatchResult.Foreground = (System.Windows.Media.Brush)FindResource("WarningBrush");
            TxtMatchResult.Visibility = Visibility.Visible;
        }

        BtnAutoMatch.IsEnabled = true;
        BtnAutoMatch.Content = originalContent;
    }

    private void OnAddRecipientAttachmentFromGrid(object sender, RoutedEventArgs e)
    {
        if (sender is FrameworkElement fe && fe.Tag is RecipientAttachment ua)
        {
            var dialog = new Microsoft.Win32.OpenFileDialog
            {
                Multiselect = true,
                Title = $"选择 {ua.RecipientName} 的专有附件"
            };
            if (dialog.ShowDialog() != true) return;
            foreach (var file in dialog.FileNames)
            {
                if (!ua.Files.Contains(file))
                    ua.Files.Add(file);
            }
            _vm.SaveConfig();
            RefreshGrid();
        }
    }

    private void OnRemoveSpecialFromGrid(object sender, RoutedEventArgs e)
    {
        if (sender is FrameworkElement fe && fe.Tag is string file)
        {
            _vm.RemoveRecipientAttachmentCommand.Execute(file);
            RefreshGrid();
        }
    }

    private void OnClearSpecial(object sender, RoutedEventArgs e)
    {
        var dlg = new ConfirmDialog { Title = "确认清空", Message = "确定清空所有专有附件？" };
        if (dlg.ShowDialog() != true) return;
        foreach (var ua in _vm.RecipientAttachments.ToList())
        {
            ua.Files.Clear();
        }
        _vm.SaveConfig();
        RefreshGrid();
    }

    private void RefreshGrid()
    {
        RecipientGrid.ItemsSource = null;
        RecipientGrid.ItemsSource = _vm.RecipientAttachments;
    }

    private void RestorePanelWidth()
    {
        var settings = App.DataService.LoadSettings();
        if (settings.AttachmentPanelWidth is > 160 and < 400)
            CommonPanelColumn.Width = new System.Windows.GridLength(settings.AttachmentPanelWidth.Value);
    }

    private void SavePanelWidth()
    {
        var settings = App.DataService.LoadSettings();
        settings.AttachmentPanelWidth = CommonPanelColumn.ActualWidth;
        App.DataService.SaveSettings(settings);
    }

    #region Drag & Drop

    private void OnDragEnter(object sender, DragEventArgs e)
    {
        if (e.Data.GetDataPresent(DataFormats.FileDrop))
        {
            e.Effects = DragDropEffects.Copy;
            if (sender is Border border)
            {
                border.BorderBrush = (System.Windows.Media.Brush)FindResource("AccentBrush");
                border.BorderThickness = new Thickness(2);
                border.Background = (System.Windows.Media.Brush)FindResource("AccentLightBrush");
            }
        }
        e.Handled = true;
    }

    private void OnDragLeave(object sender, DragEventArgs e)
    {
        if (sender is Border border)
        {
            border.BorderBrush = (System.Windows.Media.Brush)FindResource("BorderSubtleBrush");
            border.BorderThickness = new Thickness(1);
            border.Background = (System.Windows.Media.Brush)FindResource("SurfaceElevatedBrush");
        }
        e.Handled = true;
    }

    private void OnDragOver(object sender, DragEventArgs e)
    {
        if (e.Data.GetDataPresent(DataFormats.FileDrop))
        {
            e.Effects = DragDropEffects.Copy;
            e.Handled = true;
        }
    }

    private void OnDrop(object sender, DragEventArgs e)
    {
        OnCommonDrop(sender, e);
    }

    private void OnCommonDrop(object sender, DragEventArgs e)
    {
        if (!e.Data.GetDataPresent(DataFormats.FileDrop)) return;

        var files = (string[]?)e.Data.GetData(DataFormats.FileDrop);
        if (files == null) return;

        // UI-34: File count limit
        if (files.Length > 500)
        {
            App.ShowWarning("单次拖放文件数量不能超过 500 个");
            return;
        }

        foreach (var file in files)
        {
            if (System.IO.File.Exists(file) && !_vm.CommonAttachments.Contains(file))
                _vm.CommonAttachments.Add(file);
        }
        _vm.SaveConfig();
        CommonList.ItemsSource = null;
        CommonList.ItemsSource = _vm.CommonAttachments;
        UpdateEmptyStates();
    }

    #endregion

    #region Recipient Grid Drag & Drop

    private void OnRecipientDragOver(object sender, DragEventArgs e)
    {
        if (e.Data.GetDataPresent(DataFormats.FileDrop))
        {
            e.Effects = DragDropEffects.Copy;
            e.Handled = true;
        }
    }

    private void OnRecipientDrop(object sender, DragEventArgs e)
    {
        if (!e.Data.GetDataPresent(DataFormats.FileDrop)) return;

        var files = (string[]?)e.Data.GetData(DataFormats.FileDrop);
        if (files == null || files.Length == 0) return;

        // Find the row under the drop position
        var hit = System.Windows.Media.VisualTreeHelper.HitTest(RecipientGrid, e.GetPosition(RecipientGrid));
        if (hit == null) return;
        var row = FindVisualParent<DataGridRow>(hit.VisualHit);
        if (row?.DataContext is RecipientAttachment ua)
        {
            foreach (var file in files)
            {
                if (System.IO.File.Exists(file) && !ua.Files.Contains(file))
                    ua.Files.Add(file);
            }
            _vm.SaveConfig();
            RefreshGrid();
        }
    }

    private static T? FindVisualParent<T>(System.Windows.DependencyObject child) where T : System.Windows.DependencyObject
    {
        var parent = System.Windows.Media.VisualTreeHelper.GetParent(child);
        while (parent != null)
        {
            if (parent is T result) return result;
            parent = System.Windows.Media.VisualTreeHelper.GetParent(parent);
        }
        return null;
    }

    private void OnCommonListClick(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        var hit = System.Windows.Media.VisualTreeHelper.HitTest(CommonList, e.GetPosition(CommonList));
        if (hit == null) { CommonList.SelectedIndex = -1; return; }
        var item = FindVisualParent<ListBoxItem>(hit.VisualHit);
        if (item == null) CommonList.SelectedIndex = -1;
    }

    #endregion
}
