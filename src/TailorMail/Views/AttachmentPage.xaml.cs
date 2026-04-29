using System.Windows;
using System.Windows.Controls;
using TailorMail.Models;
using TailorMail.ViewModels;

namespace TailorMail.Views;

/// <summary>
/// 附件管理页面，提供公共附件和单位专属附件的添加/移除、自动匹配等功能。
/// </summary>
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
    }

    public void RefreshData()
    {
        _vm.LoadData();
        LoadLists();
    }

    private void LoadLists()
    {
        CommonList.ItemsSource = _vm.CommonAttachments;
        UnitGrid.ItemsSource = _vm.UnitAttachments;
        TxtFolder.Text = string.IsNullOrEmpty(_vm.MatchDirectory) ? "" : _vm.MatchDirectory;
    }

    private void OnAddCommon(object sender, RoutedEventArgs e)
    {
        _vm.AddCommonAttachmentCommand.Execute(null);
        CommonList.ItemsSource = null;
        CommonList.ItemsSource = _vm.CommonAttachments;

        if (string.IsNullOrEmpty(_vm.MatchDirectory) && _vm.CommonAttachments.Count > 0)
        {
            var dir = System.IO.Path.GetDirectoryName(_vm.CommonAttachments[0]);
            if (dir != null && System.IO.Directory.Exists(dir))
            {
                _vm.MatchDirectory = dir;
                TxtFolder.Text = dir;
            }
        }

        RefreshGrid();
    }

    private void OnClearCommon(object sender, RoutedEventArgs e)
    {
        if (MessageBox.Show("确定清空所有公共附件？", "确认清空",
            MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes) return;
        _vm.CommonAttachments.Clear();
        _vm.SaveConfig();
        CommonList.ItemsSource = null;
    }

    private void OnRemoveCommon(object sender, RoutedEventArgs e)
    {
        if (sender is FrameworkElement fe && fe.Tag is string file)
        {
            if (MessageBox.Show($"确定删除公共附件「{System.IO.Path.GetFileName(file)}」？", "确认删除",
                MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes) return;
            _vm.CommonAttachments.Remove(file);
            _vm.SaveConfig();
            CommonList.ItemsSource = null;
            CommonList.ItemsSource = _vm.CommonAttachments;
        }
    }

    private void OnSelectDirectory(object sender, RoutedEventArgs e)
    {
        var path = Helpers.FolderPicker.PickFolder("选择专有附件所在目录");
        if (path != null)
        {
            _vm.MatchDirectory = path;
            TxtFolder.Text = path;
            _vm.AutoMatchCommand.Execute(null);
            RefreshGrid();
        }
    }

    private void OnAutoMatch(object sender, RoutedEventArgs e)
    {
        _vm.AutoMatchCommand.Execute(null);
        RefreshGrid();
    }

    private void OnAddUnitAttachmentFromGrid(object sender, RoutedEventArgs e)
    {
        if (sender is FrameworkElement fe && fe.Tag is UnitAttachment ua)
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
            _vm.RemoveUnitAttachmentCommand.Execute(file);
            RefreshGrid();
        }
    }

    private void OnClearSpecial(object sender, RoutedEventArgs e)
    {
        if (MessageBox.Show("确定清空所有专有附件？", "确认清空",
            MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes) return;
        foreach (var ua in _vm.UnitAttachments.ToList())
        {
            ua.Files.Clear();
        }
        _vm.SaveConfig();
        RefreshGrid();
    }

    private void RefreshGrid()
    {
        UnitGrid.ItemsSource = null;
        UnitGrid.ItemsSource = _vm.UnitAttachments;
    }
}
